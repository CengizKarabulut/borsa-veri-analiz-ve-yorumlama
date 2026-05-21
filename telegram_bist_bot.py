#!/usr/bin/env python3
"""
BIST Sektör Analiz Telegram Botu
==================================
Kullanıcı bot'a bir BIST sektör Excel'i (.xlsx) gönderir.
Bot dosyayı indirir, analiz eder ve raporu Telegram'a geri gönderir:
  1) Tüm hisselerin puanlama tablosu
  2) İlk 10 hissenin detaylı analizi
  3) Tam markdown rapor (dosya olarak)

ÇALIŞMA MODLARI:
  python telegram_bist_bot.py           # sürekli mod (long-polling) — VPS/local için
  python telegram_bist_bot.py --once    # tek seferlik — GitHub Actions cron için

GEREKLİ ORTAM DEĞİŞKENLERİ:
  TELEGRAM_BOT_TOKEN     (zorunlu)
  AUTHORIZED_CHAT_IDS    (opsiyonel — virgülle ayrılmış chat ID'leri.
                          Verilmezse herkes kullanabilir)
"""

import os
import sys
import time
import json
import tempfile
import traceback
from pathlib import Path
from html import escape as html_escape

import requests
import numpy as np
import pandas as pd


# ==========================================================
# Konfigürasyon
# ==========================================================
BOT_TOKEN = os.environ.get('TELEGRAM_BOT_TOKEN', '').strip()
if not BOT_TOKEN:
    print("HATA: TELEGRAM_BOT_TOKEN ortam değişkeni boş.", file=sys.stderr)
    sys.exit(1)

API = f"https://api.telegram.org/bot{BOT_TOKEN}"
FILE_API = f"https://api.telegram.org/file/bot{BOT_TOKEN}"

AUTH_IDS = set()
_auth_env = os.environ.get('AUTHORIZED_CHAT_IDS', '').strip()
if _auth_env:
    AUTH_IDS = {int(x.strip()) for x in _auth_env.split(',') if x.strip()}

AGIRLIKLAR = {'degerleme': 0.30, 'karlilik': 0.30, 'buyume': 0.25, 'bilanco': 0.15}
TOP_N_DETAY = 10
MAX_MSG = 3800  # 4096 limit, güvenli pay


# ==========================================================
# Telegram API yardımcıları
# ==========================================================
def tg_call(method, **params):
    url = f"{API}/{method}"
    try:
        r = requests.post(url, json=params, timeout=60)
        return r.json()
    except Exception as e:
        print(f"tg_call error {method}: {e}", file=sys.stderr)
        return {'ok': False, 'error': str(e)}


def send_message(chat_id, text, parse_mode='HTML', disable_preview=True):
    """Uzun mesajları otomatik böler ve sırayla gönderir."""
    chunks = _split_message(text, MAX_MSG)
    for i, chunk in enumerate(chunks):
        res = tg_call('sendMessage',
                      chat_id=chat_id,
                      text=chunk,
                      parse_mode=parse_mode,
                      disable_web_page_preview=disable_preview)
        if not res.get('ok'):
            print(f"sendMessage başarısız (chunk {i}): {res}", file=sys.stderr)
        time.sleep(0.4)  # Telegram rate limit dostu
    return chunks


def send_document(chat_id, file_path, caption=None):
    url = f"{API}/sendDocument"
    with open(file_path, 'rb') as f:
        files = {'document': (Path(file_path).name, f)}
        data = {'chat_id': chat_id}
        if caption:
            data['caption'] = caption
            data['parse_mode'] = 'HTML'
        try:
            r = requests.post(url, data=data, files=files, timeout=120)
            return r.json()
        except Exception as e:
            return {'ok': False, 'error': str(e)}


def send_chat_action(chat_id, action='upload_document'):
    tg_call('sendChatAction', chat_id=chat_id, action=action)


def _split_message(text, limit):
    """HTML mesajı satır sınırında, gerekirse <pre> bloğu sınırında böler."""
    if len(text) <= limit:
        return [text]
    chunks = []
    cur = []
    cur_len = 0
    in_pre = False
    for line in text.split('\n'):
        if '<pre>' in line:
            in_pre = True
        line_len = len(line) + 1
        if cur_len + line_len > limit and cur:
            if in_pre and '</pre>' not in '\n'.join(cur):
                cur.append('</pre>')
            chunks.append('\n'.join(cur))
            cur = []
            cur_len = 0
            if in_pre:
                cur.append('<pre>')
                cur_len += 6
        cur.append(line)
        cur_len += line_len
        if '</pre>' in line:
            in_pre = False
    if cur:
        chunks.append('\n'.join(cur))
    return chunks


def download_telegram_file(file_id, dest_dir):
    info = tg_call('getFile', file_id=file_id)
    if not info.get('ok'):
        return None
    file_path = info['result']['file_path']
    url = f"{FILE_API}/{file_path}"
    name = Path(file_path).name
    dest = Path(dest_dir) / name
    r = requests.get(url, timeout=120)
    if r.status_code != 200:
        return None
    dest.write_bytes(r.content)
    return dest


# ==========================================================
# Analiz mantığı — 6 BOYUTLU GENİŞLETİLMİŞ SKORLAMA (v2)
# ==========================================================
# Ağırlıklar (toplam %100):
#   Değerleme         %25  → F/K, PD/DD, FD/FAVÖK, PEG
#   Karlılık          %25  → ROE, ROIC, ROA, Net/FAVÖK/Brüt Marj
#   Büyüme (REEL)     %20  → Reel satış, Reel kâr, Çeyreklik momentum
#   Bilanço Sağlığı   %15  → Borç/ÖzK, Cari, Likidite, Nakit, Borç/FAVÖK, Faiz Karş.
#   Op. Verimlilik    %10  → Stok devir, DSO, Nakit Çevirme Süresi
#   Piyasa Sinyalleri %5   → Yabancı Oran, Bilanço Sonrası Getiri, Volatilite (ters)
#
# REEL büyüme = ((1 + nominal/100) / (1 + TÜFE/100) - 1) * 100
# Kaynak: Yaşar Erdinç (19 mali oran) + CANSLIM (William O'Neil)
# ==========================================================

AGIRLIKLAR_V2 = {
    'degerleme': 0.25,
    'karlilik':  0.25,
    'buyume':    0.20,
    'bilanco':   0.15,
    'verimlilik': 0.10,
    'piyasa':    0.05,
}


def normalize(s, higher_better=True, clip_pct=(5, 95)):
    s = s.copy()
    valid = s.dropna()
    if len(valid) == 0:
        return s
    lo, hi = np.nanpercentile(valid, clip_pct)
    s = s.clip(lo, hi)
    if hi - lo < 1e-9:
        return pd.Series([50.0] * len(s), index=s.index)
    return 100 * ((s - lo) / (hi - lo) if higher_better else (hi - s) / (hi - lo))


def _reel_buyume(nominal_series, tufe_series):
    """Nominal büyümeyi TÜFE ile düzelterek reel büyüme döndürür.
    Formül: ((1 + nominal/100) / (1 + tufe/100) - 1) * 100
    Örn: nominal %150, TÜFE %32 → reel %89.4"""
    return ((1 + nominal_series/100) / (1 + tufe_series/100) - 1) * 100


def _nan_series(df):
    return pd.Series([np.nan]*len(df), index=df.index)


def skorla(df):
    out = pd.DataFrame()
    out['Hisse'] = df['Hisse Adı']
    out['Firma'] = df['Firma Adı']
    out['Sektor'] = df.get('Firma Sektörü', '')
    out['Piyasa Değeri (mn TL)'] = df['Piyasa Değeri'] / 1e6
    out['Çalışan'] = df.get('Çalışan Sayısı', 0)

    # TÜFE — makro oran, her satırda aynı değer
    tufe = df.get('TÜFE Yıllık', pd.Series([0]*len(df), index=df.index))

    # ============== 1) DEĞERLEME (%25) ==============
    fk = df['F/K Günlük'].where(df['F/K Günlük'] > 0)
    pddd = df['PD/DD Günlük'].where(df['PD/DD Günlük'] > 0)
    ev = df['FD/FAVOK Günlük'].where(df['FD/FAVOK Günlük'] > 0)
    peg = df.get('Peg (Çeyreklik)', _nan_series(df))
    peg_pos = peg.where(peg > 0)

    out['F/K'] = df['F/K Günlük']
    out['PD/DD'] = df['PD/DD Günlük']
    out['FD/FAVÖK'] = df['FD/FAVOK Günlük']
    out['PEG'] = peg

    out['Değerleme Skoru'] = pd.concat([
        normalize(fk, False),
        normalize(pddd, False),
        normalize(ev, False),
        normalize(peg_pos, False),
    ], axis=1).mean(axis=1)

    # ============== 2) KARLILIK (%25) ==============
    out['ROE %'] = df['Özsermaye Karlılığı']
    out['ROIC %'] = df['ROIC']
    out['ROA %'] = df.get('Aktif Karlılığı', _nan_series(df))
    out['Net Marj %'] = df['Net Kar Marjı (Yıllık %)']
    out['FAVÖK Marjı %'] = df['Favök Marjı (Yıllık %)']
    out['Brüt Marj %'] = df.get('Brüt Kar Marjı (Yıllık %)', _nan_series(df))

    out['Karlılık Skoru'] = pd.concat([
        normalize(df['Özsermaye Karlılığı']),
        normalize(df['ROIC']),
        normalize(out['ROA %']),
        normalize(df['Net Kar Marjı (Yıllık %)']),
        normalize(df['Favök Marjı (Yıllık %)']),
        normalize(out['Brüt Marj %']),
    ], axis=1).mean(axis=1)

    # ============== 3) BÜYÜME — REEL (%20) ==============
    # TÜFE düzeltilmiş! (CANSLIM C ve A unsurları)
    nom_satis = df['Satış Gelirleri (Yıllık %)']
    nom_kar = df['Ana Ortaklık Payları (Yıllık %)']
    reel_satis = _reel_buyume(nom_satis, tufe)
    reel_kar = _reel_buyume(nom_kar, tufe)

    out['Satış Büyüme Nominal %'] = nom_satis
    out['Kâr Büyüme Nominal %'] = nom_kar
    out['Reel Satış %'] = reel_satis
    out['Reel Kâr %'] = reel_kar
    out['Çeyreklik Satış %'] = df.get('Satış Gelirleri (Çeyreklik %)', _nan_series(df))
    out['Çeyreklik Kâr %'] = df.get('Ana Ortaklık Payları (Çeyreklik %)', _nan_series(df))

    out['Büyüme Skoru'] = pd.concat([
        normalize(reel_satis),
        normalize(reel_kar),
        normalize(out['Çeyreklik Satış %']),
        normalize(out['Çeyreklik Kâr %']),
    ], axis=1).mean(axis=1)

    # ============== 4) BİLANÇO SAĞLIĞI (%15) ==============
    out['Borç/Özkaynak'] = df['Borç/ÖzKaynak']
    out['Cari Oran'] = df['Cari Oranı']
    out['Likidite Oranı'] = df.get('Likidite Oranı', _nan_series(df))
    out['Nakit Oranı'] = df.get('Nakit Oranı', _nan_series(df))
    out['Borç/FAVÖK'] = df.get('Borç Favök', _nan_series(df))
    out['Faiz Karşılama'] = df.get('Faiz Karşılama Oranı', _nan_series(df))
    out['Net YPP'] = df.get('Net Yabancı Para Pozisyonu', _nan_series(df))

    bf_series = df.get('Borç Favök', _nan_series(df))
    out['Bilanço Skoru'] = pd.concat([
        normalize(df['Borç/ÖzKaynak'].where(df['Borç/ÖzKaynak'] >= 0), False),
        normalize(df['Cari Oranı']),
        normalize(out['Likidite Oranı']),
        normalize(out['Nakit Oranı']),
        normalize(bf_series.where(bf_series >= 0), False),
        normalize(out['Faiz Karşılama'].where(out['Faiz Karşılama'] >= 0)),
    ], axis=1).mean(axis=1)

    # ============== 5) OPERASYONEL VERİMLİLİK (%10) ==============
    out['Stok Devir'] = df.get('Stok Devir Hızı', _nan_series(df))
    out['DSO'] = df.get('Alacak Tahsilat Süresi (DSO)', _nan_series(df))
    out['Nakit Çevirme'] = df.get('Nakit Çevirme Süresi', _nan_series(df))

    out['Verimlilik Skoru'] = pd.concat([
        normalize(out['Stok Devir'].where(out['Stok Devir'] >= 0)),
        normalize(out['DSO'].where(out['DSO'] >= 0), False),
        normalize(out['Nakit Çevirme'], False),
    ], axis=1).mean(axis=1)

    # ============== 6) PİYASA SİNYALLERİ (%5) ==============
    out['Yabancı Oran %'] = df.get('Yabancı Oran', _nan_series(df))
    out['Bilanço Sonrası %'] = df.get('Bilanço Sonrası Getiri (%)', _nan_series(df))
    out['Volatilite'] = df.get('Volatilite', _nan_series(df))
    out['Fiili Dolaşım %'] = df.get('Fiili Dolaşım (%)', _nan_series(df))

    out['Piyasa Skoru'] = pd.concat([
        normalize(out['Yabancı Oran %']),
        normalize(out['Bilanço Sonrası %']),
        normalize(out['Volatilite'], False),
    ], axis=1).mean(axis=1)

    # ============== TOPLAM SKOR ==============
    out['TOPLAM SKOR'] = (
        AGIRLIKLAR_V2['degerleme'] * out['Değerleme Skoru'].fillna(0)
        + AGIRLIKLAR_V2['karlilik'] * out['Karlılık Skoru'].fillna(0)
        + AGIRLIKLAR_V2['buyume'] * out['Büyüme Skoru'].fillna(0)
        + AGIRLIKLAR_V2['bilanco'] * out['Bilanço Skoru'].fillna(0)
        + AGIRLIKLAR_V2['verimlilik'] * out['Verimlilik Skoru'].fillna(0)
        + AGIRLIKLAR_V2['piyasa'] * out['Piyasa Skoru'].fillna(0)
    )

    # Eski kod uyumu (PDF/kart kodu bu adları beklediği yerlerde)
    out['Satış Büyüme %'] = out['Reel Satış %']
    out['Kâr Büyüme %'] = out['Reel Kâr %']

    return out.sort_values('TOPLAM SKOR', ascending=False).reset_index(drop=True)


def kirmizi_bayraklar(row):
    flags = []
    nm, em = row['Net Marj %'], row['FAVÖK Marjı %']
    if pd.notna(nm) and pd.notna(em) and nm > em + 1:
        flags.append(f"Net Marj ({nm:.1f}%) &gt; FAVÖK Marjı ({em:.1f}%) — kâr kalitesi sorgulanmalı")
    if pd.notna(row['Çalışan']) and row['Çalışan'] == 0:
        flags.append("Çalışan sayısı 0 — veri eksik veya sıra dışı yapı")
    firma = str(row['Firma']).upper()
    if 'YATIRIM' in firma or 'HOLDİNG' in firma or 'GİRİŞİM' in firma:
        flags.append("Yatırım/Holding yapısı olabilir — NAV bazlı bakılmalı")
    if pd.notna(row['F/K']) and row['F/K'] < 0:
        flags.append(f"F/K negatif ({row['F/K']:.1f}) — şirket zarar ediyor")
    if pd.notna(row['Cari Oran']) and row['Cari Oran'] < 1:
        flags.append(f"Cari Oran {row['Cari Oran']:.2f} — kısa vadeli likidite riski")
    if pd.notna(row['Kâr Büyüme %']) and row['Kâr Büyüme %'] > 300:
        flags.append(f"Kâr büyümesi %{row['Kâr Büyüme %']:,.0f} — baz etkisi, sürdürülebilir saymayın")
    yp, pv = row['Net YPP'], row['Piyasa Değeri (mn TL)']
    if pd.notna(yp) and pd.notna(pv) and yp < 0 and abs(yp)/(pv*1e6) > 0.3:
        flags.append(f"Net döviz pozisyonu büyük negatif ({yp/1e9:.1f} mlr TL) — kur riski")
    if pd.notna(row['Volatilite']) and row['Volatilite'] > 8:
        flags.append(f"Volatilite yüksek ({row['Volatilite']:.1f})")
    if pd.notna(row['Fiili Dolaşım %']) and row['Fiili Dolaşım %'] < 25:
        flags.append(f"Fiili dolaşım düşük (%{row['Fiili Dolaşım %']:.1f}) — likidite sınırlı")
    return flags


SEKTOR_BAGLAMI = {
    'Bilişim': "Bilişim/yazılım: yüksek FAVÖK marjı (%30-70) normaldir, büyüme kritik, borç tipik düşük.",
    'Gıda': "Gıda: defansif, düşük marjlı (FAVÖK %5-20), emtia ve kur duyarlılığı yüksek.",
    'Kimya': "Kimya-İlaç-Petrol: çok döngüsel, emtia fiyatları belirleyici, tek dönem yanıltıcı olabilir.",
    'Banka': "Bankacılık: FAVÖK anlamsız; F/K, PD/DD, ROE öne çıkar.",
    'Holding': "Holding: NAV bakılır, konsolide marjlar yanıltıcı.",
}


def sektor_baglam(sektor):
    if not sektor:
        return ""
    for k, v in SEKTOR_BAGLAMI.items():
        if k.lower() in str(sektor).lower():
            return v
    return ""


# ==========================================================
# Telegram için HTML rapor üretimi
# ==========================================================
def he(s):
    """HTML escape, None'a karşı güvenli."""
    return html_escape(str(s)) if s is not None else ''


def baslik_mesaji(skor_df):
    n = len(skor_df)
    zarar = (skor_df['F/K'] < 0).sum()
    sektor = skor_df.iloc[0]['Sektor'] if pd.notna(skor_df.iloc[0]['Sektor']) else 'BIST'
    baglam = sektor_baglam(sektor)

    msg = [
        f"📊 <b>BIST Sektör Analizi</b>",
        f"<i>{he(sektor)}</i>",
        "",
        f"• Analiz edilen hisse: <b>{n}</b>",
        f"• Zarar eden hisse: <b>{zarar}</b> (toplamın %{zarar*100/n:.0f}'i)",
        f"• <b>6 boyutlu skorlama</b> (Yaşar Erdinç + CANSLIM)",
        f"   Değerleme %25 · Karlılık %25 · Büyüme(REEL) %20",
        f"   Bilanço %15 · Verimlilik %10 · Piyasa %5",
    ]
    if baglam:
        msg += ["", f"💡 <i>{he(baglam)}</i>"]
    return '\n'.join(msg)


def puanlama_tablosu_mesaji(skor_df):
    """Tüm hisseleri 6 alt skorla. Monospace <pre> için sıkı format."""
    lines = []
    lines.append(" #  Hisse  Topl  Değ Kâr Büy Bil Vrm Piy   F/K   ROE%")
    lines.append("─" * 58)
    for i, r in skor_df.iterrows():
        fk_s = f"{r['F/K']:6.1f}" if pd.notna(r['F/K']) else "    —"
        roe_s = f"{r['ROE %']:5.1f}" if pd.notna(r['ROE %']) else "    —"
        lines.append(
            f"{i+1:2}  {r['Hisse']:<6}{r['TOPLAM SKOR']:5.1f}  "
            f"{r['Değerleme Skoru']:3.0f} {r['Karlılık Skoru']:3.0f} "
            f"{r['Büyüme Skoru']:3.0f} {r['Bilanço Skoru']:3.0f} "
            f"{r['Verimlilik Skoru']:3.0f} {r['Piyasa Skoru']:3.0f} "
            f"{fk_s} {roe_s}"
        )
    body = '\n'.join(lines)
    return f"<b>📋 PUANLAMA TABLOSU</b>\n<pre>{he(body)}</pre>"


def hisse_kart_mesaji(row, sira):
    """Bir hisse için 6 boyutlu detaylı kart mesajı."""
    lines = []
    lines.append(f"<b>#{sira} • {he(row['Hisse'])}</b> — <i>{he(row['Firma'][:60])}</i>")
    lines.append("")
    lines.append(f"🎯 <b>Skor: {row['TOPLAM SKOR']:.1f}/100</b>")
    lines.append(f"   Değ:{row['Değerleme Skoru']:.0f} · "
                 f"Kâr:{row['Karlılık Skoru']:.0f} · "
                 f"Büy:{row['Büyüme Skoru']:.0f} · "
                 f"Bil:{row['Bilanço Skoru']:.0f} · "
                 f"Vrm:{row['Verimlilik Skoru']:.0f} · "
                 f"Piy:{row['Piyasa Skoru']:.0f}")
    pd_val = row['Piyasa Değeri (mn TL)']
    pd_str = f"{pd_val/1000:.1f} mlr TL" if pd_val >= 1000 else f"{pd_val:.0f} mn TL"
    cal = int(row['Çalışan']) if pd.notna(row['Çalışan']) else "—"
    lines.append(f"   💰 PD: {pd_str}  •  👥 Çalışan: {cal}")
    lines.append("")

    # 1) DEĞERLEME
    fk_y = ""
    if pd.notna(row['F/K']):
        if row['F/K'] < 0: fk_y = " <i>(zarar)</i>"
        elif row['F/K'] < 8: fk_y = " <i>(çok ucuz)</i>"
        elif row['F/K'] < 15: fk_y = " <i>(makul)</i>"
        elif row['F/K'] < 25: fk_y = " <i>(ortalama)</i>"
        else: fk_y = " <i>(pahalı)</i>"
    lines.append("<b>💵 Değerleme</b>")
    lines.append(f"   F/K: <b>{row['F/K']:.1f}</b>{fk_y}")
    if pd.notna(row['PD/DD']):
        lines.append(f"   PD/DD: <b>{row['PD/DD']:.1f}</b>")
    if pd.notna(row['FD/FAVÖK']):
        lines.append(f"   FD/FAVÖK: <b>{row['FD/FAVÖK']:.1f}</b>")
    if pd.notna(row['PEG']):
        peg_y = ""
        if 0 < row['PEG'] < 1: peg_y = " <i>(büyümeye göre ucuz)</i>"
        elif row['PEG'] > 2: peg_y = " <i>(büyümeye göre pahalı)</i>"
        lines.append(f"   PEG: <b>{row['PEG']:.2f}</b>{peg_y}")

    # 2) KARLILIK
    lines.append("")
    lines.append("<b>📈 Karlılık</b>")
    if pd.notna(row['ROE %']):
        roe_y = ""
        if row['ROE %'] < 0: roe_y = " <i>(zarar)</i>"
        elif row['ROE %'] < 10: roe_y = " <i>(düşük)</i>"
        elif row['ROE %'] < 20: roe_y = " <i>(iyi)</i>"
        else: roe_y = " <i>(çok güçlü)</i>"
        lines.append(f"   ROE: <b>%{row['ROE %']:.1f}</b>{roe_y}")
    if pd.notna(row['ROIC %']):
        lines.append(f"   ROIC: <b>%{row['ROIC %']:.1f}</b>")
    if pd.notna(row['ROA %']):
        lines.append(f"   ROA: <b>%{row['ROA %']:.1f}</b>")
    if pd.notna(row['Net Marj %']):
        lines.append(f"   Net Marj: <b>%{row['Net Marj %']:.1f}</b>")
    if pd.notna(row['FAVÖK Marjı %']):
        em_y = ""
        if row['FAVÖK Marjı %'] > 30: em_y = " <i>(çok yüksek)</i>"
        elif row['FAVÖK Marjı %'] > 15: em_y = " <i>(iyi)</i>"
        elif row['FAVÖK Marjı %'] > 5: em_y = " <i>(ortalama)</i>"
        else: em_y = " <i>(zayıf)</i>"
        lines.append(f"   FAVÖK Marjı: <b>%{row['FAVÖK Marjı %']:.1f}</b>{em_y}")
    if pd.notna(row['Brüt Marj %']):
        lines.append(f"   Brüt Marj: <b>%{row['Brüt Marj %']:.1f}</b>")

    # 3) BÜYÜME (REEL!)
    lines.append("")
    lines.append("<b>🚀 Büyüme (TÜFE düzeltmeli)</b>")
    if pd.notna(row['Reel Satış %']):
        sb = row['Reel Satış %']
        sb_y = ""
        if sb < 0: sb_y = " <i>(reel daralma!)</i>"
        elif sb < 10: sb_y = " <i>(zayıf)</i>"
        elif sb < 30: sb_y = " <i>(sağlıklı)</i>"
        else: sb_y = " <i>(güçlü)</i>"
        nom = row.get('Satış Büyüme Nominal %', None)
        nom_str = f" <i>(nom %{nom:.0f})</i>" if pd.notna(nom) else ""
        lines.append(f"   Reel Satış: <b>%{sb:.1f}</b>{sb_y}{nom_str}")
    if pd.notna(row['Reel Kâr %']):
        rk = row['Reel Kâr %']
        nom = row.get('Kâr Büyüme Nominal %', None)
        nom_str = f" <i>(nom %{nom:,.0f})</i>" if pd.notna(nom) else ""
        lines.append(f"   Reel Kâr: <b>%{rk:,.1f}</b>{nom_str}")
    if pd.notna(row['Çeyreklik Satış %']):
        lines.append(f"   Çeyrek Satış: <b>%{row['Çeyreklik Satış %']:.1f}</b>")
    if pd.notna(row['Çeyreklik Kâr %']):
        lines.append(f"   Çeyrek Kâr: <b>%{row['Çeyreklik Kâr %']:,.1f}</b>")

    # 4) BİLANÇO
    lines.append("")
    lines.append("<b>🏦 Bilanço Sağlığı</b>")
    if pd.notna(row['Borç/Özkaynak']):
        de = row['Borç/Özkaynak']
        de_y = ""
        if de < 0.3: de_y = " <i>(neredeyse borçsuz)</i>"
        elif de < 0.7: de_y = " <i>(düşük)</i>"
        elif de < 1.5: de_y = " <i>(orta)</i>"
        else: de_y = " <i>(yüksek risk)</i>"
        lines.append(f"   Borç/ÖzK: <b>{de:.2f}</b>{de_y}")
    if pd.notna(row['Cari Oran']):
        co = row['Cari Oran']
        co_y = " <i>(likidite riski)</i>" if co < 1 else " <i>(sıkı)</i>" if co < 1.5 else " <i>(rahat)</i>"
        lines.append(f"   Cari Oran: <b>{co:.2f}</b>{co_y}")
    if pd.notna(row['Likidite Oranı']):
        lines.append(f"   Likidite (asit-test): <b>{row['Likidite Oranı']:.2f}</b>")
    if pd.notna(row['Faiz Karşılama']):
        fk_y = ""
        if row['Faiz Karşılama'] < 1.5: fk_y = " <i>(kritik!)</i>"
        elif row['Faiz Karşılama'] < 3: fk_y = " <i>(zayıf)</i>"
        elif row['Faiz Karşılama'] > 10: fk_y = " <i>(çok güçlü)</i>"
        lines.append(f"   Faiz Karşılama: <b>{row['Faiz Karşılama']:.1f}</b>{fk_y}")

    # 5) OPERASYONEL VERİMLİLİK (YENİ)
    if any(pd.notna(row.get(k)) for k in ['Stok Devir', 'DSO', 'Nakit Çevirme']):
        lines.append("")
        lines.append("<b>⚙️ Operasyonel Verimlilik</b>")
        if pd.notna(row.get('Stok Devir')):
            lines.append(f"   Stok Devir: <b>{row['Stok Devir']:.1f}</b>x")
        if pd.notna(row.get('DSO')):
            dso_y = " <i>(hızlı tahsilat)</i>" if row['DSO'] < 60 else (" <i>(yavaş)</i>" if row['DSO'] > 120 else "")
            lines.append(f"   Tahsilat Süresi: <b>{row['DSO']:.0f} gün</b>{dso_y}")
        if pd.notna(row.get('Nakit Çevirme')):
            nc_y = " <i>(çok iyi)</i>" if row['Nakit Çevirme'] < 0 else ""
            lines.append(f"   Nakit Çevirme: <b>{row['Nakit Çevirme']:.0f} gün</b>{nc_y}")

    # 6) PİYASA SİNYALLERİ (YENİ)
    if any(pd.notna(row.get(k)) for k in ['Yabancı Oran %', 'Bilanço Sonrası %', 'Volatilite']):
        lines.append("")
        lines.append("<b>📡 Piyasa Sinyalleri</b>")
        if pd.notna(row.get('Yabancı Oran %')):
            yo = row['Yabancı Oran %']
            yo_y = " <i>(yüksek kurumsal ilgi)</i>" if yo > 20 else " <i>(düşük)</i>" if yo < 3 else ""
            lines.append(f"   Yabancı Oran: <b>%{yo:.1f}</b>{yo_y}")
        if pd.notna(row.get('Bilanço Sonrası %')):
            bs = row['Bilanço Sonrası %']
            bs_y = " <i>(piyasa beğendi)</i>" if bs > 5 else (" <i>(piyasa beğenmedi)</i>" if bs < -5 else "")
            lines.append(f"   Bilanço Sonrası: <b>%{bs:+.1f}</b>{bs_y}")
        if pd.notna(row.get('Volatilite')):
            vol_y = " <i>(yüksek oynaklık)</i>" if row['Volatilite'] > 8 else ""
            lines.append(f"   Volatilite: <b>{row['Volatilite']:.1f}</b>{vol_y}")

    # KIRMIZI BAYRAKLAR
    flags = kirmizi_bayraklar(row)
    if flags:
        lines.append("")
        lines.append("<b>⚠️ Kırmızı Bayraklar</b>")
        for f in flags:
            lines.append(f"   • {f}")

    return '\n'.join(lines)


def uyari_mesaji():
    return (
        "⚠️ <b>Uyarı:</b> Bu bir <b>yatırım tavsiyesi değildir</b>. "
        "Analiz tek bir bilanço dönemine dayalı kamuya açık verilerin "
        "sistematik kıyaslamasıdır. Geçmiş performans gelecek getiriyi "
        "garanti etmez. Kararlar kişisel risk profili, vade ve portföy "
        "çeşitlendirmesi gözetilerek; gerekirse lisanslı bir yatırım "
        "danışmanıyla alınmalıdır."
    )


def markdown_rapor(skor_df, dosya_adi):
    """Geriye dönük: hâlâ MD üretebilir ama PDF tercih edilir."""
    n = len(skor_df)
    md = [f"# BIST Sektör Analiz Raporu",
          f"\n**Kaynak:** `{dosya_adi}`  ",
          f"**Hisse Sayısı:** {n}  "]
    md.append("\n## Puanlama Tablosu\n")
    md.append("| # | Hisse | Toplam | Değ | Kâr | Büy | Bil | F/K | ROE% | FAVÖK% | Borç/ÖzK |")
    md.append("|---|-------|-------:|----:|----:|----:|----:|----:|-----:|-------:|---------:|")
    for i, r in skor_df.iterrows():
        md.append(
            f"| {i+1} | **{r['Hisse']}** | {r['TOPLAM SKOR']:.1f} | "
            f"{r['Değerleme Skoru']:.0f} | {r['Karlılık Skoru']:.0f} | "
            f"{r['Büyüme Skoru']:.0f} | {r['Bilanço Skoru']:.0f} | "
            f"{r['F/K']:.1f} | {r['ROE %']:.1f} | "
            f"{r['FAVÖK Marjı %']:.1f} | {r['Borç/Özkaynak']:.2f} |"
        )
    return '\n'.join(md)


# PDF rapor üretici (ayrı modülde — kalabalık olmasın diye)
from pdf_rapor import pdf_uret


# ==========================================================
# Dosya işleyici
# ==========================================================
def handle_document(message):
    chat_id = message['chat']['id']
    doc = message['document']
    file_name = doc.get('file_name', 'dosya')

    # Yetki kontrolü
    if AUTH_IDS and chat_id not in AUTH_IDS:
        send_message(chat_id, "⛔ Bu botu kullanma yetkiniz yok.")
        return

    # xlsx kontrolü
    if not file_name.lower().endswith(('.xlsx', '.xls')):
        send_message(chat_id,
                     "❌ Lütfen <b>.xlsx</b> uzantılı bir BIST sektör dosyası gönderin.")
        return

    send_chat_action(chat_id, 'typing')
    send_message(chat_id, f"📥 <b>{he(file_name)}</b> alındı, analiz ediliyor…")

    with tempfile.TemporaryDirectory() as tmp:
        # İndir
        path = download_telegram_file(doc['file_id'], tmp)
        if not path:
            send_message(chat_id, "❌ Dosya indirilemedi.")
            return

        try:
            df = pd.read_excel(path)
        except Exception as e:
            send_message(chat_id, f"❌ Excel okunamadı: <code>{he(str(e))}</code>")
            return

        # Gerekli sütun kontrolü
        zorunlu = ['Hisse Adı', 'Firma Adı', 'Piyasa Değeri', 'F/K Günlük',
                   'PD/DD Günlük', 'FD/FAVOK Günlük', 'Özsermaye Karlılığı',
                   'ROIC', 'Net Kar Marjı (Yıllık %)', 'Favök Marjı (Yıllık %)',
                   'Satış Gelirleri (Yıllık %)', 'Ana Ortaklık Payları (Yıllık %)',
                   'Borç/ÖzKaynak', 'Cari Oranı']
        eksik = [c for c in zorunlu if c not in df.columns]
        if eksik:
            send_message(chat_id,
                f"❌ Beklenen sütunlar eksik:\n<code>{he(', '.join(eksik))}</code>\n\n"
                f"Bu, bu botun beklediği BIST sektör formatı değil gibi görünüyor.")
            return

        try:
            skor = skorla(df)
        except Exception as e:
            traceback.print_exc()
            send_message(chat_id, f"❌ Analiz sırasında hata: <code>{he(str(e))}</code>")
            return

        # 1) Başlık
        send_message(chat_id, baslik_mesaji(skor))
        time.sleep(0.3)

        # 2) Puanlama tablosu
        send_message(chat_id, puanlama_tablosu_mesaji(skor))
        time.sleep(0.3)

        # 3) Top N detay kartları
        send_message(chat_id, f"🔍 <b>İlk {min(TOP_N_DETAY, len(skor))} Hissenin Detayı</b>")
        for i, row in skor.head(TOP_N_DETAY).iterrows():
            send_message(chat_id, hisse_kart_mesaji(row, i+1))

        # 4) Uyarı
        send_message(chat_id, uyari_mesaji())

        # 5) Tam PDF rapor
        try:
            send_chat_action(chat_id, 'upload_document')
            pdf_bytes = pdf_uret(skor, file_name, top_n=TOP_N_DETAY)
            pdf_path = Path(tmp) / (Path(file_name).stem + '_analiz.pdf')
            pdf_path.write_bytes(pdf_bytes)
            send_document(chat_id, str(pdf_path),
                          caption="📎 <b>Tam Rapor (PDF)</b>")
        except Exception as e:
            print(f"PDF üretim hatası: {e}", file=sys.stderr)
            traceback.print_exc()
            # Yedek: PDF üretilemezse MD gönder
            md_text = markdown_rapor(skor, file_name)
            md_path = Path(tmp) / (Path(file_name).stem + '_analiz.md')
            md_path.write_text(md_text, encoding='utf-8')
            send_document(chat_id, str(md_path),
                          caption="📎 Tam rapor (Markdown — PDF üretilemedi)")


def handle_text(message):
    chat_id = message['chat']['id']
    text = message.get('text', '').strip()

    if AUTH_IDS and chat_id not in AUTH_IDS:
        send_message(chat_id, f"⛔ Yetki yok. Chat ID'niz: <code>{chat_id}</code>")
        return

    if text.startswith('/start') or text.startswith('/help'):
        send_message(chat_id,
            "👋 <b>BIST Sektör Analiz Botu</b>\n\n"
            "Bana bir BIST sektör Excel dosyası (.xlsx) gönder, sana:\n"
            "• Tüm hisselerin puanlama tablosunu\n"
            "• İlk 10 hissenin detaylı analizini\n"
            "• Tam markdown raporunu\n\n"
            "geri göndereyim.\n\n"
            "⚠️ <i>Bu bot yatırım tavsiyesi vermez, sadece kamuya açık verilerin "
            "sistematik kıyaslamasını yapar.</i>"
        )
    elif text.startswith('/id'):
        send_message(chat_id, f"Chat ID'niz: <code>{chat_id}</code>")
    else:
        send_message(chat_id, "Bana bir .xlsx dosyası gönderin. Yardım için /help.")


# ==========================================================
# Update işleyici (long-polling ve --once için ortak)
# ==========================================================
def process_update(update):
    msg = update.get('message') or update.get('edited_message')
    if not msg:
        return
    try:
        if 'document' in msg:
            handle_document(msg)
        elif 'text' in msg:
            handle_text(msg)
    except Exception:
        traceback.print_exc()
        chat_id = msg.get('chat', {}).get('id')
        if chat_id:
            send_message(chat_id, "❌ Beklenmedik bir hata oluştu. Logları kontrol edin.")


def get_updates(offset=0, timeout=30):
    try:
        r = requests.get(
            f"{API}/getUpdates",
            params={'offset': offset, 'timeout': timeout},
            timeout=timeout + 10
        )
        return r.json()
    except Exception as e:
        print(f"getUpdates error: {e}", file=sys.stderr)
        return {'ok': False}


# ==========================================================
# Ana akış
# ==========================================================
def loop_mode():
    """Sürekli mod — VPS/local için."""
    print("Bot başladı (long-polling). Durdurmak için Ctrl+C.")
    offset = 0
    while True:
        try:
            res = get_updates(offset=offset, timeout=30)
            if res.get('ok'):
                for u in res.get('result', []):
                    offset = u['update_id'] + 1
                    process_update(u)
            else:
                time.sleep(3)
        except KeyboardInterrupt:
            print("\nDuruldu.")
            break
        except Exception:
            traceback.print_exc()
            time.sleep(3)


def once_mode():
    """Tek seferlik mod — GitHub Actions cron için.
    Onaylanmamış mesajları çeker, HER BİRİNİ İŞLEDİKTEN HEMEN SONRA
    Telegram'a 'aldım' bildirir. Böylece bir sonraki cron tetiklenmesinde
    aynı dosya tekrar işlenmez."""
    print("Tek seferlik mod.")

    res = get_updates(offset=0, timeout=0)
    if not res.get('ok'):
        print(f"getUpdates başarısız: {res}", file=sys.stderr)
        return

    updates = res.get('result', [])
    if not updates:
        print("Bekleyen mesaj yok.")
        return

    print(f"{len(updates)} update işlenecek.")

    for u in updates:
        update_id = u['update_id']
        try:
            process_update(u)
            print(f"  ✓ Update {update_id} işlendi.")
        except Exception as e:
            print(f"  ✗ Update {update_id} hatası: {e}", file=sys.stderr)
            traceback.print_exc()

        # KRİTİK: her update'ten HEMEN sonra acknowledge et.
        # Bu sayede script çökerse bile, işlenmiş olanlar tekrar işlenmez.
        next_offset = update_id + 1
        for attempt in range(3):  # 3 deneme
            try:
                ack = requests.get(
                    f"{API}/getUpdates",
                    params={'offset': next_offset, 'timeout': 0, 'limit': 1},
                    timeout=15
                ).json()
                if ack.get('ok'):
                    print(f"  ✓ ACK update_id={update_id} (offset={next_offset})")
                    break
                else:
                    print(f"  ⚠ ACK başarısız (deneme {attempt+1}): {ack}",
                          file=sys.stderr)
            except Exception as e:
                print(f"  ⚠ ACK exception (deneme {attempt+1}): {e}",
                      file=sys.stderr)
            time.sleep(1)
        else:
            # 3 deneme de başarısızsa devam et (bir sonraki cron'da tekrar denenir)
            print(f"  ✗ ACK kalıcı başarısız update_id={update_id}",
                  file=sys.stderr)

    print("Tamam.")


def ack_all_mode():
    """Bekleyen TÜM mesajları işleme almadan acknowledge eder.
    Kuyrukta birikmiş eski dosyaları temizlemek için kullanılır."""
    print("Kuyruk temizleme modu — bekleyen mesajlar işlenmeden silinecek.")

    res = get_updates(offset=0, timeout=0)
    if not res.get('ok'):
        print(f"getUpdates başarısız: {res}", file=sys.stderr)
        return

    updates = res.get('result', [])
    if not updates:
        print("Bekleyen mesaj zaten yok.")
        return

    last_update_id = max(u['update_id'] for u in updates)
    next_offset = last_update_id + 1
    print(f"{len(updates)} bekleyen mesaj bulundu, hepsi temizleniyor...")

    ack = requests.get(
        f"{API}/getUpdates",
        params={'offset': next_offset, 'timeout': 0, 'limit': 1},
        timeout=15
    ).json()
    if ack.get('ok'):
        print(f"✓ {len(updates)} mesaj temizlendi (offset={next_offset}).")
    else:
        print(f"✗ Temizleme başarısız: {ack}", file=sys.stderr)


if __name__ == "__main__":
    if '--ack-all' in sys.argv:
        ack_all_mode()
    elif '--once' in sys.argv:
        once_mode()
    else:
        loop_mode()
