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
# Analiz mantığı (bist_analiz.py'den uyarlanmış)
# ==========================================================
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


def skorla(df):
    out = pd.DataFrame()
    out['Hisse'] = df['Hisse Adı']
    out['Firma'] = df['Firma Adı']
    out['Sektor'] = df.get('Firma Sektörü', '')
    out['Piyasa Değeri (mn TL)'] = df['Piyasa Değeri'] / 1e6
    out['Çalışan'] = df.get('Çalışan Sayısı', 0)

    fk = df['F/K Günlük'].where(df['F/K Günlük'] > 0)
    pddd = df['PD/DD Günlük'].where(df['PD/DD Günlük'] > 0)
    ev = df['FD/FAVOK Günlük'].where(df['FD/FAVOK Günlük'] > 0)

    out['F/K'] = df['F/K Günlük']
    out['PD/DD'] = df['PD/DD Günlük']
    out['FD/FAVÖK'] = df['FD/FAVOK Günlük']
    out['PEG'] = df.get('Peg (Çeyreklik)', np.nan)

    out['Değerleme Skoru'] = pd.concat([
        normalize(fk, False), normalize(pddd, False), normalize(ev, False)
    ], axis=1).mean(axis=1)

    out['ROE %'] = df['Özsermaye Karlılığı']
    out['ROIC %'] = df['ROIC']
    out['Net Marj %'] = df['Net Kar Marjı (Yıllık %)']
    out['FAVÖK Marjı %'] = df['Favök Marjı (Yıllık %)']

    out['Karlılık Skoru'] = pd.concat([
        normalize(df['Özsermaye Karlılığı']),
        normalize(df['ROIC']),
        normalize(df['Net Kar Marjı (Yıllık %)']),
        normalize(df['Favök Marjı (Yıllık %)']),
    ], axis=1).mean(axis=1)

    out['Satış Büyüme %'] = df['Satış Gelirleri (Yıllık %)']
    out['Kâr Büyüme %'] = df['Ana Ortaklık Payları (Yıllık %)']
    out['Çeyreklik Satış %'] = df.get('Satış Gelirleri (Çeyreklik %)', np.nan)

    out['Büyüme Skoru'] = pd.concat([
        normalize(df['Satış Gelirleri (Yıllık %)']),
        normalize(df['Ana Ortaklık Payları (Yıllık %)']),
    ], axis=1).mean(axis=1)

    out['Borç/Özkaynak'] = df['Borç/ÖzKaynak']
    out['Cari Oran'] = df['Cari Oranı']
    out['Borç/FAVÖK'] = df.get('Borç Favök', np.nan)
    out['Net YPP'] = df.get('Net Yabancı Para Pozisyonu', np.nan)
    out['Volatilite'] = df.get('Volatilite', np.nan)
    out['Fiili Dolaşım %'] = df.get('Fiili Dolaşım (%)', np.nan)

    bf_series = df.get('Borç Favök', pd.Series([np.nan]*len(df), index=df.index))
    out['Bilanço Skoru'] = pd.concat([
        normalize(df['Borç/ÖzKaynak'].where(df['Borç/ÖzKaynak'] >= 0), False),
        normalize(df['Cari Oranı']),
        normalize(bf_series.where(bf_series >= 0), False),
    ], axis=1).mean(axis=1)

    out['TOPLAM SKOR'] = (
        AGIRLIKLAR['degerleme'] * out['Değerleme Skoru'].fillna(0)
        + AGIRLIKLAR['karlilik'] * out['Karlılık Skoru'].fillna(0)
        + AGIRLIKLAR['buyume'] * out['Büyüme Skoru'].fillna(0)
        + AGIRLIKLAR['bilanco'] * out['Bilanço Skoru'].fillna(0)
    )

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
        f"• Ağırlıklar: Değerleme %{int(AGIRLIKLAR['degerleme']*100)} | "
        f"Karlılık %{int(AGIRLIKLAR['karlilik']*100)} | "
        f"Büyüme %{int(AGIRLIKLAR['buyume']*100)} | "
        f"Bilanço %{int(AGIRLIKLAR['bilanco']*100)}",
    ]
    if baglam:
        msg += ["", f"💡 <i>{he(baglam)}</i>"]
    return '\n'.join(msg)


def puanlama_tablosu_mesaji(skor_df):
    """Tüm hisseleri puanlarıyla içeren tablo. Telegram için monospace <pre>."""
    lines = []
    lines.append(" #  Hisse  Toplam  Değ  Kâr  Büy  Bil    F/K   ROE%")
    lines.append("─" * 56)
    for i, r in skor_df.iterrows():
        fk = r['F/K']
        roe = r['ROE %']
        fk_s = f"{fk:6.1f}" if pd.notna(fk) else "    —"
        roe_s = f"{roe:5.1f}" if pd.notna(roe) else "    —"
        lines.append(
            f"{i+1:2}  {r['Hisse']:<6} {r['TOPLAM SKOR']:5.1f}  "
            f"{r['Değerleme Skoru']:3.0f}  {r['Karlılık Skoru']:3.0f}  "
            f"{r['Büyüme Skoru']:3.0f}  {r['Bilanço Skoru']:3.0f}  "
            f"{fk_s} {roe_s}"
        )
    body = '\n'.join(lines)
    return f"<b>📋 PUANLAMA TABLOSU</b>\n<pre>{he(body)}</pre>"


def hisse_kart_mesaji(row, sira):
    """Bir hisse için kart şeklinde detaylı mesaj."""
    lines = []
    lines.append(f"<b>#{sira} • {he(row['Hisse'])}</b> — <i>{he(row['Firma'][:60])}</i>")
    lines.append("")
    lines.append(f"🎯 <b>Skor: {row['TOPLAM SKOR']:.1f}/100</b>")
    lines.append(f"   Değ: {row['Değerleme Skoru']:.0f} | "
                 f"Kâr: {row['Karlılık Skoru']:.0f} | "
                 f"Büy: {row['Büyüme Skoru']:.0f} | "
                 f"Bil: {row['Bilanço Skoru']:.0f}")
    pd_val = row['Piyasa Değeri (mn TL)']
    pd_str = f"{pd_val/1000:.1f} mlr TL" if pd_val >= 1000 else f"{pd_val:.0f} mn TL"
    cal = int(row['Çalışan']) if pd.notna(row['Çalışan']) else "—"
    lines.append(f"   💰 PD: {pd_str}  •  👥 Çalışan: {cal}")
    lines.append("")

    # Değerleme
    fk_yorum = ""
    if pd.notna(row['F/K']):
        if row['F/K'] < 0: fk_yorum = " <i>(zarar)</i>"
        elif row['F/K'] < 8: fk_yorum = " <i>(çok ucuz)</i>"
        elif row['F/K'] < 15: fk_yorum = " <i>(makul)</i>"
        elif row['F/K'] < 25: fk_yorum = " <i>(ortalama)</i>"
        else: fk_yorum = " <i>(pahalı)</i>"
    lines.append("<b>💵 Değerleme</b>")
    lines.append(f"   F/K: <b>{row['F/K']:.1f}</b>{fk_yorum}")
    if pd.notna(row['PD/DD']):
        lines.append(f"   PD/DD: <b>{row['PD/DD']:.1f}</b>")
    if pd.notna(row['FD/FAVÖK']):
        lines.append(f"   FD/FAVÖK: <b>{row['FD/FAVÖK']:.1f}</b>")

    # Karlılık
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
    if pd.notna(row['Net Marj %']):
        lines.append(f"   Net Marj: <b>%{row['Net Marj %']:.1f}</b>")
    if pd.notna(row['FAVÖK Marjı %']):
        em_y = ""
        if row['FAVÖK Marjı %'] > 30: em_y = " <i>(çok yüksek)</i>"
        elif row['FAVÖK Marjı %'] > 15: em_y = " <i>(iyi)</i>"
        elif row['FAVÖK Marjı %'] > 5: em_y = " <i>(ortalama)</i>"
        else: em_y = " <i>(zayıf)</i>"
        lines.append(f"   FAVÖK Marjı: <b>%{row['FAVÖK Marjı %']:.1f}</b>{em_y}")

    # Büyüme
    lines.append("")
    lines.append("<b>🚀 Büyüme</b>")
    if pd.notna(row['Satış Büyüme %']):
        sb_y = ""
        if row['Satış Büyüme %'] < 0: sb_y = " <i>(daralma)</i>"
        elif row['Satış Büyüme %'] < 20: sb_y = " <i>(yavaş)</i>"
        elif row['Satış Büyüme %'] < 50: sb_y = " <i>(sağlıklı)</i>"
        else: sb_y = " <i>(güçlü)</i>"
        lines.append(f"   Satış: <b>%{row['Satış Büyüme %']:.1f}</b>{sb_y}")
    if pd.notna(row['Kâr Büyüme %']):
        lines.append(f"   Kâr: <b>%{row['Kâr Büyüme %']:,.1f}</b>")

    # Bilanço
    lines.append("")
    lines.append("<b>🏦 Bilanço</b>")
    if pd.notna(row['Borç/Özkaynak']):
        de = row['Borç/Özkaynak']
        de_y = ""
        if de < 0.3: de_y = " <i>(neredeyse borçsuz)</i>"
        elif de < 0.7: de_y = " <i>(düşük borç)</i>"
        elif de < 1.5: de_y = " <i>(orta kaldıraç)</i>"
        else: de_y = " <i>(yüksek risk)</i>"
        lines.append(f"   Borç/ÖzK: <b>{de:.2f}</b>{de_y}")
    if pd.notna(row['Cari Oran']):
        co = row['Cari Oran']
        co_y = " <i>(likidite riski)</i>" if co < 1 else " <i>(sıkı)</i>" if co < 1.5 else " <i>(rahat)</i>"
        lines.append(f"   Cari Oran: <b>{co:.2f}</b>{co_y}")

    # Kırmızı bayraklar
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
    """İlave dosya olarak gönderilecek tam markdown rapor."""
    n = len(skor_df)
    md = [f"# BIST Sektör Analiz Raporu",
          f"\n**Kaynak:** `{dosya_adi}`  ",
          f"**Hisse Sayısı:** {n}  ",
          "\n## Puanlama Tablosu\n",
          "| # | Hisse | Toplam | Değ | Kâr | Büy | Bil | F/K | ROE% | FAVÖK% | Borç/ÖzK |",
          "|---|-------|-------:|----:|----:|----:|----:|----:|-----:|-------:|---------:|"]
    for i, r in skor_df.iterrows():
        md.append(
            f"| {i+1} | **{r['Hisse']}** | {r['TOPLAM SKOR']:.1f} | "
            f"{r['Değerleme Skoru']:.0f} | {r['Karlılık Skoru']:.0f} | "
            f"{r['Büyüme Skoru']:.0f} | {r['Bilanço Skoru']:.0f} | "
            f"{r['F/K']:.1f} | {r['ROE %']:.1f} | "
            f"{r['FAVÖK Marjı %']:.1f} | {r['Borç/Özkaynak']:.2f} |"
        )
    md.append(f"\n## İlk {min(TOP_N_DETAY, n)} Detay\n")
    for i, r in skor_df.head(TOP_N_DETAY).iterrows():
        md.append(f"### {i+1}. {r['Hisse']} — {r['Firma']}")
        md.append(f"**Skor: {r['TOPLAM SKOR']:.1f}** (D:{r['Değerleme Skoru']:.0f} "
                  f"K:{r['Karlılık Skoru']:.0f} B:{r['Büyüme Skoru']:.0f} "
                  f"Bi:{r['Bilanço Skoru']:.0f})")
        md.append(f"- F/K: {r['F/K']:.1f} | PD/DD: {r['PD/DD']:.1f} | "
                  f"ROE: %{r['ROE %']:.1f} | FAVÖK Marjı: %{r['FAVÖK Marjı %']:.1f}")
        md.append(f"- Satış Büyüme: %{r['Satış Büyüme %']:.1f} | "
                  f"Borç/ÖzK: {r['Borç/Özkaynak']:.2f}")
        flags = kirmizi_bayraklar(r)
        if flags:
            md.append("- ⚠️ " + " | ".join(f.replace('&gt;', '>') for f in flags))
        md.append("")
    md.append("\n---\n*Yatırım tavsiyesi değildir.*")
    return '\n'.join(md)


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

        # 5) Tam markdown rapor dosyası
        md_text = markdown_rapor(skor, file_name)
        md_path = Path(tmp) / (Path(file_name).stem + '_analiz.md')
        md_path.write_text(md_text, encoding='utf-8')
        send_chat_action(chat_id, 'upload_document')
        send_document(chat_id, str(md_path),
                      caption="📎 Tam rapor (Markdown)")


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
