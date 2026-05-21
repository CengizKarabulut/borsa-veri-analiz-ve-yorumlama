"""
PDF rapor üretici modülü — v2 (6 boyutlu skorlama)

Boyutlar:
  1. Değerleme       (%25)
  2. Karlılık        (%25)
  3. Büyüme (REEL)   (%20)
  4. Bilanço Sağlığı (%15)
  5. Op. Verimlilik  (%10)
  6. Piyasa Sinyali  (%5)
"""

import os
from io import BytesIO

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    PageBreak, KeepTogether,
)


# ---- Font kayıt -----------------------------------------------------
_FONT_CANDIDATES = [
    ('/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf', 'DejaVu'),
    ('/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf', 'DejaVu-Bold'),
    ('/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf', 'DejaVu-Mono'),
    ('/usr/share/fonts/truetype/dejavu/DejaVuSansMono-Bold.ttf', 'DejaVu-Mono-Bold'),
]
_FONT_OK = False
for path, name in _FONT_CANDIDATES:
    if os.path.exists(path):
        try:
            pdfmetrics.registerFont(TTFont(name, path))
            _FONT_OK = True
        except Exception:
            pass

F_REG = 'DejaVu' if _FONT_OK else 'Helvetica'
F_BOLD = 'DejaVu-Bold' if _FONT_OK else 'Helvetica-Bold'
F_MONO = 'DejaVu-Mono' if _FONT_OK else 'Courier'

# ---- Renkler --------------------------------------------------------
BLUE = colors.HexColor('#1F4E78')
LIGHT_BLUE = colors.HexColor('#D9E2F3')
DARK_GREY = colors.HexColor('#595959')
LIGHT_GREY = colors.HexColor('#F2F2F2')
RED_FLAG = colors.HexColor('#C00000')
GREEN = colors.HexColor('#548235')


def _score_color(score):
    if pd.isna(score):
        return colors.white
    if score >= 75: return colors.HexColor('#63BE7B')
    if score >= 60: return colors.HexColor('#A8D08D')
    if score >= 45: return colors.HexColor('#FFEB84')
    if score >= 30: return colors.HexColor('#F8B084')
    return colors.HexColor('#F8696B')


def _fmt(v, decimals=1, suffix=''):
    if pd.isna(v):
        return '—'
    return f"{v:,.{decimals}f}{suffix}"


SEKTOR_BAGLAMI = {
    'Bilişim': "Bilişim/yazılım sektöründe yüksek FAVÖK marjı (%30-70) normaldir. "
               "Büyüme kritiktir, borç tipik olarak düşüktür.",
    'Gıda': "Gıda sektörü defansif ve düşük marjlıdır (FAVÖK %5-20). Hammadde "
            "ve kur dalgalanmalarına duyarlıdır.",
    'Kimya': "Kimya-İlaç-Petrol sektörü çok döngüseldir. Emtia fiyat hareketleri "
             "(Brent, nafta, etilen, doğal gaz, kauçuk) belirleyicidir.",
    'Banka': "Bankacılıkta FAVÖK marjı anlamsızdır. F/K, PD/DD, ROE öne çıkar.",
    'Holding': "Holdinglerde iştirak değerleri (NAV) bakılır; konsolide marjlar "
               "yanıltıcı olabilir.",
}


def _sektor_baglam(sektor):
    if not sektor:
        return ""
    for k, v in SEKTOR_BAGLAMI.items():
        if k.lower() in str(sektor).lower():
            return v
    return ""


def _styles():
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name='Baslik', fontName=F_BOLD, fontSize=22,
                              leading=26, textColor=BLUE, spaceAfter=10, alignment=1))
    styles.add(ParagraphStyle(name='AltBaslik', fontName=F_REG, fontSize=14,
                              leading=18, textColor=DARK_GREY, spaceAfter=6, alignment=1))
    styles.add(ParagraphStyle(name='SectionTitle', fontName=F_BOLD, fontSize=14,
                              leading=18, textColor=BLUE, spaceBefore=12, spaceAfter=8))
    styles.add(ParagraphStyle(name='CardTitle', fontName=F_BOLD, fontSize=12,
                              leading=15, textColor=BLUE, spaceBefore=6, spaceAfter=3))
    styles.add(ParagraphStyle(name='Normal2', fontName=F_REG, fontSize=10, leading=13))
    styles.add(ParagraphStyle(name='Small', fontName=F_REG, fontSize=8.5,
                              leading=11, textColor=DARK_GREY))
    styles.add(ParagraphStyle(name='Flag', fontName=F_REG, fontSize=8.5,
                              leading=11, textColor=RED_FLAG, leftIndent=6))
    styles.add(ParagraphStyle(name='Disclaimer', fontName=F_REG, fontSize=8.5,
                              leading=12, textColor=DARK_GREY, alignment=4))
    return styles


def _kirmizi_bayraklar(row):
    flags = []
    nm, em = row.get('Net Marj %'), row.get('FAVÖK Marjı %')
    if pd.notna(nm) and pd.notna(em) and nm > em + 1:
        flags.append(f"Net Marj (%{nm:.1f}) &gt; FAVÖK Marjı (%{em:.1f}) — kâr kalitesi sorgulanmalı.")
    if pd.notna(row.get('Çalışan')) and row['Çalışan'] == 0:
        flags.append("Çalışan sayısı 0 — veri eksik veya operasyonel yapı sıra dışı.")
    firma = str(row.get('Firma', '')).upper()
    if 'YATIRIM' in firma or 'HOLDİNG' in firma or 'GİRİŞİM' in firma:
        flags.append("Yatırım/Holding yapısı olabilir — NAV bazlı bakılmalı.")
    if pd.notna(row.get('F/K')) and row['F/K'] < 0:
        flags.append(f"F/K negatif ({row['F/K']:.1f}) — şirket zarar ediyor.")
    if pd.notna(row.get('Cari Oran')) and row['Cari Oran'] < 1:
        flags.append(f"Cari Oran {row['Cari Oran']:.2f} — kısa vadeli likidite riski.")
    rk = row.get('Reel Kâr %')
    if pd.notna(rk) and rk > 200:
        flags.append(f"Reel kâr büyümesi %{rk:,.0f} — baz etkisi, sürdürülebilir saymayın.")
    elif pd.notna(rk) and rk < -30:
        flags.append(f"Reel kâr daralması %{rk:.0f} — temel zayıflık.")
    yp = row.get('Net YPP')
    pv = row.get('Piyasa Değeri (mn TL)')
    if pd.notna(yp) and pd.notna(pv) and yp < 0 and abs(yp)/(pv*1e6) > 0.3:
        flags.append(f"Net döviz pozisyonu büyük negatif ({yp/1e9:.1f} mlr TL) — kur riski.")
    if pd.notna(row.get('Faiz Karşılama')) and 0 < row['Faiz Karşılama'] < 1.5:
        flags.append(f"Faiz Karşılama {row['Faiz Karşılama']:.1f} — faiz yükü kritik.")
    if pd.notna(row.get('Volatilite')) and row['Volatilite'] > 8:
        flags.append(f"Volatilite yüksek ({row['Volatilite']:.1f}).")
    if pd.notna(row.get('Fiili Dolaşım %')) and row['Fiili Dolaşım %'] < 25:
        flags.append(f"Fiili dolaşım düşük (%{row['Fiili Dolaşım %']:.1f}) — likidite sınırlı.")
    # YENİ: Altman Z-Score
    az = row.get('Altman Z')
    if pd.notna(az):
        if az < 1.81:
            flags.append(f"🚨 Altman Z-Score {az:.2f} — İFLAS RİSKİ BÖLGESİ.")
        elif az < 2.99:
            flags.append(f"Altman Z-Score {az:.2f} — gri bölge.")
    # YENİ: Piotroski F-Score düşük
    ps = row.get('Piotroski F-Score')
    if pd.notna(ps) and ps <= 3:
        flags.append(f"Piotroski F-Score {int(ps)}/9 — düşük kalite.")
    # YENİ: Faiz karşılama kritik
    if pd.notna(row.get('Faiz Karşılama')) and 0 < row['Faiz Karşılama'] < 1.5:
        flags.append(f"Faiz Karşılama {row['Faiz Karşılama']:.1f} — faiz yükü kritik.")
    return flags


def pdf_uret(skor_df, kaynak_dosya, top_n=10):
    buf = BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        topMargin=15*mm, bottomMargin=15*mm,
        leftMargin=15*mm, rightMargin=15*mm,
        title="BIST Sektör Analiz Raporu",
        author="BIST Analiz Botu",
    )
    styles = _styles()
    elems = []

    # ---- KAPAK ----
    n = len(skor_df)
    zarar = int((skor_df['F/K'] < 0).sum())
    sektor = skor_df.iloc[0]['Sektor'] if pd.notna(skor_df.iloc[0]['Sektor']) else 'BIST'
    baglam = _sektor_baglam(sektor)

    elems.append(Spacer(1, 25*mm))
    elems.append(Paragraph("BIST Sektör Analiz Raporu", styles['Baslik']))
    elems.append(Paragraph(str(sektor), styles['AltBaslik']))
    elems.append(Paragraph(
        "<i>6 Boyutlu Skorlama + Piotroski F-Score + Altman Z-Score</i>",
        styles['AltBaslik']))
    elems.append(Spacer(1, 15*mm))

    ozet_data = [
        ['Kaynak dosya:', os.path.basename(kaynak_dosya)],
        ['Analiz edilen hisse:', f"{n}"],
        ['Zarar eden hisse:', f"{zarar}  (toplamın %{zarar*100/n:.0f}'i)"],
        ['Skorlama ağırlıkları:', "Değerleme %25 · Karlılık %25 · Büyüme(REEL) %20"],
        ['', "Bilanço(+Altman) %15 · Op.Verimlilik %10 · Piyasa %5"],
        ['Büyüme hesabı:', "TÜFE düzeltmeli reel büyüme (CANSLIM yaklaşımı)"],
        ['Kalite metrikleri:', "Piotroski F-Score (0-9) · Altman Z-Score · Kâr Kalitesi"],
        ['Kaynaklar:', "Yaşar Erdinç + CANSLIM + Piotroski (1976) + Altman (1968)"],
    ]
    ozet_tbl = Table(ozet_data, colWidths=[50*mm, 115*mm])
    ozet_tbl.setStyle(TableStyle([
        ('FONTNAME', (0, 0), (0, -1), F_BOLD),
        ('FONTNAME', (1, 0), (1, -1), F_REG),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('TEXTCOLOR', (0, 0), (0, -1), DARK_GREY),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('LINEBELOW', (0, 0), (-1, -1), 0.3, LIGHT_GREY),
    ]))
    elems.append(ozet_tbl)

    if baglam:
        elems.append(Spacer(1, 8*mm))
        elems.append(Paragraph(f"<i>{baglam}</i>", styles['Normal2']))

    elems.append(PageBreak())

    # ---- PUANLAMA TABLOSU (6 alt skor) ----
    elems.append(Paragraph("Puanlama Tablosu", styles['SectionTitle']))
    elems.append(Paragraph(
        "Tüm hisseler toplam skora göre sıralı. 6 alt skor "
        "(Değ=Değerleme, Kâr=Karlılık, Büy=Büyüme[REEL], Bil=Bilanço, "
        "Vrm=Verimlilik, Piy=Piyasa).", styles['Small']))
    elems.append(Spacer(1, 3*mm))

    table_data = [['#', 'Hisse', 'Toplam', 'Değ', 'Kâr', 'Büy', 'Bil', 'Vrm', 'Piy',
                   'F/K', 'ROE%', 'FAVÖK%']]
    for i, r in skor_df.iterrows():
        table_data.append([
            str(i+1),
            str(r['Hisse']),
            _fmt(r['TOPLAM SKOR'], 1),
            _fmt(r['Değerleme Skoru'], 0),
            _fmt(r['Karlılık Skoru'], 0),
            _fmt(r['Büyüme Skoru'], 0),
            _fmt(r['Bilanço Skoru'], 0),
            _fmt(r['Verimlilik Skoru'], 0),
            _fmt(r['Piyasa Skoru'], 0),
            _fmt(r['F/K'], 1),
            _fmt(r['ROE %'], 1),
            _fmt(r['FAVÖK Marjı %'], 1),
        ])

    col_widths = [8*mm, 16*mm, 15*mm, 11*mm, 11*mm, 11*mm, 11*mm, 11*mm, 11*mm,
                  14*mm, 14*mm, 16*mm]
    tbl = Table(table_data, colWidths=col_widths, repeatRows=1)
    style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), BLUE),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), F_BOLD),
        ('FONTSIZE', (0, 0), (-1, 0), 8.5),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('FONTNAME', (0, 1), (-1, -1), F_REG),
        ('FONTSIZE', (0, 1), (-1, -1), 8),
        ('ALIGN', (0, 1), (0, -1), 'CENTER'),
        ('ALIGN', (1, 1), (1, -1), 'LEFT'),
        ('ALIGN', (2, 1), (-1, -1), 'RIGHT'),
        ('FONTNAME', (1, 1), (1, -1), F_BOLD),
        ('FONTNAME', (2, 1), (2, -1), F_BOLD),
        ('BOX', (0, 0), (-1, -1), 0.5, BLUE),
        ('LINEBELOW', (0, 0), (-1, 0), 1, BLUE),
        ('GRID', (0, 1), (-1, -1), 0.25, LIGHT_GREY),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
    ])
    for i in range(len(skor_df)):
        if (i+1) % 2 == 0:
            style.add('BACKGROUND', (0, i+1), (-1, i+1), LIGHT_GREY)
        sc = skor_df.iloc[i]['TOPLAM SKOR']
        style.add('BACKGROUND', (2, i+1), (2, i+1), _score_color(sc))
    tbl.setStyle(style)
    elems.append(tbl)

    elems.append(PageBreak())

    # ---- TOP N DETAY KARTLARI (6 kutu, 3x2 grid) ----
    elems.append(Paragraph(f"İlk {min(top_n, n)} Hissenin Detaylı Analizi",
                           styles['SectionTitle']))
    elems.append(Spacer(1, 3*mm))

    def _box(title_str, items):
        lines = [f"<b><font color='#1F4E78'>{title_str}</font></b>"]
        for label, val, comment in items:
            line = f"{label}: <b>{val}</b>"
            if comment:
                line += f" <font size=7 color='#595959'><i>({comment})</i></font>"
            lines.append(line)
        return Paragraph("<br/>".join(lines), styles['Small'])

    def _fk_y(v):
        if pd.isna(v): return ''
        if v < 0: return 'zarar'
        if v < 8: return 'çok ucuz'
        if v < 15: return 'makul'
        if v < 25: return 'ortalama'
        return 'pahalı'

    def _roe_y(v):
        if pd.isna(v): return ''
        if v < 0: return 'zarar'
        if v < 10: return 'düşük'
        if v < 20: return 'iyi'
        return 'çok güçlü'

    def _em_y(v):
        if pd.isna(v): return ''
        if v > 30: return 'çok yüksek'
        if v > 15: return 'iyi'
        if v > 5: return 'ortalama'
        return 'zayıf'

    def _de_y(v):
        if pd.isna(v): return ''
        if v < 0.3: return 'neredeyse borçsuz'
        if v < 0.7: return 'düşük'
        if v < 1.5: return 'orta'
        return 'yüksek risk'

    def _co_y(v):
        if pd.isna(v): return ''
        if v < 1: return 'likidite riski'
        if v < 1.5: return 'sıkı'
        return 'rahat'

    def _sb_y(v):
        if pd.isna(v): return ''
        if v < 0: return 'reel daralma'
        if v < 10: return 'zayıf'
        if v < 30: return 'sağlıklı'
        return 'güçlü'

    def _faiz_y(v):
        if pd.isna(v): return ''
        if 0 < v < 1.5: return 'kritik!'
        if 1.5 <= v < 3: return 'zayıf'
        if v > 10: return 'çok güçlü'
        if v < 0: return 'net faiz geliri'
        return ''

    for i, row in skor_df.head(top_n).iterrows():
        card_elems = []

        # Başlık satırı
        title = (f"#{i+1}  &middot;  <b>{row['Hisse']}</b>  &mdash; "
                 f"<i>{str(row['Firma'])[:55]}</i>")
        card_elems.append(Paragraph(title, styles['CardTitle']))

        # Skor satırı — 6 alt skor
        vrm = f"{row['Verimlilik Skoru']:.0f}" if pd.notna(row['Verimlilik Skoru']) else "—"
        score_text = (
            f"<b>Toplam: {row['TOPLAM SKOR']:.1f}/100</b> &nbsp; "
            f"<font size=8 color='#595959'>"
            f"Değ {row['Değerleme Skoru']:.0f} · "
            f"Kâr {row['Karlılık Skoru']:.0f} · "
            f"Büy {row['Büyüme Skoru']:.0f} · "
            f"Bil {row['Bilanço Skoru']:.0f} · "
            f"Vrm {vrm} · "
            f"Piy {row['Piyasa Skoru']:.0f}"
            f"</font>"
        )
        card_elems.append(Paragraph(score_text, styles['Normal2']))

        pv = row['Piyasa Değeri (mn TL)']
        pv_str = f"{pv/1000:.1f} mlr TL" if pv >= 1000 else f"{pv:.0f} mn TL"
        cal = int(row['Çalışan']) if pd.notna(row['Çalışan']) else '—'
        meta = (f"<font size=8 color='#595959'>PD: {pv_str} &nbsp;·&nbsp; "
                f"Çalışan: {cal}</font>")
        card_elems.append(Paragraph(meta, styles['Normal2']))

        # ⭐ Kalite Skorları satırı (Piotroski + Altman + Kâr Kalitesi)
        ps = row.get('Piotroski F-Score')
        az = row.get('Altman Z')
        kk = row.get('Kâr Kalitesi')
        if pd.notna(ps) or pd.notna(az) or pd.notna(kk):
            quality_parts = []
            if pd.notna(ps):
                ps_int = int(ps)
                ps_clr = '#548235' if ps_int >= 7 else ('#C65911' if ps_int >= 4 else '#C00000')
                quality_parts.append(
                    f"<font color='{ps_clr}'><b>Piotroski {ps_int}/9</b></font>")
            if pd.notna(az):
                if az >= 2.99: az_clr, az_lbl = '#548235', 'güvenli'
                elif az >= 1.81: az_clr, az_lbl = '#C65911', 'gri'
                else: az_clr, az_lbl = '#C00000', 'iflas riski'
                quality_parts.append(
                    f"<font color='{az_clr}'><b>Altman Z {az:.2f}</b> ({az_lbl})</font>")
            if pd.notna(kk):
                quality_parts.append(f"<b>Faal.K/Net K: {kk:.2f}</b>")
            quality_line = " &nbsp;·&nbsp; ".join(quality_parts)
            card_elems.append(Paragraph(
                f"<font size=9>⭐ {quality_line}</font>", styles['Normal2']))

        card_elems.append(Spacer(1, 2*mm))

        # 6 kutu — 3x2 grid (zenginleştirilmiş içerik + alt skor başlıkta)
        s_deg = row['Değerleme Skoru']
        s_kar = row['Karlılık Skoru']
        s_buy = row['Büyüme Skoru']
        s_bil = row['Bilanço Skoru']
        s_vrm = row['Verimlilik Skoru']
        s_piy = row['Piyasa Skoru']

        def _h(name, score):
            ss = f"{score:.0f}" if pd.notna(score) else "—"
            return f"{name} — {ss}/100"

        box_deg = _box(_h("Değerleme", s_deg), [
            ('F/K', _fmt(row['F/K'], 1), _fk_y(row['F/K'])),
            ('PD/DD', _fmt(row['PD/DD'], 1), ''),
            ('FD/FAVÖK', _fmt(row['FD/FAVÖK'], 1), ''),
            ('PEG', _fmt(row.get('PEG'), 2), ''),
        ])
        box_kar = _box(_h("Karlılık", s_kar), [
            ('ROE', f"%{_fmt(row['ROE %'], 1)}", _roe_y(row['ROE %'])),
            ('ROIC', f"%{_fmt(row['ROIC %'], 1)}", ''),
            ('ROA', f"%{_fmt(row.get('ROA %'), 1)}", ''),
            ('Net Marj', f"%{_fmt(row['Net Marj %'], 1)}", ''),
            ('FAVÖK', f"%{_fmt(row['FAVÖK Marjı %'], 1)}",
             _em_y(row['FAVÖK Marjı %'])),
            ('Faaliyet', f"%{_fmt(row.get('Faaliyet Marj %'), 1)}", ''),
            ('Brüt', f"%{_fmt(row.get('Brüt Marj %'), 1)}", ''),
        ])
        box_buy = _box(_h("Büyüme (REEL)", s_buy), [
            ('Satış', f"%{_fmt(row.get('Reel Satış %'), 1)}",
             _sb_y(row.get('Reel Satış %'))),
            ('Kâr', f"%{_fmt(row.get('Reel Kâr %'), 1)}", ''),
            ('Faal.Kâr', f"%{_fmt(row.get('Reel Faal.Kâr %'), 1)}", ''),
            ('Çey.Satış', f"%{_fmt(row.get('Çeyreklik Satış %'), 1)}", ''),
            ('Marj Tr.', f"{_fmt(row.get('Marj Trendi'), 1)} pp", ''),
        ])
        box_bil = _box(_h("Bilanço", s_bil), [
            ('Borç/ÖzK', _fmt(row['Borç/Özkaynak'], 2), _de_y(row['Borç/Özkaynak'])),
            ('Cari', _fmt(row['Cari Oran'], 2), _co_y(row['Cari Oran'])),
            ('Likidite', _fmt(row.get('Likidite Oranı'), 2), ''),
            ('Faiz Karş.', _fmt(row.get('Faiz Karşılama'), 1),
             _faiz_y(row.get('Faiz Karşılama'))),
            ('Fin.Borç', _fmt(row.get('Finansal Borç Oranı'), 2), ''),
        ])
        box_vrm = _box(_h("Op. Verimlilik", s_vrm), [
            ('Aktif Devir', _fmt(row.get('Aktif Devir'), 2) + 'x', ''),
            ('Stok Devir', _fmt(row.get('Stok Devir'), 1) + 'x', ''),
            ('DSO', _fmt(row.get('DSO'), 0) + ' gün', ''),
            ('DPO', _fmt(row.get('DPO'), 0) + ' gün', ''),
            ('Nakit Çev.', _fmt(row.get('Nakit Çevirme'), 0) + ' gün', ''),
        ])
        box_piy = _box(_h("Piyasa Sinyali", s_piy), [
            ('Yabancı', f"%{_fmt(row.get('Yabancı Oran %'), 1)}", ''),
            ('Bil.Sonr.', f"%{_fmt(row.get('Bilanço Sonrası %'), 1)}", ''),
            ('Volatilite', _fmt(row.get('Volatilite'), 1), ''),
            ('Dev.Kesici', _fmt(row.get('Devre Kesici'), 0), ''),
        ])

        # 3 sütunlu, 2 satırlı grid
        boxes = Table(
            [[box_deg, box_kar, box_buy],
             [box_bil, box_vrm, box_piy]],
            colWidths=[60*mm, 60*mm, 60*mm]
        )
        boxes.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('BOX', (0, 0), (-1, -1), 0.3, LIGHT_GREY),
            ('INNERGRID', (0, 0), (-1, -1), 0.3, LIGHT_GREY),
            ('LEFTPADDING', (0, 0), (-1, -1), 4),
            ('RIGHTPADDING', (0, 0), (-1, -1), 4),
            ('TOPPADDING', (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ]))
        card_elems.append(boxes)

        # Kırmızı bayraklar
        flags = _kirmizi_bayraklar(row)
        if flags:
            card_elems.append(Spacer(1, 2*mm))
            card_elems.append(Paragraph(
                "<b><font color='#C00000'>⚠ Kırmızı Bayraklar</font></b>",
                styles['Small']))
            for f in flags:
                card_elems.append(Paragraph(f"• {f}", styles['Flag']))

        card_elems.append(Spacer(1, 4*mm))
        elems.append(KeepTogether(card_elems))

    # ---- METODOLOJİ + METRİK SÖZLÜĞÜ + UYARI ----
    elems.append(PageBreak())
    elems.append(Paragraph("Skorlama Metodolojisi ve Metrik Sözlüğü",
                           styles['SectionTitle']))

    elems.append(Paragraph(
        "<b>Toplam skor</b>, 6 alt boyutun ağırlıklı ortalamasıdır. Her metrik sektör "
        "içi %5-95 persentilde normalize edilir (uç değerler bastırılır) ve 0-100 "
        "aralığına ölçeklenir.",
        styles['Normal2']))
    elems.append(Spacer(1, 4*mm))

    # Her boyutun detaylı tablosu
    methodology_rows = [
        ['Boyut', 'Ağırlık', 'İçerdiği Metrikler'],
        ['Değerleme', '%25',
         'F/K (düşük iyi), PD/DD (düşük iyi), FD/FAVÖK (düşük iyi), PEG (düşük iyi)'],
        ['Karlılık', '%25',
         'ROE, ROIC, ROA, Net/FAVÖK/Brüt/Faaliyet Marjı, Esas Faaliyet Kâr Kalitesi'],
        ['Büyüme (REEL)', '%20',
         'Reel Satış, Reel Kâr, Reel Faal.Kâr (hepsi TÜFE düzeltmeli), '
         'Çeyreklik Satış/Kâr momentumu, FAVÖK Marj Trendi'],
        ['Bilanço Sağlığı', '%15',
         'Borç/Özkaynak (ters), Cari Oran, Likidite Oranı, Nakit Oranı, '
         'Borç/FAVÖK (ters), Faiz Karşılama, Finansal Borç Oranı, Altman Z-Score'],
        ['Op. Verimlilik', '%10',
         'Aktif Devir Hızı, Stok Devir Hızı, DOI (ters), DSO (ters), DPO, '
         'Nakit Çevirme Süresi (ters)'],
        ['Piyasa Sinyali', '%5',
         'Yabancı Oran, Yabancı Etki, Bilanço Sonrası Getiri, '
         'Volatilite (ters), Devre Kesici (ters)'],
    ]
    meth_tbl = Table(methodology_rows, colWidths=[30*mm, 18*mm, 132*mm])
    meth_tbl.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), BLUE),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), F_BOLD),
        ('FONTSIZE', (0, 0), (-1, -1), 8.5),
        ('FONTNAME', (0, 1), (0, -1), F_BOLD),
        ('FONTNAME', (1, 1), (1, -1), F_BOLD),
        ('ALIGN', (1, 0), (1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('GRID', (0, 0), (-1, -1), 0.3, LIGHT_GREY),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('LEFTPADDING', (0, 0), (-1, -1), 5),
        ('RIGHTPADDING', (0, 0), (-1, -1), 5),
    ]))
    elems.append(meth_tbl)

    elems.append(Spacer(1, 6*mm))
    elems.append(Paragraph("Ek Kompozit Kalite Skorları", styles['SectionTitle']))
    elems.append(Paragraph(
        "<b>Piotroski F-Score (0-9)</b> — Joseph Piotroski'nin 1976-1996 verilerinde "
        "geliştirdiği 9 maddelik kalite testi. Türk Excel verisi için uyarlanmış "
        "versiyon. Her madde 1 puan: "
        "<i>(1) ROA &gt; 0 · (2) Net Kâr &gt; 0 · (3) Çeyreklik kâr büyümesi pozitif · "
        "(4) Faaliyet K./Net K. &gt; 0.7 · (5) Borç/Özkaynak &lt; 1 · "
        "(6) Cari Oran &gt; 1.5 · (7) Faiz Karşılama &gt; 2 (veya net borç negatif) · "
        "(8) Marj iyileşiyor · (9) Aktif Devir &gt; 0.5</i>. "
        "Yorum: <b>7-9 yüksek kalite</b>, 4-6 orta, <b>0-3 düşük kalite</b>.",
        styles['Normal2']))
    elems.append(Spacer(1, 3*mm))
    elems.append(Paragraph(
        "<b>Altman Z-Score</b> — Edward Altman'ın 1968'de geliştirdiği iflas tahmin "
        "modeli, 50 yıldır kullanılıyor. Formül: "
        "<i>Z = 1.2·(İşletme Sermayesi/Aktif) + 1.4·(Birikmiş Kâr/Aktif) + "
        "3.3·(EBIT/Aktif) + 0.6·(PD/Toplam Yük.) + 1.0·(Satış/Aktif)</i>. "
        "Yorum: <b>Z &gt; 2.99 güvenli</b>, 1.81-2.99 gri bölge, "
        "<b>Z &lt; 1.81 iflas risk bölgesi</b>.",
        styles['Normal2']))
    elems.append(Spacer(1, 3*mm))
    elems.append(Paragraph(
        "<b>Esas Faaliyet Kâr Kalitesi</b> = Faaliyet Kârı / Net Kâr. "
        "1'e yakınsa net kâr tamamen esas faaliyetten geliyor (yüksek kalite). "
        "0.5'in altındaysa net karın çoğu faaliyet dışı kalemlerden (kur farkı, "
        "finansal yatırım kazancı, tek seferlik gelir) — Yaşar Erdinç'in özellikle "
        "uyardığı tuzak.",
        styles['Normal2']))

    elems.append(Spacer(1, 4*mm))
    elems.append(Paragraph("Reel Büyüme Hesabı", styles['SectionTitle']))
    elems.append(Paragraph(
        "Tüm büyüme metrikleri <b>TÜFE düzeltmeli reel bazda</b> hesaplanır: "
        "<i>reel = ((1 + nominal/100) / (1 + TÜFE/100) - 1) × 100</i>. "
        "Örnek: nominal %150 büyüme, TÜFE %32 → reel %89. "
        "Bu, CANSLIM (William O'Neil) metodolojisinin C ve A unsurlarının "
        "uygulanmasını sağlar — enflasyon arındırılmış kâr büyümesi.",
        styles['Normal2']))

    elems.append(Spacer(1, 4*mm))
    elems.append(Paragraph("Otomatik Tetiklenen Kırmızı Bayraklar",
                           styles['SectionTitle']))
    elems.append(Paragraph(
        "Skorlama ile bağımsız olarak çalışan otomatik uyarılar: "
        "Net Marj &gt; FAVÖK Marjı (faaliyet dışı kâr), F/K negatif (zarar), "
        "Cari Oran &lt; 1 (likidite riski), Faiz Karşılama &lt; 1.5 (kritik), "
        "Yatırım/Holding ismi (NAV gereği), baz etkili büyüme (&gt;%300), "
        "büyük negatif Net YPP (kur riski), yüksek volatilite, düşük fiili dolaşım, "
        "Altman Z &lt; 1.81 (iflas riski), Piotroski ≤ 3 (düşük kalite), "
        "Esas faaliyet kâr kalitesi &lt; 0.5.",
        styles['Normal2']))

    elems.append(Spacer(1, 8*mm))
    elems.append(Paragraph("Önemli Uyarı", styles['SectionTitle']))
    elems.append(Paragraph(
        "Bu rapor <b>yatırım tavsiyesi değildir</b>. Analiz, tek bir bilanço "
        "dönemine dayalı kamuya açık verilerin sistematik bir kıyaslamasıdır. "
        "Geçmiş performans gelecekteki getiriyi garanti etmez. "
        "Yatırım kararları kişisel risk profili, vade ve portföy çeşitlendirmesi "
        "gözetilerek; gerekirse lisanslı bir yatırım danışmanıyla birlikte "
        "alınmalıdır.",
        styles['Disclaimer']))

    def _footer(canvas, doc):
        canvas.saveState()
        canvas.setFont(F_REG, 8)
        canvas.setFillColor(DARK_GREY)
        canvas.drawString(15*mm, 8*mm, "BIST Sektör Analiz Raporu · v2 (6 boyutlu)")
        canvas.drawRightString(A4[0] - 15*mm, 8*mm, f"Sayfa {doc.page}")
        canvas.restoreState()

    doc.build(elems, onFirstPage=_footer, onLaterPages=_footer)
    buf.seek(0)
    return buf.getvalue()
