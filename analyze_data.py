import pandas as pd

def analyze_company(row):
    company_name = row['Firma Adı']
    analysis_text = f"## {company_name} Finansal Analiz Raporu\n\n"

    # Temel Oranlar
    fk = row.get('F/K Günlük')
    pddd = row.get('PD/DD Günlük')
    cari_oran = row.get('Cari Oranı')
    likidite_oran = row.get('Likidite Oranı')
    borc_ozkaynak = row.get('Borç/ÖzKaynak')
    ozsermaye_karliligi = row.get('Öz Kaynak Karlılığı (ROE)')
    net_kar_marji = row.get('Net Kar Marjı (Yıllık %)')
    favok_marji = row.get('Favök Marjı (Yıllık %)')
    brut_kar_marji = row.get('Brüt Kar Marjı (Yıllık %)')
    satis_gelirleri = row.get('Satış Gelirleri (Yıllıklandırılmış)')

    analysis_text += "### Temel Finansal Göstergeler\n\n"
    if pd.notna(fk):
        analysis_text += f"- **F/K Oranı (Günlük):** {fk:.2f}\n"
        if fk < 10:
            analysis_text += "  *Yorum: Şirketin hisse senedi, kazançlarına göre nispeten ucuz görünüyor.*\n"
        elif fk > 20:
            analysis_text += "  *Yorum: Şirketin hisse senedi, kazançlarına göre pahalı görünüyor veya yüksek büyüme beklentisi var.*\n"
        else:
            analysis_text += "  *Yorum: Şirketin F/K oranı sektör ortalamasına yakın.*\n"

    if pd.notna(pddd):
        analysis_text += f"- **PD/DD Oranı (Günlük):** {pddd:.2f}\n"
        if pddd < 1:
            analysis_text += "  *Yorum: Şirketin piyasa değeri defter değerinin altında, potansiyel olarak değerinin altında.*\n"
        elif pddd > 3:
            analysis_text += "  *Yorum: Şirketin piyasa değeri defter değerinin oldukça üzerinde, yüksek büyüme beklentisi veya varlıklarının yüksek değerlemesi olabilir.*\n"
        else:
            analysis_text += "  *Yorum: Şirketin PD/DD oranı makul seviyelerde.*\n"

    if pd.notna(cari_oran):
        analysis_text += f"- **Cari Oran:** {cari_oran:.2f}\n"
        if cari_oran < 1.5:
            analysis_text += "  *Yorum: Şirketin kısa vadeli borçlarını ödeme kapasitesi düşük olabilir.*\n"
        elif cari_oran > 2.5:
            analysis_text += "  *Yorum: Şirketin kısa vadeli borçlarını ödeme kapasitesi güçlü.*\n"
        else:
            analysis_text += "  *Yorum: Şirketin cari oranı sağlıklı seviyelerde.*\n"

    if pd.notna(likidite_oran):
        analysis_text += f"- **Likidite Oranı:** {likidite_oran:.2f}\n"
        if likidite_oran < 1:
            analysis_text += "  *Yorum: Şirketin acil kısa vadeli borçlarını ödeme kapasitesi zayıf olabilir.*\n"
        elif likidite_oran > 1.5:
            analysis_text += "  *Yorum: Şirketin acil kısa vadeli borçlarını ödeme kapasitesi güçlü.*\n"
        else:
            analysis_text += "  *Yorum: Şirketin likidite oranı yeterli seviyede.*\n"

    if pd.notna(borc_ozkaynak):
        analysis_text += f"- **Borç/Özkaynak Oranı:** {borc_ozkaynak:.2f}\n"
        if borc_ozkaynak > 1:
            analysis_text += "  *Yorum: Şirketin finansal yapısı borca dayalı, riskli olabilir.*\n"
        elif borc_ozkaynak < 0.5:
            analysis_text += "  *Yorum: Şirketin finansal yapısı güçlü, borçluluk oranı düşük.*\n"
        else:
            analysis_text += "  *Yorum: Şirketin borç/özkaynak oranı dengeli.*\n"

    analysis_text += "\n### Karlılık Göstergeleri\n\n"
    if pd.notna(ozsermaye_karliligi):
        analysis_text += f"- **Özkaynak Karlılığı (ROE):** %{ozsermaye_karliligi:.2f}\n"
        if ozsermaye_karliligi > 15:
            analysis_text += "  *Yorum: Şirket, özkaynaklarını verimli kullanarak yüksek kar elde ediyor.*\n"
        elif ozsermaye_karliligi < 5:
            analysis_text += "  *Yorum: Şirketin özkaynak karlılığı düşük, verimlilik sorunları olabilir.*\n"
        else:
            analysis_text += "  *Yorum: Şirketin özkaynak karlılığı sektör ortalamasına yakın.*\n"

    if pd.notna(net_kar_marji):
        analysis_text += f"- **Net Kar Marjı (Yıllık):** %{net_kar_marji:.2f}\n"
        if net_kar_marji > 10:
            analysis_text += "  *Yorum: Şirketin satışlarından elde ettiği net kar oranı yüksek, operasyonel verimliliği iyi.*\n"
        elif net_kar_marji < 3:
            analysis_text += "  *Yorum: Şirketin net kar marjı düşük, maliyet yönetimi veya fiyatlandırma sorunları olabilir.*\n"
        else:
            analysis_text += "  *Yorum: Şirketin net kar marjı istikrarlı seviyelerde.*\n"

    if pd.notna(favok_marji):
        analysis_text += f"- **FAVÖK Marjı (Yıllık):** %{favok_marji:.2f}\n"
        if favok_marji > 20:
            analysis_text += "  *Yorum: Şirketin ana faaliyetlerinden elde ettiği kar marjı oldukça yüksek.*\n"
        elif favok_marji < 5:
            analysis_text += "  *Yorum: Şirketin ana faaliyetlerinden elde ettiği kar marjı düşük, operasyonel sorunlar olabilir.*\n"
        else:
            analysis_text += "  *Yorum: Şirketin FAVÖK marjı sektör ortalamasına uygun.*\n"

    if pd.notna(brut_kar_marji):
        analysis_text += f"- **Brüt Kar Marjı (Yıllık):** %{brut_kar_marji:.2f}\n"
        if brut_kar_marji > 30:
            analysis_text += "  *Yorum: Şirketin ürün veya hizmetlerinin maliyeti düşük, karlılığı yüksek.*\n"
        elif brut_kar_marji < 10:
            analysis_text += "  *Yorum: Şirketin brüt kar marjı düşük, üretim maliyetleri yüksek olabilir.*\n"
        else:
            analysis_text += "  *Yorum: Şirketin brüt kar marjı tatmin edici seviyelerde.*\n"

    analysis_text += "\n### Büyüme Göstergeleri\n\n"
    if pd.notna(satis_gelirleri):
        analysis_text += f"- **Satış Gelirleri (Yıllıklandırılmış):** {satis_gelirleri:,.2f} TL\n"
        # Yorum için satış gelirlerinin geçmiş dönemlerle karşılaştırılması gerekir, burada sadece mevcut değer gösteriliyor.
        analysis_text += "  *Yorum: Satış gelirlerinin geçmiş dönemlerle karşılaştırılması, şirketin büyüme trendi hakkında daha net bilgi verecektir.*\n"

    return analysis_text

def main():
    try:
        df = pd.read_excel('/home/ubuntu/upload/Bilişim-2026.03.xlsx')
        
        all_reports = []
        for index, row in df.iterrows():
            report = analyze_company(row)
            all_reports.append(report)
        
        with open('analysis_report.md', 'w', encoding='utf-8') as f:
            f.write("# Borsa Veri Analiz Raporu\n\n")
            f.write("Bu rapor, sağlanan Excel dosyasındaki şirketlerin finansal verilerini analiz etmektedir.\n\n")
            for report in all_reports:
                f.write(report)
                f.write("\n---\n\n") # Şirketler arası ayırıcı
        
        print("Analiz raporu 'analysis_report.md' dosyasına başarıyla yazıldı.")

    except FileNotFoundError:
        print("Hata: 'Bilişim-2026.03.xlsx' dosyası bulunamadı. Lütfen dosyanın doğru konumda olduğundan emin olun.")
    except Exception as e:
        print(f"Bir hata oluştu: {e}")

if __name__ == "__main__":
    main()
