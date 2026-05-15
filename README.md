# Borsa Veri Analiz ve Yorumlama Projesi

Bu proje, sağlanan Excel dosyasındaki borsa verilerini otomatik olarak analiz eder ve finansal göstergeler hakkında yorumlar sunar.

## Özellikler

- Excel dosyasından veri okuma
- Temel finansal oranların hesaplanması (F/K, PD/DD, Cari Oran vb.)
- Şirketlerin finansal sağlığı ve performansına ilişkin otomatik yorumlar
- Analiz sonuçlarının okunabilir bir Markdown formatında sunulması

## Kullanım

1. `Bilişim-2026.03.xlsx` adlı Excel dosyasını projenin kök dizinine yerleştirin.
2. Python bağımlılıklarını yükleyin:
   ```bash
   pip install pandas openpyxl
   ```
3. Analiz betiğini çalıştırın:
   ```bash
   python analyze_data.py
   ```
4. Analiz sonuçları `analysis_report.md` dosyasında bulunacaktır.

## Geliştirme

Proje, finansal analiz yeteneklerini genişletmek ve daha derinlemesine yorumlar sunmak için geliştirilebilir.
