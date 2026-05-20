# BIST Telegram Analiz Botu — Kurulum Rehberi

Bu rehberi takip ederek 5-10 dakika içinde botu çalışır hale getirebilirsin.

## Nasıl Çalışır?

1. Sen Telegram'da bota bir BIST sektör Excel dosyası (.xlsx) gönderirsin
2. Bot dosyayı indirir, analiz eder
3. Sana sırayla gönderir:
   - 📊 Sektör özeti
   - 📋 Tüm hisselerin puanlama tablosu (monospace, hizalı)
   - 🔍 İlk 10 hisse için detaylı analiz kartları (kırmızı bayraklar dahil)
   - ⚠️ Yatırım tavsiyesi olmadığı uyarısı
   - 📎 Tam markdown rapor (.md dosyası)

---

## Adım 1: Telegram Botunu Oluştur

1. Telegram'da [@BotFather](https://t.me/BotFather)'a git
2. `/newbot` yaz
3. Bot için bir ad ver (örn. "BIST Analiz Bot")
4. Bir kullanıcı adı ver (örn. `bist_analiz_bot` — `bot` ile bitmeli)
5. BotFather sana bir **token** verecek, şuna benzer:
   ```
   123456789:AAH...xyz
   ```
   **Bu token'ı sakla**, az sonra kullanacağız.

## Adım 2: Kendi Chat ID'ini Öğren (Opsiyonel ama önerilir)

Botunu sadece kendinle sınırlamak istersen:

1. Telegram'da [@userinfobot](https://t.me/userinfobot)'a git, `/start` yaz
2. Sana bir **ID** (örn. `123456789`) verecek — bu senin chat ID'in

## Adım 3: GitHub Repo Hazırla

```bash
# Bu klasördeki dosyaları kendi repo'na koy:
telegram_bist_bot.py
requirements.txt
.github/workflows/telegram-bot.yml
KURULUM.md  (bu dosya)
```

Kendi repo'nda commit & push et.

## Adım 4: GitHub Secrets Ekle

GitHub repo'nda:
1. **Settings** → **Secrets and variables** → **Actions** → **New repository secret**
2. Şu iki secret'ı ekle:

| Secret Adı | Değer |
|---|---|
| `TELEGRAM_BOT_TOKEN` | BotFather'dan aldığın token (Adım 1) |
| `AUTHORIZED_CHAT_IDS` | (Opsiyonel) Senin chat ID'in (Adım 2). Birden fazlaysa virgülle ayır: `123,456,789`. Boş bırakırsan herkes kullanabilir. |

## Adım 5: Workflow'u Aktive Et

GitHub repo'nda **Actions** sekmesine git, workflow'ları etkinleştir.

`.github/workflows/telegram-bot.yml` her **5 dakikada bir** otomatik çalışacak.

> ⚠️ **Önemli not:** GitHub Actions free tier'da public repolar için sınırsız, private repolar için ayda 2000 dakika ücretsiz. 5 dk'da bir × 30 sn ≈ ayda ~75 dk kullanım → free tier'da rahatça sığar.

## Adım 6: Bot'la Konuş

1. Telegram'da botunun kullanıcı adını ara (örn. `@bist_analiz_bot`)
2. `/start` yaz
3. Bir BIST sektör Excel'i gönder (`.xlsx`)
4. **En geç 5 dakika içinde** rapor gelir (GitHub Actions cron 5 dk olduğu için)

---

## Alternatif: Yerel veya VPS'te Çalıştırma (Daha hızlı)

GitHub Actions'ın 5 dakikalık gecikmesi can sıkıcı geliyorsa, botu sürekli çalışan bir ortamda barındırabilirsin — o zaman cevap **anında** gelir.

### Yerel (kendi bilgisayarında):

```bash
pip install -r requirements.txt
export TELEGRAM_BOT_TOKEN="123456789:AAH..."
export AUTHORIZED_CHAT_IDS="123456789"   # opsiyonel
python telegram_bist_bot.py
```

### VPS (Ubuntu örneği) — systemd servisi olarak:

```bash
# 1. Sunucuya bağlan, repo'yu klonla
git clone https://github.com/<KULLANICI>/<REPO>.git
cd <REPO>
pip install -r requirements.txt

# 2. /etc/systemd/system/bist-bot.service oluştur:
sudo tee /etc/systemd/system/bist-bot.service > /dev/null <<'EOF'
[Unit]
Description=BIST Telegram Bot
After=network.target

[Service]
Type=simple
User=ubuntu
WorkingDirectory=/home/ubuntu/<REPO>
Environment="TELEGRAM_BOT_TOKEN=123456789:AAH..."
Environment="AUTHORIZED_CHAT_IDS=123456789"
ExecStart=/usr/bin/python3 telegram_bist_bot.py
Restart=always
RestartSec=10

[Install]
WantedBy=multi-user.target
EOF

# 3. Etkinleştir
sudo systemctl daemon-reload
sudo systemctl enable --now bist-bot
sudo systemctl status bist-bot
```

### Ücretsiz Hosting Seçenekleri (sürekli çalışan):
- **Railway.app** — Free trial, sonra ~5 USD/ay
- **Render.com** — Background worker, free tier mevcut
- **Fly.io** — Free tier mevcut
- **Oracle Cloud Free Tier** — Ücretsiz küçük VM (Always Free)

---

## Sorun Giderme

**Bot cevap vermiyor?**
- GitHub Actions logs'a bak (repo → Actions sekmesi → son çalışmaya tıkla)
- Secret'lar doğru mu kontrol et
- Bot ile en az bir kez `/start` ile konuşmuş olmalısın

**"⛔ Bu botu kullanma yetkiniz yok" hatası alıyorum?**
- Bota `/id` yazarak chat ID'ini öğren
- O ID'i `AUTHORIZED_CHAT_IDS` secret'ına ekle (varsa virgülle ayır)
- Ya da secret'ı tamamen sil → herkese açık olur

**"Beklenen sütunlar eksik" hatası?**
- Bot, BIST sektör Excel'inin standart 141 sütunlu formatını bekliyor
- Hisse Adı, Firma Adı, F/K Günlük, PD/DD Günlük, FD/FAVOK Günlük gibi sütunlar olmalı
- Dosyan farklıysa `telegram_bist_bot.py` içindeki `zorunlu` listesini güncelle

**Mesajlar düzgün görünmüyor (HTML tag'leri görünüyor)?**
- Bu Telegram'ın HTML render hatası. Botun gönderdiği mesajlarda `<` veya `>` karakteri var olabilir
- Beklenmedik bir karakter varsa `telegram_bist_bot.py` içindeki `he()` fonksiyonu işliyor olmalı

---

## Ayarlanabilir Parametreler

`telegram_bist_bot.py` dosyasının üst kısmında:

```python
AGIRLIKLAR = {'degerleme': 0.30, 'karlilik': 0.30, 'buyume': 0.25, 'bilanco': 0.15}
TOP_N_DETAY = 10   # Kaç hisse için detaylı kart gönderilsin
```

Profiline göre ağırlıkları değiştir:
- **Büyüme odaklıyım** → `'buyume': 0.40, 'degerleme': 0.20`
- **Temettü/güvenlik odaklıyım** → `'bilanco': 0.30, 'karlilik': 0.35`
- **Değer yatırımcısıyım** → `'degerleme': 0.45`

---

## ⚠️ Önemli Uyarı

Bu bot **yatırım tavsiyesi vermez**. Sadece kamuya açık finansal verilerin sistematik bir kıyaslamasını yapar. Geçmiş performans gelecek getiriyi garanti etmez. Yatırım kararları kişisel risk profilin, vaden ve portföy çeşitlendirmen gözetilerek; gerekirse lisanslı bir yatırım danışmanıyla birlikte alınmalıdır.
