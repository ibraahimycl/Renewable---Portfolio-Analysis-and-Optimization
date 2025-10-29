## Renewable Portfolio Analysis and Optimization

### Overview
This repository contains a Streamlit application and supporting artifacts to compare two electricity generation plants of the same type (HES or RES) over a selected date range. The app authenticates to EPİAŞ Şeffaflık (Transparency) Platform, fetches market and generation data, builds per-plant analytics tables, and produces a formatted Excel workbook with detailed sheets and a comparison sheet.

The repository also includes pre-generated comparative PDF reports, example Excel outputs, and combined CSV datasets for the year 2024.

### Key Features
- EPİAŞ authentication to obtain TGT and subsequent data access
- Data retrieval for:
  - PTF (Gün Öncesi Piyasası Fiyatı)
  - SMF (Sistem Marjinal Fiyatı)
  - KGÜP (Gün Öncesi Üretim Tahmini)
  - Gerçekleşen Üretim
- Per-plant merged analytics table with derived metrics (e.g., dengesizlik, gelir ve maliyet kalemleri)
- Excel export with:
  - Two detailed sheets (one for each plant) with Turkish column headers
  - A comparison sheet aggregating monthly totals and additional monthly KPIs

### Repository Structure
- `kod/`
  - `streamlit/`
    - `streamlit_app.py`: Streamlit UI and Excel export workflow
    - `epias_client.py`: EPİAŞ API client and dataframe construction logic
    - `requirements.txt`: Python dependencies
  - `analiz/`: Jupyter notebooks used for data exploration (reference)
- `rapor/`
  - `MELKO HES vs YANBOLU HES RAPOR.pdf`
  - `EBER RES vs MASLAKTEPE RES RAPOR.pdf`
- `tablolar/`
  - Example Excel outputs, e.g. `Analiz_...xlsx`
- `veri/`
  - `MELKOM_HES_birlesik_2024.csv`
  - `YANBOLU_HES_birlesik_2024.csv`
  - `EBER_RES_birlesik_2024.csv`
  - `MASLAKTEPE_RES_birlesik_2024.csv`
  

### Data Sources (from code)
The application uses EPİAŞ Şeffaflık Platform endpoints via authenticated POST requests with the TGT token:
- TGT: `https://giris.epias.com.tr/cas/v1/tickets`
- PTF: `https://seffaflik.epias.com.tr/electricity-service/v1/markets/dam/data/mcp`
- SMF: `https://seffaflik.epias.com.tr/electricity-service/v1/markets/bpm/data/system-marginal-price`
- KGÜP: `https://seffaflik.epias.com.tr/electricity-service/v1/generation/data/dpp-first-version`
- Gerçekleşen Üretim: `https://seffaflik.epias.com.tr/electricity-service/v1/generation/data/realtime-generation`

### Running the App
1) Create a virtual environment and install dependencies
```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\\Scripts\\activate
pip install -r kod/streamlit/requirements.txt
```

2) Launch Streamlit
```bash
streamlit run kod/streamlit/streamlit_app.py
```

3) Authenticate in the sidebar
- Enter your EPİAŞ Şeffaflık Platform credentials to obtain a TGT token (valid ~2 hours).
- Alternatively, paste an existing TGT.

4) Select plants and dates, then run
- Choose two plants of the same type (HES–HES or RES–RES).
- Select start/end dates (defaults target 2024-01-01 to 2024-12-31).
- Click “Verileri Çek ve Excel İndir” to fetch data and download the Excel report.

### Excel Output Structure (from code)
The app writes a workbook to memory and serves it for download. The file name follows:
```
Analiz_{Plant1}_vs_{Plant2}_{YYYYMMDD}_{YYYYMMDD}.xlsx
```

Sheets and columns:
- `Santral_1` and `Santral_2` (detailed per-plant tables) with Turkish headers:
  - `Tarih, Ay, Saat, PTF, SMF, Pozitif Dengesizlik Fiyatı, Negatif Dengesizlik Fiyatı, Gün Öncesi Üretim Tahmini (KGÜP), Gerçekleşen Üretim, Dengesizlik Miktarı, GÖP Geliri (TL), Dengesizlik Tutarı (TL), Toplam (Net) Gelir (TL), Dengesizlik Maliyeti (TL), Birim Dengesizlik Maliyeti`
- `Karşılaştırma` (comparison) with monthly totals and KPIs:
  - Monthly totals: `Gerçekleşen Üretim (MWh), Dengesizlik Miktarı (MWh), GÖP Geliri (TL), Dengesizlik Tutarı (TL), Toplam Gelir (TL), Birim Gelir (TL/MWh), Dengesizlik Maliyeti (TL), Birim Deng Mal. (TL/MWh)`
  - Additional monthly KPIs computed in code: `Tahmin Doğruluğu (%), Maliyet Asimetrisi (Poz/Neg), Kapasite Faktörü (%), En Maliyetli 5 Gün (TL), Top 5 Gün DM Payı (%), Gelir Payı (%), Yıllık Pozitif Deng. Payı (%), Yıllık Negatif Deng. Payı (%), Üretim Saati (saat), Üretim Saat Payı (%), Üretim Payı (%)`

### Strategy and Optimization Approach
- Amaç: En riskli ve maliyetli zaman dilimlerini (ör. yüksek dengesizlik maliyeti oluşan saatler/günler) belirleyip, tahmini üretim (KGÜP) üzerinde parametre bazlı ayarlamalar yaparak beklenen net geliri iyileştirmek.
- Yaklaşım (analizden türetilir):
  - En maliyetli 5 gün ve aylık dengesizlik maliyeti dağılımlarını kullanarak “risk pencereleri” tanımlanır (saat ve ay bazında yoğunlaşan dönemler).
  - Bu pencereler için ayarlanabilir parametreler belirlenir (ör. `adjustment_factor` ve saat aralığı). Parametre, KGÜP üzerinde küçük ölçekli oynama senaryoları (yukarı/aşağı) üretmek için kullanılır.
  - Senaryo çıktıları, hesaplanan `Dengesizlik_Tutarı`, `Net_Gelir` ve `Birim_DM` metrikleri üzerinden kıyaslanır. Hedef, dengesizlik maliyetini düşürürken toplam geliri maksimize edecek parametre kombinasyonunu bulmaktır.
- Uygulama notu: Bu, karar-destek amaçlı bir senaryolama yaklaşımıdır. Operasyonel, piyasa ve mevzuat gereklilikleri mutlaka gözetilmelidir.

### Notes
- Plant list JSON: The app searches for `pp_list.json` in several candidate locations (including the repository root). Ensure `pp_list.json` is available at one of the searched paths before running the app.
- TGT cache: The app can optionally cache the TGT to a JSON file for ~2 hours. The current code uses an absolute `WORKSPACE_DIR` and writes `.tgt_cache.json` there. If you run the app from a different location, you may adjust `WORKSPACE_DIR` in `kod/streamlit/streamlit_app.py` or simply use the TGT input without caching.
- Comparisons require both plants to be of the same type (HES with HES, RES with RES). The app enforces this in the UI.

### Included Reports and Data
- `rapor/`: Two PDF comparison reports corresponding to plant pairs:
  - `MELKO HES vs YANBOLU HES RAPOR.pdf`
  - `EBER RES vs MASLAKTEPE RES RAPOR.pdf`
- `tablolar/`: Example Excel outputs created by the app.
- `veri/`: Combined 2024 CSVs for individual plants, usable for offline inspection and validation.

### Disclaimer
This project interacts with EPİAŞ Şeffaflık Platform endpoints. Access, authentication, and data use are subject to EPİAŞ terms, rate limits, and availability. Ensure you have appropriate credentials and permissions.


