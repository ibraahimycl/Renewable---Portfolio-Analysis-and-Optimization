import os
import json
from datetime import date, datetime, timedelta
from io import BytesIO
from typing import List, Optional

import numpy as np
import pandas as pd
import streamlit as st

from epias_client import EpiasClient, PlantMeta, load_plants, slugify


st.set_page_config(page_title="Gain Enerji Analiz", layout="wide")

# Constants and defaults
WORKSPACE_DIR = "/Users/ibrahimyucel/Downloads/Gain Enerji Intern Analyst Case Study_2025"
PP_JSON_CANDIDATES = [
	f"{WORKSPACE_DIR}/GAIN ENERGY/pp_list.json",
	"GAIN ENERGY/pp_list.json",
	"pp_list.json",
]
DEFAULT_START = date(2024, 1, 1)
DEFAULT_END = date(2024, 12, 31)
TGT_CACHE_PATH = f"{WORKSPACE_DIR}/.tgt_cache.json"


@st.cache_data(show_spinner=False)
def load_pp_cache() -> List[PlantMeta]:
	return load_plants(PP_JSON_CANDIDATES)


def get_client() -> Optional[EpiasClient]:
	if "tgt" in st.session_state and st.session_state["tgt"]:
		return EpiasClient(st.session_state["tgt"])
	return None

def load_cached_tgt() -> Optional[str]:
	try:
		if not os.path.exists(TGT_CACHE_PATH):
			return None
		with open(TGT_CACHE_PATH, "r", encoding="utf-8") as f:
			obj = json.load(f)
		exp = obj.get("expires_at")
		if exp:
			exp_dt = datetime.fromisoformat(exp)
			if datetime.now() > exp_dt:
				return None
		return obj.get("tgt") or None
	except Exception:
		return None

def save_cached_tgt(tgt: str, username: str = "") -> None:
	try:
		exp = datetime.now() + timedelta(hours=2)
		obj = {"tgt": tgt, "username": username, "saved_at": datetime.now().isoformat(), "expires_at": exp.isoformat()}
		with open(TGT_CACHE_PATH, "w", encoding="utf-8") as f:
			json.dump(obj, f)
	except Exception:
		pass


def build_monthly_summary(df: pd.DataFrame) -> pd.DataFrame:
	# Aggregate by month using the per-plant detailed table
	agg = (
		df.groupby("Ay", dropna=False)
		.agg({
			"URETIM": "sum",
			"Dengesizlik": "sum",
			"GOP_Geliri": "sum",
			"Dengesizlik_Tutarı": "sum",
			"Net_Gelir": "sum",
			"Dengesizlik_Maliyeti": "sum",
		})
		.reindex(range(1, 12 + 1), fill_value=0)
		.reset_index()
	)
	# Unit metrics based on sums
	agg["Birim Gelir (TL/MWh)"] = agg.apply(lambda r: (r["Net_Gelir"] / r["URETIM"]) if r["URETIM"] not in (0, 0.0) else 0.0, axis=1)
	agg["Birim Deng Mal. (TL/MWh)"] = agg.apply(lambda r: (r["Dengesizlik_Maliyeti"] / r["URETIM"]) if r["URETIM"] not in (0, 0.0) else 0.0, axis=1)
	# Build output with exact column names
	out = pd.DataFrame({
		"Ay": agg["Ay"],
		"Gerçekleşen Üretim  (MWh)": agg["URETIM"],
		"Dengesizlik Miktarı  (MWh)": agg["Dengesizlik"],
		"GÖP Geliri (TL)": agg["GOP_Geliri"],
		"Dengesizlik Tutarı (TL)": agg["Dengesizlik_Tutarı"],
		"Toplam Gelir (TL)": agg["Net_Gelir"],
		"Birim Gelir (TL/MWh)": agg["Birim Gelir (TL/MWh)"],
		"Dengesizlik Maliyeti (TL)": agg["Dengesizlik_Maliyeti"],
		"Birim Deng Mal. (TL/MWh)": agg["Birim Deng Mal. (TL/MWh)"],
	})
	# Totals row
	total_row = {
		"Ay": "Toplam",
		"Gerçekleşen Üretim  (MWh)": out["Gerçekleşen Üretim  (MWh)"].sum(),
		"Dengesizlik Miktarı  (MWh)": out["Dengesizlik Miktarı  (MWh)"].sum(),
		"GÖP Geliri (TL)": out["GÖP Geliri (TL)"].sum(),
		"Dengesizlik Tutarı (TL)": out["Dengesizlik Tutarı (TL)"].sum(),
		"Toplam Gelir (TL)": out["Toplam Gelir (TL)"].sum(),
		"Birim Gelir (TL/MWh)": 0.0,
		"Dengesizlik Maliyeti (TL)": out["Dengesizlik Maliyeti (TL)"].sum(),
		"Birim Deng Mal. (TL/MWh)": 0.0,
	}
	# Compute unit values at total using totals
	if total_row["Gerçekleşen Üretim  (MWh)"] not in (0, 0.0):
		total_row["Birim Gelir (TL/MWh)"] = total_row["Toplam Gelir (TL)"] / total_row["Gerçekleşen Üretim  (MWh)"]
		total_row["Birim Deng Mal. (TL/MWh)"] = total_row["Dengesizlik Maliyeti (TL)"] / total_row["Gerçekleşen Üretim  (MWh)"]
	out = pd.concat([out, pd.DataFrame([total_row])], ignore_index=True)
	return out


def rename_to_turkish(df: pd.DataFrame) -> pd.DataFrame:
	mapping = {
		"Poz_DF": "Pozitif Dengesizlik Fiyatı",
		"Neg_DF": "Negatif Dengesizlik Fiyatı",
		"KGUP": "Gün Öncesi Üretim Tahmini (KGÜP)",
		"URETIM": "Gerçekleşen Üretim",
		"Dengesizlik": "Dengesizlik Miktarı",
		"GOP_Geliri": "GÖP Geliri",
		"Dengesizlik_Tutarı": "Dengesizlik Tutarı",
		"Net_Gelir": "Toplam (Net) Gelir",
		"Dengesizlik_Maliyeti": "Dengesizlik Maliyeti",
		"Birim_DM": "Birim Dengesizlik Maliyeti",
	}
	order = [
		"Tarih","Ay","Saat","PTF","SMF",
		"Pozitif Dengesizlik Fiyatı","Negatif Dengesizlik Fiyatı",
		"Gün Öncesi Üretim Tahmini (KGÜP)","Gerçekleşen Üretim","Dengesizlik Miktarı",
		"GÖP Geliri","Dengesizlik Tutarı","Toplam (Net) Gelir",
		"Dengesizlik Maliyeti","Birim Dengesizlik Maliyeti",
	]
	df2 = df.rename(columns=mapping)
	return df2.loc[:, order]


def _col_letter(idx_zero_based: int) -> str:
	letters = ""
	idx = idx_zero_based
	while True:
		idx, rem = divmod(idx, 26)
		letters = chr(65 + rem) + letters
		if idx == 0:
			break
		idx -= 1
	return letters


def _plant_type(p: PlantMeta) -> str:
	name = (p.powerPlantName or "").upper()
	if "HES" in name:
		return "HES"
	if "RES" in name:
		return "RES"
	return "OTHER"

# --- Yeni: Aylık ek KPI hesaplayıcı ---

def compute_monthly_extras(df: pd.DataFrame) -> dict:
	months = list(range(1, 13))
	days_2024 = [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
	# Hazırlık
	df2 = df.copy()
	# Tarih günlük toplamlarda kullanılacak
	if not pd.api.types.is_datetime64_any_dtype(df2.get("Tarih")):
		try:
			df2["Tarih"] = pd.to_datetime(df2["Tarih"]).dt.date
		except Exception:
			pass
	# Yıllık toplamlar
	total_revenue = float(df2["Net_Gelir"].sum()) if "Net_Gelir" in df2.columns else float(df2["Toplam (Net) Gelir"].sum()) if "Toplam (Net) Gelir" in df2.columns else 0.0
	total_production = float(df2["URETIM"].sum())
	total_prod_hours = int((df2["URETIM"] > 0).sum())
	total_pos_vol = float(df2.loc[df2["Dengesizlik"] > 0, "Dengesizlik"].sum())
	total_neg_vol = float(df2.loc[df2["Dengesizlik"] < 0, "Dengesizlik"].abs().sum())
	# Çıktı dizileri
	accuracy_pct = []
	asym_ratio = []
	capacity_factor_pct = []
	top5_dm_tl = []
	top5_dm_share_pct = []
	revenue_share_pct = []
	pos_share_pct = []
	neg_share_pct = []
	prod_hours = []
	prod_hours_share_pct = []
	prod_share_pct = []
	for m in months:
		mask_m = (df2["Ay"] == m)
		sum_kgup = float(df2.loc[mask_m, "KGUP"].sum())
		sum_abs_imb = float(df2.loc[mask_m, "Dengesizlik"].abs().sum())
		acc = (1.0 - (sum_abs_imb / sum_kgup)) * 100.0 if sum_kgup > 0 else 0.0
		accuracy_pct.append(acc)
		# Asimetri oranı (Poz/Neg DM)
		pos_dm = float(df2.loc[mask_m & (df2["Dengesizlik"] > 0), "Dengesizlik_Maliyeti"].sum())
		neg_dm = float(df2.loc[mask_m & (df2["Dengesizlik"] < 0), "Dengesizlik_Maliyeti"].sum())
		asym_ratio.append((pos_dm / neg_dm) if neg_dm != 0 else None)
		# Kapasite faktörü (aylık max KGUP * 24 * gün)
		max_kgup = float(df2.loc[mask_m, "KGUP"].max()) if mask_m.any() else 0.0
		potential = max_kgup * 24.0 * days_2024[m - 1]
		month_prod = float(df2.loc[mask_m, "URETIM"].sum())
		cap_fac = (month_prod / potential) * 100.0 if potential > 0 else 0.0
		capacity_factor_pct.append(cap_fac)
		# En maliyetli 5 gün ve payı
		monthly = df2.loc[mask_m]
		if not monthly.empty and "Tarih" in monthly.columns:
			daily_dm = monthly.groupby("Tarih")["Dengesizlik_Maliyeti"].sum().sort_values(ascending=False)
			top5 = float(daily_dm.head(5).sum())
			month_dm_total = float(monthly["Dengesizlik_Maliyeti"].sum())
			top5_dm_tl.append(top5)
			top5_dm_share_pct.append((top5 / month_dm_total) * 100.0 if month_dm_total > 0 else 0.0)
		else:
			top5_dm_tl.append(0.0)
			top5_dm_share_pct.append(0.0)
		# Gelir payı
		month_rev = float(df2.loc[mask_m, "Net_Gelir"].sum()) if "Net_Gelir" in df2.columns else float(df2.loc[mask_m, "Toplam (Net) Gelir"].sum()) if "Toplam (Net) Gelir" in df2.columns else 0.0
		revenue_share_pct.append((month_rev / total_revenue) * 100.0 if total_revenue > 0 else 0.0)
		# Poz/Neg dengesizlik payları (yıllık toplam içinde)
		pos_vol_m = float(df2.loc[mask_m & (df2["Dengesizlik"] > 0), "Dengesizlik"].sum())
		neg_vol_m = float(df2.loc[mask_m & (df2["Dengesizlik"] < 0), "Dengesizlik"].abs().sum())
		pos_share_pct.append((pos_vol_m / total_pos_vol) * 100.0 if total_pos_vol > 0 else 0.0)
		neg_share_pct.append((neg_vol_m / total_neg_vol) * 100.0 if total_neg_vol > 0 else 0.0)
		# Üretim saatleri ve payı
		prod_h_m = int((df2.loc[mask_m, "URETIM"] > 0).sum())
		prod_hours.append(prod_h_m)
		prod_hours_share_pct.append((prod_h_m / total_prod_hours) * 100.0 if total_prod_hours > 0 else 0.0)
		# Üretim payı
		prod_share_pct.append((month_prod / total_production) * 100.0 if total_production > 0 else 0.0)
	return {
		"accuracy_pct": accuracy_pct,
		"asym_ratio": asym_ratio,
		"capacity_factor_pct": capacity_factor_pct,
		"top5_dm_tl": top5_dm_tl,
		"top5_dm_share_pct": top5_dm_share_pct,
		"revenue_share_pct": revenue_share_pct,
		"pos_share_pct": pos_share_pct,
		"neg_share_pct": neg_share_pct,
		"prod_hours": prod_hours,
		"prod_hours_share_pct": prod_hours_share_pct,
		"prod_share_pct": prod_share_pct,
	}


with st.sidebar:
	st.header("Giriş")
	st.write("EPİAŞ Şeffaflık Platformu hesabınızla giriş yapın ve TGT alın.")
	with st.form("login_form", clear_on_submit=False):
		username = st.text_input("Kullanıcı Adı", value="", autocomplete="username")
		password = st.text_input("Şifre", value="", type="password", autocomplete="current-password")
		col_a, col_b = st.columns(2)
		login_clicked = col_a.form_submit_button("Giriş Yap ve TGT Al")
		clear_clicked = col_b.form_submit_button("Çıkış")
	if login_clicked:
		try:
			tgt = EpiasClient.obtain_tgt(username, password)
			st.session_state["tgt"] = tgt
			st.success("TGT alındı.")
			if st.sidebar.toggle("TGT'yi yerel dosyaya kaydet (2 saat)", value=True, key="save_tgt_toggle_login"):
				save_cached_tgt(tgt, username)
		except Exception as e:
			st.error(f"Giriş/TGT hatası: {e}")
	if clear_clicked:
		st.session_state.pop("tgt", None)
		st.info("Oturum temizlendi.")
	st.divider()
	manual_tgt = st.text_input("TGT (opsiyonel, yapıştır)", value=st.session_state.get("tgt", ""))
	if manual_tgt and manual_tgt != st.session_state.get("tgt"):
		st.session_state["tgt"] = manual_tgt
		st.success("TGT güncellendi.")
		if st.sidebar.toggle("TGT'yi kaydet (2 saat)", value=True, key="save_tgt_toggle_manual"):
			save_cached_tgt(manual_tgt)

st.title("Gain Enerji - Karşılaştırmalı Analiz")

# Try auto-load cached TGT if not set
if not st.session_state.get("tgt"):
	cached = load_cached_tgt()
	if cached:
		st.session_state["tgt"] = cached
		st.toast("Kayıtlı TGT yüklendi.")

plants: List[PlantMeta] = []
try:
	plants = load_pp_cache()
except Exception as e:
	st.error(f"Santral listesi okunamadı: {e}")

if not plants:
	st.stop()

client = get_client()
if not client:
	st.warning("Lütfen sol taraftan giriş yaparak TGT oluşturun veya yapıştırın.")
	st.stop()

# UI Controls
col1, col2, col3 = st.columns([2, 2, 3])
with col1:
	pp_names = [p.powerPlantName for p in plants]
	pp1_name = st.selectbox("Santral 1", options=pp_names, index=0, key="pp1")
	pp1_obj = next(p for p in plants if p.powerPlantName == pp1_name)
	pp1_type = _plant_type(pp1_obj)
	pp2_candidates = [p.powerPlantName for p in plants if _plant_type(p) == pp1_type and p.powerPlantName != pp1_name]
	if not pp2_candidates:
		st.error("Seçilen santral tipinde (HES/RES) başka santral yok.")
		pp2_candidates = [pp1_name]
	# pick default first option
	pp2_name = st.selectbox("Santral 2", options=pp2_candidates, index=0, key="pp2")
	if pp1_name == pp2_name:
		st.warning("İki farklı santral seçmelisiniz (aynı tipten).")
with col2:
	start_date = st.date_input("Başlangıç", value=DEFAULT_START, min_value=date(2024, 1, 1), max_value=date(2024, 12, 31))
	end_date = st.date_input("Bitiş", value=DEFAULT_END, min_value=date(2024, 1, 1), max_value=date(2024, 12, 31))
	if end_date < start_date:
		st.error("Bitiş tarihi başlangıçtan küçük olamaz.")
		st.stop()
with col3:
	st.write(" ")
	download_placeholder = st.empty()

selected_pp1 = next(p for p in plants if p.powerPlantName == pp1_name)
selected_pp2 = next(p for p in plants if p.powerPlantName == pp2_name)

run_btn = st.button(
	"Verileri Çek ve Excel İndir",
	type="primary",
	disabled=(pp1_name == pp2_name or _plant_type(next(p for p in plants if p.powerPlantName == pp2_name)) != _plant_type(next(p for p in plants if p.powerPlantName == pp1_name)))
)

if run_btn:
	progress = st.progress(0, text="PTF/SMF çekiliyor...")
	try:
		start_dt = datetime.combine(start_date, datetime.min.time())
		end_dt = datetime.combine(end_date, datetime.min.time())
		ptf_df = client.fetch_ptf(start_dt, end_dt)
		smf_df = client.fetch_smf(start_dt, end_dt)
		progress.progress(20, text=f"{selected_pp1.powerPlantName} verileri çekiliyor...")
		pp1_df = client.build_plant_dataframe(selected_pp1, start_dt, end_dt, ptf_df, smf_df)
		progress.progress(60, text=f"{selected_pp2.powerPlantName} verileri çekiliyor...")
		pp2_df = client.build_plant_dataframe(selected_pp2, start_dt, end_dt, ptf_df, smf_df)
		if pp1_df.empty or pp2_df.empty:
			st.error("Boş veri döndü. Tarih aralığı veya TGT doğruluğunu kontrol edin.")
			st.stop()
		# Comparison sheets built from the two plant tables
		sum1 = build_monthly_summary(pp1_df)
		sum2 = build_monthly_summary(pp2_df)
		# Yeni: aylık ek KPI'lar
		pp1_extras = compute_monthly_extras(pp1_df)
		pp2_extras = compute_monthly_extras(pp2_df)
		# Write Excel to memory
		buf = BytesIO()
		file_name = f"Analiz_{slugify(selected_pp1.powerPlantName)}_vs_{slugify(selected_pp2.powerPlantName)}_{start_date.strftime('%Y%m%d')}_{end_date.strftime('%Y%m%d')}.xlsx"
		with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
			# Sheet names
			sheet1 = "Santral_1"
			sheet2 = "Santral_2"
			sheetc = "Karşılaştırma"
			# Per-plant detailed tables with exact Turkish columns
			pp1_t = rename_to_turkish(pp1_df)
			pp2_t = rename_to_turkish(pp2_df)
			pp1_t.to_excel(writer, sheet_name=sheet1, index=False)
			pp2_t.to_excel(writer, sheet_name=sheet2, index=False)
			wb  = writer.book
			ws1 = writer.sheets[sheet1]
			ws2 = writer.sheets[sheet2]
			# Formats
			head_fmt = wb.add_format({"bold": True, "text_wrap": True, "valign": "vcenter", "align": "center"})
			num_fmt = wb.add_format({"num_format": "#,##0.00"})
			int_fmt = wb.add_format({"num_format": "0"})
			title_fmt = wb.add_format({"bold": True, "font_size": 14})
			# Column widths map
			col_widths = {
				"Tarih": 19,
				"Ay": 6,
				"Saat": 7,
			}
			# Apply table style and widths for a detailed sheet
			def style_detailed(ws, df):
				cols = list(df.columns)
				nrows = len(df)
				ncols = len(cols)
				# Add table over the data range
				ws.add_table(0, 0, nrows, ncols - 1, {
					"style": "Table Style Light 9",
					"columns": [{"header": c} for c in cols],
				})
				# Set widths and numeric formats
				for idx, name in enumerate(cols):
					width = col_widths.get(name, 16)
					# numeric columns (all except Tarih/Ay/Saat treated as numeric)
					if name in ("Tarih", "Ay", "Saat"):
						ws.set_column(idx, idx, width)
					elif name in ("Ay",):
						ws.set_column(idx, idx, width, int_fmt)
					else:
						ws.set_column(idx, idx, width, num_fmt)
				# Freeze header
				ws.freeze_panes(1, 0)
			# Style both detailed sheets
			style_detailed(ws1, pp1_t)
			style_detailed(ws2, pp2_t)
			# Comparison sheet with formulas
			wsC = wb.add_worksheet(sheetc)
			# Build monthly summary formulas referencing Santral_1 and Santral_2
			# Find column letters on detailed sheets (use pp1_t as reference)
			headers = list(pp1_t.columns)
			col_map = {name: _col_letter(headers.index(name)) for name in headers}
			# Month is column 'Ay'
			col_month = col_map["Ay"]
			# Ranges (data starts at row 2 in Excel 1-based because headers at row 1)
			last_row1 = len(pp1_t) + 1
			last_row2 = len(pp2_t) + 1
			# Helper to build SUMIF for a given column name and month
			def sumif(sheet, col_name, month_cell_ref):
				col = col_map[col_name]
				return f"=SUMIF('{sheet}'!${col_month}$2:${col_month}${last_row1 if sheet==sheet1 else last_row2},{month_cell_ref},'{sheet}'!${col}$2:${col}${last_row1 if sheet==sheet1 else last_row2})"
			# Headers for blocks
			base_headers = [
				"Ay",
				"Gerçekleşen Üretim  (MWh)",
				"Dengesizlik Miktarı  (MWh)",
				"GÖP Geliri (TL)",
				"Dengesizlik Tutarı (TL)",
				"Toplam Gelir (TL)",
				"Birim Gelir (TL/MWh)",
				"Dengesizlik Maliyeti (TL)",
				"Birim Deng Mal. (TL/MWh)",
			]
			extra_headers = [
				"Tahmin Doğruluğu (%)",
				"Maliyet Asimetrisi (Poz/Neg)",
				"Kapasite Faktörü (%)",
				"En Maliyetli 5 Gün (TL)",
				"Top 5 Gün DM Payı (%)",
				"Gelir Payı (%)",
				"Yıllık Pozitif Deng. Payı (%)",
				"Yıllık Negatif Deng. Payı (%)",
				"Üretim Saati (saat)",
				"Üretim Saat Payı (%)",
				"Üretim Payı (%)",
			]
			headers_comp = base_headers + extra_headers
			# Santral 1 title
			wsC.write(0, 0, "Santral 1", title_fmt)
			for j, h in enumerate(headers_comp):
				wsC.write(2, j, h, head_fmt)
			# Auto column widths for comparison sheet
			comp_col_widths = {
				"Ay": 6,
				"Gerçekleşen Üretim  (MWh)": 18,
				"Dengesizlik Miktarı  (MWh)": 18,
				"GÖP Geliri (TL)": 20,
				"Dengesizlik Tutarı (TL)": 20,
				"Toplam Gelir (TL)": 20,
				"Birim Gelir (TL/MWh)": 16,
				"Dengesizlik Maliyeti (TL)": 20,
				"Birim Deng Mal. (TL/MWh)": 18,
				"Tahmin Doğruluğu (%)": 14,
				"Maliyet Asimetrisi (Poz/Neg)": 18,
				"Kapasite Faktörü (%)": 14,
				"En Maliyetli 5 Gün (TL)": 20,
				"Top 5 Gün DM Payı (%)": 16,
				"Gelir Payı (%)": 14,
				"Yıllık Pozitif Deng. Payı (%)": 18,
				"Yıllık Negatif Deng. Payı (%)": 18,
				"Üretim Saati (saat)": 16,
				"Üretim Saat Payı (%)": 16,
				"Üretim Payı (%)": 14,
			}
			for j, h in enumerate(headers_comp):
				wsC.set_column(j, j, comp_col_widths.get(h, 16))
			# Fill months 1..12 (rows 3..14)
			for i in range(12):
				row = 3 + i
				month_num = i + 1
				wsC.write(row, 0, month_num, int_fmt)
				mref = f"A{row+1}"
				wsC.write_formula(row, 1, sumif(sheet1, "Gerçekleşen Üretim", mref), num_fmt)
				wsC.write_formula(row, 2, sumif(sheet1, "Dengesizlik Miktarı", mref), num_fmt)
				wsC.write_formula(row, 3, sumif(sheet1, "GÖP Geliri", mref), num_fmt)
				wsC.write_formula(row, 4, sumif(sheet1, "Dengesizlik Tutarı", mref), num_fmt)
				wsC.write_formula(row, 5, sumif(sheet1, "Toplam (Net) Gelir", mref), num_fmt)
				wsC.write_formula(row, 6, f"=IF(B{row+1}=0,0,F{row+1}/B{row+1})", num_fmt)
				wsC.write_formula(row, 7, sumif(sheet1, "Dengesizlik Maliyeti", mref), num_fmt)
				wsC.write_formula(row, 8, f"=IF(B{row+1}=0,0,H{row+1}/B{row+1})", num_fmt)
				# Extra columns (values computed in Python)
				base_idx = len(base_headers)
				wsC.write(row, base_idx + 0, pp1_extras["accuracy_pct"][i], num_fmt)
				wsC.write(row, base_idx + 1, (pp1_extras["asym_ratio"][i] if pp1_extras["asym_ratio"][i] is not None else ""), num_fmt)
				wsC.write(row, base_idx + 2, pp1_extras["capacity_factor_pct"][i], num_fmt)
				wsC.write(row, base_idx + 3, pp1_extras["top5_dm_tl"][i], num_fmt)
				wsC.write(row, base_idx + 4, pp1_extras["top5_dm_share_pct"][i], num_fmt)
				wsC.write(row, base_idx + 5, pp1_extras["revenue_share_pct"][i], num_fmt)
				wsC.write(row, base_idx + 6, pp1_extras["pos_share_pct"][i], num_fmt)
				wsC.write(row, base_idx + 7, pp1_extras["neg_share_pct"][i], num_fmt)
				wsC.write(row, base_idx + 8, pp1_extras["prod_hours"][i], int_fmt)
				wsC.write(row, base_idx + 9, pp1_extras["prod_hours_share_pct"][i], num_fmt)
				wsC.write(row, base_idx + 10, pp1_extras["prod_share_pct"][i], num_fmt)
			# Totals row for Santral 1 (row 15)
			total_row_1 = 3 + 12
			wsC.write(total_row_1, 0, "Toplam", head_fmt)
			wsC.write_formula(total_row_1, 1, f"=SUM(B4:B15)", num_fmt)
			wsC.write_formula(total_row_1, 2, f"=SUM(C4:C15)", num_fmt)
			wsC.write_formula(total_row_1, 3, f"=SUM(D4:D15)", num_fmt)
			wsC.write_formula(total_row_1, 4, f"=SUM(E4:E15)", num_fmt)
			wsC.write_formula(total_row_1, 5, f"=SUM(F4:F15)", num_fmt)
			wsC.write_formula(total_row_1, 6, f"=IF(B{total_row_1+1}=0,0,F{total_row_1+1}/B{total_row_1+1})", num_fmt)
			wsC.write_formula(total_row_1, 7, f"=SUM(H4:H15)", num_fmt)
			wsC.write_formula(total_row_1, 8, f"=IF(B{total_row_1+1}=0,0,H{total_row_1+1}/B{total_row_1+1})", num_fmt)
			# Sum for 'En Maliyetli 5 Gün (TL)'
			top5_col_idx = len(base_headers) + 3
			top5_col_letter = _col_letter(top5_col_idx)
			wsC.write_formula(total_row_1, top5_col_idx, f"=SUM({top5_col_letter}4:{top5_col_letter}15)", num_fmt)
			# Sum for 'Üretim Saati (saat)'
			prod_hours_col_idx = len(base_headers) + 8
			prod_hours_col_letter = _col_letter(prod_hours_col_idx)
			wsC.write_formula(total_row_1, prod_hours_col_idx, f"=SUM({prod_hours_col_letter}4:{prod_hours_col_letter}15)", int_fmt)
			# Santral 2 block start row
			start2 = total_row_1 + 3
			wsC.write(start2, 0, "Santral 2", title_fmt)
			for j, h in enumerate(headers_comp):
				wsC.write(start2 + 2, j, h, head_fmt)
			for i in range(12):
				row = start2 + 3 + i
				month_num = i + 1
				wsC.write(row, 0, month_num, int_fmt)
				mref = f"A{row+1}"
				wsC.write_formula(row, 1, sumif(sheet2, "Gerçekleşen Üretim", mref), num_fmt)
				wsC.write_formula(row, 2, sumif(sheet2, "Dengesizlik Miktarı", mref), num_fmt)
				wsC.write_formula(row, 3, sumif(sheet2, "GÖP Geliri", mref), num_fmt)
				wsC.write_formula(row, 4, sumif(sheet2, "Dengesizlik Tutarı", mref), num_fmt)
				wsC.write_formula(row, 5, sumif(sheet2, "Toplam (Net) Gelir", mref), num_fmt)
				wsC.write_formula(row, 6, f"=IF(B{row+1}=0,0,F{row+1}/B{row+1})", num_fmt)
				wsC.write_formula(row, 7, sumif(sheet2, "Dengesizlik Maliyeti", mref), num_fmt)
				wsC.write_formula(row, 8, f"=IF(B{row+1}=0,0,H{row+1}/B{row+1})", num_fmt)
				# Extra columns for Santral 2
				base_idx = len(base_headers)
				wsC.write(row, base_idx + 0, pp2_extras["accuracy_pct"][i], num_fmt)
				wsC.write(row, base_idx + 1, (pp2_extras["asym_ratio"][i] if pp2_extras["asym_ratio"][i] is not None else ""), num_fmt)
				wsC.write(row, base_idx + 2, pp2_extras["capacity_factor_pct"][i], num_fmt)
				wsC.write(row, base_idx + 3, pp2_extras["top5_dm_tl"][i], num_fmt)
				wsC.write(row, base_idx + 4, pp2_extras["top5_dm_share_pct"][i], num_fmt)
				wsC.write(row, base_idx + 5, pp2_extras["revenue_share_pct"][i], num_fmt)
				wsC.write(row, base_idx + 6, pp2_extras["pos_share_pct"][i], num_fmt)
				wsC.write(row, base_idx + 7, pp2_extras["neg_share_pct"][i], num_fmt)
				wsC.write(row, base_idx + 8, pp2_extras["prod_hours"][i], int_fmt)
				wsC.write(row, base_idx + 9, pp2_extras["prod_hours_share_pct"][i], num_fmt)
				wsC.write(row, base_idx + 10, pp2_extras["prod_share_pct"][i], num_fmt)
			# Totals row for Santral 2
			total_row_2 = start2 + 3 + 12
			wsC.write(total_row_2, 0, "Toplam", head_fmt)
			wsC.write_formula(total_row_2, 1, f"=SUM(B{start2+4}:B{start2+15})", num_fmt)
			wsC.write_formula(total_row_2, 2, f"=SUM(C{start2+4}:C{start2+15})", num_fmt)
			wsC.write_formula(total_row_2, 3, f"=SUM(D{start2+4}:D{start2+15})", num_fmt)
			wsC.write_formula(total_row_2, 4, f"=SUM(E{start2+4}:E{start2+15})", num_fmt)
			wsC.write_formula(total_row_2, 5, f"=SUM(F{start2+4}:F{start2+15})", num_fmt)
			wsC.write_formula(total_row_2, 6, f"=IF(B{total_row_2+1}=0,0,F{total_row_2+1}/B{total_row_2+1})", num_fmt)
			wsC.write_formula(total_row_2, 7, f"=SUM(H{start2+4}:H{start2+15})", num_fmt)
			wsC.write_formula(total_row_2, 8, f"=IF(B{total_row_2+1}=0,0,H{total_row_2+1}/B{total_row_2+1})", num_fmt)
			# Sum for 'En Maliyetli 5 Gün (TL)' in Santral 2 block
			top5_col_idx_2 = len(base_headers) + 3
			top5_col_letter_2 = _col_letter(top5_col_idx_2)
			wsC.write_formula(total_row_2, top5_col_idx_2, f"=SUM({top5_col_letter_2}{start2+4}:{top5_col_letter_2}{start2+15})", num_fmt)
			# Sum for 'Üretim Saati (saat)' in Santral 2 block
			prod_hours_col_idx_2 = len(base_headers) + 8
			prod_hours_col_letter_2 = _col_letter(prod_hours_col_idx_2)
			wsC.write_formula(total_row_2, prod_hours_col_idx_2, f"=SUM({prod_hours_col_letter_2}{start2+4}:{prod_hours_col_letter_2}{start2+15})", int_fmt)
		# Writer is closed here; now stream the buffer
		buf.seek(0)
		progress.progress(100, text="Hazır")
		download_placeholder.download_button(
			label="Excel'i İndir",
			data=buf,
			file_name=file_name,
			mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
		)
	except Exception as e:
		st.error(f"İşlem sırasında hata: {e}")
	finally:
		progress.empty() 