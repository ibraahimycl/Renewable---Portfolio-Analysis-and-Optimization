import json
import time
from dataclasses import dataclass
from datetime import datetime, timedelta, date
from io import BytesIO
from typing import Dict, Iterable, List, Optional, Tuple

import pandas as pd
import requests


ISO_FMT = "%Y-%m-%dT%H:%M:%S+03:00"


def start_of_day(dt: date) -> datetime:
	return datetime(dt.year, dt.month, dt.day, 0, 0, 0)


def end_of_day(dt: date) -> datetime:
	return datetime(dt.year, dt.month, dt.day, 0, 0, 0)


def month_start_end_strings(start_dt: datetime, end_dt: datetime) -> List[Tuple[str, str]]:
	"""Split [start_dt, end_dt] into month-sized ISO-8601 ranges with +03:00 tz suffix."""
	if end_dt < start_dt:
		raise ValueError("end_dt must be >= start_dt")
	ranges: List[Tuple[str, str]] = []
	cursor = datetime(start_dt.year, start_dt.month, 1)
	# Ensure cursor <= start_dt month
	if cursor < datetime(start_dt.year, start_dt.month, 1):
		cursor = datetime(start_dt.year, start_dt.month, 1)
	# iterate months until we pass end_dt
	while cursor <= end_dt:
		if cursor.month == 12:
			next_month = datetime(cursor.year + 1, 1, 1)
		else:
			next_month = datetime(cursor.year, cursor.month + 1, 1)
		period_start = max(start_dt, cursor)
		period_end = min(end_dt, next_month - timedelta(days=1))
		ranges.append(
			(period_start.strftime(ISO_FMT), period_end.strftime(ISO_FMT))
		)
		cursor = next_month
	return ranges


@dataclass
class PlantMeta:
	powerPlantName: str
	organizationId: int
	powerPlantId: int
	uevcbId: int


class EpiasClient:
	"""Simple client for EPİAŞ Transparency Platform endpoints."""

	TGT_URL = "https://giris.epias.com.tr/cas/v1/tickets"
	PTF_URL = "https://seffaflik.epias.com.tr/electricity-service/v1/markets/dam/data/mcp"
	SMF_URL = "https://seffaflik.epias.com.tr/electricity-service/v1/markets/bpm/data/system-marginal-price"
	KGUP_URL = "https://seffaflik.epias.com.tr/electricity-service/v1/generation/data/dpp-first-version"
	URETIM_URL = "https://seffaflik.epias.com.tr/electricity-service/v1/generation/data/realtime-generation"

	def __init__(self, tgt: Optional[str] = None):
		self.tgt: Optional[str] = tgt

	@staticmethod
	def obtain_tgt(username: str, password: str, timeout: int = 30) -> str:
		"""Obtain TGT token using username/password. Token is valid ~2 hours."""
		headers = {"Accept": "text/plain", "Content-Type": "application/x-www-form-urlencoded"}
		resp = requests.post(EpiasClient.TGT_URL, data={"username": username, "password": password}, headers=headers, timeout=timeout)
		resp.raise_for_status()
		return resp.text.strip()

	def _post_json(self, url: str, body: Dict, timeout: int = 60) -> Dict:
		if not self.tgt:
			raise RuntimeError("TGT is not set. Please authenticate first.")
		headers = {
			"TGT": self.tgt,
			"Accept-Language": "en",
			"Accept": "application/json",
			"Content-Type": "application/json",
		}
		resp = requests.post(url, headers=headers, json=body, timeout=timeout)
		resp.raise_for_status()
		return resp.json()

	def fetch_ptf(self, start_dt: datetime, end_dt: datetime, delay_s: float = 0.1) -> pd.DataFrame:
		rows: List[Dict] = []
		for s, e in month_start_end_strings(start_dt, end_dt):
			data = self._post_json(self.PTF_URL, {"startDate": s, "endDate": e})
			for it in data.get("items", []):
				rows.append({"date": it.get("date"), "hour": it.get("hour"), "PTF": it.get("price")})
			time.sleep(delay_s)
		df = pd.DataFrame(rows)
		if df.empty:
			return df
		return self._build_datetime(df).loc[:, ["datetime", "PTF"]]

	def fetch_smf(self, start_dt: datetime, end_dt: datetime, delay_s: float = 0.1) -> pd.DataFrame:
		rows: List[Dict] = []
		for s, e in month_start_end_strings(start_dt, end_dt):
			data = self._post_json(self.SMF_URL, {"startDate": s, "endDate": e})
			for it in data.get("items", []):
				rows.append({"date": it.get("date"), "hour": it.get("hour"), "SMF": it.get("systemMarginalPrice")})
			time.sleep(delay_s)
		df = pd.DataFrame(rows)
		if df.empty:
			return df
		return self._build_datetime(df).loc[:, ["datetime", "SMF"]]

	def fetch_kgup(self, plant: PlantMeta, start_dt: datetime, end_dt: datetime, delay_s: float = 0.2) -> pd.DataFrame:
		rows: List[Dict] = []
		for s, e in month_start_end_strings(start_dt, end_dt):
			body = {
				"startDate": s,
				"endDate": e,
				"organizationId": plant.organizationId,
				"uevcbId": plant.uevcbId,
				"region": "TR1",
			}
			data = self._post_json(self.KGUP_URL, body)
			for it in data.get("items", []):
				rows.append({
					"date": it.get("date"),
					"hour": it.get("time"),
					"KGUP": it.get("toplam"),
				})
			time.sleep(delay_s)
		df = pd.DataFrame(rows)
		if df.empty:
			return df
		df = self._build_datetime(df)
		return df.loc[:, ["datetime", "Tarih", "Saat", "KGUP"]]

	def fetch_uretim(self, plant: PlantMeta, start_dt: datetime, end_dt: datetime, delay_s: float = 0.2) -> pd.DataFrame:
		rows: List[Dict] = []
		for s, e in month_start_end_strings(start_dt, end_dt):
			body = {
				"startDate": s,
				"endDate": e,
				"powerPlantId": plant.powerPlantId,
			}
			data = self._post_json(self.URETIM_URL, body)
			for it in data.get("items", []):
				rows.append({
					"date": it.get("date"),
					"hour": it.get("time"),
					"URETIM": it.get("total"),
				})
			time.sleep(delay_s)
		df = pd.DataFrame(rows)
		if df.empty:
			return df
		df = self._build_datetime(df)
		return df.loc[:, ["datetime", "URETIM"]]

	def build_plant_dataframe(
		self,
		plant: PlantMeta,
		start_dt: datetime,
		end_dt: datetime,
		ptf_df: Optional[pd.DataFrame] = None,
		smf_df: Optional[pd.DataFrame] = None,
	) -> pd.DataFrame:
		"""Create the combined per-plant dataframe with required derived columns."""
		kgup_df = self.fetch_kgup(plant, start_dt, end_dt)
		uretim_df = self.fetch_uretim(plant, start_dt, end_dt)
		if kgup_df.empty or uretim_df.empty:
			return pd.DataFrame()
		df = kgup_df.merge(uretim_df, on="datetime", how="inner")
		if ptf_df is None:
			ptf_df = self.fetch_ptf(start_dt, end_dt)
		if smf_df is None:
			smf_df = self.fetch_smf(start_dt, end_dt)
		if not ptf_df.empty:
			df = df.merge(ptf_df, on="datetime", how="left")
		if not smf_df.empty:
			df = df.merge(smf_df, on="datetime", how="left")
		# Derived columns
		df["Ay"] = df["datetime"].dt.month
		df["Poz_DF"] = (df[["PTF", "SMF"]].min(axis=1) * 0.97).astype(float)
		df["Neg_DF"] = (df[["PTF", "SMF"]].max(axis=1) * 1.03).astype(float)
		df["KGUP"] = pd.to_numeric(df["KGUP"], errors="coerce")
		df["URETIM"] = pd.to_numeric(df["URETIM"], errors="coerce")
		df["PTF"] = pd.to_numeric(df["PTF"], errors="coerce")
		df["SMF"] = pd.to_numeric(df["SMF"], errors="coerce")
		df["Dengesizlik"] = df["URETIM"] - df["KGUP"]
		df["GOP_Geliri"] = df["KGUP"] * df["PTF"]
		def _imbalance_amount(row: pd.Series) -> float:
			if pd.isna(row["Dengesizlik"]):
				return 0.0
			return row["Dengesizlik"] * (row["Poz_DF"] if row["Dengesizlik"] >= 0 else row["Neg_DF"])
		df["Dengesizlik_Tutarı"] = df.apply(_imbalance_amount, axis=1)
		df["Net_Gelir"] = df["GOP_Geliri"] + df["Dengesizlik_Tutarı"]
		df["Dengesizlik_Maliyeti"] = (df["URETIM"] * df["PTF"] - df["Net_Gelir"]).clip(lower=0)
		def _unit_cost(row: pd.Series) -> float:
			uretim_val = row["URETIM"]
			if pd.isna(uretim_val) or float(uretim_val) == 0.0:
				return 0.0
			return float(row["Dengesizlik_Maliyeti"]) / float(uretim_val)
		df["Birim_DM"] = df.apply(_unit_cost, axis=1)
		cols = [
			"Tarih", "Ay", "Saat",
			"PTF", "SMF", "Poz_DF", "Neg_DF",
			"KGUP", "URETIM", "Dengesizlik",
			"GOP_Geliri", "Dengesizlik_Tutarı", "Net_Gelir",
			"Dengesizlik_Maliyeti", "Birim_DM",
		]
		return df.loc[:, cols].copy()

	@staticmethod
	def _build_datetime(df: pd.DataFrame, date_col: str = "date", hour_col: str = "hour") -> pd.DataFrame:
		out = df.copy()
		# Normalize day as YYYY-MM-DD
		day = out[date_col].astype(str).str.extract(r'^(\d{4}-\d{2}-\d{2})')[0]
		# Normalize hour to HH:MM (strip anything like seconds or +03:00)
		hour_raw = out[hour_col].astype(str) if hour_col in out.columns else pd.Series([None] * len(out))
		hour_norm = hour_raw.str.extract(r'^(\d{2}:\d{2})')[0]
		# If missing, derive from date string
		derived_from_date = out[date_col].astype(str).str.extract(r'T(\d{2}:\d{2})')[0]
		hour_final = hour_norm.fillna(derived_from_date).fillna("00:00")
		# Build tz-naive datetime with explicit format
		dt = pd.to_datetime(day + " " + hour_final, format="%Y-%m-%d %H:%M", errors="coerce")
		out = out.assign(datetime=dt, Tarih=day, Saat=hour_final.astype(str))
		return out


def slugify(value: str) -> str:
	s = str(value).strip().replace(" ", "_")
	return "".join(ch for ch in s if ch.isalnum() or ch in ("_", "-"))


def load_plants(pp_list_path_candidates: Iterable[str]) -> List[PlantMeta]:
	"""Load plant list from JSON and normalize key naming.

	Accepts an iterable of path candidates and picks the first existing.
	"""
	import os
	path_to_use: Optional[str] = None
	for p in pp_list_path_candidates:
		if p and os.path.exists(p):
			path_to_use = p
			break
	if not path_to_use:
		raise FileNotFoundError("pp_list.json not found in provided locations")
	with open(path_to_use, "r", encoding="utf-8") as f:
		raw = json.load(f)
	def _norm(pp: Dict) -> PlantMeta:
		return PlantMeta(
			powerPlantName=pp.get("powerPlantName") or pp.get("powerplantName"),
			organizationId=int(pp.get("organizationId")),
			powerPlantId=int(pp.get("powerPlantId") or pp.get("powerplantId")),
			uevcbId=int(pp.get("uevcbId")),
		)
	return [_norm(pp) for pp in raw] 