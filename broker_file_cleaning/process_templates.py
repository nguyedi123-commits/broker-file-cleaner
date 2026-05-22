#!/usr/bin/env python3
"""
Broker meter data normaliser.

Supports two modes:
- Local batch mode using Input/ and Output/ folders
- Web mode via process_uploaded_files() for Streamlit uploads
"""
from __future__ import annotations

import io
import json
import re
import shutil
import traceback
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import BinaryIO

import pandas as pd

BASE = Path(__file__).resolve().parent
INPUT_DIR = BASE / "Input"
OUTPUT_DIR = BASE / "Output"
REGISTRY_PATH = BASE / "format_registry.json"

SUPPORTED_EXT = {".xlsx", ".csv", ".xlsb", ".xls"}
EXCLUDED_FILENAME_PATTERNS = ["Market Region Data Report"]
SUPPORTED_ENERGY_UNITS = {"KWH": 1.0, "MWH": 1000.0}


@dataclass
class ProcessedFileResult:
    filename: str
    rows: int
    source_type: str
    format_id: str
    nmi_preview: str
    aggregated_to_30min: bool


@dataclass
class FailedFileResult:
    filename: str
    status: str
    reason: str


@dataclass
class ProcessResult:
    output_csv_bytes: bytes | None
    output_filename: str | None
    processed_files: list[ProcessedFileResult]
    failed_files: list[FailedFileResult]
    warnings: list[str]
    total_rows: int
    unique_nmis: int

    def processed_rows(self) -> int:
        return sum(item.rows for item in self.processed_files)


def load_registry(registry_path: Path = REGISTRY_PATH) -> list[dict]:
    with open(registry_path, encoding="utf-8") as f:
        return json.load(f)["formats"]


def is_nem12(file_path: Path) -> bool:
    try:
        with open(file_path, encoding="utf-8", errors="ignore") as f:
            first = f.readline()
            return first.startswith("100,NEM12") or first.startswith("200,")
    except Exception:
        return False


def parse_nem12(file_path: Path) -> pd.DataFrame:
    """
    Parse a NEM12 file and return [NMI, dt, consumption_kWh, export_kWh].
    Handles multiple NMI blocks and both E (import) and B (export) registers.
    """
    blocks: dict[tuple, list] = {}
    nmi = None
    interval_min = 30
    is_export = False
    skip_stream = False

    with open(file_path, encoding="utf-8", errors="ignore") as f:
        for line in f:
            parts = line.strip().split(",")
            if not parts or not parts[0]:
                continue
            rec = parts[0]

            if rec == "200":
                offset = 1 if (len(parts) > 1 and parts[1] == "NEM12") else 0
                nmi = parts[1 + offset] if len(parts) > 1 + offset else None
                register_id = parts[3 + offset] if len(parts) > 3 + offset else "E1"
                uom = parts[7 + offset] if len(parts) > 7 + offset else "kWh"
                unit_name, unit_multiplier = _normalise_energy_unit(uom)
                skip_stream = unit_name not in SUPPORTED_ENERGY_UNITS
                try:
                    raw_interval = parts[8 + offset] if len(parts) > 8 + offset else ""
                    interval_min = int(raw_interval) if raw_interval.strip() else 30
                except (ValueError, IndexError):
                    interval_min = 30
                is_export = register_id.startswith("B")
                key = (nmi, "export" if is_export else "consume")
                if not skip_stream and key not in blocks:
                    blocks[key] = []

            elif rec == "300" and nmi and not skip_stream:
                n_intervals = 1440 // interval_min
                try:
                    date = pd.to_datetime(parts[1], format="%Y%m%d")
                except Exception:
                    continue
                values = parts[2 : 2 + n_intervals]
                key = (nmi, "export" if is_export else "consume")
                for i, value in enumerate(values):
                    try:
                        kwh = float(value) * unit_multiplier
                    except (ValueError, TypeError):
                        continue
                    dt = date + pd.Timedelta(minutes=interval_min * i)
                    blocks.setdefault(key, []).append({"dt": dt, "kwh": kwh})

    if not blocks:
        return pd.DataFrame(columns=["NMI", "dt", "consumption_kWh", "export_kWh"])

    frames = []
    for nmi_value in {key[0] for key in blocks}:
        consume = pd.DataFrame(blocks.get((nmi_value, "consume"), [])).rename(
            columns={"kwh": "consumption_kWh"}
        )
        export = pd.DataFrame(blocks.get((nmi_value, "export"), [])).rename(
            columns={"kwh": "export_kWh"}
        )

        if consume.empty and export.empty:
            continue
        if consume.empty:
            export["consumption_kWh"] = 0.0
            merged = export
        elif export.empty:
            consume["export_kWh"] = 0.0
            merged = consume
        else:
            merged = consume.merge(export, on="dt", how="outer").fillna(0)

        merged["NMI"] = nmi_value
        frames.append(merged[["NMI", "dt", "consumption_kWh", "export_kWh"]])

    if not frames:
        return pd.DataFrame(columns=["NMI", "dt", "consumption_kWh", "export_kWh"])

    df = pd.concat(frames, ignore_index=True)
    if interval_min < 30:
        df = _resample_30min(df)
    return df[["NMI", "dt", "consumption_kWh", "export_kWh"]]


def _headers(file_path: Path, skip_rows: int, sheet_index=None) -> set[str]:
    try:
        ext = file_path.suffix.lower()
        if ext in (".xlsx", ".xlsb"):
            engine = "pyxlsb" if ext == ".xlsb" else None
            if sheet_index is not None:
                df = pd.read_excel(file_path, header=skip_rows, nrows=0, sheet_name=sheet_index, engine=engine)
                return set(df.columns.astype(str))
            sheets = pd.read_excel(file_path, header=skip_rows, nrows=0, sheet_name=None, engine=engine)
            all_cols: set[str] = set()
            for df in sheets.values():
                all_cols |= set(df.columns.astype(str))
            return all_cols
        if ext == ".xls":
            with open(file_path, encoding="utf-8", errors="ignore") as f:
                chunk = f.read(16384)
            tr = re.search(r"<tr[^>]*>(.*?)</tr>", chunk, re.IGNORECASE | re.DOTALL)
            if tr:
                cells = re.findall(r"<t[hd][^>]*>(.*?)</t[hd]>", tr.group(1), re.IGNORECASE | re.DOTALL)
                return {cell.strip() for cell in cells if cell.strip()}
            return set()
        df = pd.read_csv(file_path, header=skip_rows, nrows=0)
        return set(df.columns.astype(str))
    except Exception:
        return set()


def detect_format(file_path: Path, registry: list[dict]) -> dict | None:
    ext = file_path.suffix.lower().lstrip(".")
    for fmt in registry:
        if fmt.get("file_type") and ext != fmt["file_type"]:
            continue
        skip = fmt.get("skip_rows", 0)
        headers = _headers(file_path, skip, sheet_index=fmt.get("sheet_index"))
        if set(fmt["fingerprint_columns"]).issubset(headers):
            col_pattern = fmt.get("col_pattern")
            if col_pattern and not any(re.search(col_pattern, header) for header in headers):
                continue
            req_sheets = fmt.get("fingerprint_sheets")
            if req_sheets:
                try:
                    xl_sheets = pd.ExcelFile(file_path).sheet_names
                    if not all(sheet in xl_sheets for sheet in req_sheets):
                        continue
                except Exception:
                    continue
            return fmt
    return None


def _read_raw(file_path: Path, fmt: dict) -> pd.DataFrame:
    skip = fmt.get("skip_rows", 0)
    ext = file_path.suffix.lower()
    if ext in (".xlsx", ".xlsb"):
        engine = "pyxlsb" if ext == ".xlsb" else None
        sheet_index = fmt.get("sheet_index")
        if sheet_index is not None:
            return pd.read_excel(file_path, header=skip, sheet_name=sheet_index, engine=engine)
        sheets = pd.read_excel(file_path, header=skip, sheet_name=None, engine=engine)
        fingerprint = set(fmt["fingerprint_columns"])
        matching = [df for df in sheets.values() if fingerprint.issubset(set(df.columns.astype(str)))]
        if not matching:
            return next(iter(sheets.values()))
        return pd.concat(matching, ignore_index=True)
    if ext == ".xls":
        tables = pd.read_html(file_path, header=skip)
        fingerprint = set(fmt["fingerprint_columns"])
        for table in tables:
            if fingerprint.issubset(set(table.columns.astype(str))):
                return table
        return tables[0]
    return pd.read_csv(file_path, header=skip)


def _parse_dt(df: pd.DataFrame, fmt: dict) -> pd.Series:
    cols = fmt["datetime_cols"]
    dt_type = fmt.get("datetime_type", "auto")
    raw = df[cols[0]].astype(str) + " " + df[cols[1]].astype(str) if len(cols) == 2 else df[cols[0]]
    if dt_type == "auto":
        try:
            return pd.to_datetime(raw)
        except Exception:
            return pd.to_datetime(raw, format="mixed")
    if dt_type == "dayfirst":
        midnight_mask = None
        if len(cols) == 2:
            midnight_mask = raw.str.contains(r"\s24:00", regex=True, na=False)
            if midnight_mask.any():
                raw = raw.str.replace(r"\s24:00(:00)?$", " 00:00", regex=True)
        try:
            result = pd.to_datetime(raw, dayfirst=True)
        except Exception:
            result = pd.to_datetime(raw, dayfirst=True, format="mixed")
        if midnight_mask is not None and midnight_mask.any():
            result = result + pd.to_timedelta(midnight_mask.astype(int), unit="D")
        return result
    if dt_type == "xlsb_serial":
        return pd.Timestamp("1899-12-30") + pd.to_timedelta(df[cols[0]], unit="D")
    if dt_type == "xlsb_date_time":
        serial = pd.to_numeric(df[cols[0]], errors="coerce") + pd.to_numeric(df[cols[1]], errors="coerce")
        return pd.Timestamp("1899-12-30") + pd.to_timedelta(serial, unit="D")
    if dt_type == "date_plus_time":
        dates = pd.to_datetime(df[cols[0]], errors="coerce")
        times = pd.to_timedelta(df[cols[1]].astype(str), errors="coerce")
        return dates + times
    return pd.to_datetime(raw, format=dt_type)


def _resample_30min(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["dt"] = pd.to_datetime(df["dt"]).dt.floor("30min")
    return df.groupby(["NMI", "dt"], as_index=False)[["consumption_kWh", "export_kWh"]].sum()


def _nmi_from_filename(file_path: Path) -> str:
    stem = file_path.stem
    matches = re.findall(r"(?<![A-Z0-9])([A-Z]{0,4}[0-9]{7,11})(?![A-Z0-9])", stem, re.IGNORECASE)
    for match in matches:
        if 10 <= len(match) <= 11:
            return match.upper()
    matches = re.findall(r"(?<![A-Z0-9])([A-Z0-9]{10,11})(?![A-Z0-9])", stem, re.IGNORECASE)
    for match in matches:
        if 10 <= len(match) <= 11:
            return match.upper()
    return "UNKNOWN"


def _normalise_energy_unit(value: object) -> tuple[str | None, float]:
    if pd.isna(value):
        return None, 1.0
    unit = str(value).strip().upper()
    multiplier = SUPPORTED_ENERGY_UNITS.get(unit)
    return unit, multiplier if multiplier is not None else 1.0


def _filter_supported_energy_rows(
    df: pd.DataFrame,
    unit_col: str,
    value_col: str,
) -> tuple[pd.DataFrame, set[str]]:
    units = df[unit_col].apply(_normalise_energy_unit)
    unit_names = units.str[0]
    multipliers = units.str[1]
    supported_mask = unit_names.isin(SUPPORTED_ENERGY_UNITS)
    ignored_units = {
        unit for unit in unit_names[~supported_mask].dropna().astype(str).unique() if unit
    }

    filtered = df.loc[supported_mask].copy()
    if filtered.empty:
        return filtered, ignored_units

    filtered[value_col] = pd.to_numeric(filtered[value_col], errors="coerce").fillna(0)
    filtered[value_col] = filtered[value_col] * multipliers[supported_mask].astype(float)
    return filtered, ignored_units


def normalise(file_path: Path, fmt: dict) -> pd.DataFrame:
    if fmt.get("multi_sheet_wide"):
        dt_col = fmt["datetime_col"]
        cons_df = pd.read_excel(file_path, sheet_name=fmt["consumption_sheet"])
        exp_df = pd.read_excel(file_path, sheet_name=fmt["export_sheet"])
        cons_df.columns = cons_df.columns.astype(str)
        exp_df.columns = exp_df.columns.astype(str)
        cons_melted = cons_df.melt(id_vars=[dt_col], var_name="NMI", value_name="consumption_kWh")
        exp_melted = exp_df.melt(id_vars=[dt_col], var_name="NMI", value_name="export_kWh")
        cons_melted["dt"] = pd.to_datetime(cons_melted[dt_col])
        exp_melted["dt"] = pd.to_datetime(exp_melted[dt_col])
        cons_melted["consumption_kWh"] = pd.to_numeric(cons_melted["consumption_kWh"], errors="coerce").fillna(0)
        exp_melted["export_kWh"] = pd.to_numeric(exp_melted["export_kWh"], errors="coerce").fillna(0)
        result = (
            cons_melted[["NMI", "dt", "consumption_kWh"]]
            .merge(exp_melted[["NMI", "dt", "export_kWh"]], on=["NMI", "dt"], how="outer")
            .fillna(0)
        )
        result = result[result["NMI"].str.match(r"^[A-Z0-9]{7,11}$", na=False)]
        return result[["NMI", "dt", "consumption_kWh", "export_kWh"]]

    if fmt.get("nmi_from_sheet"):
        skip = fmt.get("skip_rows", 0)
        sheets = pd.read_excel(file_path, header=skip, sheet_name=None)
        nmi_pat = fmt.get("sheet_nmi_pattern", r"^[A-Z0-9]{10,11}$")
        frames = []
        for sheet_name, sheet_df in sheets.items():
            if not re.match(nmi_pat, str(sheet_name)):
                continue
            sheet_df = sheet_df.copy()
            sheet_df["NMI"] = str(sheet_name)
            frames.append(sheet_df)
        if not frames:
            return pd.DataFrame(columns=["NMI", "dt", "consumption_kWh", "export_kWh"])
        df = pd.concat(frames, ignore_index=True)
        df["dt"] = _parse_dt(df, fmt)
        if fmt.get("datetime_is_end"):
            df["dt"] -= pd.Timedelta(minutes=fmt.get("interval_min", 30))
        watts_col = fmt.get("watts_col")
        if watts_col:
            interval = fmt.get("interval_min", 30)
            divisor = 1000 if fmt.get("watts_unit", "kW") == "W" else 1
            df["consumption_kWh"] = pd.to_numeric(df[watts_col], errors="coerce").fillna(0) * (interval / 60) / divisor
        else:
            df["consumption_kWh"] = pd.to_numeric(df[fmt["consumption_col"]], errors="coerce").fillna(0)
        df["export_kWh"] = 0.0
        df = df[["NMI", "dt", "consumption_kWh", "export_kWh"]]
        if fmt.get("interval_min", 30) < 30:
            df = _resample_30min(df)
        return df[["NMI", "dt", "consumption_kWh", "export_kWh"]]

    df = _read_raw(file_path, fmt)

    if fmt.get("wide_interval_cols"):
        nmi_col = fmt["nmi_col"]
        date_col = fmt["date_col"]
        ct_col = fmt.get("consumption_type_col", "CONSUMPTION_TYPE")
        unit_col = fmt.get("unit_col")
        unit_filt = fmt.get("unit_filter", "KWH")
        imp_type = fmt.get("import_type", "Import")
        exp_type = fmt.get("export_type", "Export")

        df["NMI"] = df[nmi_col].astype(str)
        if unit_col and unit_filt.upper() not in SUPPORTED_ENERGY_UNITS:
            df = df[df[unit_col].astype(str).str.upper() == unit_filt.upper()]
        df["_date"] = pd.to_datetime(df[date_col], dayfirst=True, errors="coerce")
        df = df.dropna(subset=["_date"])

        time_cols = [col for col in df.columns if re.match(r"^\d{2}:\d{2} - \d{2}:\d{2}$", str(col))]
        if unit_col and unit_filt.upper() in SUPPORTED_ENERGY_UNITS:
            unit_pairs = df[unit_col].apply(_normalise_energy_unit)
            supported_mask = unit_pairs.str[0].isin(SUPPORTED_ENERGY_UNITS)
            df = df.loc[supported_mask].copy()
            if df.empty:
                return pd.DataFrame(columns=["NMI", "dt", "consumption_kWh", "export_kWh"])
            multipliers = unit_pairs.loc[supported_mask].str[1].astype(float)
            for col in time_cols:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0) * multipliers
        melted = df[["NMI", "_date", ct_col] + time_cols].melt(
            id_vars=["NMI", "_date", ct_col], var_name="_interval", value_name="_kWh"
        )
        melted["_kWh"] = pd.to_numeric(melted["_kWh"], errors="coerce").fillna(0)

        def interval_end_td(value: str) -> pd.Timedelta:
            end = str(value).split(" - ")[1].strip()
            hour, minute = map(int, end.split(":"))
            return pd.Timedelta(hours=hour if hour else 24, minutes=minute)

        melted["dt"] = melted["_date"] + melted["_interval"].map(interval_end_td)

        imp = (
            melted[melted[ct_col] == imp_type]
            .groupby(["NMI", "dt"])["_kWh"]
            .sum()
            .rename("consumption_kWh")
            .reset_index()
        )
        exp = (
            melted[melted[ct_col] == exp_type]
            .groupby(["NMI", "dt"])["_kWh"]
            .sum()
            .rename("export_kWh")
            .reset_index()
        )
        result = imp.merge(exp, on=["NMI", "dt"], how="outer").fillna(0)
        return result[["NMI", "dt", "consumption_kWh", "export_kWh"]]

    if fmt.get("wide_date_only"):
        df.columns = df.columns.astype(str)
        date_col = fmt["date_col"]
        nmi_pat = fmt.get("col_pattern", r"^[A-Z0-9]{10,11}$")
        nmi_cols = [col for col in df.columns if re.match(nmi_pat, col)]
        interval = fmt.get("interval_min", 5)

        df["_date"] = pd.to_datetime(df[date_col], dayfirst=True, errors="coerce")
        df = df.dropna(subset=["_date"])
        df["_slot"] = df.groupby("_date").cumcount()
        offset_min = (df["_slot"] + 1) * interval if fmt.get("datetime_is_end") else df["_slot"] * interval
        df["dt"] = df["_date"] + pd.to_timedelta(offset_min, unit="min")

        melted = df[["dt"] + nmi_cols].melt(id_vars="dt", var_name="NMI", value_name="consumption_kWh")
        melted["export_kWh"] = 0.0
        melted["consumption_kWh"] = pd.to_numeric(melted["consumption_kWh"], errors="coerce")
        melted = melted.dropna(subset=["consumption_kWh"])
        melted["consumption_kWh"] = melted["consumption_kWh"].fillna(0)
        melted = melted[["NMI", "dt", "consumption_kWh", "export_kWh"]]
        if interval < 30:
            melted = _resample_30min(melted)
        return melted[["NMI", "dt", "consumption_kWh", "export_kWh"]]

    if fmt.get("nmi_in_header"):
        col_pat = fmt.get("col_pattern", r"^([A-Z0-9]{10,11})$")
        nmi_col_match = next((col for col in df.columns if re.search(col_pat, str(col), re.IGNORECASE)), None)
        if not nmi_col_match:
            return pd.DataFrame(columns=["NMI", "dt", "consumption_kWh", "export_kWh"])
        nmi = re.search(col_pat, str(nmi_col_match), re.IGNORECASE).group(1)
        df["NMI"] = nmi
        df["dt"] = _parse_dt(df, fmt)
        if fmt.get("datetime_is_end"):
            df["dt"] -= pd.Timedelta(minutes=fmt.get("interval_min", 30))
        import_suffix = fmt.get("import_col_suffix")
        import_col = f"{nmi}{import_suffix}" if import_suffix else nmi_col_match
        export_suffix = fmt.get("export_col_suffix")
        export_col = f"{nmi}{export_suffix}" if export_suffix else None
        df["consumption_kWh"] = pd.to_numeric(df[import_col], errors="coerce").fillna(0)
        if export_col and export_col in df.columns:
            df["export_kWh"] = pd.to_numeric(df[export_col], errors="coerce").fillna(0)
        else:
            df["export_kWh"] = 0.0
        df = df[["NMI", "dt", "consumption_kWh", "export_kWh"]]
        if fmt.get("interval_min", 30) < 30:
            df = _resample_30min(df)
        return df[["NMI", "dt", "consumption_kWh", "export_kWh"]]

    if fmt.get("nmi_from_filename"):
        df["NMI"] = _nmi_from_filename(file_path)
    elif fmt.get("nmi_as_int"):
        df["NMI"] = df[fmt["nmi_col"]].apply(
            lambda value: str(int(float(value))) if pd.notna(value) and str(value) not in ("nan", "") else None
        )
    else:
        df["NMI"] = df[fmt["nmi_col"]].astype(str).str.replace(r"\.0$", "", regex=True)
        if fmt.get("nmi_strip_after"):
            df["NMI"] = df["NMI"].str.split(fmt["nmi_strip_after"]).str[0]

    df["dt"] = _parse_dt(df, fmt)
    if fmt.get("datetime_is_end"):
        df["dt"] -= pd.Timedelta(minutes=fmt.get("interval_min", 30))

    if fmt.get("long_format"):
        unit_col = fmt["unit_col"]
        val_col = fmt["value_col"]
        dir_col = fmt["direction_col"]
        kwh_rows, _ = _filter_supported_energy_rows(df, unit_col, val_col)
        consume = (
            kwh_rows[kwh_rows[dir_col] == fmt["import_direction"]]
            .groupby(["NMI", "dt"])[val_col]
            .sum()
            .rename("consumption_kWh")
            .reset_index()
        )
        export = (
            kwh_rows[kwh_rows[dir_col] == fmt["export_direction"]]
            .groupby(["NMI", "dt"])[val_col]
            .sum()
            .rename("export_kWh")
            .reset_index()
        )
        df = consume.merge(export, on=["NMI", "dt"], how="outer").fillna(0)
        return df[["NMI", "dt", "consumption_kWh", "export_kWh"]]

    if fmt.get("channel_col"):
        ch_col = fmt["channel_col"]
        val_col = fmt["value_col"]
        imp_ch = fmt["import_channel"]
        exp_ch = fmt["export_channel"]
        unit_col = fmt.get("unit_col")
        if unit_col:
            unit_filter = fmt.get("unit_filter", "KWH")
            if unit_filter.upper() in SUPPORTED_ENERGY_UNITS:
                df, _ = _filter_supported_energy_rows(df, unit_col, val_col)
            else:
                df = df[df[unit_col].astype(str).str.upper() == unit_filter.upper()].copy()
                df[val_col] = pd.to_numeric(df[val_col], errors="coerce").fillna(0)
        else:
            df[val_col] = pd.to_numeric(df[val_col], errors="coerce").fillna(0)
        consume = (
            df[df[ch_col] == imp_ch]
            .groupby(["NMI", "dt"])[val_col]
            .sum()
            .rename("consumption_kWh")
            .reset_index()
        )
        export = (
            df[df[ch_col] == exp_ch]
            .groupby(["NMI", "dt"])[val_col]
            .sum()
            .rename("export_kWh")
            .reset_index()
        )
        df = consume.merge(export, on=["NMI", "dt"], how="outer").fillna(0)
        if fmt.get("interval_min", 30) < 30:
            df = _resample_30min(df)
        return df[["NMI", "dt", "consumption_kWh", "export_kWh"]]

    if fmt.get("register_mode"):
        reg_col = fmt["register_col"]
        energy_col = fmt["energy_col"]
        df["_e"] = pd.to_numeric(df[energy_col], errors="coerce").fillna(0)
        is_export = df[reg_col].str.startswith(fmt.get("export_prefix", "B"))
        consume = (
            df[~is_export]
            .groupby(["NMI", "dt"])["_e"]
            .sum()
            .rename("consumption_kWh")
            .reset_index()
        )
        export = (
            df[is_export]
            .groupby(["NMI", "dt"])["_e"]
            .sum()
            .rename("export_kWh")
            .reset_index()
        )
        df = consume.merge(export, on=["NMI", "dt"], how="outer").fillna(0)
    elif fmt.get("watts_col"):
        interval = fmt.get("interval_min", 30)
        divisor = 1000 if fmt.get("watts_unit", "kW") == "W" else 1
        df["consumption_kWh"] = (
            pd.to_numeric(df[fmt["watts_col"]], errors="coerce").fillna(0) * (interval / 60) / divisor
        )
        df["export_kWh"] = 0.0
        df = df[["NMI", "dt", "consumption_kWh", "export_kWh"]]
    else:
        con_col = fmt.get("consumption_col")
        exp_col = fmt.get("export_col")
        df["consumption_kWh"] = pd.to_numeric(df[con_col], errors="coerce").fillna(0) if con_col else 0.0
        df["export_kWh"] = pd.to_numeric(df[exp_col], errors="coerce").fillna(0) if exp_col else 0.0
        df = df[["NMI", "dt", "consumption_kWh", "export_kWh"]]

    if fmt.get("interval_min", 30) < 30:
        df = _resample_30min(df)

    return df[["NMI", "dt", "consumption_kWh", "export_kWh"]]


def _avg_profile(day_df: pd.DataFrame, is_bd: bool) -> pd.DataFrame:
    mask = (day_df["dt"].dt.weekday < 5) == is_bd
    subset = day_df[mask].copy()
    if subset.empty:
        return pd.DataFrame({"consumption_kWh": 0.0, "export_kWh": 0.0}, index=range(48))
    subset["slot"] = subset["dt"].dt.hour * 2 + subset["dt"].dt.minute // 30
    return (
        subset.groupby("slot")[["consumption_kWh", "export_kWh"]]
        .mean()
        .reindex(range(48), fill_value=0.0)
    )


def _generate_estimated_month(
    nmi: str,
    period: "pd.Period",
    bd_prof: pd.DataFrame,
    nbd_prof: pd.DataFrame,
) -> pd.DataFrame:
    rows = []
    for day in pd.date_range(period.to_timestamp(), periods=period.days_in_month, freq="D"):
        profile = bd_prof if day.weekday() < 5 else nbd_prof
        for slot in range(48):
            rows.append(
                {
                    "NMI": nmi,
                    "dt": day + pd.Timedelta(minutes=30 * slot),
                    "consumption_kWh": round(float(profile.loc[slot, "consumption_kWh"]), 5),
                    "export_kWh": round(float(profile.loc[slot, "export_kWh"]), 5),
                    "notes": "estimate",
                }
            )
    return pd.DataFrame(rows)


def fill_missing_months(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["dt"] = pd.to_datetime(df["dt"])
    estimated_frames = []

    for nmi, nmi_df in df.groupby("NMI"):
        nmi_df = nmi_df.dropna(subset=["dt"])
        if nmi_df.empty:
            continue
        months_present = set(nmi_df["dt"].dt.to_period("M").unique())
        max_month = max(months_present)
        min_target = max_month - 11
        target_range = set(pd.period_range(min_target, max_month, freq="M"))
        missing = sorted(target_range - months_present)

        if not missing:
            continue

        runs, run = [], [missing[0]]
        for month in missing[1:]:
            if month == run[-1] + 1:
                run.append(month)
            else:
                runs.append(run)
                run = [month]
        runs.append(run)

        for missing_run in runs:
            if len(missing_run) <= 3:
                before = [month for month in months_present if month < missing_run[0]]
                after = [month for month in months_present if month > missing_run[-1]]
                if before and after:
                    before_month, after_month = max(before), min(after)
                    nearest = (
                        before_month
                        if (missing_run[0].ordinal - before_month.ordinal)
                        <= (after_month.ordinal - missing_run[-1].ordinal)
                        else after_month
                    )
                elif before:
                    nearest = max(before)
                else:
                    nearest = min(after)
                template = nmi_df[nmi_df["dt"].dt.to_period("M") == nearest]
            else:
                template = nmi_df

            bd_prof = _avg_profile(template, is_bd=True)
            nbd_prof = _avg_profile(template, is_bd=False)

            for month in missing_run:
                estimated_frames.append(_generate_estimated_month(nmi, month, bd_prof, nbd_prof))

    if estimated_frames:
        df = pd.concat([df] + estimated_frames, ignore_index=True)

    return df.sort_values(["NMI", "dt"]).reset_index(drop=True)


def _fmt_datetime(dt_series: pd.Series) -> pd.Series:
    return (
        dt_series.dt.day.astype(str)
        + "/"
        + dt_series.dt.month.astype(str).str.zfill(2)
        + "/"
        + dt_series.dt.year.astype(str)
        + " "
        + dt_series.dt.hour.astype(str)
        + ":"
        + dt_series.dt.minute.astype(str).str.zfill(2)
    )


def _nmi_preview(df: pd.DataFrame) -> str:
    preview = ", ".join(sorted(df["NMI"].astype(str).unique())[:3])
    if df["NMI"].nunique() > 3:
        preview += " ..."
    return preview


def _build_output(frames: list[pd.DataFrame]) -> tuple[pd.DataFrame, list[str]]:
    combined = pd.concat(frames, ignore_index=True)
    combined["dt"] = pd.to_datetime(combined["dt"])
    if "notes" not in combined.columns:
        combined["notes"] = "actual"
    combined["notes"] = combined["notes"].fillna("actual")

    combined = combined.dropna(subset=["dt"])
    combined = fill_missing_months(combined)

    combined["net_kWh"] = (combined["consumption_kWh"] - combined["export_kWh"].abs()).round(3)
    combined["CE_kWh"] = combined["consumption_kWh"].round(3)
    combined["SOE_kWh"] = combined["export_kWh"].abs().round(3)
    combined["Date_time"] = _fmt_datetime(combined["dt"])
    combined["the_date"] = combined["dt"].dt.strftime("%Y-%m-%d")
    combined["the_interval"] = combined["dt"].dt.hour * 2 + combined["dt"].dt.minute // 30 + 1

    output = (
        combined.sort_values(["NMI", "dt"])[
            ["NMI", "Date_time", "the_date", "the_interval", "net_kWh", "CE_kWh", "SOE_kWh", "notes"]
        ]
        .reset_index(drop=True)
    )

    actual_rows = output[output["notes"] == "actual"]
    warnings = []
    nmi_totals = actual_rows.groupby("NMI")[["net_kWh", "CE_kWh"]].sum()
    zero_ce = nmi_totals[nmi_totals["CE_kWh"] == 0].index.tolist()
    neg_net = nmi_totals[nmi_totals["net_kWh"] < 0].index.tolist()
    if zero_ce:
        warnings.append(f"CE_kWh total = 0 (check mapping): {', '.join(zero_ce)}")
    if neg_net:
        warnings.append(f"net_kWh total < 0 (excess solar or mapping issue): {', '.join(neg_net)}")

    return output, warnings


def _iter_supported_files(input_dir: Path) -> list[Path]:
    return sorted(path for path in input_dir.iterdir() if path.is_file() and path.suffix.lower() in SUPPORTED_EXT)


def _process_candidates(candidates: list[Path], registry: list[dict]) -> tuple[list[pd.DataFrame], list[ProcessedFileResult], list[FailedFileResult]]:
    frames: list[pd.DataFrame] = []
    processed_files: list[ProcessedFileResult] = []
    failed_files: list[FailedFileResult] = []

    for file_path in candidates:
        if file_path.suffix.lower() == ".csv" and is_nem12(file_path):
            try:
                df = parse_nem12(file_path)
                if df.empty or df["NMI"].isna().all():
                    failed_files.append(
                        FailedFileResult(file_path.name, "no_valid_nmi", "No valid NMI data in NEM12")
                    )
                    continue
                frames.append(df)
                processed_files.append(
                    ProcessedFileResult(
                        filename=file_path.name,
                        rows=len(df),
                        source_type="NEM12",
                        format_id="NEM12",
                        nmi_preview=_nmi_preview(df),
                        aggregated_to_30min=False,
                    )
                )
            except Exception as exc:
                failed_files.append(FailedFileResult(file_path.name, "nem12_error", f"NEM12 processing error: {exc}"))
            continue

        fmt = detect_format(file_path, registry)
        if fmt is None:
            failed_files.append(
                FailedFileResult(file_path.name, "unrecognised_format", "No matching format in registry")
            )
            continue

        try:
            df = normalise(file_path, fmt)
            df = df[df["NMI"].notna() & (df["NMI"] != "UNKNOWN")]
            if df.empty:
                failed_files.append(
                    FailedFileResult(file_path.name, "no_valid_nmi", "No valid NMI data extracted")
                )
                continue

            frames.append(df)
            processed_files.append(
                ProcessedFileResult(
                    filename=file_path.name,
                    rows=len(df),
                    source_type="registry",
                    format_id=fmt["id"],
                    nmi_preview=_nmi_preview(df),
                    aggregated_to_30min=fmt.get("interval_min", 30) < 30,
                )
            )
        except Exception as exc:
            failed_files.append(FailedFileResult(file_path.name, "processing_error", f"Processing error: {exc}"))

    return frames, processed_files, failed_files


def process_input_directory(
    input_dir: Path = INPUT_DIR,
    output_dir: Path = OUTPUT_DIR,
    registry_path: Path = REGISTRY_PATH,
    archive_processed: bool = True,
    move_failed_to_unprocessed: bool = True,
) -> ProcessResult:
    run_ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_dir.mkdir(parents=True, exist_ok=True)
    registry = load_registry(registry_path)

    all_files = _iter_supported_files(input_dir)
    excluded = [path for path in all_files if any(pattern in path.name for pattern in EXCLUDED_FILENAME_PATTERNS)]
    candidates = [path for path in all_files if path not in excluded]

    failed_files = [
        FailedFileResult(path.name, "excluded", "Excluded by filename pattern")
        for path in excluded
    ]

    frames, processed_files, candidate_failures = _process_candidates(candidates, registry)
    failed_files.extend(candidate_failures)

    output_csv_bytes = None
    output_filename = None
    total_rows = 0
    unique_nmis = 0
    warnings: list[str] = []

    if frames:
        output_df, warnings = _build_output(frames)
        total_rows = len(output_df)
        unique_nmis = output_df["NMI"].nunique()
        output_filename = f"clean_output_{run_ts}.csv"
        output_path = output_dir / output_filename
        output_df.to_csv(output_path, index=False)
        output_csv_bytes = output_path.read_bytes()

    if archive_processed and processed_files:
        proc_dir = input_dir / "processed" / run_ts
        proc_dir.mkdir(parents=True, exist_ok=True)
        for item in processed_files:
            source = input_dir / item.filename
            if source.exists():
                shutil.move(str(source), str(proc_dir / item.filename))

    if move_failed_to_unprocessed and failed_files:
        unproc_dir = input_dir / "unprocessed"
        unproc_dir.mkdir(parents=True, exist_ok=True)
        for item in failed_files:
            source = input_dir / item.filename
            if source.exists():
                shutil.move(str(source), str(unproc_dir / item.filename))

    return ProcessResult(
        output_csv_bytes=output_csv_bytes,
        output_filename=output_filename,
        processed_files=processed_files,
        failed_files=failed_files,
        warnings=warnings,
        total_rows=total_rows,
        unique_nmis=unique_nmis,
    )


def _write_uploaded_file(upload, destination: Path) -> None:
    if hasattr(upload, "getbuffer"):
        destination.write_bytes(bytes(upload.getbuffer()))
        return
    if hasattr(upload, "read"):
        content = upload.read()
        destination.write_bytes(content if isinstance(content, bytes) else bytes(content))
        return
    raise TypeError(f"Unsupported upload object for {destination.name}")


def process_uploaded_files(files: list[object], registry_path: Path = REGISTRY_PATH) -> ProcessResult:
    with TemporaryDirectory() as tmp_dir_name:
        tmp_dir = Path(tmp_dir_name)
        input_dir = tmp_dir / "Input"
        output_dir = tmp_dir / "Output"
        input_dir.mkdir(parents=True, exist_ok=True)
        output_dir.mkdir(parents=True, exist_ok=True)

        for upload in files:
            filename = Path(getattr(upload, "name", "upload.bin")).name
            _write_uploaded_file(upload, input_dir / filename)

        return process_input_directory(
            input_dir=input_dir,
            output_dir=output_dir,
            registry_path=registry_path,
            archive_processed=False,
            move_failed_to_unprocessed=False,
        )


def process_uploaded_file_buffers(files: list[tuple[str, bytes]], registry_path: Path = REGISTRY_PATH) -> ProcessResult:
    uploads = []
    for name, content in files:
        uploads.append(type("Upload", (), {"name": name, "getbuffer": lambda self, data=content: memoryview(data)})())
    return process_uploaded_files(uploads, registry_path=registry_path)


def _print_local_summary(result: ProcessResult) -> None:
    if not result.processed_files and not result.failed_files:
        print("No files found in Input/  -- nothing to do.")
        return

    print(f"Files processed : {len(result.processed_files)}")
    print(f"Files failed    : {len(result.failed_files)}")

    for item in result.processed_files:
        agg_note = "  (aggregated to 30-min)" if item.aggregated_to_30min else ""
        print(f"  {item.filename:<40}  {item.rows:>10,} rows  [{item.format_id}]{agg_note}")
        print(f"  {'':40}  NMI: {item.nmi_preview}")

    for item in result.failed_files:
        print(f"  {item.filename:<40}  *** {item.status.upper()} -- {item.reason}")

    if result.output_filename:
        print(
            f"\nTotal rows  : {result.total_rows:,}\n"
            f"Unique NMIs : {result.unique_nmis:,}\n"
            f"Output      : {OUTPUT_DIR / result.output_filename}"
        )
    else:
        print("\nNo data to output.")

    if result.warnings:
        print("\n*** Data quality warnings:")
        for warning in result.warnings:
            print(f"  {warning}")


def result_to_dict(result: ProcessResult) -> dict:
    return {
        "output_csv_bytes": result.output_csv_bytes,
        "output_filename": result.output_filename,
        "processed_files": [asdict(item) for item in result.processed_files],
        "failed_files": [asdict(item) for item in result.failed_files],
        "warnings": result.warnings,
        "total_rows": result.total_rows,
        "unique_nmis": result.unique_nmis,
    }


def main() -> None:
    try:
        result = process_input_directory()
        _print_local_summary(result)
    except Exception:
        traceback.print_exc()
        raise


if __name__ == "__main__":
    main()
