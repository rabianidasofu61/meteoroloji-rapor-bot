from pathlib import Path
from typing import Dict, List, Tuple, Optional
import openpyxl
from config import EXCEL_PATH, MONTH_NAMES_TR, REGION_ORDER


def _format_num(val, use_comma: bool = False) -> str:
    if val is None:
        return ""
    s = str(val).strip()
    if not s:
        return ""
    try:
        num = float(s.replace(",", "."))
        formatted = f"{num:.1f}"
        return formatted.replace(".", ",") if use_comma else formatted
    except ValueError:
        return s


def load_seasonal_normals(
    month: int, excel_path: Optional[Path] = None, use_comma: bool = False
) -> List[Dict]:
    """
    mevsim_normalleri.xlsx dosyasından belirtilen ay (1..12) için
    7 bölgenin il bazlı Ortalama Sıcaklık (°C) ve Aylık Toplam Yağış (mm)
    verilerini okur.

    Dönen yapı:
    [
        {
            "table_no": 1,
            "sheet_name": "Akdeniz",
            "region_title": "Akdeniz Bölgesi",
            "caption": "Tablo 1. Akdeniz Bölgesi Ortalama Sıcaklıkları (°C) ve Aylık Toplam Yağış Miktarı Ortalamaları (mm)",
            "rows": [
                ("Adana", "29.2", "10.4"),
                ...
            ]
        },
        ...
    ]
    """
    path = excel_path or EXCEL_PATH
    if not path.exists():
        raise FileNotFoundError(f"Mevsim normalleri Excel dosyası bulunamadı: {path}")

    month_name = MONTH_NAMES_TR[month]
    wb = openpyxl.load_workbook(path, data_only=True)
    result = []

    for idx, (sheet_key, region_title) in enumerate(REGION_ORDER, start=1):
        # Eşleşen sayfayı bul
        sheet_name = None
        for s in wb.sheetnames:
            if s.strip().lower() == sheet_key.lower():
                sheet_name = s
                break
        if not sheet_name:
            continue

        ws = wb[sheet_name]
        all_rows = list(ws.iter_rows(values_only=True))

        # Tablo 1: Ortalama Sıcaklık (91-2020)
        # Tablo 2: Aylık Toplam Yağış Miktarı Ort 1991-2020
        mode = None  # "temp" veya "precip"
        month_col_idx = None
        temps: Dict[str, str] = {}
        precips: Dict[str, str] = {}
        provinces_order: List[str] = []

        for row in all_rows:
            if not row or all(c is None or str(c).strip() == "" for c in row):
                continue
            first_cell = str(row[0]).strip() if row[0] is not None else ""

            if "Ortalama Sıcaklık" in first_cell:
                mode = "temp"
                month_col_idx = None
                continue
            elif "Toplam Yağış" in first_cell or "Yağış Miktarı" in first_cell:
                mode = "precip"
                month_col_idx = None
                continue

            if first_cell == "İl Adı":
                # Ay sütununu bul
                for c_i, cell_val in enumerate(row):
                    if cell_val is not None and str(cell_val).strip().lower() == month_name.lower():
                        month_col_idx = c_i
                        break
                if month_col_idx is None:
                    # Standart sıra: İl Adı(0), Ocak(1)..Aralık(12)
                    month_col_idx = month
                continue

            if mode in ("temp", "precip") and first_cell:
                col = month_col_idx if month_col_idx is not None else month
                val = row[col] if col < len(row) else None
                formatted_val = _format_num(val, use_comma=use_comma)
                if mode == "temp":
                    if first_cell not in provinces_order:
                        provinces_order.append(first_cell)
                    temps[first_cell] = formatted_val
                elif mode == "precip":
                    if first_cell not in provinces_order:
                        provinces_order.append(first_cell)
                    precips[first_cell] = formatted_val

        table_rows: List[Tuple[str, str, str]] = []
        for prov in provinces_order:
            t_val = temps.get(prov, "")
            p_val = precips.get(prov, "")
            table_rows.append((prov, t_val, p_val))

        result.append(
            {
                "table_no": idx,
                "sheet_name": sheet_key,
                "region_title": region_title,
                "caption": f"Tablo {idx}. {region_title} Ortalama Sıcaklıkları (°C) ve Aylık Toplam Yağış Miktarı Ortalamaları (mm)",
                "rows": table_rows,
            }
        )

    return result


def get_national_normal_temp(month: int, excel_path: Optional[Path] = None) -> float:
    """Belirtilen ay için Türkiye genelindeki illerin mevsim normali ortalama sıcaklığını döndürür."""
    tables = load_seasonal_normals(month, excel_path=excel_path)
    vals = []
    for t in tables:
        for _, temp_s, _ in t["rows"]:
            if temp_s:
                try:
                    vals.append(float(temp_s.replace(",", ".")))
                except ValueError:
                    pass
    return sum(vals) / len(vals) if vals else 15.0
