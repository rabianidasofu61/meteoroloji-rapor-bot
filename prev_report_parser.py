import re
import zipfile
import xml.etree.ElementTree as ET
from datetime import date, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from config import DOCUMENTS_DIR, OUTPUT_DIR, MONTH_NAMES_TR


PAST_TENSE_REPLACEMENTS = [
    (r"\btahmin edilmektedir\b", "tahmin edilmekteydi"),
    (r"\btahmin edilmemektedir\b", "tahmin edilmemekteydi"),
    (r"\bbeklenmektedir\b", "beklenmekteydi"),
    (r"\bbeklenmemektedir\b", "beklenmemekteydi"),
    (r"\bön görülmektedir\b", "ön görülmekteydi"),
    (r"\böngörülmektedir\b", "öngörülmekteydi"),
    (r"\bön görülmemektedir\b", "ön görülmemekteydi"),
    (r"\böngörülmemektedir\b", "öngörülmemekteydi"),
    (r"\bbulunmaktadır\b", "bulunmaktaydı"),
    (r"\bbulunmamaktadır\b", "bulunmamaktaydı"),
    (r"\bdüşünülmektedir\b", "düşünülmekteydi"),
    (r"\bgörülmektedir\b", "görülmekteydi"),
    (r"\bgörülmemektedir\b", "görülmemekteydi"),
    (r"\bdevam etmektedir\b", "devam etmekteydi"),
    (r"\bgerçekleşmektedir\b", "gerçekleşmekteydi"),
    (r"\bşu şekildedir\s*:", "şu şekildeydi:"),
    (r"\başağıdaki gibidir\s*:", "aşağıdaki gibiydi:"),
]


def convert_to_past_tense(text: str) -> str:
    """
    'tahmin edilmektedir', 'beklenmektedir', 'ön görülmektedir' vb.
    ifadeleri geçmiş zaman kipine ('tahmin edilmekteydi' vb.) dönüştürür.
    """
    if not text:
        return text
    result = text
    for pattern, repl in PAST_TENSE_REPLACEMENTS:
        result = re.sub(pattern, repl, result, flags=re.IGNORECASE)
    return result


def format_week_range_tr(start_date: date, end_date: date, include_year: bool = True) -> str:
    """
    Örn:
    Aynı ay: '24 – 30 Ağustos 2026'
    Farklı ay: '23 Şubat – 01 Mart 2026'
    """
    d1 = f"{start_date.day:02d}"
    d2 = f"{end_date.day:02d}"
    m1 = MONTH_NAMES_TR[start_date.month]
    m2 = MONTH_NAMES_TR[end_date.month]
    yr = f" {end_date.year}" if include_year else ""
    if start_date.month == end_date.month and start_date.year == end_date.year:
        return f"{d1} – {d2} {m2}{yr}"
    elif start_date.year == end_date.year:
        return f"{d1} {m1} – {d2} {m2}{yr}"
    else:
        return f"{d1} {m1} {start_date.year} – {d2} {m2} {end_date.year}"


def find_previous_week_docx(prev_monday: date) -> Optional[Path]:
    """
    Önceki haftanın Pazartesi-Pazar tarih aralığına uyan .docx dosyasını
    Üretilen Raporlar veya Haftalık Rapor Bot Documents klasörlerinde arar.
    """
    prev_sunday = prev_monday + timedelta(days=6)
    candidates: List[Path] = []
    for folder in (OUTPUT_DIR, DOCUMENTS_DIR):
        if folder.exists():
            candidates.extend(folder.glob("*.docx"))

    # Dosya adında önceki haftanın gün/ay bilgisi geçenleri önceliklendir
    d1_str = f"{prev_monday.day:02d}"
    d2_str = f"{prev_sunday.day:02d}"
    m_str = f"{prev_sunday.month:02d}"
    y_str = f"{prev_sunday.year}"

    for p in candidates:
        if p.name.startswith("~$"):
            continue
        if d1_str in p.name and d2_str in p.name and y_str in p.name:
            return p

    return None


def extract_section2_from_docx(docx_path: Path) -> Optional[Dict]:
    """
    Verilen .docx raporunun 2. Bölümünü ('2) ... Haftalık Meteorolojik Değerlendirme')
    okur, özet paragrafını ve günlük maddelerini geçmiş zamana çevirerek döndürür.
    """
    if not docx_path or not docx_path.exists():
        return None

    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    try:
        with zipfile.ZipFile(docx_path) as z:
            doc = ET.fromstring(z.read("word/document.xml"))
    except Exception:
        return None

    body = doc.find("w:body", ns)
    if body is None:
        return None

    in_sec2 = False
    intro_text = ""
    daily_bullets: List[Tuple[str, str]] = []

    for child in body:
        tag = child.tag.split("}")[-1]
        if tag != "p":
            continue
        text = "".join(child.itertext()).strip()
        if not text:
            continue

        # 2. Bölüm başlığı (İçindekiler'den sonraki gerçek 2. Bölüm başlığı)
        if re.match(r"^2\)\s+", text) and "Değerlendirme" in text:
            # İçindekiler satırı mı yoksa ana başlık mı kontrol et:
            # İçindekiler'de hemen ardından '3) ... Mevsim Normalleri' gelir,
            # gerçek 2. Bölümde ise açıklama paragrafı gelir.
            in_sec2 = True
            intro_text = ""
            daily_bullets = []
            continue

        if in_sec2:
            if re.match(r"^3\)\s+", text):
                # Eğer henüz intro_text yoksa bu İçindekiler'deki 2) ve 3) satırıdır
                if not intro_text and not daily_bullets:
                    in_sec2 = False
                    continue
                else:
                    break

            # Şekil yazısı ise atla
            if re.match(r"^Şekil\s*\d+", text, flags=re.IGNORECASE):
                continue

            if not intro_text and ":" in text and ("tahmin" in text.lower() or "haftasında" in text.lower()):
                intro_text = convert_to_past_tense(text)
            elif ":" in text:
                parts = text.split(":", 1)
                date_label = parts[0].strip()
                desc = convert_to_past_tense(parts[1].strip())
                daily_bullets.append((date_label, desc))

    if not intro_text and not daily_bullets:
        return None

    return {
        "source_file": docx_path.name,
        "intro": intro_text,
        "bullets": daily_bullets,
    }
