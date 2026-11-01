import os
import shutil
from pathlib import Path
from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")

TEMPLATES_DIR = BASE_DIR / "templates"
DOCUMENTS_DIR = BASE_DIR / "Haftalık Rapor Bot Documents"
OUTPUT_DIR = BASE_DIR / "Üretilen Raporlar"
MAPS_DIR = BASE_DIR / "Haritalar"
STATE_DIR = BASE_DIR / "data"

for d in (TEMPLATES_DIR, DOCUMENTS_DIR, OUTPUT_DIR, MAPS_DIR, STATE_DIR):
    d.mkdir(parents=True, exist_ok=True)

EXCEL_PATH = DOCUMENTS_DIR / "mevsim_normalleri.xlsx"
FALLBACK_EXCEL_PATH = TEMPLATES_DIR / "mevsim_normalleri.xlsx"

# Bulut ortamında ilk kez ayağa kalktığında şablon klasöründeki Excel'i Documents'a kopyala
if not EXCEL_PATH.exists() and FALLBACK_EXCEL_PATH.exists():
    shutil.copy2(FALLBACK_EXCEL_PATH, EXCEL_PATH)
elif EXCEL_PATH.exists() and not FALLBACK_EXCEL_PATH.exists():
    shutil.copy2(EXCEL_PATH, FALLBACK_EXCEL_PATH)

TEMPLATE_CANDIDATES = [
    TEMPLATES_DIR / "tarsim_master_template.docx",
    DOCUMENTS_DIR / "02-08.03.2026 Meteorolojik Rapor.docx",
    DOCUMENTS_DIR / "18-24.05.2026 Meteorolojik Rapor.docx",
]


def get_template_docx() -> Path:
    for p in TEMPLATE_CANDIDATES:
        if p.exists():
            return p
    for folder in (TEMPLATES_DIR, DOCUMENTS_DIR):
        docx_files = sorted(folder.glob("*.docx"))
        if docx_files:
            return docx_files[0]
    raise FileNotFoundError(f"Şablon .docx dosyası bulunamadı: {TEMPLATES_DIR} / {DOCUMENTS_DIR}")


TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "").strip()
PORT = int(os.getenv("PORT", "7860"))
PUBLIC_URL = os.getenv("PUBLIC_URL", os.getenv("RENDER_EXTERNAL_URL", "")).strip()

MONTH_NAMES_TR = {
    1: "Ocak",
    2: "Şubat",
    3: "Mart",
    4: "Nisan",
    5: "Mayıs",
    6: "Haziran",
    7: "Temmuz",
    8: "Ağustos",
    9: "Eylül",
    10: "Ekim",
    11: "Kasım",
    12: "Aralık",
}

DAY_NAMES_TR = {
    0: "Pazartesi",
    1: "Salı",
    2: "Çarşamba",
    3: "Perşembe",
    4: "Cuma",
    5: "Cumartesi",
    6: "Pazar",
}

REGION_ORDER = [
    ("Akdeniz", "Akdeniz Bölgesi"),
    ("Doğu Anadolu", "Doğu Anadolu Bölgesi"),
    ("Ege", "Ege Bölgesi"),
    ("Güneydoğu Anadolu", "Güneydoğu Anadolu Bölgesi"),
    ("İç Anadolu", "İç Anadolu Bölgesi"),
    ("Karadeniz", "Karadeniz Bölgesi"),
    ("Marmara", "Marmara Bölgesi"),
]
