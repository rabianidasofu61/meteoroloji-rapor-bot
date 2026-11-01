import argparse
import sys
from datetime import date, timedelta

from weather_analyzer import build_full_report_draft
from windy_scraper import capture_windy_maps_sync
from docx_generator import generate_weekly_report_docx
from telegram_bot import _parse_user_date, _get_monday

sys.stdout.reconfigure(encoding="utf-8")


def main():
    parser = argparse.ArgumentParser(description="TARSİM Haftalık Meteorolojik Rapor Üretici")
    parser.add_argument(
        "--tarih",
        type=str,
        default="",
        help="Rapor haftası Pazartesi tarihi (Örn: 24.08.2026 veya 2026-08-24). Boş bırakılırsa gelecek Pazartesi seçilir.",
    )
    parser.add_argument(
        "--haritasiz",
        action="store_true",
        help="Windy ekran görüntüsü çekmeden hızlıca .docx oluşturur.",
    )
    parser.add_argument(
        "--mgm",
        type=str,
        default="",
        help="2. Bölüm MGM haftalık sıcaklık cümlesi.",
    )
    args = parser.parse_args()

    if args.tarih:
        target_mon = _parse_user_date(args.tarih)
        if not target_mon:
            print(f"Hata: '{args.tarih}' geçerli bir tarih değil. Örn: 24.08.2026")
            sys.exit(1)
    else:
        today = date.today()
        target_mon = _get_monday(today)

    print(f"[1/3] {target_mon.isoformat()} haftası için meteorolojik veriler ve mevsim normalleri analiz ediliyor...")
    draft = build_full_report_draft(
        target_mon,
        custom_mgm_sentence=args.mgm if args.mgm else None,
    )
    print(f"      1. Bölüm Kaynağı: {draft['sec1_source']}")
    print(f"      3. Bölüm Mevsim Normalleri Ayı: {draft['normals_month_name']}")

    if not args.haritasiz:
        print(f"[2/3] Windy.com üzerinden 8 adet 06:00 Sıcaklık Haritası (+ varsa yağış/CAPE haritası) çekiliyor...")
        capture_windy_maps_sync(draft["temp_maps"], draft.get("sec2_extra_maps", []))
    else:
        print("[2/3] Windy harita çekimi atlandı (--haritasiz).")

    print("[3/3] TARSİM şablonlu Word (.docx) raporu oluşturuluyor...")
    out_path = generate_weekly_report_docx(draft, save_to_documents=True)
    print(f"\n✅ Rapor başarıyla oluşturuldu:\n   -> {out_path}")


if __name__ == "__main__":
    main()
