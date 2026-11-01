import asyncio
import os
import subprocess
import sys
import threading
from datetime import date, datetime, timedelta

from config import TELEGRAM_BOT_TOKEN, EXCEL_PATH
from weather_analyzer import build_full_report_draft
from windy_scraper import capture_windy_maps_sync
from docx_generator import generate_weekly_report_docx
from telegram_bot import (
    start_cmd,
    rapor_cmd,
    callback_handler,
    document_handler,
    text_handler,
    _parse_user_date,
    _get_monday,
)
from telegram.ext import (
    Application,
    CommandHandler,
    CallbackQueryHandler,
    MessageHandler,
    filters,
)


def ensure_playwright_chromium():
    """Bulut (Gradio / GitHub Actions) ortamında Chromium yüklü değilse otomatik kurar."""
    try:
        subprocess.run(
            [sys.executable, "-m", "playwright", "install", "chromium"],
            check=False,
        )
    except Exception as e:
        print("Playwright kurulum uyarısı:", e)


def _run_telegram_bot_thread():
    """Arka planda 7/24 Telegram Botunu çalıştırır."""
    if not TELEGRAM_BOT_TOKEN:
        print("UYARI: TELEGRAM_BOT_TOKEN henüz tanımlanmamış!")
        return

    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)

    async def _runner():
        tg_app = Application.builder().token(TELEGRAM_BOT_TOKEN).build()
        tg_app.add_handler(CommandHandler("start", start_cmd))
        tg_app.add_handler(CommandHandler("menu", start_cmd))
        tg_app.add_handler(CommandHandler("yardim", start_cmd))
        tg_app.add_handler(CommandHandler("rapor", rapor_cmd))
        tg_app.add_handler(CallbackQueryHandler(callback_handler))
        tg_app.add_handler(MessageHandler(filters.Document.ALL, document_handler))
        tg_app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, text_handler))

        await tg_app.initialize()
        await tg_app.start()
        await tg_app.updater.start_polling(drop_pending_updates=True)
        print("✅ Telegram Botu arka planda 7/24 aktif!")
        while True:
            await asyncio.sleep(3600)

    loop.run_until_complete(_runner())


def web_generate_report(tarih_str: str, mgm_cumlesi: str, windy_cek: bool):
    """Web arayüzünden de tek tıkla rapor üretilebilmesini sağlar."""
    if tarih_str and tarih_str.strip():
        target_mon = _parse_user_date(tarih_str.strip())
        if not target_mon:
            return None, f"❌ Hata: '{tarih_str}' geçerli bir tarih değil. Örn: 24.08.2026"
    else:
        target_mon = _get_monday(date.today()) + timedelta(days=7)

    draft = build_full_report_draft(
        target_mon,
        custom_mgm_sentence=mgm_cumlesi.strip() if mgm_cumlesi and mgm_cumlesi.strip() else None,
    )
    if windy_cek:
        capture_windy_maps_sync(draft["temp_maps"], draft.get("sec2_extra_maps", []))

    out_path = generate_weekly_report_docx(draft, save_to_documents=True)
    return str(out_path), f"✅ Rapor başarıyla oluşturuldu: {out_path.name}"


if __name__ == "__main__":
    ensure_playwright_chromium()

    # Telegram botunu arka plan thread'inde başlat
    bot_thread = threading.Thread(target=_run_telegram_bot_thread, daemon=True)
    bot_thread.start()

    # Eğer Gradio yüklüyse Gradio arayüzünü başlat, değilse doğrudan telegram_bot.py çalıştır
    try:
        import gradio as gr

        token_status = (
            "✅ TELEGRAM_BOT_TOKEN Aktif — Telegram üzerinden 7/24 kullanabilirsiniz!"
            if TELEGRAM_BOT_TOKEN
            else "⚠️ TELEGRAM_BOT_TOKEN Eksik — Settings -> Variables and secrets kısmından ekleyin."
        )

        with gr.Blocks(title="TARSİM Haftalık Meteorolojik Rapor Botu") as demo:
            gr.Markdown(
                f"## 🌦️ TARSİM Haftalık Meteorolojik Değerlendirme Rapor Botu\n"
                f"**Durum:** {token_status}\n\n"
                f"Raporlarınızı **Telegram botunuz üzerinden** veya doğrudan aşağıdaki panelden oluşturabilirsiniz."
            )
            with gr.Row():
                tarih_inp = gr.Textbox(
                    label="Rapor Haftası Tarihi (Örn: 24.08.2026 — Boş bırakılırsa gelecek hafta)",
                    placeholder="24.08.2026",
                )
                windy_chk = gr.Checkbox(label="Windy Haritalarını Otomatik Çek", value=True)
            mgm_inp = gr.Textbox(
                label="2. Bölüm MGM Sıcaklık Cümlesi (İsteğe Bağlı)",
                placeholder="Hava sıcaklıklarının ülkemiz genelinde mevsim normalleri civarında seyredeceği",
            )
            btn = gr.Button("📄 Word (.docx) Raporunu Oluştur")
            status_out = gr.Textbox(label="İşlem Durumu")
            file_out = gr.File(label="Üretilen Word (.docx) Raporu")

            btn.click(
                fn=web_generate_report,
                inputs=[tarih_inp, mgm_inp, windy_chk],
                outputs=[file_out, status_out],
            )

        demo.launch(server_name="0.0.0.0", server_port=int(os.getenv("PORT", "7860")))
    except ImportError:
        from telegram_bot import main_async
        asyncio.run(main_async())
