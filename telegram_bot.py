import asyncio
import json
import logging
import os
import re
import shutil
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Dict, Optional

from aiohttp import web, ClientSession
from telegram import (
    Update,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    InputFile,
)
from telegram.ext import (
    Application,
    CommandHandler,
    CallbackQueryHandler,
    MessageHandler,
    ContextTypes,
    filters,
)

from config import (
    TELEGRAM_BOT_TOKEN,
    PORT,
    PUBLIC_URL,
    DOCUMENTS_DIR,
    OUTPUT_DIR,
    MAPS_DIR,
    STATE_DIR,
    EXCEL_PATH,
    FALLBACK_EXCEL_PATH,
    MONTH_NAMES_TR,
)
from prev_report_parser import extract_section2_from_docx, convert_to_past_tense
from weather_analyzer import build_full_report_draft
from windy_scraper import capture_windy_maps_async
from docx_generator import generate_weekly_report_docx

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)

USER_DRAFTS: Dict[int, Dict] = {}
USER_STATE: Dict[int, str] = {}
START_TIME = datetime.now()


def _get_monday(d: date) -> date:
    return d - timedelta(days=d.weekday())


def _parse_user_date(text: str) -> Optional[date]:
    text = text.strip()
    m = re.search(r"(\d{1,2})\s*[-–]\s*\d{1,2}\.(\d{1,2})\.(\d{4})", text)
    if m:
        d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        return _get_monday(date(y, mo, d))

    m = re.search(r"(\d{1,2})[./](\d{1,2})[./](\d{4})", text)
    if m:
        d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        return _get_monday(date(y, mo, d))

    m = re.search(r"(\d{4})-(\d{1,2})-(\d{1,2})", text)
    if m:
        y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
        return _get_monday(date(y, mo, d))

    return None


def _format_draft_preview(draft: Dict) -> str:
    lines = [
        f"📋 *{draft['curr_range_str']} Haftalık Meteorolojik Değerlendirme Taslağı*",
        f"📅 *Kapak Tarihi:* {draft['prep_date_str']}",
        f"📊 *3. Bölüm Mevsim Normalleri Ayı:* {MONTH_NAMES_TR[int(draft['normals_month'])]}",
        f"📂 *1. Bölüm Kaynağı:* `{draft.get('sec1_source', 'Otomatik')}`",
        "",
        f"🔹 *1) {draft['prev_range_str']} (Geçen Hafta):*",
        f"_{draft['sec1_intro']}_",
    ]
    for lbl, desc in draft["sec1_bullets"][:3]:
        lines.append(f"  • *{lbl}:* {desc[:90]}...")
    if len(draft["sec1_bullets"]) > 3:
        lines.append(f"  _... (+{len(draft['sec1_bullets']) - 3} gün daha)_")

    lines.append("")
    lines.append(f"🔹 *2) {draft['curr_range_str']} (Tahmin Haftası):*")
    lines.append(f"_{draft['sec2_intro']}_")
    for lbl, desc in draft["sec2_bullets"]:
        lines.append(f"  • *{lbl}:* {desc}")

    extra_maps = draft.get("sec2_extra_maps", [])
    if extra_maps:
        lines.append("  🗺️ *Eklenecek Yağış / CAPE Haritaları:*")
        for em in extra_maps:
            lines.append(f"    - `{em['caption_suffix']}`")

    lines.append("")
    lines.append("🔹 *4) Önemli Uyarılar (Rüzgâr & Fırtına):*")
    ww = draft.get("wind_warnings", {})
    for r_key in ("Marmara", "Ege", "Akdeniz", "Karadeniz"):
        w_list = ww.get(r_key, [])
        if not w_list:
            lines.append(f"  • *{r_key}:* Fırtına riski bulunmamaktadır.")
        else:
            lines.append(f"  • *{r_key}:* {'; '.join(w_list)}")

    lines.append("")
    lines.append("🔹 *5) 06:00 Sıcaklık Haritaları & Don Riski:*")
    lines.append(f"  • 8 adet 06:00 Windy Sıcaklık Haritası ({draft['temp_maps'][0]['label_pair']} .. {draft['temp_maps'][-1]['label_pair']})")
    lines.append(f"  • *Değerlendirme:* _{draft['frost_summary']}_")

    return "\n".join(lines)


def _get_draft_keyboard() -> InlineKeyboardMarkup:
    keyboard = [
        [
            InlineKeyboardButton(
                "📸 Windy Haritalarını Çek & Word (.docx) İndir",
                callback_data="act_generate_full",
            )
        ],
        [
            InlineKeyboardButton(
                "⚡ Hızlı Word (.docx) Oluştur (Mevcut Haritalarla)",
                callback_data="act_generate_fast",
            )
        ],
        [
            InlineKeyboardButton("🌡️ MGM Sıcaklık Cümlesi", callback_data="edit_mgm"),
            InlineKeyboardButton("🌦️ 2. Bölüm Günleri Düzenle", callback_data="edit_sec2"),
        ],
        [
            InlineKeyboardButton("⏪ 1. Bölüm (Geçen Hafta)", callback_data="edit_sec1"),
            InlineKeyboardButton("💨 4. Bölüm Fırtına", callback_data="edit_wind"),
        ],
        [
            InlineKeyboardButton("❄️ 5. Bölüm Don / Tablo 8", callback_data="edit_frost"),
            InlineKeyboardButton("📊 Normaller Ayını Değiştir", callback_data="edit_month"),
        ],
        [
            InlineKeyboardButton("🗺️ Yağış/CAPE Haritası Ekle", callback_data="add_extra_map"),
            InlineKeyboardButton("🏠 Ana Menü", callback_data="main_menu"),
        ],
    ]
    return InlineKeyboardMarkup(keyboard)


def _get_main_menu_keyboard() -> InlineKeyboardMarkup:
    today = date.today()
    this_mon = _get_monday(today)
    next_mon = this_mon + timedelta(days=7)
    keyboard = [
        [
            InlineKeyboardButton(
                f"📅 Bu Hafta ({this_mon.strftime('%d.%m')} - {(this_mon+timedelta(days=6)).strftime('%d.%m.%Y')})",
                callback_data=f"new_week_{this_mon.isoformat()}",
            )
        ],
        [
            InlineKeyboardButton(
                f"⏭️ Gelecek Hafta ({next_mon.strftime('%d.%m')} - {(next_mon+timedelta(days=6)).strftime('%d.%m.%Y')})",
                callback_data=f"new_week_{next_mon.isoformat()}",
            )
        ],
        [
            InlineKeyboardButton("🗓️ Özel Tarih Gir (Örn: 24.08.2026)", callback_data="ask_custom_date"),
        ],
        [
            InlineKeyboardButton("📂 Kayıtlı Raporları İndir", callback_data="list_reports"),
            InlineKeyboardButton("📊 Normaller Excel İndir", callback_data="dl_excel"),
        ],
    ]
    return InlineKeyboardMarkup(keyboard)


async def start_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    USER_STATE.pop(user_id, None)
    welcome = (
        "🌦️ *TARSİM Haftalık Meteorolojik Değerlendirme Rapor Botu (7/24 Bulut Destekli)*\n\n"
        "Bu bot ile bilgisayarınız kapalıyken bile telefondan veya her yerden:\n"
        "1️⃣ Geçen haftanın raporunu otomatik okuyup geçmiş zaman kipine (*tahmin edilmekteydi, beklenmekteydi*) çevirebilir,\n"
        "2️⃣ Yeni haftanın saatlik yağış, rüzgâr/fırtına ve don analizlerini otomatik oluşturabilir,\n"
        "3️⃣ *Windy.com* üzerinden Türkiye'yi kapsayan 8 adet sabah *06:00 Sıcaklık Haritasını* ve şiddetli yağış/CAPE haritalarını otomatik çekebilir,\n"
        "4️⃣ *mevsim_normalleri.xlsx* dosyasından ilgili ayın bölgesel tablolarını (Tablo 1-7) ekleyerek TARSİM formatında *Word (.docx)* çıktısı alabilirsiniz.\n\n"
        "💡 *Bulut İpucu:* Ay başında yeni `mevsim_normalleri.xlsx` dosyasını veya önceki haftanın `.docx` raporunu doğrudan bu Telegram sohbetine dosya olarak göndermeniz yeterlidir!\n\n"
        "Lütfen rapor hazırlamak istediğiniz haftayı seçin:"
    )
    await update.message.reply_text(welcome, parse_mode="Markdown", reply_markup=_get_main_menu_keyboard())


async def rapor_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    args = context.args
    if args:
        parsed = _parse_user_date(" ".join(args))
        if parsed:
            await _create_and_send_draft(update.message, user_id, parsed)
            return
    today = date.today()
    next_mon = _get_monday(today) + timedelta(days=7)
    await _create_and_send_draft(update.message, user_id, next_mon)


async def _create_and_send_draft(msg_obj, user_id: int, target_monday: date):
    status_msg = await msg_obj.reply_text(
        f"⏳ *{target_monday.strftime('%d.%m.%Y')}* haftası için meteorolojik veriler, "
        f"önceki hafta raporu ve mevsim normalleri analiz ediliyor...",
        parse_mode="Markdown",
    )
    loop = asyncio.get_running_loop()
    draft = await loop.run_in_executor(None, build_full_report_draft, target_monday)
    USER_DRAFTS[user_id] = draft
    USER_STATE.pop(user_id, None)

    preview_text = _format_draft_preview(draft)
    if len(preview_text) > 4000:
        preview_text = preview_text[:3950] + "\n..."

    await status_msg.edit_text(
        preview_text,
        parse_mode="Markdown",
        reply_markup=_get_draft_keyboard(),
    )


async def callback_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = query.from_user.id
    data = query.data

    if data == "main_menu":
        USER_STATE.pop(user_id, None)
        await query.message.reply_text(
            "🏠 *Ana Menü* — Rapor hazırlamak istediğiniz haftayı seçin:",
            parse_mode="Markdown",
            reply_markup=_get_main_menu_keyboard(),
        )
        return

    if data.startswith("new_week_"):
        iso_d = data.replace("new_week_", "")
        target_mon = date.fromisoformat(iso_d)
        await _create_and_send_draft(query.message, user_id, target_mon)
        return

    if data == "ask_custom_date":
        USER_STATE[user_id] = "WAIT_CUSTOM_DATE"
        await query.message.reply_text(
            "🗓️ Lütfen rapor haftasının başlangıç tarihini veya tarih aralığını yazın:\n"
            "Örnek: `24.08.2026` veya `24-30.08.2026`",
            parse_mode="Markdown",
        )
        return

    if data == "dl_excel":
        if EXCEL_PATH.exists():
            with open(EXCEL_PATH, "rb") as fp:
                await query.message.reply_document(
                    document=InputFile(fp, filename=EXCEL_PATH.name),
                    caption=(
                        "📊 Mevcut `mevsim_normalleri.xlsx` dosyası.\n"
                        "Güncellemek istediğinizde yeni `.xlsx` dosyasını bu sohbete göndermeniz yeterlidir."
                    ),
                    parse_mode="Markdown",
                )
        else:
            await query.message.reply_text("⚠️ Mevsim normalleri Excel dosyası bulunamadı. Lütfen `.xlsx` dosyasını sohbete gönderin.")
        return

    if data == "list_reports":
        files = sorted(
            [p for p in list(OUTPUT_DIR.glob("*.docx")) + list(DOCUMENTS_DIR.glob("*.docx")) if not p.name.startswith("~$")],
            key=lambda x: x.stat().st_mtime,
            reverse=True,
        )
        seen = set()
        unique_files = []
        for f in files:
            if f.name not in seen:
                seen.add(f.name)
                unique_files.append(f)

        if not unique_files:
            await query.message.reply_text("📂 Henüz kayıtlı bir rapor bulunamadı.")
            return

        kb = []
        for idx, f in enumerate(unique_files[:8]):
            kb.append([InlineKeyboardButton(f"📄 {f.name}", callback_data=f"dl_doc_{idx}")])
        context.user_data["listed_files"] = [str(f) for f in unique_files[:8]]
        await query.message.reply_text(
            "📂 *Kayıtlı Raporlar (İndirmek için tıklayın):*",
            parse_mode="Markdown",
            reply_markup=InlineKeyboardMarkup(kb),
        )
        return

    if data.startswith("dl_doc_"):
        idx = int(data.replace("dl_doc_", ""))
        flist = context.user_data.get("listed_files", [])
        if 0 <= idx < len(flist):
            fpath = Path(flist[idx])
            if fpath.exists():
                with open(fpath, "rb") as fp:
                    await query.message.reply_document(
                        document=InputFile(fp, filename=fpath.name),
                        caption=f"📄 {fpath.name}",
                    )
        return

    draft = USER_DRAFTS.get(user_id)
    if not draft:
        await query.message.reply_text(
            "⚠️ Aktif bir rapor taslağı bulunamadı. Lütfen önce haftayı seçin:",
            reply_markup=_get_main_menu_keyboard(),
        )
        return

    if data == "edit_mgm":
        USER_STATE[user_id] = "WAIT_MGM_TEXT"
        await query.message.reply_text(
            "🌡️ *2. Bölüm MGM Haftalık Sıcaklık Açıklaması Düzenleme*\n\n"
            f"Mevcut metin:\n`{draft['sec2_intro']}`\n\n"
            "MGM'nin Instagram/Web haftalık tahmininde yer alan sıcaklık cümlesini gönderin.\n"
            "Örnek: `Hava sıcaklıklarının ülkemiz genelinde mevsim normallerinin 2 ila 4 °C üzerinde seyredeceği tahmin edilmektedir.`",
            parse_mode="Markdown",
        )
        return

    if data == "edit_sec2":
        USER_STATE[user_id] = "WAIT_SEC2_BULLETS"
        current_lines = [f"{lbl}: {desc}" for lbl, desc in draft["sec2_bullets"]]
        joined = "\n\n".join(current_lines)
        await query.message.reply_text(
            "🌦️ *2. Bölüm Günlük Tahmin Maddelerini Düzenleme*\n\n"
            "Aşağıdaki metni kopyalayıp istediğiniz günü/günleri düzenleyerek tekrar gönderin "
            "(Her satır `Tarih: Açıklama` formatında olmalıdır):\n\n"
            f"```\n{joined}\n```",
            parse_mode="Markdown",
        )
        return

    if data == "edit_sec1":
        USER_STATE[user_id] = "WAIT_SEC1_BULLETS"
        current_lines = [f"{lbl}: {desc}" for lbl, desc in draft["sec1_bullets"]]
        joined = "\n\n".join(current_lines)
        await query.message.reply_text(
            "⏪ *1. Bölüm (Geçen Hafta) Maddelerini Düzenleme*\n\n"
            "İsterseniz doğrudan önceki haftanın `.docx` dosyasını bu sohbete gönderebilirsiniz (otomatik geçmiş zamana çevrilir) "
            "veya aşağıdaki maddeleri düzenleyip gönderebilirsiniz:\n\n"
            f"```\n{joined}\n```",
            parse_mode="Markdown",
        )
        return

    if data == "edit_wind":
        USER_STATE[user_id] = "WAIT_WIND_TEXT"
        ww = draft.get("wind_warnings", {})
        lines = []
        for k in ("Marmara", "Ege", "Akdeniz", "Karadeniz"):
            vals = ww.get(k, [])
            if not vals:
                lines.append(f"{k}: Yok")
            else:
                for v in vals:
                    lines.append(f"{k} | {v}")
        joined = "\n".join(lines)
        await query.message.reply_text(
            "💨 *4. Bölüm (Rüzgâr & Fırtına Uyarıları) Düzenleme*\n\n"
            "Mevcut uyarılar:\n"
            f"```\n{joined}\n```\n\n"
            "Yeni uyarıları her satırda `Bölge | Tarih: Açıklama` (veya fırtına yoksa `Bölge: Yok`) şeklinde gönderebilirsiniz.\n"
            "Örnek:\n"
            "`Marmara | 27 Ağustos 2026: Kuzeyden 45-50 km/saat şiddetinde gün boyu devam edeceği ön görülmektedir.`\n"
            "`Ege: Yok`",
            parse_mode="Markdown",
        )
        return

    if data == "edit_frost":
        USER_STATE[user_id] = "WAIT_FROST_TEXT"
        await query.message.reply_text(
            "❄️ *5. Bölüm Don Riski Özeti & Tablo 8 Düzenleme*\n\n"
            f"Mevcut özet cümlesi:\n`{draft['frost_summary']}`\n\n"
            "Don beklenmiyorsa sadece `Yok` yazabilirsiniz. Don beklenen tarih ve bölgeleri girmek için:\n"
            "İlk satıra özet cümlesini, sonraki satırlara `Tarih | Bölge` yazıp gönderin.",
            parse_mode="Markdown",
        )
        return

    if data == "edit_month":
        kb = []
        row = []
        for m_num, m_name in MONTH_NAMES_TR.items():
            row.append(InlineKeyboardButton(m_name, callback_data=f"set_month_{m_num}"))
            if len(row) == 3:
                kb.append(row)
                row = []
        if row:
            kb.append(row)
        await query.message.reply_text(
            "📊 *3. Bölümde yer alacak Mevsim Normalleri ayını seçin:*",
            parse_mode="Markdown",
            reply_markup=InlineKeyboardMarkup(kb),
        )
        return

    if data.startswith("set_month_"):
        m_num = int(data.replace("set_month_", ""))
        draft["normals_month"] = m_num
        draft["normals_month_name"] = MONTH_NAMES_TR[m_num]
        await query.message.reply_text(
            f"✅ 3. Bölüm Mevsim Normalleri ayı *{MONTH_NAMES_TR[m_num]}* olarak güncellendi!",
            parse_mode="Markdown",
            reply_markup=_get_draft_keyboard(),
        )
        return

    if data == "add_extra_map":
        USER_STATE[user_id] = "WAIT_EXTRA_MAP"
        await query.message.reply_text(
            "🗺️ *2. Bölüme Şiddetli Yağış veya CAPE Haritası Ekleme*\n\n"
            "Windy'den otomatik çekilmesi için şu formatta yazın:\n"
            "`rain | 27.08.2026 | 15:00` (Yağış Haritası için)\n"
            "`cape | 23.05.2026 | 14:00` (CAPE İndeksi Haritası için)\n\n"
            "Veya tüm ekstra haritaları temizlemek için `temizle` yazın.",
            parse_mode="Markdown",
        )
        return

    if data in ("act_generate_full", "act_generate_fast"):
        do_scrape = data == "act_generate_full"
        progress_msg = await query.message.reply_text(
            "🚀 *Rapor hazırlanıyor...*\n"
            + ("🌐 Windy.com üzerinden Türkiye harita ekran görüntüleri alınıyor (yaklaşık 25-35 sn)..." if do_scrape else "📄 Word (.docx) dosyası oluşturuluyor..."),
            parse_mode="Markdown",
        )

        if do_scrape:
            async def prog_cb(done_c, total_c, label_str):
                try:
                    await progress_msg.edit_text(
                        f"📸 *Windy Haritaları Çekiliyor ({done_c}/{total_c})*\n"
                        f"👉 `{label_str}`",
                        parse_mode="Markdown",
                    )
                except Exception:
                    pass

            await capture_windy_maps_async(
                draft["temp_maps"],
                draft.get("sec2_extra_maps", []),
                progress_callback=prog_cb,
            )

        await progress_msg.edit_text("📝 *Word (.docx) raporu TARSİM şablonuyla birleştiriliyor...*", parse_mode="Markdown")
        loop = asyncio.get_running_loop()
        out_docx_path = await loop.run_in_executor(None, generate_weekly_report_docx, draft, True)

        await progress_msg.edit_text(
            f"✅ *Rapor başarıyla oluşturuldu!*\n📂 Dosya: `{out_docx_path.name}`\n⬇️ Hemen aşağıdan indirebilirsiniz:",
            parse_mode="Markdown",
        )
        with open(out_docx_path, "rb") as fp:
            await query.message.reply_document(
                document=InputFile(fp, filename=out_docx_path.name),
                caption=f"📊 {draft['curr_range_str']} Haftalık Meteorolojik Değerlendirme",
            )
        return


async def text_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    text = (update.message.text or "").strip()
    state = USER_STATE.get(user_id)

    if state == "WAIT_CUSTOM_DATE":
        parsed = _parse_user_date(text)
        if not parsed:
            await update.message.reply_text(
                "⚠️ Tarih formatı anlaşılamadı. Lütfen `24.08.2026` gibi bir tarih yazın:",
                parse_mode="Markdown",
            )
            return
        await _create_and_send_draft(update.message, user_id, parsed)
        return

    draft = USER_DRAFTS.get(user_id)
    if not draft:
        parsed = _parse_user_date(text)
        if parsed:
            await _create_and_send_draft(update.message, user_id, parsed)
            return
        await start_cmd(update, context)
        return

    if state == "WAIT_MGM_TEXT":
        USER_STATE.pop(user_id, None)
        cleaned = text.strip()
        if not cleaned.endswith("Günlere göre haftanın diğer meteorolojik durumları ile ilgili açıklamalar ise şu şekildedir:"):
            if not cleaned.lower().startswith(draft["curr_range_str"].lower()[:2]):
                cleaned = re.sub(r"^hava sıcaklıklarının\s+", "hava sıcaklıklarının ", cleaned, flags=re.IGNORECASE)
                if not cleaned.endswith("."):
                    cleaned += "."
                draft["sec2_intro"] = (
                    f"{draft['curr_range_str']} haftasında {cleaned} "
                    f"Günlere göre haftanın diğer meteorolojik durumları ile ilgili açıklamalar ise şu şekildedir:"
                )
            else:
                draft["sec2_intro"] = (
                    f"{cleaned.rstrip('.')}."
                    f" Günlere göre haftanın diğer meteorolojik durumları ile ilgili açıklamalar ise şu şekildedir:"
                )
        else:
            draft["sec2_intro"] = cleaned

        await update.message.reply_text(
            f"✅ *2. Bölüm giriş paragrafı güncellendi:*\n\n_{draft['sec2_intro']}_",
            parse_mode="Markdown",
            reply_markup=_get_draft_keyboard(),
        )
        return

    if state in ("WAIT_SEC2_BULLETS", "WAIT_SEC1_BULLETS"):
        is_sec1 = state == "WAIT_SEC1_BULLETS"
        USER_STATE.pop(user_id, None)
        new_bullets = []
        for raw_line in text.splitlines():
            line = raw_line.strip().lstrip("•-*").strip()
            if not line or ":" not in line:
                continue
            d_part, desc_part = line.split(":", 1)
            desc_clean = convert_to_past_tense(desc_part.strip()) if is_sec1 else desc_part.strip()
            new_bullets.append((d_part.strip(), desc_clean))

        if new_bullets:
            if is_sec1:
                draft["sec1_bullets"] = new_bullets
            else:
                draft["sec2_bullets"] = new_bullets
            await update.message.reply_text(
                f"✅ *{'1. Bölüm' if is_sec1 else '2. Bölüm'} günlük maddeleri ({len(new_bullets)} adet) güncellendi!*",
                parse_mode="Markdown",
                reply_markup=_get_draft_keyboard(),
            )
        else:
            await update.message.reply_text(
                "⚠️ Geçerli `Tarih: Açıklama` satırı bulunamadı.",
                parse_mode="Markdown",
                reply_markup=_get_draft_keyboard(),
            )
        return

    if state == "WAIT_WIND_TEXT":
        USER_STATE.pop(user_id, None)
        new_ww = {"Marmara": [], "Ege": [], "Akdeniz": [], "Karadeniz": []}
        for raw_line in text.splitlines():
            line = raw_line.strip()
            if not line:
                continue
            if "|" in line:
                reg, rest = line.split("|", 1)
                reg = reg.strip()
                rest = rest.strip()
                for k in new_ww:
                    if k.lower() in reg.lower():
                        if rest.lower() not in ("yok", "-"):
                            new_ww[k].append(rest)
        draft["wind_warnings"] = new_ww
        await update.message.reply_text(
            "✅ *4. Bölüm (Rüzgâr & Fırtına) uyarıları güncellendi!*",
            parse_mode="Markdown",
            reply_markup=_get_draft_keyboard(),
        )
        return

    if state == "WAIT_FROST_TEXT":
        USER_STATE.pop(user_id, None)
        if text.lower() in ("yok", "-"):
            draft["frost_summary"] = (
                "Tüm hafta boyunca, yurdumuz genelinde saat 06:00’daki sıcaklıklar incelendiğinde "
                "hafta boyunca yurt genelinde sıcaklıkların 0 °C üzerinde olması beklendiğinden "
                "don hadisesi beklenmemektedir"
            )
            all_dates = "\n".join(tm["label_pair"] for tm in draft["temp_maps"][:7])
            draft["table8_rows"] = [(all_dates, "-")]
        else:
            lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
            if lines:
                draft["frost_summary"] = lines[0]
            rows = []
            for ln in lines[1:]:
                if "|" in ln:
                    d_p, r_p = ln.split("|", 1)
                    rows.append((d_p.strip(), r_p.strip()))
            if rows:
                draft["table8_rows"] = rows

        await update.message.reply_text(
            "✅ *5. Bölüm Don Riski ve Tablo 8 güncellendi!*",
            parse_mode="Markdown",
            reply_markup=_get_draft_keyboard(),
        )
        return

    if state == "WAIT_EXTRA_MAP":
        USER_STATE.pop(user_id, None)
        if text.lower() == "temizle":
            draft["sec2_extra_maps"] = []
            await update.message.reply_text(
                "✅ 2. Bölüm ekstra haritaları temizlendi.",
                reply_markup=_get_draft_keyboard(),
            )
            return

        parts = [p.strip() for p in text.split("|")]
        if len(parts) >= 2:
            m_type = "cape" if "cape" in parts[0].lower() else "rain"
            d_str = parts[1]
            hr_str = parts[2] if len(parts) >= 3 else ("14:00" if m_type == "cape" else "15:00")
            m = re.search(r"(\d{1,2})[./](\d{1,2})[./](\d{4})", d_str)
            if m:
                d_obj = date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
                start_mon = date.fromisoformat(draft["target_monday"])
                b_idx = max(0, min(6, (d_obj - start_mon).days))
                label_type = "CAPE İndeksi Haritası" if m_type == "cape" else "Yağış Haritası"
                draft.setdefault("sec2_extra_maps", []).append(
                    {
                        "type": m_type,
                        "date": d_obj.isoformat(),
                        "hour": hr_str,
                        "after_bullet_idx": b_idx,
                        "caption_suffix": f"{d_obj.strftime('%d.%m.%Y')} {hr_str} (Yerel Saat) {label_type}",
                    }
                )
                await update.message.reply_text(
                    f"✅ Ekstra harita eklendi: `{d_obj.strftime('%d.%m.%Y')} {hr_str} {label_type}`",
                    parse_mode="Markdown",
                    reply_markup=_get_draft_keyboard(),
                )
                return

        await update.message.reply_text(
            "⚠️ Format anlaşılamadı. Örnek: `rain | 27.08.2026 | 15:00`",
            parse_mode="Markdown",
            reply_markup=_get_draft_keyboard(),
        )
        return

    await update.message.reply_text(
        "👇 Aktif taslağınız üzerinde işlem yapmak için aşağıdaki butonları kullanabilirsiniz:",
        reply_markup=_get_draft_keyboard(),
    )


async def document_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    Kullanıcı Telegram sohbetine:
    - .xlsx dosyası gönderirse -> mevsim_normalleri.xlsx dosyasını günceller.
    - .docx dosyası gönderirse -> önceki hafta raporu olarak kaydeder ve 1. Bölüme geçmiş zamanla aktarır.
    """
    user_id = update.effective_user.id
    doc = update.message.document
    if not doc:
        return

    fname = doc.file_name or ""
    tg_file = await doc.get_file()

    if fname.lower().endswith(".xlsx"):
        await tg_file.download_to_drive(custom_path=str(EXCEL_PATH))
        try:
            shutil.copy2(EXCEL_PATH, FALLBACK_EXCEL_PATH)
        except Exception:
            pass
        await update.message.reply_text(
            f"✅ *Mevsim Normalleri Excel dosyası (`{EXCEL_PATH.name}`) başarıyla güncellendi!*",
            parse_mode="Markdown",
        )
        return

    if fname.lower().endswith(".docx"):
        save_path = DOCUMENTS_DIR / fname
        await tg_file.download_to_drive(custom_path=str(save_path))
        sec2 = extract_section2_from_docx(save_path)
        if sec2 and user_id in USER_DRAFTS:
            draft = USER_DRAFTS[user_id]
            draft["sec1_source"] = fname
            draft["sec1_intro"] = sec2["intro"]
            draft["sec1_bullets"] = sec2["bullets"]
            await update.message.reply_text(
                f"✅ `{fname}` kaydedildi ve 2. Bölümü geçmiş zaman kipine (*-mekteydi*) çevrilerek "
                f"aktif taslağın *1. Bölümüne* aktarıldı!",
                parse_mode="Markdown",
                reply_markup=_get_draft_keyboard(),
            )
        else:
            await update.message.reply_text(
                f"✅ `{fname}` dosyası bulut/sistem klasörüne kaydedildi. Yeni rapor oluşturduğunuzda 1. Bölüm için otomatik kullanılacaktır.",
                parse_mode="Markdown",
            )


# --- BULUT 7/24 HEALTH-CHECK & KEEP-ALIVE WEB SUNUCUSU ---

async def _handle_health(request: web.Request) -> web.Response:
    uptime_sec = int((datetime.now() - START_TIME).total_seconds())
    token_ok = bool(TELEGRAM_BOT_TOKEN)
    html = f"""<!DOCTYPE html>
<html lang="tr">
<head>
<meta charset="utf-8">
<title>TARSİM Meteoroloji Haftalık Rapor Botu — 7/24 Bulut Servisi</title>
<style>
body {{ font-family: Arial, sans-serif; background: #0f172a; color: #f8fafc; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; }}
.card {{ background: #1e293b; padding: 32px; border-radius: 14px; max-width: 560px; box-shadow: 0 10px 30px rgba(0,0,0,0.4); }}
.badge {{ display: inline-block; padding: 6px 12px; border-radius: 999px; font-weight: bold; font-size: 14px; background: {'#16a34a' if token_ok else '#dc2626'}; }}
</style>
</head>
<body>
<div class="card">
  <h2>🌦️ TARSİM Haftalık Meteorolojik Rapor Botu</h2>
  <p><span class="badge">{'✅ 7/24 AKTİF ÇALIŞIYOR' if token_ok else '⚠️ TELEGRAM_BOT_TOKEN BEKLENİYOR'}</span></p>
  <p><strong>Çalışma Süresi (Uptime):</strong> {uptime_sec} saniye</p>
  <p><strong>Mevsim Normalleri Excel:</strong> {'Mevcut' if EXCEL_PATH.exists() else 'Eksik'}</p>
  {'' if token_ok else '<p style="color:#fca5a5;">Lütfen Bulut (Hugging Face Spaces / Render) <b>Settings &rarr; Variables / Environment</b> kısmına <code>TELEGRAM_BOT_TOKEN</code> değerinizi ekleyin.</p>'}
</div>
</body>
</html>"""
    return web.Response(text=html, content_type="text/html")


async def _keep_alive_loop():
    """Bulut servisinin uykuya geçmemesi için periyodik self-ping döngüsü."""
    if not PUBLIC_URL:
        return
    ping_url = PUBLIC_URL.rstrip("/") + "/health"
    async with ClientSession() as session:
        while True:
            await asyncio.sleep(480)  # 8 dakikada bir
            try:
                async with session.get(ping_url, timeout=10) as resp:
                    _ = resp.status
            except Exception:
                pass


async def main_async():
    # 1) HTTP Health-Check Sunucusunu başlat (Hugging Face Spaces & Render 7/24 uyumu)
    web_app = web.Application()
    web_app.router.add_get("/", _handle_health)
    web_app.router.add_get("/health", _handle_health)
    runner = web.AppRunner(web_app)
    await runner.setup()
    site = web.TCPSite(runner, host="0.0.0.0", port=PORT)
    await site.start()
    logger.info(f"🌐 7/24 Bulut Health-Check Sunucusu aktif: http://0.0.0.0:{PORT}")

    asyncio.create_task(_keep_alive_loop())

    if not TELEGRAM_BOT_TOKEN:
        logger.warning(
            "TELEGRAM_BOT_TOKEN henüz tanımlanmamış! Bulut panelinden (Settings -> Variables/Secrets) "
            "TELEGRAM_BOT_TOKEN eklediğinizde bot otomatik bağlanacaktır."
        )
        while True:
            await asyncio.sleep(3600)

    # 2) Telegram Botunu aynı event loop içinde başlat
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
    logger.info("✅ Meteoroloji Haftalık Rapor Telegram Botu 7/24 dinlemede!")

    try:
        while True:
            await asyncio.sleep(3600)
    finally:
        await tg_app.updater.stop()
        await tg_app.stop()
        await tg_app.shutdown()
        await runner.cleanup()


if __name__ == "__main__":
    asyncio.run(main_async())
