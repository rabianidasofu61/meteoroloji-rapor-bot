import asyncio
from datetime import datetime, timezone, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from PIL import Image, ImageDraw
from playwright.async_api import async_playwright

from config import MAPS_DIR

# Türkiye'yi Edirne/Çeşme'den Kars/Hakkâri'ye tam ekran kapsayan Windy koordinat ve zoom ayarı
TURKEY_LAT = 38.850
TURKEY_LON = 35.550
TURKEY_ZOOM = 7


def _local_to_windy_utc_str(date_iso: str, hour_str: str) -> Tuple[str, int]:
    """
    Türkiye yerel saatini (UTC+3) Windy URL formatına ('YYYYMMDDHH' UTC)
    ve Unix epoch milisaniyesine çevirir.
    """
    dt_local = datetime.strptime(f"{date_iso} {hour_str}", "%Y-%m-%d %H:%M")
    tz_tr = timezone(timedelta(hours=3))
    dt_aware = dt_local.replace(tzinfo=tz_tr)
    dt_utc = dt_aware.astimezone(timezone.utc)
    return dt_utc.strftime("%Y%m%d%H"), int(dt_utc.timestamp() * 1000)


def _create_fallback_map_image(out_path: Path, title: str, subtitle: str) -> Path:
    img = Image.new("RGB", (1913, 864), color=(42, 58, 74))
    draw = ImageDraw.Draw(img)
    draw.rectangle([40, 40, 1873, 824], outline=(100, 140, 180), width=3)
    draw.text((80, 380), title, fill=(255, 255, 255))
    draw.text((80, 440), subtitle, fill=(200, 220, 240))
    img.save(out_path, "PNG")
    return out_path


async def capture_windy_maps_async(
    temp_maps: List[Dict],
    extra_maps: Optional[List[Dict]] = None,
    progress_callback=None,
) -> Tuple[List[Dict], List[Dict]]:
    """
    Playwright kullanarak Windy.com üzerinden Türkiye'yi tam kapsayacak şekilde (Zoom=7, 1913x864):
    - Her haftanın raporunda sabit yer alan 8 adet 06:00 Türkiye Sıcaklık Haritasını (temp katmanı),
    - Türkiye genelinde yağışlı günlerde Yağış Haritasını (rain katmanı),
    - CAPE değerlerinin yüksek olduğu günlerde CAPE İndeksi Haritasını (cape katmanı)
    otomatik olarak ekran görüntüsü alır.
    """
    extra_maps = extra_maps or []
    total_count = len(temp_maps) + len(extra_maps)
    done_count = 0

    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=True,
            args=[
                "--no-sandbox",
                "--disable-setuid-sandbox",
                "--disable-dev-shm-usage",
                "--disable-gpu",
            ],
        )
        context = await browser.new_context(
            viewport={"width": 1913, "height": 864},
            locale="tr-TR",
            timezone_id="Europe/Istanbul",
            device_scale_factor=1,
        )
        page = await context.new_page()

        first_date = temp_maps[0]["morning_date"] if temp_maps else datetime.now().strftime("%Y-%m-%d")
        utc_str, ts_ms = _local_to_windy_utc_str(first_date, "06:00")
        init_url = f"https://www.windy.com/?temp,{utc_str},{TURKEY_LAT:.3f},{TURKEY_LON:.3f},{TURKEY_ZOOM}"

        try:
            await page.goto(init_url, wait_until="domcontentloaded", timeout=35000)
            await asyncio.sleep(4.5)
            await page.evaluate(
                """([lat, lon, zoom]) => {
                    const closeSelectors = [
                        '#CybotCookiebotDialogBodyLevelButtonLevelOptinAllowAll',
                        '.cookie-banner button',
                        '.onboarding-close',
                        '[data-do="close"]'
                    ];
                    for (const sel of closeSelectors) {
                        const el = document.querySelector(sel);
                        if (el) { try { el.click(); } catch(e) {} }
                    }
                    // Sol üst arama çubuğu ve logoyu gizle (örnek görsellerdeki gibi temiz görünüm)
                    for (const s of ['#search', '#logo']) {
                        const el = document.querySelector(s);
                        if (el) el.style.display = 'none';
                    }
                    if (window.W && window.W.store) {
                        try {
                            window.W.store.set('product', 'ecmwf');
                            window.W.store.set('isolines', 'pressure');
                        } catch(e) {}
                    }
                    if (window.W && window.W.map) {
                        try { window.W.map.setView([lat, lon], zoom); } catch(e) {}
                    }
                }""",
                [TURKEY_LAT, TURKEY_LON, TURKEY_ZOOM],
            )
        except Exception:
            pass

        # 1) 8 adet 06:00 Türkiye Sıcaklık Haritası (Her haftalık raporda zorunlu)
        for m in temp_maps:
            done_count += 1
            m_date = m["morning_date"]
            m_hour = m.get("hour", "06:00")
            out_file = MAPS_DIR / f"temp_{m['label_pair'].replace('.', '_')}_0600.png"

            if progress_callback:
                try:
                    await progress_callback(done_count, total_count, f"{m['label_pair']} 06:00 Sıcaklık Haritası")
                except Exception:
                    pass

            utc_str, ts_ms = _local_to_windy_utc_str(m_date, m_hour)
            try:
                changed = await page.evaluate(
                    """([ts, overlay, lat, lon, zoom]) => {
                        if (window.W && window.W.store) {
                            window.W.store.set('overlay', overlay);
                            window.W.store.set('product', 'ecmwf');
                            window.W.store.set('isolines', 'pressure');
                            window.W.store.set('timestamp', ts);
                            if (window.W.map) {
                                window.W.map.setView([lat, lon], zoom);
                            }
                            return true;
                        }
                        return false;
                    }""",
                    [ts_ms, "temp", TURKEY_LAT, TURKEY_LON, TURKEY_ZOOM],
                )
                if not changed:
                    url = f"https://www.windy.com/?temp,{utc_str},{TURKEY_LAT:.3f},{TURKEY_LON:.3f},{TURKEY_ZOOM}"
                    await page.goto(url, wait_until="domcontentloaded", timeout=25000)

                await asyncio.sleep(2.8)
                await page.screenshot(path=str(out_file), full_page=False)
                m["image_path"] = str(out_file)
            except Exception as e:
                _create_fallback_map_image(out_file, m["caption_suffix"], str(e))
                m["image_path"] = str(out_file)

        # 2) Ekstra Yağış ('rain') ve CAPE ('cape') Haritaları (Bölüm 2 için)
        for em in extra_maps:
            done_count += 1
            em_date = em["date"]
            em_hour = em.get("hour", "15:00")
            em_type = em.get("type", "rain")
            out_file = MAPS_DIR / f"extra_{em_type}_{em_date}_{em_hour.replace(':', '')}.png"

            if progress_callback:
                try:
                    await progress_callback(done_count, total_count, em["caption_suffix"])
                except Exception:
                    pass

            utc_str, ts_ms = _local_to_windy_utc_str(em_date, em_hour)
            try:
                changed = await page.evaluate(
                    """([ts, overlay, lat, lon, zoom]) => {
                        if (window.W && window.W.store) {
                            window.W.store.set('overlay', overlay);
                            window.W.store.set('product', 'ecmwf');
                            window.W.store.set('isolines', 'gh');
                            window.W.store.set('timestamp', ts);
                            if (window.W.map) {
                                window.W.map.setView([lat, lon], zoom);
                            }
                            return true;
                        }
                        return false;
                    }""",
                    [ts_ms, em_type, TURKEY_LAT, TURKEY_LON, TURKEY_ZOOM],
                )
                if not changed:
                    url = f"https://www.windy.com/?{em_type},{utc_str},{TURKEY_LAT:.3f},{TURKEY_LON:.3f},{TURKEY_ZOOM}"
                    await page.goto(url, wait_until="domcontentloaded", timeout=25000)

                await asyncio.sleep(3.0)
                await page.screenshot(path=str(out_file), full_page=False)
                em["image_path"] = str(out_file)
            except Exception as e:
                _create_fallback_map_image(out_file, em["caption_suffix"], str(e))
                em["image_path"] = str(out_file)

        await browser.close()

    return temp_maps, extra_maps


def capture_windy_maps_sync(
    temp_maps: List[Dict],
    extra_maps: Optional[List[Dict]] = None,
) -> Tuple[List[Dict], List[Dict]]:
    return asyncio.run(capture_windy_maps_async(temp_maps, extra_maps))
