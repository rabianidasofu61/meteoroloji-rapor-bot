import re
from datetime import date, datetime, timedelta
from typing import Dict, List, Optional, Tuple
import requests

from config import MONTH_NAMES_TR, DAY_NAMES_TR
from excel_service import get_national_normal_temp
from prev_report_parser import (
    find_previous_week_docx,
    extract_section2_from_docx,
    format_week_range_tr,
    convert_to_past_tense,
)

# Türkiye bölgelerini ve alt bölgelerini temsil eden koordinatlar (Windy ECMWF analizi için)
REPRESENTATIVE_POINTS = [
    {"id": "marmara", "name": "Marmara", "region": "Marmara", "sub": "Marmara", "lat": 40.75, "lon": 28.90},
    {"id": "marmara_kd", "name": "Marmara’nın kuzeydoğusu", "region": "Marmara", "sub": "Kocaeli-Sakarya", "lat": 40.80, "lon": 30.10},
    {"id": "ege_kuzey", "name": "Ege", "region": "Ege", "sub": "İzmir-Manisa", "lat": 38.45, "lon": 27.20},
    {"id": "ege_ic", "name": "İç Ege", "region": "Ege", "sub": "Afyon-Kütahya", "lat": 38.75, "lon": 30.50},
    {"id": "akdeniz_bati", "name": "Akdeniz’in batısı", "region": "Akdeniz", "sub": "Antalya-Isparta", "lat": 37.05, "lon": 30.75},
    {"id": "akdeniz_dogu", "name": "Akdeniz", "region": "Akdeniz", "sub": "Adana-Mersin-Hatay", "lat": 36.95, "lon": 35.35},
    {"id": "ic_anadolu_bati", "name": "İç Anadolu’nun kuzeybatısı", "region": "İç Anadolu", "sub": "Ankara-Eskişehir-Konya", "lat": 39.50, "lon": 32.50},
    {"id": "ic_anadolu_dogu", "name": "İç Anadolu’nun doğusu", "region": "İç Anadolu", "sub": "Kayseri-Sivas-Niğde", "lat": 39.10, "lon": 36.00},
    {"id": "bati_karadeniz", "name": "Batı Karadeniz", "region": "Karadeniz", "sub": "Düzce-Bolu-Kastamonu", "lat": 41.20, "lon": 32.60},
    {"id": "orta_karadeniz", "name": "Orta Karadeniz", "region": "Karadeniz", "sub": "Samsun-Ordu-Amasya", "lat": 41.25, "lon": 36.35},
    {"id": "dogu_karadeniz", "name": "Doğu Karadeniz", "region": "Karadeniz", "sub": "Trabzon-Rize", "lat": 41.00, "lon": 40.50},
    {"id": "dogu_anadolu_kuzey", "name": "Doğu Anadolu’nun kuzeyi", "region": "Doğu Anadolu", "sub": "Erzurum-Kars-Ardahan", "lat": 40.15, "lon": 41.60},
    {"id": "dogu_anadolu_guney", "name": "Doğu Anadolu", "region": "Doğu Anadolu", "sub": "Van-Muş-Bitlis-Elazığ", "lat": 38.60, "lon": 41.50},
    {"id": "guneydogu", "name": "Güneydoğu Anadolu", "region": "Güneydoğu Anadolu", "sub": "Gaziantep-Şanlıurfa-Diyarbakır", "lat": 37.50, "lon": 39.30},
]


def deg_to_direction_tr(deg: float) -> str:
    """Rüzgâr açısını Türkçe yön ifadesine çevirir."""
    dirs = [
        "Kuzeyden",
        "Kuzey-kuzeydoğudan",
        "Kuzeydoğudan",
        "Doğu-kuzeydoğudan",
        "Doğudan",
        "Doğu-güneydoğudan",
        "Güneydoğudan",
        "Güney-güneydoğudan",
        "Güneyden",
        "Güney-güneybatıdan",
        "Güneybatıdan",
        "Batı-güneybatıdan",
        "Batıdan",
        "Batı-kuzeybatıdan",
        "Kuzeybatıdan",
        "Kuzey-kuzeybatıdan",
    ]
    idx = int((deg + 11.25) // 22.5) % 16
    return dirs[idx]


def fetch_open_meteo_hourly(start_d: date, end_d: date) -> Dict[str, Dict]:
    """
    Belirtilen tarih aralığı için tüm temsil noktalarının saatlik meteorolojik
    verilerini Open-Meteo API üzerinden tek bir toplu sorguyla çeker.
    """
    lats = ",".join(str(p["lat"]) for p in REPRESENTATIVE_POINTS)
    lons = ",".join(str(p["lon"]) for p in REPRESENTATIVE_POINTS)

    today = date.today()
    if end_d < today - timedelta(days=5):
        url = "https://archive-api.open-meteo.com/v1/archive"
        hourly_vars = "temperature_2m,precipitation,rain,snowfall,pressure_msl,wind_speed_10m,wind_gusts_10m,wind_direction_10m"
    else:
        url = "https://api.open-meteo.com/v1/forecast"
        hourly_vars = "temperature_2m,precipitation,rain,snowfall,pressure_msl,wind_speed_10m,wind_gusts_10m,wind_direction_10m,cape"

    params = {
        "latitude": lats,
        "longitude": lons,
        "hourly": hourly_vars,
        "timezone": "Europe/Istanbul",
        "start_date": start_d.isoformat(),
        "end_date": end_d.isoformat(),
    }

    try:
        resp = requests.get(url, params=params, timeout=20)
        resp.raise_for_status()
        data = resp.json()
        results = {}
        if isinstance(data, list):
            for p, item in zip(REPRESENTATIVE_POINTS, data):
                results[p["id"]] = item.get("hourly", {})
        elif isinstance(data, dict) and "hourly" in data:
            results[REPRESENTATIVE_POINTS[0]["id"]] = data["hourly"]
        return results
    except Exception:
        return {}


def _join_regions_tr(regions: List[str]) -> str:
    if not regions:
        return ""
    if len(regions) == 1:
        return regions[0]
    return ", ".join(regions[:-1]) + " ve " + regions[-1]


def synthesize_daily_bullet(
    day_date: date,
    day_idx: int,
    hourly_by_point: Dict[str, Dict],
    past_tense: bool = False,
) -> Tuple[str, str, List[Dict]]:
    """
    Tek bir gün için saatlik yağış, kar, basınç ve CAPE verilerini inceleyerek
    TARSİM rapor diline uygun günlük değerlendirme maddesini üretir.
    Ayrıca:
    - Türkiye genelinde yağışlı bir gün veya şiddetli yağış varsa Yağış Haritası,
    - Türkiye'nin belirli ya da genel bölgelerinde CAPE değerleri yüksekse CAPE İndeksi Haritası
      ve örnek Word raporundaki standart CAPE/dolu yorumunu ekler.
    """
    date_label = f"{day_date.day:02d} {MONTH_NAMES_TR[day_date.month]} {day_date.year}"
    target_prefix = day_date.isoformat()

    morning_rain_regions = []
    afternoon_rain_regions = []
    evening_rain_regions = []
    heavy_rain_regions = []
    snow_regions = []
    cape_regions = []
    max_cape = 0.0
    max_precip_day = 0.0

    for p in REPRESENTATIVE_POINTS:
        pid = p["id"]
        h = hourly_by_point.get(pid, {})
        times = h.get("time", [])
        if not times:
            continue

        idxs = [i for i, t in enumerate(times) if t.startswith(target_prefix)]
        if not idxs:
            continue

        precips = [h.get("precipitation", [0]*len(times))[i] or 0.0 for i in idxs]
        snows = [h.get("snowfall", [0]*len(times))[i] or 0.0 for i in idxs]
        capes = [h.get("cape", [0]*len(times))[i] or 0.0 for i in idxs] if "cape" in h else [0.0]*len(idxs)

        morning_p = sum(precips[6:12])
        afternoon_p = sum(precips[12:18])
        evening_p = sum(precips[18:24])
        total_p = sum(precips)
        total_s = sum(snows)

        # Öğleden sonra (12:00 - 18:00) maksimum CAPE
        afternoon_capes = capes[11:18] if len(capes) >= 18 else capes
        point_max_cape = max(afternoon_capes) if afternoon_capes else 0.0

        if total_p > max_precip_day:
            max_precip_day = total_p
        if point_max_cape > max_cape:
            max_cape = point_max_cape

        reg_name = p["region"]
        if morning_p >= 1.0 and reg_name not in morning_rain_regions:
            morning_rain_regions.append(reg_name)
        if afternoon_p >= 1.0 and reg_name not in afternoon_rain_regions:
            afternoon_rain_regions.append(reg_name)
        if evening_p >= 1.0 and reg_name not in evening_rain_regions:
            evening_rain_regions.append(reg_name)
        if total_p >= 14.0 and p["name"] not in heavy_rain_regions:
            heavy_rain_regions.append(p["name"])
        if total_s >= 0.5 and reg_name not in snow_regions:
            snow_regions.append(reg_name)

        # Windy CAPE ölçeğinde yeşil/sarı/kırmızı alanlar (~450+ J/kg)
        if point_max_cape >= 450.0 and reg_name not in cape_regions:
            cape_regions.append(reg_name)

    all_rain_regions = list(dict.fromkeys(morning_rain_regions + afternoon_rain_regions + evening_rain_regions))

    # Cümle sentezi
    if not all_rain_regions:
        text = (
            "Yurt genelinde yüksek basınç sisteminin hâkim olması nedeniyle "
            "açık ve yağışsız hava koşulları beklenmektedir."
        )
    else:
        precip_word = "yağış ve kar yağışı" if snow_regions else "yağış"
        parts = []

        if len(all_rain_regions) >= 5:
            if morning_rain_regions and len(morning_rain_regions) <= 3:
                m_str = _join_regions_tr(morning_rain_regions)
                parts.append(
                    f"Sabah saatlerinde {m_str} kesimlerinde, öğle saatlerinden sonra ise "
                    f"neredeyse tüm yurtta {precip_word}lı havanın hâkim olması ön görülmektedir."
                )
            else:
                parts.append(
                    f"Etkili olan alçak basınç sistemi nedeniyle gün içerisinde "
                    f"neredeyse tüm yurtta {precip_word}lı havanın hâkim olması beklenmektedir."
                )
        else:
            if morning_rain_regions and afternoon_rain_regions:
                m_str = _join_regions_tr(morning_rain_regions)
                a_only = [r for r in afternoon_rain_regions if r not in morning_rain_regions]
                if a_only:
                    a_str = _join_regions_tr(a_only)
                    parts.append(
                        f"Sabah saatlerinde {m_str} bölgesinde başlayan {precip_word}ların, "
                        f"öğle saatlerinden sonra etki alanını genişleterek {a_str} kesimlerinde de etkili olması beklenmektedir."
                    )
                else:
                    parts.append(
                        f"Gün içerisinde {m_str} kesimlerinde yerel {precip_word} geçişleri ön görülmektedir."
                    )
            elif afternoon_rain_regions:
                a_str = _join_regions_tr(afternoon_rain_regions)
                parts.append(
                    f"Öğle saatlerinden sonra {a_str} kesimlerinde yerel {precip_word} geçişleri beklenmektedir."
                )
            elif morning_rain_regions:
                m_str = _join_regions_tr(morning_rain_regions)
                parts.append(
                    f"Sabah saatlerinde {m_str} kesimlerinde {precip_word} geçişleri ön görülmekte olup, "
                    f"öğle saatlerinden sonra yağışların etkisini kaybetmesi beklenmektedir."
                )
            elif evening_rain_regions:
                e_str = _join_regions_tr(evening_rain_regions)
                parts.append(
                    f"Akşam saatlerine doğru {e_str} kesimlerinde yerel {precip_word} geçişleri beklenmektedir."
                )

        if heavy_rain_regions:
            h_str = _join_regions_tr(heavy_rain_regions[:3])
            parts.append(
                f"Özellikle {h_str} çevresinde diğer bölgelere kıyasla daha şiddetli yağışların gerçekleşeceği ön görülmektedir."
            )

        # CAPE (Dolu Riski) Yorumu — Örnek Word raporundaki birebir metin yapısı
        if not past_tense and cape_regions and max_cape >= 450.0:
            if len(cape_regions) >= 4:
                cape_scope_str = "yurdun büyük bir bölümünde"
            else:
                cape_scope_str = f"{_join_regions_tr(cape_regions)} bölgelerinde"
            parts.append(
                f"Ek olarak dolu hadisesi potansiyelinin belirlenmesinde kullanılan "
                f"CAPE - Convective Available Potential Energy (Konvektif Kullanılabilir Potansiyel Enerji) "
                f"indeksi haritasında öğle saatlerinden sonra {cape_scope_str} dolu riskinin olduğu görülmektedir. "
                f"Dolu hadisesi anlık gelişmesi nedeniyle ileri zamanlı (bir hafta sonrası gibi)  tahmin edilmesi oldukça zordur. "
                f"Ek olarak bu indeks tek başına doluyu direkt tespit etmede yeterli olamamaktadır. "
                f"Radardan anlık olarak takip edilmesi de gerekmektedir."
            )

        text = " ".join(parts)

    if past_tense:
        text = convert_to_past_tense(text)

    # Harita önerileri:
    # 1) Türkiye genelinde yağışlı bir gün (en az 4 bölgede yağış) veya şiddetli yağış (>= 15 mm) varsa Yağış Haritası
    # 2) Belirli veya genel bölgelerde CAPE yüksekse (>= 450 J/kg) CAPE İndeksi Haritası
    extra_maps_for_day: List[Dict] = []
    if not past_tense:
        if len(all_rain_regions) >= 4 or max_precip_day >= 15.0:
            extra_maps_for_day.append(
                {
                    "type": "rain",
                    "date": day_date.isoformat(),
                    "hour": "15:00",
                    "rain_regions_count": len(all_rain_regions),
                    "max_precip": max_precip_day,
                    "caption_suffix": f"{day_date.strftime('%d.%m.%Y')} 15:00 (Yerel Saat) Yağış Haritası",
                }
            )
        if cape_regions and max_cape >= 450.0:
            extra_maps_for_day.append(
                {
                    "type": "cape",
                    "date": day_date.isoformat(),
                    "hour": "14:00",
                    "max_cape": max_cape,
                    "caption_suffix": f"{day_date.strftime('%d.%m.%Y')} 14:00 (Yerel Saat) CAPE İndeksi Haritası",
                }
            )

    return date_label, text, extra_maps_for_day


def analyze_wind_warnings(start_date: date, hourly_by_point: Dict[str, Dict]) -> Dict[str, List[str]]:
    """
    4. Bölüm (Önemli Uyarılar -> 1. Rüzgâr & Fırtına) için
    Marmara, Ege, Akdeniz ve Karadeniz bölgelerinde saatlik rüzgâr hızını inceler.
    """
    region_point_map = {
        "Marmara": [("marmara", ""), ("marmara_kd", "")],
        "Ege": [("ege_kuzey", "")],
        "Akdeniz": [("akdeniz_bati", ""), ("akdeniz_dogu", "")],
        "Karadeniz": [("bati_karadeniz", "Batı Karadeniz’de "), ("dogu_karadeniz", "Doğu Karadeniz’de ")],
    }

    warnings: Dict[str, List[str]] = {
        "Marmara": [],
        "Ege": [],
        "Akdeniz": [],
        "Karadeniz": [],
    }

    for reg_key, pid_list in region_point_map.items():
        for d_offset in range(7):
            cur_d = start_date + timedelta(days=d_offset)
            d_prefix = cur_d.isoformat()
            d_label = f"{cur_d.day:02d} {MONTH_NAMES_TR[cur_d.month]} {cur_d.year}"

            for pid, sub_prefix in pid_list:
                h = hourly_by_point.get(pid, {})
                times = h.get("time", [])
                if not times:
                    continue
                idxs = [i for i, t in enumerate(times) if t.startswith(d_prefix)]
                if not idxs:
                    continue

                speeds = [h.get("wind_speed_10m", [0]*len(times))[i] or 0.0 for i in idxs]
                gusts = [h.get("wind_gusts_10m", [0]*len(times))[i] or 0.0 for i in idxs]
                dirs = [h.get("wind_direction_10m", [0]*len(times))[i] or 0.0 for i in idxs]

                eff_speeds = [max(s, g * 0.82) for s, g in zip(speeds, gusts)]
                max_spd = max(eff_speeds) if eff_speeds else 0.0

                if max_spd >= 38.0:
                    peak_idx = eff_speeds.index(max_spd)
                    dir_tr = deg_to_direction_tr(dirs[peak_idx])
                    low_b = int((max_spd // 5) * 5)
                    low_b = max(40, min(low_b, 75))
                    high_b = low_b + 5

                    strong_hours = [hr for hr, s in enumerate(eff_speeds) if s >= 35.0]
                    if len(strong_hours) >= 10:
                        period_str = "gün boyu devam edeceği ön görülmektedir."
                    elif strong_hours and max(strong_hours) <= 13:
                        period_str = "öğle saatlerine kadar devam edeceği ön görülmektedir."
                    elif strong_hours and min(strong_hours) >= 15:
                        period_str = "akşam saatlerinden sonra başlayıp gece saatlerinde kesilmesi ön görülmektedir."
                    else:
                        period_str = "gün içerisinde etkili olacağı ön görülmektedir."

                    prefix_lower = dir_tr.lower() if sub_prefix else dir_tr
                    bullet = f"{d_label}: {sub_prefix}{prefix_lower} {low_b}-{high_b} km/saat şiddetinde {period_str}"
                    warnings[reg_key].append(bullet)

    return warnings


def analyze_frost_risk_0600(
    start_date: date, hourly_by_point: Dict[str, Dict]
) -> Tuple[str, List[Tuple[str, str]], List[Dict]]:
    """
    5. Bölüm için 8 ardışık sabah 06:00 Türkiye Sıcaklık Haritalarını ve
    Tablo 8 (Don Riski Beklenen Bölge(ler)) verilerini üretir.
    """
    map_configs = []
    frost_by_night: List[Tuple[str, str]] = []
    any_frost = False

    for idx in range(8):
        night_start = start_date - timedelta(days=1) + timedelta(days=idx)
        morning_date = night_start + timedelta(days=1)
        target_iso_06 = f"{morning_date.isoformat()}T06:00"

        if night_start.month == morning_date.month:
            label_pair = f"{night_start.day:02d}-{morning_date.day:02d}.{morning_date.month:02d}.{morning_date.year}"
        else:
            label_pair = f"{night_start.day:02d}.{night_start.month:02d}-{morning_date.day:02d}.{morning_date.month:02d}.{morning_date.year}"

        frost_regions = []
        for p in REPRESENTATIVE_POINTS:
            pid = p["id"]
            h = hourly_by_point.get(pid, {})
            times = h.get("time", [])
            temps = h.get("temperature_2m", [])
            if target_iso_06 in times:
                t_i = times.index(target_iso_06)
                t_val = temps[t_i]
                if t_val is not None and t_val <= 0.5:
                    if p["region"] not in frost_regions:
                        frost_regions.append(p["region"])

        if frost_regions:
            any_frost = True
            frost_by_night.append((label_pair, ", ".join(frost_regions)))
        else:
            frost_by_night.append((label_pair, "-"))

        map_configs.append(
            {
                "index": idx + 1,
                "morning_date": morning_date.isoformat(),
                "hour": "06:00",
                "label_pair": label_pair,
                "caption_suffix": f"{label_pair} 06:00 (Yerel Saat) Sıcaklık Haritası",
            }
        )

    if not any_frost:
        summary_text = (
            "Tüm hafta boyunca, yurdumuz genelinde saat 06:00’daki sıcaklıklar incelendiğinde "
            "hafta boyunca yurt genelinde sıcaklıkların 0 °C üzerinde olması beklendiğinden "
            "don hadisesi beklenmemektedir"
        )
        all_dates_str = "\n".join(d for d, _ in frost_by_night[:7])
        table8_rows = [(all_dates_str, "-")]
    else:
        summary_text = (
            "Tüm hafta boyunca, yurdumuz genelinde saat 06:00’daki sıcaklıklar incelendiğinde "
            "iç ve doğu bölgelerimiz başta olmak üzere sıcaklığın 0 °C ve altına düştüğü "
            "kesimlerde zirai don riski bulunmaktadır"
        )
        table8_rows = []
        for d_lbl, reg_str in frost_by_night:
            if table8_rows and table8_rows[-1][1] == reg_str:
                prev_d, prev_r = table8_rows[-1]
                table8_rows[-1] = (f"{prev_d}\n{d_lbl}", prev_r)
            else:
                table8_rows.append((d_lbl, reg_str))

    return summary_text, table8_rows, map_configs


def build_full_report_draft(
    target_monday: date,
    custom_prev_docx=None,
    custom_mgm_sentence: Optional[str] = None,
) -> Dict:
    target_sunday = target_monday + timedelta(days=6)
    prev_monday = target_monday - timedelta(days=7)
    prev_sunday = target_monday - timedelta(days=1)

    prev_range_str = format_week_range_tr(prev_monday, prev_sunday, include_year=True)
    curr_range_str = format_week_range_tr(target_monday, target_sunday, include_year=True)

    prev_docx = custom_prev_docx or find_previous_week_docx(prev_monday)
    sec1_data = extract_section2_from_docx(prev_docx) if prev_docx else None

    fetch_start = prev_monday if not sec1_data else (target_monday - timedelta(days=1))
    fetch_end = target_sunday + timedelta(days=1)
    hourly_data = fetch_open_meteo_hourly(fetch_start, fetch_end)

    if not sec1_data:
        sec1_intro = (
            f"{prev_range_str} haftasında hava sıcaklıklarının yurt genelinde mevsim normalleri "
            f"civarında seyredeceği tahmin edilmekteydi. Günlere göre haftanın diğer meteorolojik "
            f"durumları ile ilgili açıklamalar ise şu şekildeydi:"
        )
        sec1_bullets = []
        for i in range(7):
            d = prev_monday + timedelta(days=i)
            lbl, txt, _ = synthesize_daily_bullet(d, i, hourly_data, past_tense=True)
            sec1_bullets.append((lbl, txt))
        sec1_source = "ECMWF Arşiv Verisi (Otomatik)"
    else:
        sec1_intro = sec1_data["intro"]
        sec1_bullets = sec1_data["bullets"]
        sec1_source = sec1_data["source_file"]

    if custom_mgm_sentence:
        mgm_temp_clause = custom_mgm_sentence.strip()
    else:
        normal_t = get_national_normal_temp(target_monday.month)
        all_temps = []
        for p in REPRESENTATIVE_POINTS:
            t_list = hourly_data.get(p["id"], {}).get("temperature_2m", [])
            all_temps.extend([x for x in t_list if x is not None])
        avg_forecast_t = sum(all_temps) / len(all_temps) if all_temps else normal_t
        diff = avg_forecast_t - normal_t

        if diff >= 2.0:
            mgm_temp_clause = "hava sıcaklıklarının ülkemiz genelinde mevsim normallerinin 2 ila 4 °C üzerinde seyredeceği"
        elif diff <= -2.0:
            mgm_temp_clause = "hava sıcaklıklarının ülkemiz genelinde azalarak mevsim normallerinin 2 ila 4 °C altında seyredeceği"
        else:
            mgm_temp_clause = "hava sıcaklıklarının ülkemiz genelinde mevsim normalleri civarında, yer yer üzerinde seyredeceği"

    sec2_intro = (
        f"{curr_range_str} haftasında {mgm_temp_clause} tahmin edilmektedir. "
        f"Günlere göre haftanın diğer meteorolojik durumları ile ilgili açıklamalar ise şu şekildedir:"
    )

    sec2_bullets = []
    rain_candidates = []
    cape_candidates = []

    for i in range(7):
        d = target_monday + timedelta(days=i)
        lbl, txt, day_maps = synthesize_daily_bullet(d, i, hourly_data, past_tense=False)
        sec2_bullets.append((lbl, txt))
        for m_item in day_maps:
            m_item["after_bullet_idx"] = i
            if m_item["type"] == "rain":
                rain_candidates.append(m_item)
            elif m_item["type"] == "cape":
                cape_candidates.append(m_item)

    # En yağışlı/yaygın en fazla 2 yağış haritası ve CAPE yüksek olan en belirgin 1-2 CAPE haritasını seç
    rain_candidates.sort(key=lambda x: (x.get("rain_regions_count", 0), x.get("max_precip", 0.0)), reverse=True)
    cape_candidates.sort(key=lambda x: x.get("max_cape", 0.0), reverse=True)

    sec2_extra_maps = sorted(
        rain_candidates[:2] + cape_candidates[:1],
        key=lambda x: x["after_bullet_idx"],
    )

    wind_warnings = analyze_wind_warnings(target_monday, hourly_data)
    frost_summary, table8_rows, temp_maps_config = analyze_frost_risk_0600(target_monday, hourly_data)

    prep_friday = target_monday - timedelta(days=3)
    prep_date_str = f"{prep_friday.day:02d} {MONTH_NAMES_TR[prep_friday.month]} {prep_friday.year}"

    return {
        "target_monday": target_monday.isoformat(),
        "target_sunday": target_sunday.isoformat(),
        "prev_monday": prev_monday.isoformat(),
        "prev_sunday": prev_sunday.isoformat(),
        "prep_date_str": prep_date_str,
        "prev_range_str": prev_range_str,
        "curr_range_str": curr_range_str,
        "normals_month": target_monday.month,
        "normals_month_name": MONTH_NAMES_TR[target_monday.month],
        "sec1_source": sec1_source,
        "sec1_intro": sec1_intro,
        "sec1_bullets": sec1_bullets,
        "sec2_intro": sec2_intro,
        "sec2_bullets": sec2_bullets,
        "sec2_extra_maps": sec2_extra_maps,
        "wind_warnings": wind_warnings,
        "frost_summary": frost_summary,
        "table8_rows": table8_rows,
        "temp_maps": temp_maps_config,
    }
