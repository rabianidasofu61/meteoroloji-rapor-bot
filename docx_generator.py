import copy
import shutil
from datetime import date
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

from config import get_template_docx, OUTPUT_DIR, DOCUMENTS_DIR, MONTH_NAMES_TR
from excel_service import load_seasonal_normals


def _set_run_font(run, size_pt: int = 12, bold: bool = False):
    run.font.name = "Arial"
    run.font.size = Pt(size_pt)
    run.bold = bold
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), "Arial")
    rFonts.set(qn("w:hAnsi"), "Arial")
    rFonts.set(qn("w:cs"), "Arial")


def _format_paragraph(
    p,
    align=WD_ALIGN_PARAGRAPH.JUSTIFY,
    line_spacing: float = 1.5,
    first_line_indent_pt: Optional[float] = None,
    left_indent_pt: Optional[float] = None,
    space_before_pt: float = 0,
    space_after_pt: float = 0,
    num_id: Optional[int] = None,
    ilvl: int = 0,
):
    p.alignment = align
    pf = p.paragraph_format
    pf.line_spacing = line_spacing
    pf.space_before = Pt(space_before_pt)
    pf.space_after = Pt(space_after_pt)
    if first_line_indent_pt is not None:
        pf.first_line_indent = Pt(first_line_indent_pt)
    if left_indent_pt is not None:
        pf.left_indent = Pt(left_indent_pt)

    if num_id is not None:
        pPr = p._p.get_or_add_pPr()
        numPr = OxmlElement("w:numPr")
        ilvl_el = OxmlElement("w:ilvl")
        ilvl_el.set(qn("w:val"), str(ilvl))
        numId_el = OxmlElement("w:numId")
        numId_el.set(qn("w:val"), str(num_id))
        numPr.append(ilvl_el)
        numPr.append(numId_el)
        pPr.append(numPr)


def _update_cover_page_sdt(body, prep_date_str: str, report_title: str):
    """
    Kapak sayfasındaki (w:sdt) tarih ve rapor başlığı metin kutularını günceller.
    TARSİM logosu ve AR-GE müdürlüğü yazısını olduğu gibi korur.
    """
    if len(body) < 2:
        return
    sdt = body[1]
    t_elems = sdt.findall(".//" + qn("w:t"))
    for t in t_elems:
        val = t.text or ""
        if "2026" in val and ("Şubat" in val or "Mayıs" in val) and "Değerlendirme" not in val:
            t.text = prep_date_str
        elif "Haftalık Meteorolojik Değerlendirme" in val:
            t.text = report_title


def _set_cell_borders(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = parse_xml(
        f"<w:tcBorders {nsdecls('w')}>"
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        "</w:tcBorders>"
    )
    tcPr.append(tcBorders)


def _create_normals_table(doc, rows_data: List[Tuple[str, str, str]]):
    """
    3. Bölümdeki bölgesel Mevsim Normalleri tablosunu (Tablo 1..7)
    orijinal rapordaki sütun genişlikleri, kenarlık ve fontlarla oluşturur.
    """
    tbl = doc.add_table(rows=1 + len(rows_data), cols=3)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER

    headers = ["İl Adı", "Sıcaklık (°C)", "Toplam Yağış Miktarı Ortalamaları (mm)"]
    col_widths = [Inches(1.44), Inches(1.07), Inches(2.01)]

    # Başlık satırı
    hdr_cells = tbl.rows[0].cells
    for i, h_text in enumerate(headers):
        cell = hdr_cells[i]
        cell.width = col_widths[i]
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        _set_cell_borders(cell)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if i > 0 else WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        r = p.add_run(h_text)
        _set_run_font(r, size_pt=11, bold=True)

    # Veri satırları
    for r_idx, (prov, temp_val, precip_val) in enumerate(rows_data, start=1):
        row_cells = tbl.rows[r_idx].cells
        vals = [prov, temp_val, precip_val]
        for c_i, val in enumerate(vals):
            cell = row_cells[c_i]
            cell.width = col_widths[c_i]
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            _set_cell_borders(cell)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_i > 0 else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(1)
            r = p.add_run(val)
            _set_run_font(r, size_pt=11, bold=(c_i == 0))

    return tbl._tbl


def _create_frost_table8(doc, table8_rows: List[Tuple[str, str]]):
    """
    5. Bölüm sonundaki Tablo 8 (Don Riski Beklenen Bölge(ler)) tablosunu oluşturur.
    """
    tbl = doc.add_table(rows=1 + len(table8_rows), cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER

    headers = ["Tarih (Gece)", "Beklenen Bölge(ler):"]
    col_widths = [Inches(2.1), Inches(4.2)]

    for i, h_text in enumerate(headers):
        cell = tbl.rows[0].cells[i]
        cell.width = col_widths[i]
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        _set_cell_borders(cell)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(h_text)
        _set_run_font(r, size_pt=12, bold=True)

    for r_idx, (dates_str, regions_str) in enumerate(table8_rows, start=1):
        row_cells = tbl.rows[r_idx].cells
        for c_i, val in enumerate([dates_str, regions_str]):
            cell = row_cells[c_i]
            cell.width = col_widths[c_i]
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            _set_cell_borders(cell)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            r = p.add_run(val)
            _set_run_font(r, size_pt=12, bold=False)

    return tbl._tbl


def generate_weekly_report_docx(draft: Dict, save_to_documents: bool = True) -> Path:
    """
    Verilen draft sözlüğünden TARSİM şablonuna birebir uygun Word (.docx)
    raporunu üretir ve dosya yolunu döndürür.
    """
    template_path = get_template_docx()
    start_d = date.fromisoformat(draft["target_monday"])
    end_d = date.fromisoformat(draft["target_sunday"])

    if start_d.month == end_d.month:
        file_name = f"{start_d.day:02d}-{end_d.day:02d}.{end_d.month:02d}.{end_d.year} Meteorolojik Rapor.docx"
    else:
        file_name = f"{start_d.day:02d}.{start_d.month:02d}-{end_d.day:02d}.{end_d.month:02d}.{end_d.year} Meteorolojik Rapor.docx"

    out_path = OUTPUT_DIR / file_name
    shutil.copy2(template_path, out_path)

    doc = docx.Document(str(out_path))
    body = doc._body._body

    # 1) Kapak sayfasını güncelle
    report_title = f"{draft['curr_range_str']} Haftalık Meteorolojik Değerlendirme"
    _update_cover_page_sdt(body, draft["prep_date_str"], report_title)

    # 2) Sabit 6. ve 7. bölümlerin başlangıç indeksini bul ('6) Zirai Don Hadisesinde Kritik Saatler')
    sec6_anchor = None
    to_remove = []
    for idx in range(2, len(body)):
        child = body[idx]
        tag = child.tag.split("}")[-1]
        if tag == "sectPr":
            continue
        txt = "".join(child.itertext()).strip()
        if txt.startswith("6) Zirai Don Hadisesinde Kritik Saatler"):
            sec6_anchor = child
            break
        to_remove.append(child)

    for el in to_remove:
        body.remove(el)

    # Yeni oluşturulan paragraf ve tabloları sırayla tutup sec6_anchor öncesine ekleyeceğiz
    new_elements = []

    def add_p(
        runs: List[Tuple[str, int, bool]],
        align=WD_ALIGN_PARAGRAPH.JUSTIFY,
        first_line_pt: Optional[float] = None,
        left_pt: Optional[float] = None,
        space_before_pt: float = 0,
        space_after_pt: float = 0,
        num_id: Optional[int] = None,
        ilvl: int = 0,
        page_break_before: bool = False,
    ):
        p = doc.add_paragraph()
        _format_paragraph(
            p,
            align=align,
            line_spacing=1.5,
            first_line_indent_pt=first_line_pt,
            left_indent_pt=left_pt,
            space_before_pt=space_before_pt,
            space_after_pt=space_after_pt,
            num_id=num_id,
            ilvl=ilvl,
        )
        if page_break_before:
            p.paragraph_format.page_break_before = True
        for text, sz, is_bold in runs:
            r = p.add_run(text)
            _set_run_font(r, size_pt=sz, bold=is_bold)
        new_elements.append(p._p)
        return p

    def add_image_p(img_path: str, width_inches: float = 6.5):
        p = doc.add_paragraph()
        _format_paragraph(p, align=WD_ALIGN_PARAGRAPH.CENTER, space_before_pt=6, space_after_pt=4)
        r = p.add_run()
        r.add_picture(img_path, width=Inches(width_inches))
        new_elements.append(p._p)
        return p

    normals_month = int(draft.get("normals_month", start_d.month))
    normals_month_name = MONTH_NAMES_TR[normals_month]

    # --- İÇİNDEKİLER ---
    add_p([("İçindekiler", 14, True)], space_before_pt=6, space_after_pt=6)
    toc_lines = [
        f"1) {draft['prev_range_str']} Haftalık Meteorolojik Değerlendirme",
        f"2) {draft['curr_range_str']} Haftalık Meteorolojik Değerlendirme",
        f"3) {normals_month_name} Ayı Mevsim Normalleri Sıcaklık ve Yağış Miktarları",
        "4) Önemli Uyarılar",
        f"5) {draft['curr_range_str']} 06:00  (Lokal Saat) Türkiye Sıcaklık Haritaları",
        "6) Zirai Don Hadisesinde Kritik Saatler",
        "7) Yağış, Rüzgâr Hızı ve Hava Kütleleri Sınıflandırması",
    ]
    for line in toc_lines:
        add_p([(line, 12, False)], space_before_pt=3, space_after_pt=3)

    add_p([("", 12, False)])

    # --- 1. BÖLÜM: GEÇEN HAFTA ---
    add_p([(f"1) {draft['prev_range_str']} Haftalık Meteorolojik Değerlendirme", 14, True)], space_before_pt=6, space_after_pt=4)
    add_p([(draft["sec1_intro"], 12, False)], first_line_pt=18)
    for d_label, desc in draft["sec1_bullets"]:
        add_p([(f"{d_label}: ", 12, True), (desc, 12, False)], num_id=24, ilvl=0)

    add_p([("", 12, False)])

    # --- 2. BÖLÜM: TAHMİN HAFTASI ---
    add_p([(f"2)  {draft['curr_range_str']} Haftalık Meteorolojik Değerlendirme", 14, True)], space_before_pt=6, space_after_pt=4)
    add_p([(draft["sec2_intro"], 12, False)], first_line_pt=18)

    fig_counter = 1
    extra_maps_by_idx = {}
    for em in draft.get("sec2_extra_maps", []):
        b_idx = em.get("after_bullet_idx", -1)
        extra_maps_by_idx.setdefault(b_idx, []).append(em)

    for b_idx, (d_label, desc) in enumerate(draft["sec2_bullets"]):
        add_p([(f"{d_label}: ", 12, True), (desc, 12, False)], num_id=24, ilvl=0)
        if b_idx in extra_maps_by_idx:
            for em in extra_maps_by_idx[b_idx]:
                img_p = em.get("image_path")
                if img_p and Path(img_p).exists():
                    add_image_p(img_p, width_inches=6.2)
                    add_p(
                        [(f"Şekil {fig_counter}. {em['caption_suffix']}", 12, True)],
                        align=WD_ALIGN_PARAGRAPH.LEFT,
                        space_after_pt=6,
                    )
                    fig_counter += 1

    add_p([("", 12, False)])

    # --- 3. BÖLÜM: MEVSİM NORMALLERİ ---
    add_p(
        [(f"3) {normals_month_name} Ayı Mevsim Normalleri Sıcaklık ve Yağış Miktarları", 14, True)],
        space_before_pt=10,
        space_after_pt=4,
        page_break_before=True,
    )
    add_p(
        [
            (
                f"{normals_month_name} ayı için mevsim normalleri (1991-2020) sıcaklıkları ve aylık toplam yağış miktarı ortalamaları bölgesel olarak aşağıdaki tablolar gibidir:",
                12,
                False,
            )
        ],
        first_line_pt=35,
        space_after_pt=6,
    )

    normals_tables = load_seasonal_normals(normals_month)
    for t_info in normals_tables:
        t_no = t_info["table_no"]
        reg_title = t_info["region_title"]
        add_p(
            [
                (f"Tablo {t_no}. ", 12, True),
                (f"{reg_title} Ortalama Sıcaklıkları (°C) ve Aylık Toplam Yağış Miktarı Ortalamaları (mm)", 12, False),
            ],
            align=WD_ALIGN_PARAGRAPH.CENTER,
            space_before_pt=6,
            space_after_pt=2,
        )
        tbl_el = _create_normals_table(doc, t_info["rows"])
        new_elements.append(tbl_el)
        add_p([("", 10, False)])

    # --- 4. BÖLÜM: ÖNEMLİ UYARILAR ---
    add_p([("4) Önemli Uyarılar", 14, True)], space_before_pt=10, space_after_pt=4)
    add_p([("1. Rüzgâr & Fırtına:  ", 12, True)], space_after_pt=2)

    wind_sections = [
        ("1.1 ", "Marmara"),
        ("1.2 ", "Ege"),
        ("1.3 ", "Akdeniz"),
        ("1.4 ", "Karadeniz"),
    ]
    wind_warnings = draft.get("wind_warnings", {})
    for num_prefix, reg_key in wind_sections:
        reg_bullets = wind_warnings.get(reg_key, [])
        if not reg_bullets:
            add_p(
                [
                    (f"{num_prefix}{reg_key}: ", 12, True),
                    (" Hafta içerisinde herhangi bir fırtına riski bulunmamaktadır.", 12, False),
                ],
                left_pt=36,
            )
        else:
            add_p([(f"{num_prefix}{reg_key}: ", 12, True)], left_pt=36)
            for wb in reg_bullets:
                if ":" in wb:
                    w_date, w_rest = wb.split(":", 1)
                    add_p([(f"{w_date.strip()}: ", 12, True), (w_rest.strip(), 12, False)], num_id=31, ilvl=1)
                else:
                    add_p([(wb, 12, False)], num_id=31, ilvl=1)

    add_p([("", 12, False)])

    # --- 5. BÖLÜM: 06:00 TÜRKİYE SICAKLIK HARİTALARI & TABLO 8 ---
    add_p(
        [(f"5)  {draft['curr_range_str']} 06:00  (Lokal Saat) Türkiye Sıcaklık Haritaları", 14, True)],
        space_before_pt=10,
        space_after_pt=4,
        page_break_before=True,
    )
    add_p(
        [
            (
                "Sıcaklık haritalarının değerlendirilmesinde sabah 06:00 (yerel saat) dikkate alınmıştır. Bunun nedeni günlük sıcaklık değerlerinin en düşük olduğu saat aralığı gün doğumundan birkaç saat sonrasında gerçekleşir. İçerisinde bulunduğumuz dönem itibari ile gün doğumu saati değerlendirildiğinde yaklaşık olarak 06:00 saatinin uygun olunabileceği düşünülmektedir.",
                12,
                False,
            )
        ],
        first_line_pt=35,
    )
    add_p(
        [
            (
                "Gökyüzünün bulutlu olup olmaması don ihtimali üzerinde belirleyici bir faktördür. Eğer 0 °C altında bulutluluk bulunuyorsa burada radyasyon kaybı açık gökyüzüyle kıyasla daha az olacağından, don riski daha düşük olacaktır. Don riski, özellikle erken çiçeklenen meyve türlerinde verim kayıplarına neden olabilir.",
                12,
                False,
            )
        ],
        first_line_pt=35,
        space_after_pt=6,
    )

    first_temp_fig = fig_counter
    for tm in draft.get("temp_maps", []):
        img_p = tm.get("image_path")
        if img_p and Path(img_p).exists():
            add_image_p(img_p, width_inches=6.5)
        add_p(
            [(f"Şekil {fig_counter}. {tm['caption_suffix']}", 12, True)],
            align=WD_ALIGN_PARAGRAPH.LEFT,
            space_after_pt=6,
        )
        fig_counter += 1
    last_temp_fig = fig_counter - 1

    frost_summary = draft.get("frost_summary", "").rstrip(".")
    fig_ref = f" (Şekil {first_temp_fig}-{last_temp_fig})." if last_temp_fig >= first_temp_fig else "."
    add_p([(f"{frost_summary}{fig_ref}", 12, False)], first_line_pt=35, space_before_pt=6, space_after_pt=6)

    add_p(
        [
            ("Tablo 8.", 12, True),
            (f"  {draft['curr_range_str']} Don Riski Beklenen Bölge(ler)", 12, False),
        ],
        align=WD_ALIGN_PARAGRAPH.CENTER,
        space_before_pt=6,
        space_after_pt=2,
    )
    tbl8_el = _create_frost_table8(doc, draft.get("table8_rows", []))
    new_elements.append(tbl8_el)
    add_p([("", 12, False)])

    # Tüm yeni elemanları 6. Bölümün (sec6_anchor) hemen önüne taşı
    if sec6_anchor is not None:
        for el in new_elements:
            sec6_anchor.addprevious(el)

    doc.save(str(out_path))

    # Ayrıca Haftalık Rapor Bot Documents klasörüne de kopyala ki bir sonraki hafta otomatik bulunsun
    if save_to_documents:
        try:
            shutil.copy2(out_path, DOCUMENTS_DIR / file_name)
        except Exception:
            pass

    return out_path
