import os
import sys
import subprocess
import docx
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def create_master_proposal():
    doc = docx.Document()
    
    # 1. Page Setup: F4 (Folio) 21.5 cm x 33.0 cm
    section = doc.sections[0]
    section.page_width = Cm(21.5)
    section.page_height = Cm(33.0)
    section.top_margin = Cm(2.54)
    section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(2.54)
    section.right_margin = Cm(2.54)
    
    # Enable different first page for clean cover page
    section.different_first_page_header_footer = True
    
    # 2. Executive Color Palette
    COLOR_PRIMARY = RGBColor(27, 54, 93)     # Deep Executive Navy #1B365D
    COLOR_SECONDARY = RGBColor(61, 43, 31)   # Rich Warm Coffee #3D2B1F
    COLOR_TEXT = RGBColor(45, 45, 45)        # Charcoal Text #2D2D2D
    COLOR_MUTED = RGBColor(105, 115, 125)    # Neutral Slate #69737D
    COLOR_GOLD = RGBColor(185, 130, 10)      # Amber Gold #B9820A
    COLOR_WHITE = RGBColor(255, 255, 255)
    COLOR_SUCCESS = RGBColor(25, 120, 50)    # Forest Green for profit margins
    
    # Base Normal Style
    style_normal = doc.styles['Normal']
    style_normal.font.name = 'Arial'
    style_normal.font.size = Pt(10)
    style_normal.font.color.rgb = COLOR_TEXT
    style_normal.paragraph_format.line_spacing = 1.15
    style_normal.paragraph_format.space_after = Pt(3.5)
    
    # Setup Footer for subsequent pages
    footer = section.footer
    f_p = footer.paragraphs[0]
    f_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    
    f_run_left = f_p.add_run("ZeroSpot Cafe — Proposal Investasi Bisnis  |  ")
    f_run_left.font.name = 'Arial'
    f_run_left.font.size = Pt(8.5)
    f_run_left.font.color.rgb = COLOR_MUTED
    
    f_run_page = f_p.add_run("Halaman ")
    f_run_page.font.name = 'Arial'
    f_run_page.font.size = Pt(8.5)
    f_run_page.font.color.rgb = COLOR_MUTED
    
    # Dynamic Page
    fld1 = parse_xml(r'<w:fldChar %s w:fldCharType="begin"/>' % nsdecls('w'))
    inst1 = parse_xml(r'<w:instrText %s xml:space="preserve"> PAGE </w:instrText>' % nsdecls('w'))
    fld2 = parse_xml(r'<w:fldChar %s w:fldCharType="separate"/>' % nsdecls('w'))
    fld3 = parse_xml(r'<w:fldChar %s w:fldCharType="end"/>' % nsdecls('w'))
    f_run_page._r.append(fld1)
    f_run_page._r.append(inst1)
    f_run_page._r.append(fld2)
    f_run_page._r.append(fld3)
    
    f_run_mid = f_p.add_run(" dari ")
    f_run_mid.font.name = 'Arial'
    f_run_mid.font.size = Pt(8.5)
    f_run_mid.font.color.rgb = COLOR_MUTED
    
    # Dynamic NumPages
    fld4 = parse_xml(r'<w:fldChar %s w:fldCharType="begin"/>' % nsdecls('w'))
    inst2 = parse_xml(r'<w:instrText %s xml:space="preserve"> NUMPAGES </w:instrText>' % nsdecls('w'))
    fld5 = parse_xml(r'<w:fldChar %s w:fldCharType="separate"/>' % nsdecls('w'))
    fld6 = parse_xml(r'<w:fldChar %s w:fldCharType="end"/>' % nsdecls('w'))
    f_run_mid._r.append(fld4)
    f_run_mid._r.append(inst2)
    f_run_mid._r.append(fld5)
    f_run_mid._r.append(fld6)

    # Header for subsequent pages
    header = section.header
    h_p = header.paragraphs[0]
    h_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    h_run = h_p.add_run("DOKUMEN PENAWARAN INVESTASI & KEMITRAAN  —  ZEROSPOT CAFE CIKEDAL")
    h_run.font.name = 'Arial'
    h_run.font.size = Pt(8)
    h_run.font.color.rgb = RGBColor(160, 160, 160)

    # Helper Functions
    def add_p(text="", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=3.5, line_spacing=1.15, bold=False, italic=False, color=COLOR_TEXT, size=Pt(10)):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = line_spacing
        if text:
            r = p.add_run(text)
            r.font.name = 'Arial'
            r.font.size = size
            r.font.bold = bold
            r.font.italic = italic
            r.font.color.rgb = color
        return p

    def add_h1(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(text)
        r.font.name = 'Arial'
        r.font.size = Pt(13.5)
        r.font.bold = True
        r.font.color.rgb = COLOR_PRIMARY
        return p

    def add_h2(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(2.5)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(text)
        r.font.name = 'Arial'
        r.font.size = Pt(11)
        r.font.bold = True
        r.font.color.rgb = COLOR_SECONDARY
        return p

    def add_bullet(bold_prefix, text):
        p = doc.add_paragraph(style='List Bullet')
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.15
        if bold_prefix:
            r_b = p.add_run(bold_prefix + ": ")
            r_b.font.name = 'Arial'
            r_b.font.size = Pt(10)
            r_b.font.bold = True
            r_b.font.color.rgb = COLOR_PRIMARY
        r_t = p.add_run(text)
        r_t.font.name = 'Arial'
        r_t.font.size = Pt(10)
        r_t.font.color.rgb = COLOR_TEXT
        return p

    def set_cell_shading(cell, color_hex):
        shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
        cell._tc.get_or_add_tcPr().append(shading)

    def set_cell_padding(cell, top=60, bottom=60, left=90, right=90):
        tcPr = cell._tc.get_or_add_tcPr()
        tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
        tcPr.append(tcMar)

    def set_cell_borders(cell, top=None, bottom=None, left=None, right=None):
        tcPr = cell._tc.get_or_add_tcPr()
        borders = parse_xml(f'<w:tcBorders {nsdecls("w")}/>')
        for side, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
            if val:
                b = parse_xml(f'<w:{side} {nsdecls("w")} w:val="{val.get("val","single")}" w:sz="{val.get("sz","4")}" w:space="0" w:color="{val.get("color","CCCCCC")}"/>')
                borders.append(b)
            else:
                b = parse_xml(f'<w:{side} {nsdecls("w")} w:val="none"/>')
                borders.append(b)
        tcPr.append(borders)

    def set_row_props(row, is_header=False):
        trPr = row._tr.get_or_add_trPr()
        cantSplit = parse_xml(f'<w:cantSplit {nsdecls("w")}/>')
        trPr.append(cantSplit)
        if is_header:
            tblHeader = parse_xml(f'<w:tblHeader {nsdecls("w")}/>')
            trPr.append(tblHeader)

    def add_callout(title, body_text):
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = tbl.cell(0, 0)
        cell.width = Cm(16.42)
        set_cell_shading(cell, "F3F6FA")
        set_cell_padding(cell, top=80, bottom=80, left=140, right=140)
        set_cell_borders(cell, left={"val": "single", "sz": "20", "color": "1B365D"},
                               top={"val": "single", "sz": "4", "color": "D8E2EC"},
                               bottom={"val": "single", "sz": "4", "color": "D8E2EC"},
                               right={"val": "single", "sz": "4", "color": "D8E2EC"})
        
        # Paragraph judul callout (Rata Kiri, Bold, Warna Navy)
        p_title = cell.paragraphs[0]
        p_title.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p_title.paragraph_format.space_before = Pt(0)
        p_title.paragraph_format.space_after = Pt(3)
        p_title.paragraph_format.line_spacing = 1.15
        r_title = p_title.add_run(title)
        r_title.font.name = 'Arial'
        r_title.font.size = Pt(9.5)
        r_title.font.bold = True
        r_title.font.color.rgb = COLOR_PRIMARY
        
        # Baris butir isi (Setiap baris menjadi paragraf rata kiri tersendiri)
        lines = [line.strip() for line in body_text.strip().split('\n') if line.strip()]
        for i, line in enumerate(lines):
            p = cell.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(1.5) if i < len(lines) - 1 else Pt(0)
            p.paragraph_format.line_spacing = 1.15
            
            if ":" in line and not line.startswith("http"):
                prefix, rest = line.split(":", 1)
                r_pre = p.add_run(prefix + ":")
                r_pre.font.name = 'Arial'
                r_pre.font.size = Pt(8.5)
                r_pre.font.bold = True
                r_pre.font.color.rgb = COLOR_PRIMARY
                
                r_rest = p.add_run(rest)
                r_rest.font.name = 'Arial'
                r_rest.font.size = Pt(8.5)
                r_rest.font.color.rgb = COLOR_TEXT
            elif line.startswith(("1.", "2.", "3.", "4.", "5.")):
                num, rest = line.split(" ", 1)
                r_num = p.add_run(num + " ")
                r_num.font.name = 'Arial'
                r_num.font.size = Pt(8.5)
                r_num.font.bold = True
                r_num.font.color.rgb = COLOR_PRIMARY
                
                r_rest = p.add_run(rest)
                r_rest.font.name = 'Arial'
                r_rest.font.size = Pt(8.5)
                r_rest.font.color.rgb = COLOR_TEXT
            else:
                r = p.add_run(line)
                r.font.name = 'Arial'
                r.font.size = Pt(8.5)
                r.font.color.rgb = COLOR_TEXT
        
        p_spacer = doc.add_paragraph()
        p_spacer.paragraph_format.space_before = Pt(0)
        p_spacer.paragraph_format.space_after = Pt(2)

    # ==================== PAGE 1: COVER PAGE ====================
    p_cov_space = doc.add_paragraph()
    p_cov_space.paragraph_format.space_before = Pt(36)
    
    add_p("DOKUMEN RAHASIA & PROPOSAL INVESTASI STRATEGIS", align=WD_ALIGN_PARAGRAPH.CENTER, space_after=8, bold=True, color=COLOR_PRIMARY, size=Pt(11))
    add_p("PROPOSAL INVESTASI & PENGEMBANGAN BISNIS", align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4, bold=True, color=COLOR_PRIMARY, size=Pt(19))
    add_p("ZEROSPOT CAFE & ARTISAN ANGKRINGAN", align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6, bold=True, color=COLOR_SECONDARY, size=Pt(21))
    add_p("Satu Tempat, Satu Rasa, Satu Cerita", align=WD_ALIGN_PARAGRAPH.CENTER, space_after=14, italic=True, bold=True, color=COLOR_GOLD, size=Pt(13))
    
    div_tbl = doc.add_table(rows=1, cols=1)
    div_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    div_cell = div_tbl.cell(0, 0)
    div_cell.width = Cm(16.42)
    set_cell_shading(div_cell, "1B365D")
    set_cell_padding(div_cell, top=8, bottom=8, left=0, right=0)
    div_cell.paragraphs[0].text = ""
    
    add_p(
        "Inovasi Bisnis F&B Pedesaan: Memadukan Konsep Cafe Modern Perkotaan dengan Angkringan Tradisional (Sate Taichan Bakar, Nasi Bakar Daun Pisang, Olahan Susu Sapi Murni Segar, dan Furnitur Outdoor Camping Portable) DIPADUKAN DENGAN Bar Kopi Barista Profesional & Manual Brew Berbasis Biji Kopi Lokal Banten serta Efisiensi Model Sewa Lahan Terbuka.",
        align=WD_ALIGN_PARAGRAPH.CENTER, space_before=14, space_after=24, italic=False, color=COLOR_TEXT, size=Pt(10)
    )
    
    meta_tbl = doc.add_table(rows=6, cols=2)
    meta_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_widths = [Cm(5.2), Cm(11.22)]
    meta_data = [
        ("Nama Entitas Usaha", "ZeroSpot Cafe & Outdoor Camp Angkringan"),
        ("Lokasi Basis Operasional", "Kecamatan Cikedal, Kabupaten Pandeglang, Banten"),
        ("Inisiator & Pengelola", "Fatih Farhat Asshidiq (Founder & Managing Director)"),
        ("Kontak Resmi Inisiator", "Telp/WA: +68 821-1150-0190 / +62 895-0610-0075 | Email: kazokuhairy@gmail.com"),
        ("Skema Kerja Sama", "Mudharabah Muqayyadah / Bagi Hasil Laba Bersih (Nisbah 40% : 60%)"),
        ("Target Nilai Investasi & ROI", "Rp 40.000.000 (Target Balik Modal 10 Bulan Skenario Moderat)")
    ]
    for i, (k, v) in enumerate(meta_data):
        c0 = meta_tbl.cell(i, 0)
        c1 = meta_tbl.cell(i, 1)
        c0.width, c1.width = meta_widths[0], meta_widths[1]
        set_cell_shading(c0, "F2F5F9")
        set_cell_shading(c1, "FAFAFA")
        set_cell_padding(c0, top=55, bottom=55, left=90, right=90)
        set_cell_padding(c1, top=55, bottom=55, left=90, right=90)
        set_cell_borders(c0, bottom={"val": "single", "sz": "4", "color": "E2E8F0"})
        set_cell_borders(c1, bottom={"val": "single", "sz": "4", "color": "E2E8F0"})
        set_row_props(meta_tbl.rows[i])
        
        p0 = c0.paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r0 = p0.add_run(k)
        r0.font.name = 'Arial'
        r0.font.size = Pt(8.5)
        r0.font.bold = True
        r0.font.color.rgb = COLOR_PRIMARY
        
        p1 = c1.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r1 = p1.add_run(v)
        r1.font.name = 'Arial'
        r1.font.size = Pt(8.5)
        r1.font.bold = (i in [2, 3, 5])
        r1.font.color.rgb = COLOR_TEXT
        
    add_p("Diterbitkan di Pandeglang, Banten  |  Kuartal IV 2026", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=30, space_after=0, italic=True, color=COLOR_MUTED, size=Pt(8.5))

    doc.add_page_break()

    # ==================== PAGE 2: EXECUTIVE SUMMARY & PASAR ====================
    add_h1("1. RINGKASAN EKSEKUTIF (EXECUTIVE SUMMARY)")
    add_p(
        "ZeroSpot Cafe & Artisan Angkringan adalah inisiatif bisnis ruang komunal pedesaan (third place) inovatif yang berlokasi di sekitar Kecamatan Cikedal, Pandeglang, Banten. ZeroSpot hadir sebagai pionir yang memadukan kehangatan angkringan tradisional dengan kenyamanan cafe modern perkotaan. Mengusung konsep outdoor street camping (kursi lipat kemah TrailTop dan meja roll-top hitam), tempat ini menjadi simpul temu favorit yang santai bagi pemuda desa, santri, pekerja, hingga keluarga."
    )
    add_p(
        "ZeroSpot Cafe menghadirkan Diferensiasi Kunci yang belum pernah ada di Pandeglang: Stasiun Kopi Barista Profesional (Mesin Espresso Listrik Ferratti Ferro 15 Bar & Manual Brew V60) yang menyeduh biji kopi Robusta asli Gunung Karang Pandeglang dan blend lokal, bersanding dengan aneka sate bakar dan gorengan hangat. Perpaduan ini melahirkan menu andalan 'Kopi Susu Aren Spesial ZeroSpot' (Espresso Robusta + Susu Murni Segar + Gula Aren Asli Malingping)."
    )
    add_p("Kekuatan fundamental ZeroSpot Cafe bertumpu pada 4 pilar bisnis:")
    
    add_bullet("Ambiance Outdoor Street Camping", "Menghadirkan 24 kursi lipat camping portable TrailTop dan 6 meja roll-top hitam di area lahan terbuka yang sangat santai, estetik, dan Instagramable.")
    add_bullet("Gerobak Custom & Live Infrared Roaster", "Gerobak kayu jati/mahoni beretalase kaca lampu warm white dengan pemanggang gas inframerah tanpa asap yang memanggang sate taichan dan aneka sate higienis.")
    add_bullet("Menu Autentik Sate, Susu & Kopi Barista", "Kombinasi Sate Taichan bakar pedas jeruk limau, nasi bakar daun pisang, aneka olahan susu murni segar, dan kopi specialty barista.")
    add_bullet("Efisiensi Sewa Lahan Terbuka (Bukan Ruko)", "Memilih sewa lahan terbuka/pekarangan pinggir jalan yang berbiaya sangat terjangkau dibanding ruko komersial mahal, memangkas biaya tetap dan mempercepat pencapaian titik impas (BEP).")
    
    add_callout(
        "IKHTISAR NILAI INVESTASI & POTENSI PENGEMBALIAN (CAPITAL TARGET: RP 40.000.000)",
        "Total Kebutuhan Permodalan: Rp 40.000.000 (Opsi 1 Investor Penuh, 2 Slot @ Rp 20 Jt, 4 Slot @ Rp 10 Jt, atau 5 Slot Sindikasi @ Rp 8 Jt).\n"
        "Alokasi Belanja Modal Utama: Sewa Lahan 1 Tahun & Kanopi All-Weather (Rp 9,5 Jt), Gerobak Custom & Panggangan Sate Infrared (Rp 7,0 Jt), Furnitur Camping Kursi & Meja (Rp 4,44 Jt), Stasiun Bar Kopi Listrik Ferratti Ferro 15 Bar & Manual Brew (Rp 4,7 Jt), Showcase Chiller & Wadah Saji (Rp 3,21 Jt), Listrik 2200VA/Sound/Deposit (Rp 3,75 Jt), Stok Bahan Baku Awal & Kas Runway (Rp 7,4 Jt).\n"
        "Skema Akad Kerja Sama: Mudharabah Muqayyadah (Bagi Hasil Laba Bersih: 40% Investor : 60% Pengelola Operasional).\n"
        "Target Pengembalian Modal (Payback Period): 10,0 Bulan (Skenario Moderat) hingga 5,1 Bulan (Skenario Agresif).\n"
        "Proyeksi Dividen Bulanan Investor: Rp 4.000.000 (Moderat) hingga Rp 7.840.000 (Agresif) per bulan setelah masa rintisan."
    )

    add_h1("2. ANALISIS PASAR, GEOGRAFIS & DAYA BELI LOKAL")
    add_h2("2.1 Keunggulan Geografis Koridor Cikedal - Menes")
    add_p(
        "Kecamatan Cikedal berada di persimpangan strategis koridor tengah Kabupaten Pandeglang, berdampingan langsung dengan Kecamatan Menes sebagai pusat pendidikan pesantren (Mathla'ul Anwar dan UNMA). Jalur Cikedal juga merupakan perlintasan alternatif wisatawan menuju kawasan pantai Labuan, Carita, dan Tanjung Lesung. Suasana pedesaan yang sejuk, pemandangan persawahan dan Situ Cikedal menjadikan lokasi ini sangat ideal untuk konsep outdoor night coffee camp."
    )
    
    add_h2("2.2 Segmentasi Pasar & Bauran Konsumen")
    add_bullet("Pemuda Desa & Komunitas Motor (Pangsa 40%)", "Mencari spot nongkrong malam yang santai, bebas asap rokok pengap, ada WiFi, stopkontak, dan sate taichan murah.")
    add_bullet("Santri Senior & Mahasiswa UNMA/Menes (Pangsa 25%)", "Konsumen setia kopi seduh malam hari, wedang jahe susu penghangat badan, dan nasi bakar porsi hemat.")
    add_bullet("Pegawai Kantor, Guru & Aparatur Desa (Pangsa 20%)", "Makan malam santai bersama rekan kerja menikmati sate bakar hangat, mendoan panas, dan susu segar.")
    add_bullet("Wisatawan & Pelintas Akhir Pekan (Pangsa 15%)", "Keluarga dan pelancong yang singgah menikmati sensasi nongkrong street camp dan kopi susu gula aren Banten.")

    add_h2("2.3 Daya Beli & Psikologi Harga Mikro-Transaksional")
    add_p(
        "Rata-rata pengeluaran nongkrong di Cikedal-Menes adalah Rp 12.000 - Rp 25.000 per orang. ZeroSpot menerapkan rentang harga bersahabat: Sate usus/kulit Rp 2.500, Sate Taichan porsi Rp 13.000, Nasi bakar Rp 8.000, Susu murni mug Rp 9.000, serta Kopi Susu Aren & Manual Brew Rp 15.000. Rentang harga ini sangat mudah dijangkau namun menghasilkan margin keuntungan tinggi melalui volume pemesanan ganda."
    )

    # ==================== SECTION 3: KONSEP BISNIS ====================
    add_h1("3. KONSEP BISNIS: PERPADUAN CAFE MODERN & ANGKRINGAN TRADISIONAL")
    add_h2("3.1 Tata Ruang 4 Zona Terpadu (Outdoor Street Camp)")
    add_p(
        "Mengadopsi tata letak terbuka yang memaksimalkan sirkulasi udara pedesaan Cikedal di area lahan terbuka yang asri dan representatif:"
    )
    
    add_bullet("Zona 1: Gerobak Kayu Custom & Live Infrared Grill", "Gerobak kayu jati/mahoni klasik-modern beratap lengkung dengan etalase kaca display sate berpenerangan lampu LED warm white (3000K). Di atas meja gerobak terpasang mesin gas infrared roaster tanpa asap untuk memanggang sate taichan dan sate angkringan secara higienis.")
    add_bullet("Zona 2: Stasiun Barista Espresso Mesin & Showcase Chiller", "Bar counter kayu tempat barista meracik espresso mesin Ferratti Ferro 15 Bar (cappuccino, latte art, es kopi susu aren), manual brew V60, cup sealer, serta showcase chiller pendingin susu segar dan sate.")
    add_bullet("Zona 3: Prasmanan Nasi Bungkus & Gorengan Hangat", "Meja display keranjang nasi bakar daun pisang, nasi kucing kertas coklat, dan kuali penggorengan mendoan panas dadakan.")
    add_bullet("Zona 4: The Camping Yard (Area Duduk Outdoor)", "24 unit kursi lipat camping portable TrailTop dan 6 meja roll-top alumunium hitam beralas paving/kerikil koral dengan lampu gantung pijar festoon, proyektor nobar, dan akustik santai.")

    add_h2("3.2 Keunggulan Kompetitif (Parit Pertahanan Bisnis)")
    add_bullet(
        "Perpaduan Cafe Modern & Angkringan Tradisional",
        "Satu-satunya konsep bisnis di Pandeglang yang memadukan konsep cafe modern perkotaan dengan angkringan tradisional. Pengunjung bisa menikmati kopi hitam dan gorengan hangat tanpa harus berpindah dari cafe ke angkringan berbeda."
    )
    add_bullet("Sewa Lahan Terbuka yang Efisien & Fleksibel", "Memilih sewa lahan terbuka ketimbang ruko komersial, menekan biaya sewa hingga 70% lebih hemat, menghindarkan biaya fit-out bangunan yang mahal, serta memaksimalkan sirkulasi udara alami untuk konsep outdoor street camp.")
    add_bullet("Estetika Street Camping Kekinian", "Set kursi meja camping TrailTop menciptakan daya tarik visual tinggi yang viral di media sosial pemuda lokal.")

    # ==================== SECTION 4: CAPEX 40 JUTA ====================
    add_h1("4. RENCANA ANGGARAN BIAYA & ALOKASI BELANJA MODAL (CAPEX)")
    add_p(
        "Pendanaan modal sebesar Rp 40.000.000 dialokasikan secara presisi dan terukur untuk sewa lahan terbuka 1 tahun di muka, penyiapan kanopi all-weather, pengadaan gerobak custom, pemanggang infrared roaster, set kursi-meja camping, mesin espresso semi-komersial Ferratti Ferro 15 Bar, chiller pendingin, serta modal kerja awal:"
    )

    add_h2("4.1 Rincian Alokasi Belanja Modal Investasi (Rp 40.000.000)")
    
    capex_items = [
        ("Sewa Lahan Terbuka Strategis (1 Tahun di Muka)", "Sewa pekarangan terbuka pinggir jalan koridor Menes-Cikedal (12 bln)", "1 tahun", "Rp 5.000.000"),
        ("Penyiapan Lahan & Saung Kanopi All-Weather", "Pondasi koral split bebas becek, saung atap rustic & terpal penahan angin", "1 paket", "Rp 4.500.000"),
        ("Gerobak Angkringan Kayu Custom Mahoni", "P 200cm x L 85cm x T 195cm, atap lengkung, roda, laci kasir", "1 unit", "Rp 4.200.000"),
        ("Etalase Kaca Display Sate + Strip LED Warm", "Kaca 5mm 2 rak, pintu geser higienis, lampu LED tube 3000K", "1 set", "Rp 650.000"),
        ("Gas Infrared Smokeless BBQ Roaster 4 Burner", "Panggangan sate gas keramik tanpa asap pekat (Fomac/Getra)", "1 unit", "Rp 1.150.000"),
        ("Peralatan Masak Dapur & Wedangan Lengkap", "2 Ceret stainless 5L, kuali baja cor mendoan, kompor gas, regulator, 2 LPG", "1 paket", "Rp 1.000.000"),
        ("Kursi Lipat Camping Portable (TrailTop Style)", "Rangka pipa baja silang X hitam, kain Oxford 600D, cup holder", "24 unit", "Rp 2.400.000"),
        ("Meja Lipat Alumunium Roll-Top Hitam", "Black alloy roll-up table (95x55x50 cm) anti karat & tahan panas", "6 unit", "Rp 1.680.000"),
        ("Karpet Lipat Outdoor Lesehan Cadangan", "Karpet spons lipat waterproof untuk area lesehan santai", "3 lembar", "Rp 360.000"),
        ("Mesin Espresso Listrik Ferratti Ferro FCM3605", "15 Bar, 1450W, steam wand latte art profesional, tangki 1.7L, fast serve", "1 unit", "Rp 2.950.000"),
        ("Grinder Kopi Listrik N600 Burr 60mm", "Penggiling biji kopi presisi hemat listrik 150 Watt", "1 unit", "Rp 650.000"),
        ("Alat Seduh Manual Brew (V60, French Press, Vietnam)", "2 set V60 Hario + server, 4 Vietnam Drip stainless, 1 French Press", "1 paket", "Rp 550.000"),
        ("Kettle Gooseneck, Milk Jug Stainless & Scale", "Teko leher angsa 1L, milk jug frothing 350ml/600ml + timbangan 0.1g", "1 paket", "Rp 550.000"),
        ("Showcase Chiller Pendingin Susu Segar 120L", "Pendingin kaca transparan untuk susu sapi murni & marinasi sate", "1 unit", "Rp 2.400.000"),
        ("Wadah Saji Piring Seng, Mug Bir & Piring Anyam", "Piring seng taichan blirik, mug bir kaca bertangkai, piring rotan & baki", "1 paket", "Rp 810.000"),
        ("Instalasi Listrik Daya 2200VA, Festoon & Sign", "Kabel waterproof NYYHY tebal daya 1450W, MCB box, festoon warm, neon sign", "1 paket", "Rp 1.500.000"),
        ("Mini Projector Nobar & Sound System Akustik", "Proyektor portabel nobar bola/film dan pengeras suara santai", "1 paket", "Rp 1.150.000"),
        ("Stok Bahan Baku Awal (Initial Inventory)", "Susu murni segar 80L, fillet ayam taichan, blend espresso, biji lokal, sate", "1 paket", "Rp 4.000.000"),
        ("Cadangan Kas Operasional (Buffer Runway)", "Cadangan kas darurat kontinjensi operasional & likuiditas 1-2 bulan", "1 paket", "Rp 3.400.000"),
        ("Deposit Tambah Daya PLN 2200VA & Izin Desa", "Penambahan daya PLN ke 2200 VA untuk mesin 1450W + deposit token + izin", "1 paket", "Rp 1.100.000"),
    ]
    
    tbl_capex = doc.add_table(rows=len(capex_items) + 2, cols=4)
    tbl_capex.alignment = WD_TABLE_ALIGNMENT.CENTER
    c_widths = [Cm(5.0), Cm(6.5), Cm(2.0), Cm(2.92)]
    
    headers = ["Komponen Belanja Modal", "Deskripsi & Spesifikasi", "Volume", "Total Biaya"]
    for j, h in enumerate(headers):
        c = tbl_capex.cell(0, j)
        c.width = c_widths[j]
        set_cell_shading(c, "1B365D")
        set_cell_padding(c, top=65, bottom=65, left=90, right=90)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if j >= 2 else WD_ALIGN_PARAGRAPH.LEFT
        r = p.add_run(h)
        r.font.name = 'Arial'
        r.font.size = Pt(8.5)
        r.font.bold = True
        r.font.color.rgb = COLOR_WHITE
    set_row_props(tbl_capex.rows[0], is_header=True)
        
    for i, row in enumerate(capex_items):
        bg = "FAFAFA" if i % 2 == 0 else "FFFFFF"
        set_row_props(tbl_capex.rows[i + 1])
        for j, val in enumerate(row):
            c = tbl_capex.cell(i + 1, j)
            c.width = c_widths[j]
            set_cell_shading(c, bg)
            set_cell_padding(c, top=35, bottom=35, left=70, right=70)
            set_cell_borders(c, bottom={"val": "single", "sz": "4", "color": "E2E8F0"})
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if j == 2 else (WD_ALIGN_PARAGRAPH.RIGHT if j == 3 else WD_ALIGN_PARAGRAPH.LEFT)
            r = p.add_run(val)
            r.font.name = 'Arial'
            r.font.size = Pt(8)
            r.font.color.rgb = COLOR_TEXT
            if j == 0:
                r.font.bold = True
                
    # Total Row
    set_row_props(tbl_capex.rows[len(capex_items) + 1])
    c_tot_label = tbl_capex.cell(len(capex_items) + 1, 0)
    c_tot_desc = tbl_capex.cell(len(capex_items) + 1, 1)
    c_tot_vol = tbl_capex.cell(len(capex_items) + 1, 2)
    c_tot_val = tbl_capex.cell(len(capex_items) + 1, 3)
    
    c_tot_label.width, c_tot_desc.width, c_tot_vol.width, c_tot_val.width = c_widths[0], c_widths[1], c_widths[2], c_widths[3]
    for cell in [c_tot_label, c_tot_desc, c_tot_vol, c_tot_val]:
        set_cell_shading(cell, "EBF3FC")
        set_cell_padding(cell, top=65, bottom=65, left=90, right=90)
        set_cell_borders(cell, top={"val": "single", "sz": "12", "color": "1B365D"}, bottom={"val": "single", "sz": "12", "color": "1B365D"})
        
    p_tot1 = c_tot_label.paragraphs[0]
    r_tot1 = p_tot1.add_run("TOTAL KEBUTUHAN MODAL")
    r_tot1.font.name = 'Arial'
    r_tot1.font.size = Pt(8.5)
    r_tot1.font.bold = True
    r_tot1.font.color.rgb = COLOR_PRIMARY
    
    p_tot4 = c_tot_val.paragraphs[0]
    p_tot4.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_tot4 = p_tot4.add_run("Rp 40.000.000")
    r_tot4.font.name = 'Arial'
    r_tot4.font.size = Pt(9)
    r_tot4.font.bold = True
    r_tot4.font.color.rgb = COLOR_PRIMARY

    # ==================== SECTION 5: STRUKTUR HPP & PROYEKSI ====================
    add_h1("5. STRUKTUR BIAYA, KALKULASI HPP & UNIT ECONOMICS")
    add_p(
        "Menu ZeroSpot memadukan hidangan favorit angkringan malam (Sate Taichan, aneka sate bakar, nasi bakar daun pisang, gorengan mendoan, olahan susu murni segar) dengan seduhan kopi barista profesional. Rata-rata margin laba kotor berada pada kisaran optimal 50% hingga 68%."
    )
    
    add_h2("5.1 Rincian HPP dan Margin Keuntungan Produk Kunci")
    hpp_data = [
        ("Sate Taichan Paha Ayam (5 Tusuk)", "Daging paha ayam 100g, bumbu marinasi, sambal rawit uleg, jeruk limau", "Rp 6.400", "Rp 13.000", "Rp 6.600 (50,8%)"),
        ("Sate Usus Ayam Bumbu Kuning", "Usus ayam rebus bumbu kuning, bumbu oles bakar, tusuk lidi", "Rp 890", "Rp 2.500", "Rp 1.610 (64,4%)"),
        ("Sate Telur Puyuh Bacem (4 Butir)", "Telur puyuh 4 butir bumbu bacem kecap rempah, bakar hangat", "Rp 2.200", "Rp 4.000", "Rp 1.800 (45,0%)"),
        ("Sate Sosis Jumbo / Bakso Sapi", "Sosis jumbo / bakso sapi bakar oles saus BBQ & sambal", "Rp 2.050", "Rp 4.000", "Rp 1.950 (48,8%)"),
        ("Nasi Bakar Ayam Kemangi Daun Pisang", "Nasi gurih pandan, ayam suwir pedas kemangi, bakar daun", "Rp 4.000", "Rp 8.000", "Rp 4.000 (50,0%)"),
        ("Nasi Kucing Kertas Coklat (Teri/Orek)", "Nasi pulen, sambal tongkol balado / tempe orek pedas", "Rp 1.850", "Rp 3.500", "Rp 1.650 (47,1%)"),
        ("Tempe Mendoan Panas Dadakan (3 Lbr)", "Tempe tipis mendoan, tepung daun bawang, kecap rawit pedas", "Rp 3.000", "Rp 7.000", "Rp 4.000 (57,1%)"),
        ("Susu Murni Segar (Gelas Mug Jadul)", "Susu sapi murni segar pasteurisasi 200ml, hangat/es", "Rp 3.900", "Rp 9.000", "Rp 5.100 (56,7%)"),
        ("Susu Jahe Merah Rempah (Wedang)", "Susu murni 180ml, jahe merah bakar geprek, serai wangi, cengkeh", "Rp 4.720", "Rp 10.000", "Rp 5.280 (52,8%)"),
        ("Es Susu Strawberry Pink / Coklat", "Susu sapi segar 180ml, sari stroberi / pasta coklat pekat", "Rp 4.220", "Rp 10.000", "Rp 5.780 (57,8%)"),
        ("Es Susu Keju Melimpah / Regal", "Susu segar murni manis creamy + taburan keju parut / biskuit Regal", "Rp 5.300", "Rp 12.000", "Rp 6.700 (55,8%)"),
        ("Kopi Susu Aren Spesial ZeroSpot", "Double espresso 15 bar, susu murni, gula aren, cup 16oz + seal + es", "Rp 7.500", "Rp 15.000", "Rp 7.500 (50,0%)"),
        ("Manual Brew V60 Single Origin", "Biji Robusta/Arabika Pandeglang 15g, filter Hario, pouring barista", "Rp 5.500", "Rp 15.000", "Rp 9.500 (63,3%)"),
        ("Vietnam Drip Susu Murni Segar", "Kopi Robusta lokal giling medium, susu murni segar & krimer kental", "Rp 4.500", "Rp 12.000", "Rp 7.500 (62,5%)"),
    ]
    
    tbl_hpp = doc.add_table(rows=len(hpp_data) + 1, cols=5)
    tbl_hpp.alignment = WD_TABLE_ALIGNMENT.CENTER
    h_widths = [Cm(4.0), Cm(5.42), Cm(2.2), Cm(2.2), Cm(2.6)]
    
    headers_hpp = ["Nama Menu", "Komposisi Bahan Pokok", "HPP Riil", "Harga Jual", "Laba Kotor (Margin)"]
    for j, h in enumerate(headers_hpp):
        c = tbl_hpp.cell(0, j)
        c.width = h_widths[j]
        set_cell_shading(c, "3D2B1F")
        set_cell_padding(c, top=60, bottom=60, left=70, right=70)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if j >= 2 else WD_ALIGN_PARAGRAPH.LEFT
        r = p.add_run(h)
        r.font.name = 'Arial'
        r.font.size = Pt(8.5)
        r.font.bold = True
        r.font.color.rgb = COLOR_WHITE
    set_row_props(tbl_hpp.rows[0], is_header=True)
        
    for i, row in enumerate(hpp_data):
        bg = "FAFAFA" if i % 2 == 0 else "FFFFFF"
        set_row_props(tbl_hpp.rows[i + 1])
        for j, val in enumerate(row):
            c = tbl_hpp.cell(i + 1, j)
            c.width = h_widths[j]
            set_cell_shading(c, bg)
            set_cell_padding(c, top=32, bottom=32, left=60, right=60)
            set_cell_borders(c, bottom={"val": "single", "sz": "4", "color": "E2E8F0"})
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if j in [2, 3, 4] else WD_ALIGN_PARAGRAPH.LEFT
            r = p.add_run(val)
            r.font.name = 'Arial'
            r.font.size = Pt(8)
            r.font.color.rgb = COLOR_TEXT
            if j == 0:
                r.font.bold = True
            if j == 4:
                r.font.bold = True
                r.font.color.rgb = COLOR_SUCCESS

    add_h2("5.2 Proyeksi Laba Rugi Bulanan Berdasarkan 3 Skenario (Modal Rp 40 Juta)")
    add_p(
        "Kombinasi sate taichan, aneka sate angkringan, olahan susu murni, dan sajian kopi espresso mesin Ferratti Ferro menghasilkan kecepatan saji tinggi dan rata-rata belanja Rp 18.000 - Rp 23.000 per orang. Beban sewa lahan tahun ke-1 telah terdanai penuh di muka melalui CAPEX, mengamankan arus kas operasional bulanan dari risiko penagihan sewa berjalan:"
    )
    
    projections = [
        ("Asumsi Kunjungan & Transaksi", "22 Transaksi / Hari", "40 Transaksi / Hari", "65 Transaksi / Hari"),
        ("Rata-rata Nilai Belanja / Orang", "Rp 18.000", "Rp 21.000", "Rp 23.000"),
        ("Estimasi Omset Harian", "Rp 396.000", "Rp 840.000", "Rp 1.495.000"),
        ("Total Omset Bulanan (30 Hari)", "Rp 12.000.000", "Rp 25.000.000", "Rp 45.000.000"),
        ("Beban Pokok Penjualan (HPP ~42%)", "(Rp 5.040.000)", "(Rp 10.500.000)", "(Rp 18.900.000)"),
        ("Laba Kotor Usaha (Gross Profit)", "Rp 6.960.000", "Rp 14.500.000", "Rp 26.100.000"),
        ("Beban Listrik, Gas & Internet WiFi", "(Rp 750.000)", "(Rp 950.000)", "(Rp 1.200.000)"),
        ("Beban Tenaga Kerja Operasional", "(Rp 1.600.000)", "(Rp 2.800.000)", "(Rp 4.200.000)"),
        ("Beban Perawatan & Penyusutan Mebel", "(Rp 250.000)", "(Rp 450.000)", "(Rp 600.000)"),
        ("Beban Pemasaran & Kas Komunitas", "(Rp 150.000)", "(Rp 300.000)", "(Rp 500.000)"),
        ("TOTAL LABA BERSIH BULANAN", "Rp 4.210.000", "Rp 10.000.000", "Rp 19.600.000"),
        ("Porsi Bagi Hasil Investor (40%)", "Rp 1.684.000", "Rp 4.000.000", "Rp 7.840.000"),
        ("Porsi Bagi Hasil Pengelola (60%)", "Rp 2.526.000", "Rp 6.000.000", "Rp 11.760.000"),
        ("Estimasi Titik Impas Modal Investor (BEP)", "23,7 Bulan", "10,0 Bulan", "5,1 Bulan"),
    ]
    
    tbl_proj = doc.add_table(rows=len(projections) + 1, cols=4)
    tbl_proj.alignment = WD_TABLE_ALIGNMENT.CENTER
    p_widths = [Cm(5.82), Cm(3.4), Cm(3.6), Cm(3.6)]
    
    headers_proj = ["Parameter Finansial", "Skenario Konservatif", "Skenario Moderat", "Skenario Agresif"]
    for j, h in enumerate(headers_proj):
        c = tbl_proj.cell(0, j)
        c.width = p_widths[j]
        set_cell_shading(c, "1B365D")
        set_cell_padding(c, top=60, bottom=60, left=70, right=70)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if j >= 1 else WD_ALIGN_PARAGRAPH.LEFT
        r = p.add_run(h)
        r.font.name = 'Arial'
        r.font.size = Pt(8.5)
        r.font.bold = True
        r.font.color.rgb = COLOR_WHITE
    set_row_props(tbl_proj.rows[0], is_header=True)
        
    for i, row in enumerate(projections):
        is_highlight = (i in [3, 5, 10, 11, 13])
        bg = "EBF3FC" if is_highlight else ("FAFAFA" if i % 2 == 0 else "FFFFFF")
        set_row_props(tbl_proj.rows[i + 1])
        for j, val in enumerate(row):
            c = tbl_proj.cell(i + 1, j)
            c.width = p_widths[j]
            set_cell_shading(c, bg)
            set_cell_padding(c, top=35, bottom=35, left=65, right=65)
            set_cell_borders(c, bottom={"val": "single", "sz": "4", "color": "D8E2EC"})
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if j >= 1 else WD_ALIGN_PARAGRAPH.LEFT
            r = p.add_run(val)
            r.font.name = 'Arial'
            r.font.size = Pt(8)
            r.font.color.rgb = COLOR_PRIMARY if is_highlight else COLOR_TEXT
            if is_highlight or j == 0:
                r.font.bold = True

    # ==================== SECTION 6: STRUKTUR PENAWARAN INVESTASI ====================
    add_h1("6. STRUKTUR PENAWARAN INVESTASI & HAK INVESTOR")
    add_p(
        "Kemitraan permodalan ZeroSpot Cafe dirancang dengan tata kelola syariah transparan dan akuntabel:"
    )
    
    add_h2("6.1 Struktur Permodalan & Pembagian Slot Investasi (Target Rp 40 Juta)")
    add_bullet("Nilai Investasi yang Dibutuhkan", "Total Rp 40.000.000 (Empat Puluh Juta Rupiah).")
    add_bullet("Opsi Partisipasi Modal", "Tersedia pilihan 1 Investor Penuh (Rp 40.000.000), Sindikasi 2 Slot (@ Rp 20.000.000), Sindikasi 4 Slot (@ Rp 10.000.000), atau Sindikasi 5 Slot (@ Rp 8.000.000 per slot).")
    add_bullet("Skema Akad Kerja Sama", "Mudharabah Muqayyadah (Bagi Hasil Usaha). Investor sebagai Shahibul Maal (pemilik modal 100%) dan Pengelola sebagai Mudharib (pelaksana operasional penuh).")
    add_bullet("Nisbah Pembagian Keuntungan", "40% Laba Bersih untuk Investor dan 60% Laba Bersih untuk Pengelola Operasional.")
    add_bullet("Periode Pembagian Dividen", "Dibagikan secara berkala setiap tanggal 5 awal bulan kalender berikutnya, langsung ditransfer ke rekening investor.")

    add_h2("6.2 Transparansi & Tata Kelola Keuangan (Governance)")
    add_bullet("Sistem Kasir Digital (POS Cloud)", "Seluruh transaksi pesanan, stok bahan baku, dan arus kas harian dicatat menggunakan aplikasi kasir digital modern. Investor diberikan akun dashboard pengawas untuk memantau penjualan secara langsung (real-time) melalui smartphone.")
    add_bullet("Laporan Keuangan Bulanan Komprehensif", "Setiap akhir bulan, pengelola menyusun Laporan Laba Rugi (Income Statement) lengkap dengan rekonsiliasi kas, nota pembelian bahan baku, dan bukti pengeluaran operasional.")
    add_bullet("Hak Audit Fisik Terbuka", "Investor berhak melakukan peninjauan lokasi operasional, pemeriksaan stok fisik bahan baku (stock opname), dan verifikasi kas kapan saja secara terbuka.")

    add_callout(
        "KEISTIMEWAAN BAGI INVESTOR PERTAMA (EARLY INVESTOR PRIVILEGE)",
        "1. Prioritas Hak Opsi Replikasi (Right of First Refusal) untuk mendanai pembukaan cabang ke-2 dan ke-3 di kawasan Menes / Labuan.\n"
        "2. Kartu Keanggotaan Kehormatan 'ZeroSpot Founder Card' yang memberikan sajian kopi dan susu gratis seumur hidup di seluruh jaringan kedai ZeroSpot.\n"
        "3. Keterlibatan dalam Rapat Strategis Tahunan penentuan arah pengembangan dan kemitraan waralaba mikro."
    )

    # ==================== SECTION 7: ROADMAP EKSPANSI & RISIKO ====================
    add_h1("7. STRATEGI PERTUMBUHAN CEPAT & ROADMAP EKSPANSI (SCALE-UP)")
    add_p(
        "Format 'Outdoor Street Camp Angkringan' ZeroSpot memiliki keunggulan fleksibilitas ruang dan mobilitas tinggi yang sangat mudah direplikasi ke lokasi lain:"
    )
    
    add_h2("7.1 Roadmap 4 Fase Pengembangan 18 Bulan")
    add_bullet("Fase 1: Validasi Unit Pertama di Cikedal (Bulan 1 - 3)", "Mematangkan SOP marinasi taichan, bumbu sate bakar, pasokan susu sapi murni segar, dan roasting profile kopi Robusta lokal. Menargetkan 500+ pelanggan tetap bulanan.")
    add_bullet("Fase 2: Diversifikasi Produk Ritel & Katering Desa (Bulan 4 - 6)", "Meluncurkan produk kopi bubuk kemasan dan susu botolan siap minum untuk wisatawan pantai barat Banten, serta paket katering besek nasi bakar untuk pengajian/rapat warga.")
    add_bullet("Fase 3: Ekspansi Cabang Kedua di Menes / Labuan (Bulan 7 - 12)", "Setelah modal awal kembali penuh (~bulan ke-6), membuka cabang kedua di titik sewa lahan terbuka/pekarangan strategis di sekitar Alun-alun Menes atau Situ Cikedal.")
    add_bullet("Fase 4: Pembentukan Central Kitchen & Roastery Terpadu (Bulan 13 - 18)", "Membangun dapur sentral penusukan sate taichan, pengolahan bumbu ungkep, dan sangrai kopi mandiri untuk memasok seluruh cabang di Pandeglang.")

    add_h1("8. ANALISIS RISIKO & STRATEGI MITIGASI")
    risk_data = [
        ("Faktor Cuaca Hujan (Area Terbuka)", "Hujan deras berpotensi mengganggu area duduk outdoor.", "Fasilitas saung kanopi pelindung cuaca hujan serta terpal roll memastikan kapasitas duduk tetap terlindungi 100%."),
        ("Daya Simpan Susu Segar Murni", "Susu sapi murni mudah basi jika suhu pendingin tidak terjaga.", "Showcase chiller khusus 2C - 4C dan sistem perebusan teko bertahap sesuai kebutuhan harian memastikan kualitas susu selalu segar."),
        ("Fluktuasi Harga Daging Ayam & Cabai", "Kenaikan harga bahan baku sate taichan di pasar lokal.", "Kerja sama kemitraan langsung dengan peternak ayam potong di Menes dan petani cabai lokal Pandeglang."),
        ("Dinamika Sosial Lingkungan Desa", "Kekhawatiran kebisingan warga sekitar perkampungan.", "Pemberlakuan SOP ramah warga: mematikan musik saat adzan berkumandang, pembatasan volume malam hari, dan perekrutan pemuda desa sebagai staf."),
    ]
    
    tbl_risk = doc.add_table(rows=len(risk_data) + 1, cols=3)
    tbl_risk.alignment = WD_TABLE_ALIGNMENT.CENTER
    r_widths = [Cm(4.5), Cm(5.0), Cm(6.92)]
    
    headers_risk = ["Kategori Risiko", "Deskripsi Dampak Potensial", "Strategi Mitigasi Terukur"]
    for j, h in enumerate(headers_risk):
        c = tbl_risk.cell(0, j)
        c.width = r_widths[j]
        set_cell_shading(c, "1B365D")
        set_cell_padding(c, top=60, bottom=60, left=70, right=70)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if j == 0 else WD_ALIGN_PARAGRAPH.LEFT
        r = p.add_run(h)
        r.font.name = 'Arial'
        r.font.size = Pt(8.5)
        r.font.bold = True
        r.font.color.rgb = COLOR_WHITE
    set_row_props(tbl_risk.rows[0], is_header=True)
        
    for i, row in enumerate(risk_data):
        bg = "FAFAFA" if i % 2 == 0 else "FFFFFF"
        set_row_props(tbl_risk.rows[i + 1])
        for j, val in enumerate(row):
            c = tbl_risk.cell(i + 1, j)
            c.width = r_widths[j]
            set_cell_shading(c, bg)
            set_cell_padding(c, top=35, bottom=35, left=65, right=65)
            set_cell_borders(c, bottom={"val": "single", "sz": "4", "color": "E2E8F0"})
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            r = p.add_run(val)
            r.font.name = 'Arial'
            r.font.size = Pt(8)
            r.font.color.rgb = COLOR_TEXT
            if j == 0:
                r.font.bold = True

    # Explicit page break before Section 9 to ensure dedicated pristine sign-off page
    doc.add_page_break()

    # ==================== PAGE 8: KESIMPULAN, KONTAK & LEMBAR KOMITMEN ====================
    add_h1("9. KESIMPULAN, KONTAK MANAJEMEN & LEMBAR KOMITMEN")
    add_p(
        "ZeroSpot Cafe sebagai pionir perpaduan cafe modern perkotaan dengan angkringan tradisional menghadirkan keunikan bisnis F&B yang belum pernah ada di Pandeglang. Didukung belanja modal realistis Rp 40.000.000, mesin espresso listrik Ferratti Ferro 15 Bar untuk sajian cepat & latte art, efisiensi sewa lahan terbuka yang terjangkau, serta antusiasme tinggi masyarakat, ZeroSpot diproyeksikan mencapai titik impas dalam 10 bulan (skenario moderat) dan siap berkembang cepat."
    )
    
    add_h2("Informasi Kontak Resmi Inisiator & Pengelola Usaha")
    contact_data = [
        ("Nama Lengkap Inisiator", "Fatih Farhat Asshidiq (Founder & Managing Director)"),
        ("Nomor Telepon / WhatsApp", "+68 821-1150-0190   |   +62 895-0610-0075"),
        ("Alamat Email Resmi", "kazokuhairy@gmail.com"),
        ("Basis Operasional Proyek", "Sekitaran Kecamatan Cikedal, Kabupaten Pandeglang, Banten")
    ]
    tbl_contact = doc.add_table(rows=len(contact_data), cols=2)
    tbl_contact.alignment = WD_TABLE_ALIGNMENT.CENTER
    cnt_widths = [Cm(5.5), Cm(10.92)]
    for i, (k, v) in enumerate(contact_data):
        c0 = tbl_contact.cell(i, 0)
        c1 = tbl_contact.cell(i, 1)
        c0.width, c1.width = cnt_widths[0], cnt_widths[1]
        set_cell_shading(c0, "F2F5F9")
        set_cell_shading(c1, "FAFAFA")
        set_cell_padding(c0, top=40, bottom=40, left=85, right=85)
        set_cell_padding(c1, top=40, bottom=40, left=85, right=85)
        set_cell_borders(c0, bottom={"val": "single", "sz": "4", "color": "D8E2EC"})
        set_cell_borders(c1, bottom={"val": "single", "sz": "4", "color": "D8E2EC"})
        set_row_props(tbl_contact.rows[i])
        
        p0 = c0.paragraphs[0]
        r0 = p0.add_run(k)
        r0.font.name = 'Arial'
        r0.font.size = Pt(8.5)
        r0.font.bold = True
        r0.font.color.rgb = COLOR_PRIMARY
        
        p1 = c1.paragraphs[0]
        r1 = p1.add_run(v)
        r1.font.name = 'Arial'
        r1.font.size = Pt(8.5)
        r1.font.bold = (i in [0, 1, 2])
        r1.font.color.rgb = COLOR_TEXT

    add_h2("Ringkasan Ketentuan Investasi (Term Sheet Summary)")
    terms_data = [
        ("Total Nilai Investasi Proyek", "Rp 40.000.000 (Empat Puluh Juta Rupiah)"),
        ("Bentuk Partisipasi Modal", "[  ] 1 Investor Penuh (Rp 40 Jt)   |   [  ] Sindikasi (1 Slot = Rp 8 Jt / Rp 10 Jt)"),
        ("Nisbah Pembagian Hasil Usaha", "40% Laba Bersih Investor  :  60% Laba Bersih Pengelola Operasional"),
        ("Frekuensi Distribusi Dividen", "Bulanan (Ditransfer setiap tanggal 5 awal bulan kalender)"),
        ("Mekanisme Pelaporan", "Dashboard POS Real-Time + Laporan Laba Rugi Resmi Setiap Akhir Bulan")
    ]
    tbl_terms = doc.add_table(rows=len(terms_data), cols=2)
    tbl_terms.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_widths = [Cm(5.5), Cm(10.92)]
    for i, (k, v) in enumerate(terms_data):
        c0 = tbl_terms.cell(i, 0)
        c1 = tbl_terms.cell(i, 1)
        c0.width, c1.width = t_widths[0], t_widths[1]
        set_cell_shading(c0, "F2F5F9")
        set_cell_shading(c1, "FAFAFA")
        set_cell_padding(c0, top=40, bottom=40, left=85, right=85)
        set_cell_padding(c1, top=40, bottom=40, left=85, right=85)
        set_cell_borders(c0, bottom={"val": "single", "sz": "4", "color": "D8E2EC"})
        set_cell_borders(c1, bottom={"val": "single", "sz": "4", "color": "D8E2EC"})
        set_row_props(tbl_terms.rows[i])
        
        p0 = c0.paragraphs[0]
        r0 = p0.add_run(k)
        r0.font.name = 'Arial'
        r0.font.size = Pt(8.5)
        r0.font.bold = True
        r0.font.color.rgb = COLOR_PRIMARY
        
        p1 = c1.paragraphs[0]
        r1 = p1.add_run(v)
        r1.font.name = 'Arial'
        r1.font.size = Pt(8.5)
        r1.font.color.rgb = COLOR_TEXT

    add_h2("Formulir Pernyataan Minat / Komitmen Calon Investor")
    form_data = [
        ("Nama Lengkap Calon Investor", ": ....................................................................................................................."),
        ("Nomor KTP / WhatsApp", ": ....................................................................................................................."),
        ("Komitmen Alokasi Modal", ": [  ] 1 Slot (Rp 8.000.000)   |   [  ] 1 Slot (Rp 10.000.000)   |   [  ] Penuh (Rp 40.000.000)")
    ]
    tbl_form = doc.add_table(rows=len(form_data), cols=2)
    tbl_form.alignment = WD_TABLE_ALIGNMENT.CENTER
    f_widths = [Cm(5.5), Cm(10.92)]
    for i, (k, v) in enumerate(form_data):
        c0 = tbl_form.cell(i, 0)
        c1 = tbl_form.cell(i, 1)
        c0.width, c1.width = f_widths[0], f_widths[1]
        set_cell_shading(c0, "FAFAFA")
        set_cell_shading(c1, "FFFFFF")
        set_cell_padding(c0, top=35, bottom=35, left=80, right=80)
        set_cell_padding(c1, top=35, bottom=35, left=80, right=80)
        set_cell_borders(c0, bottom={"val": "single", "sz": "4", "color": "E2E8F0"})
        set_cell_borders(c1, bottom={"val": "single", "sz": "4", "color": "E2E8F0"})
        set_row_props(tbl_form.rows[i])
        
        p0 = c0.paragraphs[0]
        r0 = p0.add_run(k)
        r0.font.name = 'Arial'
        r0.font.size = Pt(8.5)
        r0.font.bold = True
        r0.font.color.rgb = COLOR_TEXT
        
        p1 = c1.paragraphs[0]
        r1 = p1.add_run(v)
        r1.font.name = 'Arial'
        r1.font.size = Pt(8.5)
        r1.font.color.rgb = COLOR_MUTED

    add_p("Pandeglang, 29 September 2026", space_before=8, space_after=3)
    
    # Signature Table
    sig_tbl = doc.add_table(rows=1, cols=2)
    sig_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    s_widths = [Cm(8.21), Cm(8.21)]
    
    c_sig0 = sig_tbl.cell(0, 0)
    c_sig1 = sig_tbl.cell(0, 1)
    c_sig0.width, c_sig1.width = s_widths[0], s_widths[1]
    set_cell_shading(c_sig0, "FAFAFA")
    set_cell_shading(c_sig1, "FAFAFA")
    set_cell_padding(c_sig0, top=70, bottom=70, left=90, right=90)
    set_cell_padding(c_sig1, top=70, bottom=70, left=90, right=90)
    set_cell_borders(c_sig0, top={"val": "single", "sz": "8", "color": "1B365D"}, bottom={"val": "single", "sz": "8", "color": "1B365D"})
    set_cell_borders(c_sig1, top={"val": "single", "sz": "8", "color": "1B365D"}, bottom={"val": "single", "sz": "8", "color": "1B365D"})
    set_row_props(sig_tbl.rows[0])
    
    p_s0 = c_sig0.paragraphs[0]
    p_s0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_s0 = p_s0.add_run("PIHAK PENGELOLA & INISIATOR\nZeroSpot Cafe\n\n\n")
    r_s0.font.name = 'Arial'
    r_s0.font.size = Pt(8.5)
    r_s0.font.color.rgb = COLOR_TEXT
    
    r_s0_name = p_s0.add_run("FATIH FARHAT ASSHIDIQ\n")
    r_s0_name.font.name = 'Arial'
    r_s0_name.font.size = Pt(9.5)
    r_s0_name.font.bold = True
    r_s0_name.font.underline = True
    r_s0_name.font.color.rgb = COLOR_PRIMARY
    
    r_s0_title = p_s0.add_run("Founder & Managing Director\n")
    r_s0_title.font.name = 'Arial'
    r_s0_title.font.size = Pt(8)
    r_s0_title.font.color.rgb = COLOR_MUTED
    
    r_s0_cnt = p_s0.add_run("WA: +68 821-1150-0190 / +62 895-0610-0075\nEmail: kazokuhairy@gmail.com")
    r_s0_cnt.font.name = 'Arial'
    r_s0_cnt.font.size = Pt(7.5)
    r_s0_cnt.font.color.rgb = COLOR_MUTED
    
    p_s1 = c_sig1.paragraphs[0]
    p_s1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_s1 = p_s1.add_run("PIHAK MITRA / CALON INVESTOR\nMitra Permodalan Strategis\n\n\n")
    r_s1.font.name = 'Arial'
    r_s1.font.size = Pt(8.5)
    r_s1.font.color.rgb = COLOR_TEXT
    
    r_s1_name = p_s1.add_run("( ........................................................... )\n")
    r_s1_name.font.name = 'Arial'
    r_s1_name.font.size = Pt(9.5)
    r_s1_name.font.bold = True
    r_s1_name.font.color.rgb = COLOR_PRIMARY
    
    r_s1_title = p_s1.add_run("Investor / Shahibul Maal\n\n")
    r_s1_title.font.name = 'Arial'
    r_s1_title.font.size = Pt(8)
    r_s1_title.font.color.rgb = COLOR_MUTED

    # Save DOCX
    out_dir = "/media/fatihfarhat/New Volume1/FATIH DATA/ZeroSpot Cafe"
    os.makedirs(out_dir, exist_ok=True)
    docx_path = os.path.join(out_dir, "PROPOSAL_INVESTASI_ZEROSPOT_CAFE.docx")
    doc.save(docx_path)
    print(f"Berhasil membuat master proposal DOCX di: {docx_path}")
    return docx_path

if __name__ == "__main__":
    create_master_proposal()
