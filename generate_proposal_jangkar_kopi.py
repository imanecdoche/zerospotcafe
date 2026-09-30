import os
import sys
import subprocess
import docx
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def create_jangkar_kopi_proposal():
    doc = docx.Document()
    
    # 1. Page Setup: F4 (Folio) 21.5 cm x 33.0 cm
    section = doc.sections[0]
    section.page_width = Cm(21.5)
    section.page_height = Cm(33.0)
    section.top_margin = Cm(2.2)
    section.bottom_margin = Cm(2.2)
    section.left_margin = Cm(2.2)
    section.right_margin = Cm(2.2)
    
    # Enable different first page for clean cover page
    section.different_first_page_header_footer = True
    section.first_page_header.paragraphs[0].text = ''
    section.first_page_footer.paragraphs[0].text = '' 
    
    # 2. Executive Nautical & Earthy Color Palette
    COLOR_PRIMARY = RGBColor(16, 42, 67)     # Deep Ocean Navy #102A43 (Maritime Anchor)
    COLOR_SECONDARY = RGBColor(180, 83, 9)   # Warm Amber Terracotta #B45309 (Lantern & Fire)
    COLOR_TEXT = RGBColor(36, 59, 83)        # Dark Charcoal Slate #243B53
    COLOR_MUTED = RGBColor(98, 125, 152)     # Muted Blue-Gray #627D98
    COLOR_GOLD = RGBColor(197, 137, 23)      # Rich Amber Gold #C58917
    COLOR_WHITE = RGBColor(255, 255, 255)
    COLOR_SUCCESS = RGBColor(22, 101, 52)    # Forest Green for profit margins
    
    # Base Normal Style
    style_normal = doc.styles['Normal']
    style_normal.font.name = 'Arial'
    style_normal.font.size = Pt(9.5)
    style_normal.font.color.rgb = COLOR_TEXT
    style_normal.paragraph_format.line_spacing = 1.15
    style_normal.paragraph_format.space_after = Pt(3.0)
    
    # Setup Footer for subsequent pages
    footer = section.footer
    f_p = footer.paragraphs[0]
    f_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    
    f_run_left = f_p.add_run("Jangkar Kopi & Angkringan — Proposal Investasi Bisnis  |  ")
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
    h_run = h_p.add_run("DOKUMEN PENAWARAN INVESTASI & KEMITRAAN LEAN MVP  —  JANGKAR KOPI & ANGKRINGAN")
    h_run.font.name = 'Arial'
    h_run.font.size = Pt(8)
    h_run.font.color.rgb = RGBColor(140, 160, 180)

    # Helper Functions
    def add_p(text="", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=3.0, line_spacing=1.15, bold=False, italic=False, color=COLOR_TEXT, size=Pt(9.5)):
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

    def add_h1(text, space_before=11, space_after=4):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(text)
        r.font.name = 'Arial'
        r.font.size = Pt(13)
        r.font.bold = True
        r.font.color.rgb = COLOR_PRIMARY
        return p

    def add_h2(text, space_before=8, space_after=3):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(text)
        r.font.name = 'Arial'
        r.font.size = Pt(10.5)
        r.font.bold = True
        r.font.color.rgb = COLOR_SECONDARY
        return p

    def add_callout(title, items, border_color="102A43", bg_color="F0F4F8"):
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = tbl.cell(0, 0)
        cell.width = Cm(17.1)
        set_cell_shading(cell, bg_color)
        set_cell_padding(cell, top=70, bottom=70, left=110, right=110)
        set_cell_borders(cell, left={"val": "single", "sz": "24", "color": border_color})
        
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(4)
        r_title = p.add_run(title)
        r_title.font.name = 'Arial'
        r_title.font.size = Pt(10)
        r_title.font.bold = True
        r_title.font.color.rgb = COLOR_PRIMARY
        
        for item in items:
            p_item = cell.add_paragraph()
            p_item.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p_item.paragraph_format.space_before = Pt(1.5)
            p_item.paragraph_format.space_after = Pt(2.0)
            p_item.paragraph_format.line_spacing = 1.15
            
            if isinstance(item, tuple):
                bold_txt, reg_txt = item
                r_b = p_item.add_run(bold_txt)
                r_b.font.name = 'Arial'
                r_b.font.size = Pt(9)
                r_b.font.bold = True
                r_b.font.color.rgb = COLOR_SECONDARY
                
                r_r = p_item.add_run(reg_txt)
                r_r.font.name = 'Arial'
                r_r.font.size = Pt(9)
                r_r.font.color.rgb = COLOR_TEXT
            else:
                r_r = p_item.add_run(str(item))
                r_r.font.name = 'Arial'
                r_r.font.size = Pt(9)
                r_r.font.color.rgb = COLOR_TEXT
                
        p_space = doc.add_paragraph()
        p_space.paragraph_format.space_before = Pt(0)
        p_space.paragraph_format.space_after = Pt(2)

    def set_cell_shading(cell, color_hex):
        shd = parse_xml(r'<w:shd %s w:val="clear" w:color="auto" w:fill="%s"/>' % (nsdecls('w'), color_hex))
        cell._tc.get_or_add_tcPr().append(shd)

    def set_cell_padding(cell, top=60, bottom=60, left=80, right=80):
        mar = parse_xml(r'<w:tcMar %s><w:top w:w="%d" w:type="dxa"/><w:bottom w:w="%d" w:type="dxa"/><w:left w:w="%d" w:type="dxa"/><w:right w:w="%d" w:type="dxa"/></w:tcMar>' % (nsdecls('w'), top, bottom, left, right))
        cell._tc.get_or_add_tcPr().append(mar)

    def set_cell_borders(cell, **kwargs):
        tcPr = cell._tc.get_or_add_tcPr()
        tcBorders = parse_xml(r'<w:tcBorders %s/>' % nsdecls('w'))
        for edge in ('top', 'left', 'bottom', 'right'):
            edge_data = kwargs.get(edge)
            if edge_data:
                tag = r'<w:%s %s w:val="%s" w:sz="%s" w:space="0" w:color="%s"/>' % (
                    edge, nsdecls('w'), edge_data.get('val', 'single'),
                    edge_data.get('sz', '4'), edge_data.get('color', 'auto')
                )
            else:
                tag = r'<w:%s %s w:val="none"/>' % (edge, nsdecls('w'))
            tcBorders.append(parse_xml(tag))
        tcPr.append(tcBorders)

    def set_row_props(row, is_header=False):
        trPr = row._tr.get_or_add_trPr()
        trPr.append(parse_xml(r'<w:cantSplit %s/>' % nsdecls('w')))
        if is_header:
            trPr.append(parse_xml(r'<w:tblHeader %s/>' % nsdecls('w')))

    # ==================== PAGE 1: COVER PAGE ====================
    p_inst = doc.add_paragraph()
    p_inst.paragraph_format.space_before = Pt(8)
    p_inst.paragraph_format.space_after = Pt(2)
    r_inst = p_inst.add_run("PENAWARAN INVESTASI & KEMITRAAN USAHA MIKRO BERBASIS LEAN STARTUP")
    r_inst.font.name = 'Arial'
    r_inst.font.size = Pt(9.5)
    r_inst.font.bold = True
    r_inst.font.color.rgb = COLOR_SECONDARY
    
    p_line = doc.add_paragraph()
    p_line.paragraph_format.space_before = Pt(0)
    p_line.paragraph_format.space_after = Pt(14)
    r_l = p_line.add_run("—" * 52)
    r_l.font.color.rgb = COLOR_PRIMARY
    
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(6)
    p_title.paragraph_format.space_after = Pt(4)
    r_title = p_title.add_run("JANGKAR KOPI")
    r_title.font.name = 'Arial'
    r_title.font.size = Pt(32)
    r_title.font.bold = True
    r_title.font.color.rgb = COLOR_PRIMARY
    
    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(6)
    r_sub = p_sub.add_run("& ARTISAN ANGKRINGAN RAKYAT")
    r_sub.font.name = 'Arial'
    r_sub.font.size = Pt(17)
    r_sub.font.bold = True
    r_sub.font.color.rgb = COLOR_SECONDARY
    
    p_tag = doc.add_paragraph()
    p_tag.paragraph_format.space_before = Pt(2)
    p_tag.paragraph_format.space_after = Pt(16)
    r_tag = p_tag.add_run("Tagline: \"Labuhkan Lelah, Seduh Cerita\"")
    r_tag.font.name = 'Arial'
    r_tag.font.size = Pt(11)
    r_tag.font.italic = True
    r_tag.font.bold = True
    r_tag.font.color.rgb = COLOR_GOLD

    add_p(
        "Proposal Rencana Bisnis Sederhana (Lean MVP) — Menghadirkan Ruang Temu Lesehan Rakyat yang Hangat, Bersahaja, dan Egaliter di Koridor Strategis Cikedal - Menes, Pandeglang, Banten. Memadukan Kelezatan Aneka Sate Frozen Food Bakar Arang (Dumpling Keju, Otak-Otak Ikan, Sate Cumi Olahan, Sosis Bakar, Bakso Sapi Bakar), Nasi Kucing & Mendoan Hangat, Kopi Tubruk Robusta Khas Daerah Banten, dan Wedangan Susu Murni Segar Tanpa Beban Biaya Mesin Mewah.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY,
        space_after=14,
        line_spacing=1.2
    )

    add_callout(
        "RINGKASAN EKSEKUTIF PROYEK INVESTASI (LEAN MVP MODEL)",
        [
            ("• Nilai Total Permodalan Awal : ", "Rp 15.000.000 (Lima Belas Juta Rupiah) — Seimbang, Kokoh & Terukur."),
            ("• Opsi Partisipasi Mitra       : ", "1 Mitra Penuh (Rp 15 Jt)  |  2 Slot (@ Rp 7,5 Jt)  |  3 Slot (@ Rp 5 Jt)  |  5 Slot (@ Rp 3 Jt)."),
            ("• Alokasi Sewa Lahan Sementara : ", "Rp 6.000.000,- (Sewa Lahan Terbuka 1 Tahun di Muka ~5x5 Meter Koridor Cikedal-Menes)."),
            ("• Skema Kemitraan Usaha        : ", "Syirkah Mudharabah (Bagi Hasil Laba Bersih: 40% Investor : 60% Pengelola)."),
            ("• Proyeksi Balik Modal (BEP)   : ", "7,5 Bulan pada Skenario Moderat; 4,7 Bulan pada Skenario Agresif (Harga Merakyat)."),
            ("• Konsep Ruang & Duduk         : ", "100% Lesehan Alas Tebal Nyaman (Karpet Busa Empuk) + Meja Lipat Portabel."),
            ("• Konsep Dapur & Minuman       : ", "Panggangan Arang Batok Tradisional + Aneka Frozen Food Bakar, Kopi Tubruk Lokal & Sachet."),
            ("• Lokasi Basis Operasional     : ", "Lahan Terbuka Strategis Koridor Cikedal - Menes, Kabupaten Pandeglang, Banten."),
            ("• Inisiator & Penanggung Jawab : ", "Fatih Farhat Asshidiq (Founder & Managing Director).")
        ],
        border_color="102A43",
        bg_color="F0F4F8"
    )

    p_meta = doc.add_paragraph()
    p_meta.paragraph_format.space_before = Pt(16)
    p_meta.paragraph_format.space_after = Pt(0)
    p_meta.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_meta1 = p_meta.add_run("Dokumen Resmi Penawaran Kemitraan  •  Edisi Lean Startup  •  Oktober 2026\n")
    r_meta1.font.size = Pt(8.5)
    r_meta1.font.color.rgb = COLOR_MUTED
    r_meta2 = p_meta.add_run("Disusun untuk Calon Mitra Investor Strategis & Rekan Kemitraan")
    r_meta2.font.size = Pt(8.5)
    r_meta2.font.bold = True
    r_meta2.font.color.rgb = COLOR_PRIMARY

    doc.add_page_break()

    # ==================== PAGE 2: BAB 1 LATAR BELAKANG & FILOSOFI ====================
    add_h1("1. LATAR BELAKANG, PRINSIP LEAN STARTUP & FILOSOFI MEREK")
    
    add_h2("1.1 Mengapa Model Lean MVP Diterapkan?")
    add_p(
        "Berdasarkan evaluasi objektif bersama dewan penasihat dan calon mitra investor strategis, langkah paling bijak dalam memulai usaha F&B pedesaan adalah menerapkan prinsip Lean Startup: memangkas belanja modal yang belum mendesak, menghilangkan risiko kelebihan kapasitas (over-capitalization), dan menguji penerimaan pasar secara langsung dengan modal ringan yang cepat menghasilkan arus kas (cashflow engine)."
    )
    add_p(
        "Daya beli riil masyarakat di koridor Cikedal dan Menes terbukti sangat kuat pada rentang harga merakyat. Konsumen lokal terbiasa membelanjakan Rp 10.000 hingga Rp 20.000 per sesi nongkrong saat jajan aneka sate arang (Rp 2.500–Rp 3.500), nasi kucing, dan kopi hangat (Rp 3.000). Faktor penentu keberhasilan utama bukanlah kemewahan alat, melainkan suasana tempat yang merakyat, ramah, dan harga bersahabat yang tidak menimbulkan rasa segan (kagok/takut mahal). Oleh karena itu, Jangkar Kopi dirancang bersahaja, bersih, dan berakar pada kenyamanan warga lokal."
    )
    
    add_h2("1.2 Filosofi Mendalam \"Jangkar Kopi\" — Nama adalah Doa dan Harapan")
    add_p(
        "Nama \"Jangkar Kopi\" dipilih bukan dari istilah asing yang mengada-ada, melainkan berakar pada filosofi kehidupan, kearifan lokal, dan kedekatan geografis wilayah Pandeglang barat:"
    )

    add_callout(
        "EMPAT PILAR FILOSOFI DAN DOA DI BALIK NAMA JANGKAR KOPI",
        [
            ("1. Tempat Berlabuh & Melepas Lelah (The Safe Harbor): ", "Fungsi utama jangkar adalah diturunkan saat perahu tiba di dermaga yang tenang. Setiap hari warga bekerja keras mengarungi ombak kehidupan; Jangkar Kopi hadir sebagai tempat mereka berlabuh, menghela napas, dan melepas penat di malam hari."),
            ("2. Kaitan Geografis Jalur Maritim Banten Barat: ", "Cikedal dan Menes adalah gerbang pelintas menuju kawasan pesisir Labuan, Carita, dan Tanjung Lesung. Jangkar sangat lekat dengan ketangguhan masyarakat maritim Banten yang bersahaja dan pekerja keras."),
            ("3. Keteguhan yang Menancap Membumi (Grounded): ", "Jangkar tidak pernah melayang di langit; jangkar menancap kuat di bumi. Ini adalah doa agar usaha ini selalu membumi, tidak sombong, berakar pada realitas rakyat, dan tahan banting menghadapi ujian usaha."),
            ("4. Pengikat Tali Silaturahmi: ", "Jangkar menahan perahu agar tidak hanyut terseret arus. Usaha ini diniatkan menjadi pengikat tali persaudaraan antarwarga, pemuda, santri, dan tetangga agar tetap rukun dan guyub.")
        ],
        border_color="B45309",
        bg_color="FEF3C7"
    )

    add_h2("1.3 Sederhana, Cepat Buka & Minim Risiko")
    add_p(
        "Dengan menunda pembelian set kursi camping dan mesin espresso impor, serta memfokuskan fasilitas duduk pada alas tebal yang nyaman (karpet busa empuk waterproof) dan meja lipat kecil portabel, Jangkar Kopi menekan belanja modal awal secara drastis sehingga dapat dieksekusi dalam tempo 10–14 hari kerja. Risiko kerugian ditekan seminimal mungkin, sementara perputaran uang harian langsung aktif sejak hari pertama."
    )

    doc.add_page_break()

    # ==================== PAGE 3: BAB 2 KONSEP & TATA RUANG LAHAN TERBUKA ====================
    add_h1("2. KONSEP BISNIS, PRODUK & TATA RUANG LAHAN TERBUKA")
    
    add_h2("2.1 Konsep Tempat: 100% Lesehan Alas Tebal Nyaman & 4 Meja Lipat Portabel")
    add_p(
        "Jangkar Kopi mengusung konsep 100% lesehan bersih beralas tebal yang nyaman (karpet spons/busa empuk waterproof dilapisi tikar rapi) dipadu 4 unit meja lipat kecil portabel yang praktis dan kokoh. Hidangan disajikan di atas piring anyaman rotan plastik bahan PP tebal food-grade anti-jamur hasil riset pasar e-commerce (dipesan paket grosir lusinan hemat Rp 15rb-19rb/lusin) beralas kertas nasi coklat bulat. Pilihan ini menghadirkan estetika angkringan otentik, anti-pecah, tidak menyimpan bau, sekaligus higienis dan praktis tanpa repot mencuci piring berminyak di malam hari (zero washing). Format duduk lesehan melingkar terbukti secara sosiologis mampu meruntuhkan batas status sosial. Pelanggan dapat duduk santai berselonjor menikmati hidangan tanpa merasa canggung atau terintimidasi."
    )
    add_p(
        "Pencahayaan dirancang menggunakan lampu gantung festoon warm white (3000K) yang temaram lembut, dipadu aroma gurih aneka sate frozen food (sosis, bakso, dumpling, otak-otak, cumi) bakar arang batok kelapa yang menyebar ke jalan raya, menciptakan daya pikat panca indra (sensory branding) yang mengundang pengendara untuk menepi."
    )

    add_h2("2.2 Zonasi Tata Ruang Lahan Terbuka (~5 × 5 Meter / 25 m²)")
    add_p(
        "Rencana pemanfaatan lahan terbuka strategis seluas estimasi 5 × 5 meter (25 m²) diatur secara efisien, rapi, dan cepat dibersihkan:"
    )

    add_callout(
        "SKEMA TATA RUANG DAN ALUR KERJA LAHAN TERBUKA STRATEGIS (5 × 5 METER)",
        [
            ("• Zonasi Depan (1,5 m × 5,0 m) — Dapur Display & Panggangan Arang: ", "Menempatkan gerobak kayu etalase kaca display sate frozen food higienis, panggangan arang batok stainless dengan hembusan kipas angin kecil, ceret wedangan, dan kompor mendoan menghadap ke jalan raya agar aroma asap bakaran gurih memancing selera."),
            ("• Zonasi Tengah & Belakang (3,5 m × 5,0 m) — Area Lesehan Tamu: ", "Hamparan alas tebal nyaman (karpet busa empuk waterproof dilapisi tikar rapi), dilengkapi 4 unit meja lipat kecil portabel yang ringan, kokoh, dan praktis dipindahkan. Mampu menampung 16 hingga 20 orang tamu sekaligus."),
            ("• Perlindungan Cuaca (Tenda 2×6m Milik Sendiri + Atap Terpal Baru): ", "Memanfaatkan aset rangka tenda 2×6 meter milik pribadi inisiator yang dilengkapi atap kain terpal tebal waterproof (A12 heavy duty) baru untuk memproteksi area lesehan dan dapur dari embun malam serta gerimis hujan tanpa perlu biaya beli rangka baru."),
            ("• Zonasi Parkir & Akses Bersih: ", "Area depan pinggir jalan dimanfaatkan untuk parkir 8–10 sepeda motor, serta dilengkapi tempat cuci tangan (wastafel portabel injak) dan tempat sampah tertutup.")
        ],
        border_color="102A43",
        bg_color="F0F4F8"
    )

    add_h2("2.3 Menu Makanan Awal: Aneka Sate Frozen Food Bakar Arang & Kudapan")
    add_p(
        "Menu makanan awal Jangkar Kopi difokuskan pada olahan siap bakar yang praktis, higienis, dan zero food waste dengan harga jual ramah kantong:"
    )
    add_p(
        "• Aneka Sate Frozen Food Bakar Arang: Dumpling ayam & keju lumer (Rp 3.500), sate cumi flower (Rp 3.500), otak-otak singapura kerat silang (Rp 2.500), sosis sapi bakar BBQ (Rp 3.000), bakso sapi bakar arang lada hitam (Rp 2.500), dan fish roll (Rp 2.500). Seluruhnya dipanggang hangat di atas bara batok kelapa beralas piring rotan plastik PP & kertas nasi coklat dengan cocolan sambal kecap rawit pedas dan saus bakar spesial dalam mangkuk plastik kecil."
    )
    add_p(
        "• Kudapan Tradisional: Nasi kucing teri balado / orek tempe bungkus daun pisang (Rp 2.500) yang dihangatkan di atas bara panggangan, serta tempe mendoan hangat daun bawang (Rp 3.000 isi 3 lembar) yang renyah dan gurih."
    )

    add_h2("2.4 Signature Menu Kopi Jangkar & Minuman Rakyat (Ikon Minuman Usaha)")
    add_p(
        "Sebagai pembeda utama, Jangkar Kopi menghadirkan menu signature kopi khas daerah dan minuman rakyat berharga terjangkau:"
    )
    add_p(
        "1. Kopi Hitam Tubruk 'Jangkar Asli' (Signature Utama): Bubuk kopi Robusta sangrai tradisional khas lereng Gunung Karang Pandeglang, diseduh tubruk panas dengan air mendidih murni. Berkarakter aroma smoky alami, body pekat (bold), dan rasa pahit gurih mantap khas Banten Barat, dengan pilihan gula aren Malingping murni atau gula pasir terpisah (Rp 3.000)."
    )
    add_p(
        "2. Kopi Susu 'Jangkar Dermaga' (Signature Modern): Perpaduan seduhan pekat Robusta lokal Gunung Karang dengan susu kental manis gurih dan lelehan gula aren murni khas Pandeglang. Menghadirkan rasa manis legit yang lembut dan seimbang, disajikan hangat maupun es dingin segar (Rp 5.000)."
    )
    add_p(
        "3. Kopi Jahe Rempah 'Jangkar Samudra' (Signature Penghangat): Seduhan kopi Robusta lokal berpadu rebusan jahe merah geprek, serai, dan kayu manis alami dalam mug jadul. Menjadi primadona pekerja malam, santri, dan pelintas jalur Labuan-Pandeglang untuk memulihkan stamina (Rp 6.000)."
    )
    add_p(
        "4. Minuman Rakyat Pendamping: Susu murni segar, wedang jahe susu, es sirup sachet (Nutrisari, Extra Joss Susu), dan varian kopi sachet populer (Good Day, Kapal Api, Indocafe, Luwak White Koffie) seharga Rp 3.000 – Rp 6.000."
    )

    doc.add_page_break()

    # ==================== PAGE 4: BAB 3 ANALISIS PASAR & STRATEGI OPERASIONAL ====================
    add_h1("3. ANALISIS PASAR, SEGMENTASI & STRATEGI OPERASIONAL")
    
    add_h2("3.1 Segmentasi Sasaran Konsumen Koridor Cikedal - Menes")
    add_p(
        "Koridor jalan raya penghubung Menes dan Cikedal merupakan jalur dengan arus lalu lintas malam yang stabil, dikelilingi kantong pemukiman padat dan lembaga pendidikan pesantren:"
    )

    add_callout(
        "EMPAT KLASTER KONSUMEN UTAMA JANGKAR KOPI",
        [
            ("1. Pemuda Desa & Komunitas Nongkrong Lokal (40%): ", "Kelompok usia 17–30 tahun yang membutuhkan tempat jagongan santai tanpa beban, bebas main game online/obrolan santai, dan memesan aneka sate tusuk serta kopi sachet."),
            ("2. Santri, Alumni & Pelajar Lembaga Pendidikan (25%): ", "Ekosistem Menes yang kental dengan tradisi santri menjadikan format lesehan bersarung sangat digemari untuk berkumpul malam selepas kegiatan mengaji/belajar."),
            ("3. Pekerja Malam, Guru, Pedagang & Petugas Ronda (20%): ", "Masyarakat yang mencari santapan malam mengenyangkan berharga murah (nasi bakar daun pisang, gorengan mendoan hangat, wedang jahe susu penambah stamina)."),
            ("4. Pelintas Jalur Wisata & Logistik Pantai Barat (15%): ", "Pengendara mobil dan motor rute Pandeglang-Labuan-Carita yang memerlukan tempat istirahat (rest point) sejenak yang aman dan bersih di pinggir jalan raya.")
        ],
        border_color="B45309",
        bg_color="FEF3C7"
    )

    add_h2("3.2 Keunggulan Kompetitif Dibanding Warkop Konvensional")
    add_p(
        "Meskipun warung kopi sachet banyak bertebaran di pedesaan, Jangkar Kopi memiliki keunggulan pembeda yang tegas:"
    )
    add_p(
        "• Standar Kebersihan Unggul: Alas lesehan dilap bersih setiap pergantian tamu, tempat sampah tertutup di setiap sudut, dan etalase sate tertutup kaca higienis bebas debu jalanan."
    )
    add_p(
        "• Daya Tarik Aneka Sate Frozen Food Bakar Arang: Menghadirkan ragam sate sosis bakar, bakso sapi bakar bumbu lada hitam, dumpling keju lumer, otak-otak singapura, dan sate cumi olahan yang dibakar hangat di atas arang batok kelapa dengan saus bakar gurih yang sangat disukai pemuda desa dan santri."
    )
    add_p(
        "• Keramahan Pelayanan & Musik Akustik: Diiringi alunan musik santai melalui speaker bluetooth kecil yang menciptakan suasana rileks tanpa kebisingan yang mengganggu."
    )

    add_h2("3.3 Waktu Operasional & Pola Layanan")
    add_p(
        "Jangkar Kopi beroperasi setiap hari mulai pukul 16.30 WIB (menjelang maghrib) hingga pukul 24.00 WIB malam. Jam operasional ini mengoptimalkan jam makan malam keluarga dan waktu nongkrong santai pemuda desa."
    )

    doc.add_page_break()

    # ==================== PAGE 5: BAB 4 RENCANA ANGGARAN BIAYA (CAPEX RP 15 JT) ====================
    add_h1("4. RENCANA ANGGARAN BIAYA & ALOKASI MODAL LEAN MVP (RP 15 JT)")
    
    add_p(
        "Total kebutuhan belanja modal (CAPEX) dan modal kerja awal Jangkar Kopi dialokasikan tepat sebesar Rp 15.000.000 (Lima Belas Juta Rupiah). Anggaran ini mengalokasikan sewa lahan terbuka selama 1 tahun di muka sebesar Rp 6.000.000 (40%), serta memanfaatkan sisa dana Rp 9.000.000 (60%) secara optimal untuk pengadaan atap terpal waterproof pada rangka tenda 2×6 meter milik pribadi inisiator, gerobak kayu custom etalase kaca, panggangan arang batok dengan kipas angin kecil, peralatan dapur wedangan, fasilitas lesehan alas tebal nyaman dengan 4 meja lipat portabel, wadah saji piring anyaman rotan plastik PP tebal anti-jamur (grosir lusinan) beralas kertas nasi & mangkuk plastik kecil sambal, instalasi listrik festoon, stok bahan baku awal yang diperkuat termasuk bubuk kopi Robusta lokal khas daerah (Rp 1.600.000), dan cadangan kas darurat yang kokoh (Rp 2.000.000)."
    )
    
    add_h2("4.1 Tabel Rincian Belanja Modal Awal (Tepat Rp 15.000.000)")

    capex_data = [
        ("Sewa Lahan Terbuka 1 Tahun di Muka", "Alokasi sewa pekarangan/lahan terbuka ~5x5 meter koridor Cikedal-Menes", "1 tahun", "Rp 6.000.000"),
        ("Atap Terpal Tenda 2×6m & Lahan", "Terpal tebal waterproof A12 rangka tenda 2x6m (rangka milik sendiri) & koral", "1 paket", "Rp 400.000"),
        ("Gerobak Angkringan Kayu Custom", "Gerobak kayu custom 160x75 cm + etalase kaca display sate frozen food", "1 unit", "Rp 2.000.000"),
        ("Panggangan Arang Stainless & Dapur", "Panggangan arang 65cm, kipas angin kecil, kompor gas, wajan mendoan, ceret, LPG", "1 paket", "Rp 1.000.000"),
        ("Alas Tebal Nyaman & 4 Meja Lipat", "3 Karpet busa/spons tebal empuk waterproof 2x2m, tikar rapi, 4 meja portabel", "1 paket", "Rp 800.000"),
        ("Cooler Box, Piring Rotan & Mangkuk", "Cooler box 35-45L, 2 lusin piring rotan plastik PP tebal (grosir), kertas nasi, mangkuk sambal", "1 paket", "Rp 600.000"),
        ("Instalasi Listrik, Festoon & Audio", "Kabel outdoor waterproof, stopkontak cas hp, lampu festoon warm, mini speaker", "1 paket", "Rp 600.000"),
        ("Stok Bahan Baku Awal (Inventory)", "Frozen food sate, bubuk kopi Robusta lokal khas daerah, kopi sachet, susu", "1 paket", "Rp 1.600.000"),
        ("Cadangan Kas Operasional (Buffer)", "Dana cadangan kas darurat kontinjensi operasional & likuiditas awal 1-2 bulan", "1 paket", "Rp 2.000.000")
    ]

    tbl_capex = doc.add_table(rows=len(capex_data) + 2, cols=4)
    tbl_capex.alignment = WD_TABLE_ALIGNMENT.CENTER
    c_widths = [Cm(4.8), Cm(6.5), Cm(2.2), Cm(3.6)]
    
    headers = ["Komponen Belanja Modal", "Deskripsi & Spesifikasi", "Volume", "Total Biaya"]
    for j, h in enumerate(headers):
        c = tbl_capex.cell(0, j)
        c.width = c_widths[j]
        set_cell_shading(c, "102A43")
        set_cell_padding(c, top=60, bottom=60, left=80, right=80)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if j >= 2 else WD_ALIGN_PARAGRAPH.LEFT
        r = p.add_run(h)
        r.font.name = 'Arial'
        r.font.size = Pt(8.5)
        r.font.bold = True
        r.font.color.rgb = COLOR_WHITE
    set_row_props(tbl_capex.rows[0], is_header=True)

    for i, row in enumerate(capex_data):
        bg = "F0F4F8" if i % 2 == 0 else "FFFFFF"
        set_row_props(tbl_capex.rows[i + 1])
        for j, val in enumerate(row):
            c = tbl_capex.cell(i + 1, j)
            c.width = c_widths[j]
            set_cell_shading(c, bg)
            set_cell_padding(c, top=32, bottom=32, left=60, right=60)
            set_cell_borders(c, bottom={"val": "single", "sz": "4", "color": "D9E2EC"})
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if j == 2 else (WD_ALIGN_PARAGRAPH.RIGHT if j == 3 else WD_ALIGN_PARAGRAPH.LEFT)
            r = p.add_run(val)
            r.font.name = 'Arial'
            r.font.size = Pt(8)
            r.font.color.rgb = COLOR_TEXT
            if j == 0:
                r.font.bold = True

    # Total Row
    set_row_props(tbl_capex.rows[len(capex_data) + 1])
    c_tot_label = tbl_capex.cell(len(capex_data) + 1, 0)
    c_tot_desc = tbl_capex.cell(len(capex_data) + 1, 1)
    c_tot_vol = tbl_capex.cell(len(capex_data) + 1, 2)
    c_tot_val = tbl_capex.cell(len(capex_data) + 1, 3)
    
    for c, w in zip([c_tot_label, c_tot_desc, c_tot_vol, c_tot_val], c_widths):
        c.width = w
        set_cell_shading(c, "FEF3C7")
        set_cell_padding(c, top=60, bottom=60, left=80, right=80)
        set_cell_borders(c, top={"val": "single", "sz": "12", "color": "102A43"}, bottom={"val": "single", "sz": "12", "color": "102A43"})

    p_tot1 = c_tot_label.paragraphs[0]
    p_tot1.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r_tot1 = p_tot1.add_run("TOTAL PERMODALAN AWAL")
    r_tot1.font.name = 'Arial'
    r_tot1.font.size = Pt(8.5)
    r_tot1.font.bold = True
    r_tot1.font.color.rgb = COLOR_PRIMARY

    p_tot4 = c_tot_val.paragraphs[0]
    p_tot4.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_tot4 = p_tot4.add_run("Rp 15.000.000")
    r_tot4.font.name = 'Arial'
    r_tot4.font.size = Pt(9)
    r_tot4.font.bold = True
    r_tot4.font.color.rgb = COLOR_PRIMARY

    doc.add_page_break()

    # ==================== PAGE 6: BAB 5 STRUKTUR HPP & PROYEKSI FINANSIAL ====================
    add_h1("5. STRUKTUR HPP, SIMULASI KINERJA FINANSIAL & BEP")
    
    add_h2("5.1 Rincian HPP dan Margin Keuntungan Menu Kunci")
    add_p(
        "Kalkulasi HPP divalidasi langsung dari harga grosir riil distributor (CEDEA Dumpling Keju 500g @Rp 29rb-35rb isi 25 pcs, CEDEA/Minaku Otak-Otak Singapore @Rp 25rb-35rb, CEDEA Fish Roll 250g @Rp 15.900, kopi Kapal Api dus @Rp 1.083/sachet, serta piring rotan plastik PP tebal grosir @Rp 15rb-19rb/lusin). Strategi penetrasi awal mengutamakan harga jual realistis dan terjangkau bagi warga lokal, tanpa mematok margin berlebihan (rata-rata laba kotor 35%–50%):"
    )

    hpp_data = [
        ("Sosis Bakar Sapi/Ayam (Saus BBQ & Mayo)", "Sosis sapi/ayam olahan, oles mentega, saus BBQ bakar arang", "Rp 1.600", "Rp 3.000", "Rp 1.400 (46,7%)"),
        ("Bakso Sapi Bakar Arang (3 Butir)", "Bakso sapi olahan kenyal, bumbu kecap lada hitam bakar arang", "Rp 1.400", "Rp 2.500", "Rp 1.100 (44,0%)"),
        ("Dumpling Ayam / Keju Lumer (2 Pcs)", "CEDEA Dumpling beku isi keju lumer bakar arang batok", "Rp 2.400", "Rp 3.500", "Rp 1.100 (31,4%)"),
        ("Otak-Otak Singapore Bakar Arang", "CEDEA / Minaku otak-otak bakar kerat silang pedas manis", "Rp 1.500", "Rp 2.500", "Rp 1.000 (40,0%)"),
        ("Sate Cumi Olahan Bakar (Squid Flower)", "Olahan daging cumi flower bakar arang saus kecap pedas gurih", "Rp 2.000", "Rp 3.500", "Rp 1.500 (42,9%)"),
        ("Fish Roll / Crab Stick Olahan Bakar", "CEDEA Fish roll / kani stik rempah arang batok kelapa", "Rp 1.400", "Rp 2.500", "Rp 1.100 (44,0%)"),
        ("Nasi Kucing Teri / Orek Tempe", "Nasi pulen sambal teri balado / orek tempe bungkus daun pisang", "Rp 1.500", "Rp 2.500", "Rp 1.000 (40,0%)"),
        ("Tempe Mendoan Hangat (3 Lembar)", "Tempe kedelai lokal, adonan tepung daun bawang, kecap rawit", "Rp 2.000", "Rp 3.000", "Rp 1.000 (33,3%)"),
        ("Kopi Tubruk 'Jangkar Asli' (Signature)", "Bubuk Robusta sangrai lereng Gn. Karang, gula aren/pasir", "Rp 1.200", "Rp 3.000", "Rp 1.800 (60,0%)"),
        ("Kopi Sachet Populer (Renceng/Dus)", "Good Day / Kapal Api / Indocafe / Luwak White Koffie", "Rp 1.400", "Rp 3.000", "Rp 1.600 (53,3%)"),
        ("Es / Hangat Minuman Sachet Segar", "Nutrisari Jeruk / Extra Joss Susu / Kuku Bima Susu", "Rp 2.000", "Rp 3.500", "Rp 1.500 (42,9%)"),
        ("Susu Murni / Kopi Rempah Samudra", "Susu sapi murni / Robusta jahe merah rempah (Mug jadul)", "Rp 3.500", "Rp 6.000", "Rp 2.500 (41,7%)"),
    ]

    tbl_hpp = doc.add_table(rows=len(hpp_data) + 1, cols=5)
    tbl_hpp.alignment = WD_TABLE_ALIGNMENT.CENTER
    h_widths = [Cm(4.3), Cm(5.6), Cm(2.2), Cm(2.2), Cm(2.8)]
    
    headers_hpp = ["Nama Menu", "Komposisi Bahan Pokok", "HPP Riil", "Harga Jual", "Laba Kotor (Margin)"]
    for j, h in enumerate(headers_hpp):
        c = tbl_hpp.cell(0, j)
        c.width = h_widths[j]
        set_cell_shading(c, "102A43")
        set_cell_padding(c, top=45, bottom=45, left=60, right=60)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if j >= 2 else WD_ALIGN_PARAGRAPH.LEFT
        r = p.add_run(h)
        r.font.name = 'Arial'
        r.font.size = Pt(8)
        r.font.bold = True
        r.font.color.rgb = COLOR_WHITE
    set_row_props(tbl_hpp.rows[0], is_header=True)

    for i, row in enumerate(hpp_data):
        bg = "F0F4F8" if i % 2 == 0 else "FFFFFF"
        set_row_props(tbl_hpp.rows[i + 1])
        for j, val in enumerate(row):
            c = tbl_hpp.cell(i + 1, j)
            c.width = h_widths[j]
            set_cell_shading(c, bg)
            set_cell_padding(c, top=16, bottom=16, left=45, right=45)
            set_cell_borders(c, bottom={"val": "single", "sz": "4", "color": "D9E2EC"})
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if j in (2, 3) else (WD_ALIGN_PARAGRAPH.RIGHT if j == 4 else WD_ALIGN_PARAGRAPH.LEFT)
            r = p.add_run(val)
            r.font.name = 'Arial'
            r.font.size = Pt(7.5)
            r.font.color.rgb = COLOR_TEXT
            if j == 0:
                r.font.bold = True
            elif j == 4:
                r.font.bold = True
                r.font.color.rgb = COLOR_SUCCESS

    add_h2("5.2 Proyeksi Kinerja Keuangan & Waktu Balik Modal (BEP)")
    
    fin_scenarios = [
        ("Kinerja Harian", "Konservatif (Hujan/Sepi)", "Moderat (Target Realistis)", "Agresif (Ramai Akhir Pekan)"),
        ("Estimasi Tamu / Hari", "25 orang / hari", "40 orang / hari", "55 orang / hari"),
        ("Rata-rata Belanja / Tamu", "Rp 11.000 / orang", "Rp 12.500 / orang", "Rp 14.000 / orang"),
        ("Pendapatan Harian", "Rp 275.000 / hari", "Rp 500.000 / hari", "Rp 770.000 / hari"),
        ("Omzet Bulanan (30 Hari)", "Rp 8.250.000", "Rp 15.000.000", "Rp 23.100.000"),
        ("HPP Bahan Pokok (~52%)", "Rp 4.290.000", "Rp 7.800.000", "Rp 12.012.000"),
        ("Laba Kotor Usaha", "Rp 3.960.000", "Rp 7.200.000", "Rp 11.088.000"),
        ("Beban Operasional (OPEX)", "Rp 1.760.000", "Rp 2.200.000", "Rp 3.088.000"),
        ("Laba Bersih Usaha / Bulan", "Rp 2.200.000", "Rp 5.000.000", "Rp 8.000.000"),
        ("Dividen Investor (40%)", "Rp 880.000 / bulan", "Rp 2.000.000 / bulan", "Rp 3.200.000 / bulan"),
        ("Bagian Pengelola (60%)", "Rp 1.320.000 / bulan", "Rp 3.000.000 / bulan", "Rp 4.800.000 / bulan"),
        ("Waktu Balik Modal (BEP)", "17,0 Bulan", "7,5 Bulan (~7,5 Bulan)", "4,7 Bulan")
    ]

    tbl_fin = doc.add_table(rows=len(fin_scenarios), cols=4)
    tbl_fin.alignment = WD_TABLE_ALIGNMENT.CENTER
    f_widths = [Cm(5.1), Cm(4.0), Cm(4.0), Cm(4.0)]

    for i, row in enumerate(fin_scenarios):
        is_hdr = (i == 0)
        is_bep = (i == len(fin_scenarios) - 1)
        is_div = (i == len(fin_scenarios) - 3)
        bg = "102A43" if is_hdr else ("FEF3C7" if (is_bep or is_div) else ("F0F4F8" if i % 2 == 1 else "FFFFFF"))
        set_row_props(tbl_fin.rows[i], is_header=is_hdr)
        for j, val in enumerate(row):
            c = tbl_fin.cell(i, j)
            c.width = f_widths[j]
            set_cell_shading(c, bg)
            set_cell_padding(c, top=18 if not is_hdr else 32, bottom=18 if not is_hdr else 32, left=50, right=50)
            if not is_hdr:
                set_cell_borders(c, bottom={"val": "single", "sz": "4", "color": "D9E2EC"})
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if j == 0 else WD_ALIGN_PARAGRAPH.RIGHT
            r = p.add_run(val)
            r.font.name = 'Arial'
            r.font.size = Pt(8)
            r.font.bold = is_hdr or is_bep or is_div or (j == 0)
            r.font.color.rgb = COLOR_WHITE if is_hdr else (COLOR_PRIMARY if (is_bep or is_div) else COLOR_TEXT)

    doc.add_page_break()

    # ==================== PAGE 7: BAB 6 TATA KELOLA, KESIMPULAN & LEMBAR KOMITMEN ====================
    add_h1("6. SKEMA KEMITRAAN MUDHARABAH & FORMULIR KOMITMEN")
    
    add_h2("6.1 Struktur Penawaran Kemitraan Modal Rp 15.000.000")
    add_p(
        "Penawaran kemitraan Jangkar Kopi dijalankan menggunakan akad Syirkah Mudharabah yang adil, berkah, dan transparan:"
    )

    add_callout(
        "OPSI PARTISIPASI MODAL & HAK INVESTOR",
        [
            ("• Pilihan Slot Permodalan: ", "Terbuka opsi 1 Mitra Tunggal Penuh (Rp 15.000.000), 2 Slot Kemitraan (@ Rp 7.500.000), 3 Slot Kemitraan (@ Rp 5.000.000), atau Sindikasi 5 Slot Ringan (@ Rp 3.000.000/slot)."),
            ("• Nisbah Bagi Hasil Laba Bersih: ", "40% dialokasikan untuk Pihak Pemodal (Shahibul Maal) dan 60% dialokasikan untuk Pihak Pengelola Operasional (Mudharib)."),
            ("• Transparansi Kasir Digital: ", "Setiap transaksi dicatat real-time melalui aplikasi POS Cloud di ponsel. Laporan keuangan bulanan dan rekonsiliasi kas dibagikan resmi pada tanggal 1 setiap bulannya."),
            ("• Roadmap Ekspansi Bertahap: ", "Fasilitas kursi lipat camping dan mesin espresso komersial sengaja ditunda di awal. Pengadaan fasilitas lanjutan tersebut akan didanai mandiri dari laba ditahan operasional setelah 3 bulan berjalan stabil, tanpa membebani modal awal investor.")
        ],
        border_color="102A43",
        bg_color="F0F4F8"
    )

    add_h2("6.2 Kesimpulan Eksekutif")
    add_p(
        "Jangkar Kopi & Angkringan adalah jawaban atas kebutuhan ruang temu warga yang bersahaja, bersih, dan berbiaya terjangkau di koridor Cikedal - Menes. Dengan modal lean Rp 15.000.000, beban sewa lahan terkunci aman 1 tahun (Rp 6 Jt), stok awal dan cadangan kas yang solid, serta menu arang tradisional murah meriah tanpa mematok margin berlebihan di awal, usaha ini diproyeksikan mencapai titik impas dalam tempo 7,5 bulan pada skenario moderat (4,7 bulan skenario agresif) dan siap berkembang menjadi ikon kuliner malam rakyat."
    )

    add_h2("6.3 Lembar Pernyataan Komitmen Kemitraan Investasi")
    add_p("Saya yang bertanda tangan di bawah ini menyatakan persetujuan dan komitmen awal untuk berpartisipasi dalam permodalan usaha Jangkar Kopi & Angkringan:")

    add_callout(
        "FORMULIR DATA MITRA INVESTOR",
        [
            ("Nama Lengkap Calon Mitra : ", "..........................................................................................................."),
            ("Nomor WhatsApp / Kontak   : ", "..........................................................................................................."),
            ("Alamat / Domisili          : ", "..........................................................................................................."),
            ("Pilihan Partisipasi Modal : ", "[   ] 1 Slot Sindikasi (Rp 3.000.000)      [   ] 1 Slot Kemitraan (Rp 5.000.000)\n                              [   ] 1 Slot Kemitraan (Rp 7.500.000)      [   ] Mitra Tunggal Penuh (Rp 15.000.000)")
        ],
        border_color="B45309",
        bg_color="FEF3C7"
    )

    tbl_sign = doc.add_table(rows=1, cols=2)
    tbl_sign.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell_l, cell_r = tbl_sign.cell(0, 0), tbl_sign.cell(0, 1)
    cell_l.width = Cm(8.5)
    cell_r.width = Cm(8.5)
    set_cell_padding(cell_l, top=20, bottom=20, left=40, right=40)
    set_cell_padding(cell_r, top=20, bottom=20, left=40, right=40)

    p_sl = cell_l.paragraphs[0]
    p_sl.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_sl.add_run("Pandeglang, Oktober 2026\nPIHAK PENGELOLA USAHA\nJangkar Kopi & Angkringan\n\n\n\n\n").font.size = Pt(8.5)
    r_sl_name = p_sl.add_run("FATIH FARHAT ASSHIDIQ\n")
    r_sl_name.bold = True
    r_sl_name.font.size = Pt(9)
    p_sl.add_run("Founder & Managing Director\nWA: +68 821-1150-0190 / +62 895-0610-0075\nEmail: kazokuhairy@gmail.com").font.size = Pt(8)

    p_sr = cell_r.paragraphs[0]
    p_sr.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_sr.add_run("Disetujui dan Diterima Oleh,\nPIHAK MITRA INVESTOR\nJangkar Kopi & Angkringan\n\n\n\n\n").font.size = Pt(8.5)
    r_sr_name = p_sr.add_run("( ........................................................... )\n")
    r_sr_name.bold = True
    r_sr_name.font.size = Pt(9)
    p_sr.add_run("Calon Mitra Investor / Shahibul Maal\nTanggal: .......................................................").font.size = Pt(8)

    # Save DOCX
    out_dir = "/media/fatihfarhat/New Volume1/FATIH DATA/ZeroSpot Cafe"
    docx_path = os.path.join(out_dir, "PROPOSAL_INVESTASI_JANGKAR_KOPI.docx")
    pdf_path = os.path.join(out_dir, "PROPOSAL_INVESTASI_JANGKAR_KOPI.pdf")
    
    doc.save(docx_path)
    print(f"DOCX saved to: {docx_path}")

    # Convert to PDF using LibreOffice
    cmd = [
        "libreoffice",
        "--headless",
        "--convert-to", "pdf",
        docx_path,
        "--outdir", out_dir
    ]
    res = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    if res.returncode == 0:
        print(f"PDF successfully generated: {pdf_path}")
    else:
        print(f"Error during PDF conversion: {res.stderr}")

if __name__ == "__main__":
    create_jangkar_kopi_proposal()
