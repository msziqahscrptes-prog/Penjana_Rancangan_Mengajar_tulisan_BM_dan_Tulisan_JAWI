import streamlit as st
import google.generativeai as genai
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from io import BytesIO

# --- 1. CONFIGURATION & CORE SETTINGS ---
st.set_page_config(page_title="Penjana RPH IRK Brunei (2047)", layout="wide")
st.title("🕌 PENJANA RANCANGAN MENGAJAR IRK BRUNEI [2047]")
st.caption("Sistem dwi-bahasa (Bahasa Melayu & Tulisan Jawi) khusus untuk sukatan pelajaran Cambridge GCE O Level IRK.")

# Top Bar API Authentication
user_api_key = st.text_input(
    "🔑 MASUKKAN KUNCI API GEMINI ANDA:", 
    type="password", 
    help="Dapatkan kunci API anda dari Google AI Studio."
)

def get_working_model(api_key):
    try:
        genai.configure(api_key=api_key)
        for m in genai.list_models():
            if 'generateContent' in m.supported_generation_methods:
                return m.name
    except Exception as e:
        st.error(f"KUNCI API TIDAK SAH ATAU RALAT SAMBUNGAN: {str(e)}")
        return None
    return "models/gemini-1.5-flash"

selected_model_name = None
if user_api_key:
    selected_model_name = get_working_model(user_api_key)
    if selected_model_name:
        st.info(f"SISTEM DISAMBUNGKAN. MODEL AKTIF: {selected_model_name.upper()}")
else:
    st.warning("⚠️ SILA MASUKKAN KUNCI API GEMINI PERIBADI ANDA DI ATAS UNTUK BERMULA.")


# --- 2. DUAL-LANGUAGE AI ENGINE ROUTINES ---
def generate_irk_plan(topic, extra_context, api_key, model_name, language_mode):
    """
    Generates a lesson plan for Islamic Religious Knowledge (IRK) Brunei 2047.
    language_mode can be 'MALAY' or 'JAWI'
    """
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel(model_name)
    
    jawi_instruction = (
        "Output MUST be fully written in TULISAN JAWI characters/script except for the labels specified. "
        "Do not use Romanized Malay for the content blocks; translate everything properly into standard Jawi script."
    ) if language_mode == 'JAWI' else "Output must be completely written in standard Romanized Bahasa Melayu."

    prompt = f"""
    Topik/Tajuk: {topic}. 
    Sukatan Pelajaran: Islamic Religious Knowledge (IRK) Brunei Syllabus Code 2047 (Cambridge GCE O Level).
    Konteks Tambahan: {extra_context}.
    Bahasa Output: {language_mode}.
    
    {jawi_instruction}

    PERATURAN FORMAT TEKS KRITIKAL:
    1. JANGAN gunakan dua tanda asteris (**) di mana-mana bahagian output.
    2. Pastikan setiap tajuk bahagian (SECTION:) ditulis dalam HURUF BESAR SEPENUHNYA menggunakan rumi/huruf inggeris sebagai pembahagi teknikal.
    3. JANGAN gunakan senarai bulet atau titik. Gunakan sistem penomboran (1, 2, 3...) untuk semua senarai item.
    4. JANGAN gunakan perkataan MURID. Digantikan dengan perkataan PELAJAR untuk isi kandungan menyeluruh.
    
    Strukturkan rancangan pengajaran harian ini dengan penanda eksak berikut:
    
    SECTION: TOPIK DAN KOD SUKATAN
    {topic.upper()} (IRK BRUNEI 2047)
    
    SECTION: OBJEKTIF PEMBELAJARAN
    [4 mata menggunakan nombor 1 hingga 4]
    
    SECTION: HASIL PEMBELAJARAN
    [4 mata menggunakan nombor 1 hingga 4]
    
    SECTION: DALIL NAQLI DAN RUJUKAN KITAB
    [Ayat Al-Quran, Hadis, atau rujukan kitab utama yang berkaitan dengan topik ini menggunakan nombor 1, 2]
    
    SECTION: KRITERIA KEJAYAAN
    [4 mata menggunakan nombor 1 hingga 4]
    
    SECTION: KATA KUNCI SYARIAT / ISTILAH
    [6 istilah penting menggunakan nombor 1 hingga 6]
    
    SECTION: SOALAN KEMAHIRAN BERFIKIR ARAS TINGGI (KBAT)
    [4 soalan berbeza menggunakan nombor 1 hingga 4]
    
    SECTION: AKTIVITI PENGAJARAN TERDEZA (KUMPULAN TINGGI, SEDERHANA, RENDAH)
    1. Aktiviti Tahap Tinggi: 
    2. Aktiviti Tahap Sederhana:
    3. Aktiviti Tahap Rendah:
    
    SECTION: AKTIVITI PEMBELAJARAN UTAMA (TERADUN & DIGITAL)
    [Langkah-langkah pengajaran berasaskan teknologi, persediaan ustaz/ustazah, dan tugasan murid menggunakan nombor 1, 2, 3...]
    
    SECTION: PENILAIAN & TUGASAN RUMAH
    [Bentuk pentaksiran serta kuiz pengukuhan menggunakan nombor 1, 2]
    """
    try:
        response = model.generate_content(prompt)
        return response.text.replace("**", "")
    except Exception as e:
        return f"RALAT SISTEM ({language_mode}): {str(e)}"


# --- 3. DYNAMIC WORD EXPORT ENGINE WITH ALIGNMENT CONTROLS ---
def add_page_number(run):
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = "PAGE"
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')
    fldChar3 = OxmlElement('w:fldChar')
    fldChar3.set(qn('w:fldCharType'), 'end')
    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)
    run._r.append(fldChar3)

def create_word_export(topic, text, is_jawi=False):
    doc = Document()
    
    # Page setup adjustments
    section_geo = doc.sections[0]
    section_geo.page_width = Inches(8.5)
    section_geo.page_height = Inches(11.5)
    section_geo.top_margin = Inches(0.4)
    section_geo.bottom_margin = Inches(0.4)
    section_geo.left_margin = Inches(0.4)
    section_geo.right_margin = Inches(0.4)
    
    # Page Number Header
    header = section_geo.header
    header_p = header.paragraphs[0]
    header_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    header_run = header_p.add_run()
    header_run.font.name = 'Arial' if not is_jawi else 'Traditional Arabic'
    header_run.font.size = Pt(10)
    add_page_number(header_run)

    # Document Typography styles
    style = doc.styles['Normal']
    font = style.font
    if is_jawi:
        font.name = 'Traditional Arabic'
        font.size = Pt(18)  # Jawi script scales better at larger font sizes
    else:
        font.name = 'Arial'
        font.size = Pt(12)
        
    p_format = style.paragraph_format
    p_format.line_spacing = 1.15
    p_format.space_after = Pt(12)

    # Document Title Block
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if is_jawi else WD_ALIGN_PARAGRAPH.LEFT
    lang_label = "TULISAN JAWI" if is_jawi else "BAHASA MELAYU"
    run_title = title_p.add_run(f'RANCANGAN PENGAJARAN HARIAN ({lang_label}): {topic.upper()}')
    run_title.bold = True

    # Administrative Table
    admin_table = doc.add_table(rows=3, cols=4)
    admin_table.style = 'Table Grid'
    labels = [
        ["TARIKH:", "HARI:"], 
        ["TEMPAT / BILIK:", "MINGGU NO:"], 
        ["BILANGAN MURID:", "TEMPOH (MINIT):"]
    ]
    for r in range(3):
        admin_table.cell(r, 0).paragraphs[0].add_run(labels[r][0]).bold = True
        admin_table.cell(r, 2).paragraphs[0].add_run(labels[r][1]).bold = True
    doc.add_paragraph()

    # Dynamic Section Grid parsing
    sections = text.split('SECTION:')
    for section in sections:
        if not section.strip(): 
            continue
            
        lines = section.strip().split('\n')
        title = lines[0].strip().upper().replace("**", "")
        body_content = "\n".join(lines[1:])

        p_sec = doc.add_paragraph()
        p_sec.alignment = WD_ALIGN_PARAGRAPH.RIGHT if is_jawi else WD_ALIGN_PARAGRAPH.LEFT
        p_sec.add_run(title).bold = True

        table = doc.add_table(rows=1, cols=1)
        table.style = 'Table Grid'
        cell_p = table.cell(0, 0).paragraphs[0]
        cell_p.paragraph_format.line_spacing = 1.15
        cell_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if is_jawi else WD_ALIGN_PARAGRAPH.LEFT
        
        cleaned_body = body_content.strip().replace("**", "")
        cell_p.add_run(cleaned_body if cleaned_body else "Tiada kandungan.")
        doc.add_paragraph()

    # Principal/HOD Evaluation Box
    doc.add_page_break()
    p_hod = doc.add_paragraph()
    p_hod.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_hod.add_run("PENGESAHAN DAN ULASAN PENGETUA / KETUA JABATAN (HOD)").bold = True
    
    hod_table = doc.add_table(rows=3, cols=2)
    hod_table.style = 'Table Grid'
    hod_table.cell(0, 0).paragraphs[0].add_run("ULASAN / REMARK:").bold = True
    hod_table.cell(0, 1).paragraphs[0].add_run("TANDATANGAN & COP JABATAN:").bold = True
    hod_table.rows[1].height = Pt(50)
    hod_table.cell(2, 0).paragraphs[0].add_run("TARIKH SEMAKAN:").bold = True
    hod_table.cell(2, 1).paragraphs[0].add_run("NAMA PENYEMAK:").bold = True

    # Adjust inner spacing properties
    for row in admin_table.rows:
        for cell in row.cells: cell.paragraphs[0].paragraph_format.line_spacing = 1.0
    for row in hod_table.rows:
        for cell in row.cells: cell.paragraphs[0].paragraph_format.line_spacing = 1.0

    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio


# --- 4. MAIN USER INTERFACE GRAPHICS ---
st.write("---")
u_topic = st.text_input("TOPIK / TAJUK PELAJARAN IRK:", placeholder="Contoh: Pemantapan Akidah: Kesan Syirik Dalam Kehidupan")
u_extra = st.text_area("KONTEKS TAMBAHAN / NOTA KHUSUS (PILIHAN):", placeholder="Sebutkan rujukan Surah, Kitab Fiqh, atau aktiviti berkumpulan tertentu di sini...")

if st.button("🚀 JANA DWU-VERSI RPH (MALAY & JAWI)", type="primary"):
    if not user_api_key:
        st.error("❌ SILA MASUKKAN KUNCI API GEMINI ANDA DI BAHAGIAN ATAS HALAMAN TERLEBIH DAHULU.")
    elif not u_topic:
        st.error("❌ SILA ISI RUANGAN TOPIK / TAJUK PELAJARAN.")
    else:
        # Generate Romanized Malay Version
        with st.spinner("AI Sedang merangka RPH Bahasa Melayu..."):
            malay_result = generate_irk_plan(u_topic, u_extra, user_api_key, selected_model_name, "MALAY")
            st.session_state['irk_malay_out'] = malay_result
            
        # Generate Tulisan Jawi Version
        with st.spinner("AI Sedang menterjemah & merangka RPH Tulisan Jawi..."):
            jawi_result = generate_irk_plan(u_topic, u_extra, user_api_key, selected_model_name, "JAWI")
            st.session_state['irk_jawi_out'] = jawi_result

# --- 5. SIDE-BY-SIDE DISPLAY & DUAL EXPORT BUTTONS ---
if 'irk_malay_out' in st.session_state and 'irk_jawi_out' in st.session_state:
    st.divider()
    st.subheader("👁️ PRATONTON DWU-VERSI RPH (PREVIEW SIDES)")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("### 📝 Versi Rumi (Bahasa Melayu)")
        st.text_area("PREVIEW TULISAN RUMI", st.session_state['irk_malay_out'], height=400)
        
        malay_doc = create_word_export(u_topic, st.session_state['irk_malay_out'], is_jawi=False)
        st.download_button(
            label="📥 DOWNLOAD WORD: VERSI RUMI (.DOCX)",
            data=malay_doc,
            file_name=f"RPH_IRK_2047_RUMI_{u_topic.replace(' ', '_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key="btn_dl_malay"
        )
        
    with col2:
        st.markdown("### 🕌 Versi TULISAN JAWI")
        # Injection of custom HTML text-area styling to force true Right-to-Left alignment in browser preview
        jawi_display_text = st.session_state['irk_jawi_out']
        st.markdown(
            f'<textarea style="width:100%; height:400px; direction:rtl; text-align:right; font-family:\'Traditional Arabic\', sans-serif; font-size:16px;" readonly>{jawi_display_text}</textarea>', 
            unsafe_allow_html=True
        )
        
        jawi_doc = create_word_export(u_topic, st.session_state['irk_jawi_out'], is_jawi=True)
        st.download_button(
            label="📥 DOWNLOAD WORD: VERSI JAWI (.DOCX)",
            data=jawi_doc,
            file_name=f"RPH_IRK_2047_JAWI_{u_topic.replace(' ', '_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key="btn_dl_jawi"
        )

# --- FOOTER SECTION ---
st.markdown("---") 
st.markdown(
    """
    <div style='text-align: center; color: grey; font-size: 0.8em;'>
        <p><b>Sistem Penjana RPH Pintar IRK Brunei (Syllabus 2047) v1.0 [Developer: H Nurul Haziqah]</b></p>
        <p>© 2026 Disediakan untuk Kegunaan Pegawai Pendidikan / Guru Pendidikan Islam Brunei</p>
    </div>
    """,
    unsafe_allow_html=True
)
