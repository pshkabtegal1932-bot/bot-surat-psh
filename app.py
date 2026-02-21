import streamlit as st
import datetime
import io
import re
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
import google.generativeai as genai

# --- KONEKSI AMAN ---
try:
    GEMINI_API_KEY = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=GEMINI_API_KEY)
except:
    st.error("API Key bermasalah di Secrets!")
    st.stop()

@st.cache_resource
def load_ai_model():
    available = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
    return genai.GenerativeModel('gemini-1.5-flash' if 'models/gemini-1.5-flash' in available else available[0])

model = load_ai_model()

# --- FUNGSI GANTI TEKS (FONT TIMES NEW ROMAN & POSISI TETAP) ---
def ganti_teks_presisi(doc, tag, text_isi, is_agenda=False):
    for paragraph in doc.paragraphs:
        if tag in paragraph.text:
            # Bersihkan tag
            paragraph.text = paragraph.text.replace(tag, "")
            
            lines = text_isi.split('\n')
            for i, line in enumerate(lines):
                clean_line = re.sub(r'[*#_]', '', line).strip()
                if not clean_line: continue
                
                # Tambah text ke paragraf yang ada atau buat baru tepat setelahnya
                target_p = paragraph if i == 0 else paragraph.insert_paragraph_before("")
                target_p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                
                if is_agenda and ":" in clean_line:
                    # Logika Titik Dua Lurus (Posisi 2.0 Inci) sesuai image_6aaa44
                    label, isi = clean_line.split(":", 1)
                    target_p.paragraph_format.left_indent = Inches(0.5)
                    target_p.paragraph_format.tab_stops.add_tab_stop(Inches(2.0), WD_TAB_ALIGNMENT.LEFT)
                    run = target_p.add_run(f"{label.strip()}\t: {isi.strip()}")
                else:
                    # Paragraf Narasi (Menjorok 0.5 Inci)
                    target_p.paragraph_format.first_line_indent = Inches(0.5)
                    run = target_p.add_run(clean_line)
                
                # KUNCI FONT TIMES NEW ROMAN 11-12PT
                run.font.name = 'Times New Roman'
                run.font.size = Pt(11)

# --- UI DASHBOARD ---
st.title("🛡️ Sekretaris Digital PSH Tegal")

with st.form("input_psh"):
    col1, col2 = st.columns(2)
    with col1:
        nomor = st.text_input("Nomor Surat", value="005/PSH/II/2026")
        hal = st.text_input("Perihal", value="Undangan Halal Bi Halal")
    with col2:
        yth = st.text_input("Kepada Yth", value="Seluruh Warga PSH Tegal")
    
    arahan = st.text_area("Instruksi (Contoh: Rapat Minggu besok jam 8 malam di TC):")
    submit = st.form_submit_button("✨ Susun Surat")

if submit:
    with st.spinner("AI sedang merakit kalimat resmi..."):
        prompt = (f"Buat draf surat PSH Tegal dari: {arahan}. "
                  "Bagi jadi 2 bagian dipisah tanda '==='. "
                  "Bagian 1: Kalimat pembuka formal dan luwes. "
                  "Bagian 2: Rincian agenda 'Label : Isi'. "
                  "Tanpa salam dan nomor.")
        st.session_state['draf_raw'] = model.generate_content(prompt).text

# --- FITUR EDIT & RESET ---
if 'draf_raw' in st.session_state:
    st.subheader("📝 Edit & Review Draf")
    draf_edit = st.text_area("Sesuaikan draf di bawah (Jaga pemisah ===):", 
                             value=st.session_state['draf_raw'], height=250)
    st.session_state['draf_raw'] = draf_edit

    col_btn1, col_btn2 = st.columns(2)
    with col_btn1:
        if st.button("💾 Generate File Word"):
            try:
                doc = Document("template_psh.docx")
                pembuka_ai, agenda_ai = draf_edit.split("===")

                # Isi Header
                header = {"{{nomor}}": nomor, "{{hal}}": hal, "{{yth}}": yth, "{{tanggal}}": "21 Februari 2026"}
                for p in doc.paragraphs:
                    for k, v in header.items():
                        if k in p.text: p.text = p.text.replace(k, v)

                # Isi Konten (Kunci: Agenda di atas TTD)
                ganti_teks_presisi(doc, "{{pembuka}}", pembuka_ai.strip())
                ganti_teks_presisi(doc, "{{agenda}}", agenda_ai.strip(), is_agenda=True)

                out = io.BytesIO()
                doc.save(out)
                st.session_state['file_final'] = out.getvalue()
                st.success("Surat siap!")
            except:
                st.error("Gagal! Cek file template_psh.docx lo di GitHub.")

    with col_btn2:
        if st.button("🗑️ Hapus / Reset"):
            del st.session_state['draf_raw']
            if 'file_final' in st.session_state: del st.session_state['file_final']
            st.rerun()

    if 'file_final' in st.session_state:
        st.download_button("📥 Download Surat", data=st.session_state['file_final'], 
                           file_name=f"Surat_PSH_{nomor.replace('/','_')}.docx")
