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
except Exception:
    st.error("Waduh, API Key belum ada di Secrets Streamlit!")
    st.stop()

@st.cache_resource
def load_ai_model():
    try:
        available = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
        name = 'models/gemini-1.5-flash' if 'models/gemini-1.5-flash' in available else available[0]
        return genai.GenerativeModel(name)
    except:
        st.error("Koneksi AI Gagal!")
        st.stop()

model = load_ai_model()

# --- FUNGSI PENGISI TEMPLATE ---
def isi_template_psh(doc, arahan_ai):
    """
    Membagi draf AI menjadi Pembuka dan Agenda, lalu memasukkannya ke template.
    """
    # Pisahkan Pembuka dan Agenda berdasarkan pola
    parts = arahan_ai.split("---AGENDA---")
    pembuka_text = parts[0].strip()
    agenda_text = parts[1].strip() if len(parts) > 1 else ""

    for paragraph in doc.paragraphs:
        # 1. Mengisi Pembuka (Menjorok & Rapi)
        if "{{pembuka}}" in paragraph.text:
            paragraph.text = paragraph.text.replace("{{pembuka}}", "")
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            run = paragraph.add_run(pembuka_text)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(12)
            paragraph.paragraph_format.first_line_indent = Inches(0.5)

        # 2. Mengisi Agenda (Titik Dua Lurus)
        if "{{agenda}}" in paragraph.text:
            paragraph.text = paragraph.text.replace("{{agenda}}", "")
            lines = agenda_text.split('\n')
            for line in lines:
                new_p = doc.add_paragraph()
                new_p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                if ":" in line:
                    label, isi = line.split(":", 1)
                    new_p.paragraph_format.left_indent = Inches(0.5)
                    new_p.paragraph_format.tab_stops.add_tab_stop(Inches(2.0), WD_TAB_ALIGNMENT.LEFT)
                    run = new_p.add_run(f"{label.strip()}\t: {isi.strip()}")
                else:
                    run = new_p.add_run(line.strip())
                run.font.name = 'Times New Roman'
                run.font.size = Pt(12)

# --- UI DASHBOARD ---
st.title("🛡️ Sekretaris Digital PSH Tegal")

with st.form("input_psh"):
    col1, col2 = st.columns(2)
    with col1:
        nomor = st.text_input("Nomor Surat", value="005/PSH/II/2026")
        hal = st.text_input("Perihal", value="Halal bi Halal")
    with col2:
        yth = st.text_input("Kepada Yth", value="Seluruh Warga dan Kandidat PSH Tegal")
    
    arahan = st.text_area("Inti Pesan:", placeholder="Contoh: Halal bihalal tgl 29 maret jam 10 pagi di TC...")
    submit = st.form_submit_button("✨ Proses ke Template")

if submit:
    with st.spinner("AI sedang mengetik..."):
        prompt = (f"Bertindaklah sebagai Sekretaris PSH Tegal. Buat draf surat dari arahan: {arahan}. "
                  "Gunakan format ini: "
                  "[Tulis kalimat pembuka yang sopan dan persaudaraan] "
                  "---AGENDA--- "
                  "[Tulis poin-poin agenda dengan format Label : Isi]. "
                  "Jangan tulis salam pembuka/penutup.")
        response = model.generate_content(prompt)
        st.session_state['draf_raw'] = response.text.strip()

if 'draf_raw' in st.session_state:
    st.subheader("📝 Review Hasil")
    isi_edit = st.text_area("Edit draf sebelum dicetak:", value=st.session_state['draf_raw'], height=300)
    
    if st.button("💾 Download Surat (.docx)"):
        try:
            doc = Document("template_psh.docx")
            
            # Ganti Header
            mapping = {"{{nomor}}": nomor, "{{hal}}": hal, "{{yth}}": yth}
            for p in doc.paragraphs:
                for old, new in mapping.items():
                    if old in p.text:
                        p.text = p.text.replace(old, new)
                        for run in p.runs: run.font.name = 'Times New Roman'

            # Masukkan Pembuka & Agenda ke posisi yang benar
            isi_template_psh(doc, isi_edit)
            
            out = io.BytesIO()
            doc.save(out)
            st.download_button("📥 Download Sekarang", data=out.getvalue(), file_name=f"Surat_PSH_{nomor.replace('/','_')}.docx")
        except Exception as e:
            st.error(f"Gagal: Pastikan 'template_psh.docx' sudah diupload ke GitHub! Error: {e}")
