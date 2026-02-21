import streamlit as st
import datetime
import io
import re
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
import google.generativeai as genai

# --- KONFIGURASI ---
GEMINI_API_KEY = "AIzaSyBtoq-CLs6GMZYzMFS6tYrBrefXRJYG5Bo"
genai.configure(api_key=GEMINI_API_KEY)

# SOLUSI AMPUH: Mencari model yang tersedia secara otomatis agar tidak 404
@st.cache_resource
def load_ai_model():
    available_models = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
    # Prioritas 1: 1.5 Flash
    if 'models/gemini-1.5-flash' in available_models:
        return genai.GenerativeModel('gemini-1.5-flash')
    # Prioritas 2: 1.0 Pro
    elif 'models/gemini-pro' in available_models:
        return genai.GenerativeModel('gemini-pro')
    # Terakhir: Ambil apa saja yang ada
    return genai.GenerativeModel(available_models[0].replace('models/', ''))

model = load_ai_model()

st.set_page_config(page_title="PSH Tegal Dashboard", page_icon="📝")

# --- FUNGSI PENDUKUNG ---
def get_tanggal_indo():
    bulan_indo = {1: "Januari", 2: "Februari", 3: "Maret", 4: "April", 5: "Mei", 6: "Juni",
                  7: "Juli", 8: "Agustus", 9: "September", 10: "Oktober", 11: "November", 12: "Desember"}
    now = datetime.datetime.now()
    return f"{now.day} {bulan_indo[now.month]} {now.year}"

def format_word_pro(doc, tag, content):
    for paragraph in doc.paragraphs:
        if tag in paragraph.text:
            paragraph.text = paragraph.text.replace(tag, "")
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            tab_stops = paragraph.paragraph_format.tab_stops
            tab_stops.add_tab_stop(Inches(1.5), WD_TAB_ALIGNMENT.LEFT)
            lines = content.split('\n')
            for i, line in enumerate(lines):
                if ":" in line:
                    label, detail = line.split(":", 1)
                    run = paragraph.add_run(f"{label.strip()}\t: {detail.strip()}")
                else:
                    run = paragraph.add_run(line)
                run.font.name = 'Times New Roman'
                run.font.size = Pt(12)
                if i < len(lines) - 1:
                    paragraph.add_run("\n")

# --- UI DASHBOARD ---
st.title("🛡️ Generator Surat PSH Tegal")
st.info("Status: Siap digunakan. Sistem otomatis mencari model AI terbaik.")

with st.form("form_surat"):
    col1, col2 = st.columns(2)
    with col1:
        nomor = st.text_input("Nomor Surat", placeholder="Contoh: 001/PSH/2026")
        hal = st.text_input("Perihal", placeholder="Contoh: Undangan Rapat")
    with col2:
        yth = st.text_input("Kepada Yth", placeholder="Contoh: Seluruh Anggota PSH")
    
    agenda_input = st.text_area("Rincian Agenda / Ide Surat", placeholder="Tulis jadwal, tempat, atau inti pesan di sini...")
    submit_ai = st.form_submit_button("✨ Susun Draf dengan AI")

# --- PROSES AI ---
if submit_ai:
    if not agenda_input:
        st.error("Isi dulu detail agendanya, Bro!")
    else:
        with st.spinner("AI sedang merakit surat..."):
            try:
                prompt = (f"Buat isi surat resmi PSH Tegal dari: {agenda_input}. "
                          "Formal, indentasi, format LABEL : ISI, penutup sopan, tanpa salam, tanpa markdown.")
                response = model.generate_content(prompt)
                draf_hasil = re.sub(r'[*#_]', '', response.text).strip()
                st.session_state['draf_final'] = draf_hasil
                st.success("Draf berhasil disusun!")
            except Exception as e:
                st.error(f"Error AI: {str(e)}")

# --- EDITOR & DOWNLOAD ---
if 'draf_final' in st.session_state:
    isi_edit = st.text_area("Edit Hasil AI (Jika perlu):", value=st.session_state['draf_final'], height=250)
    st.session_state['draf_final'] = isi_edit

    if st.button("📄 Buat File Word"):
        try:
            doc = Document("template_psh.docx")
            mapping = {
                "{{nomor}}": nomor, "{{hal}}": hal, "{{yth}}": yth,
                "{{pembuka}}": "Assalamu’alaikum Warahmatullahi Wabarakatuh,\nSalam Persaudaraan,",
                "{{tanggal}}": get_tanggal_indo()
            }
            for old, new in mapping.items():
                for p in doc.paragraphs:
                    if old in p.text:
                        for run in p.runs:
                            if old in run.text:
                                run.text = run.text.replace(old, str(new))
                                run.font.name = 'Times New Roman'

            format_word_pro(doc, "{{agenda}}", st.session_state['draf_final'])
            format_word_pro(doc, "{{isi}}", st.session_state['draf_final'])

            output = io.BytesIO()
            doc.save(output)
            output.seek(0)

            st.download_button(
                label="📥 Klik untuk Download Surat",
                data=output,
                file_name=f"Surat_PSH_{nomor.replace('/', '-')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        except Exception as e:
            st.error(f"Gagal cetak: {str(e)}")

