import streamlit as st
import datetime
import io
import re
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
import google.generativeai as genai

# --- KONFIGURASI AMAN ---
try:
    GEMINI_API_KEY = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=GEMINI_API_KEY)
except Exception:
    st.error("Waduh, API Key belum disetting di Secrets Streamlit!")
    st.stop()

@st.cache_resource
def load_ai_model():
    """
    Sistem deteksi model otomatis agar tidak kena error 404.
    """
    try:
        # Mencari daftar model yang memang aktif di API Key kamu
        available_models = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
        
        # Cek satu per satu mana yang ada
        if 'models/gemini-1.5-flash' in available_models:
            return genai.GenerativeModel('gemini-1.5-flash')
        elif 'models/gemini-pro' in available_models:
            return genai.GenerativeModel('gemini-pro')
        else:
            # Jika tidak ada yang cocok, ambil apa saja yang tersedia pertama kali
            return genai.GenerativeModel(available_models[0])
    except Exception as e:
        st.error(f"Gagal koneksi ke AI: {str(e)}. Pastikan API Key benar.")
        st.stop()

# Inisialisasi Model
model = load_ai_model()

st.set_page_config(page_title="PSH Tegal Dashboard", page_icon="📝")

# --- FUNGSI PENDUKUNG ---
def get_tanggal_indo():
    bulan_indo = {1: "Januari", 2: "Februari", 3: "Maret", 4: "April", 5: "Mei", 6: "Juni",
                  7: "Juli", 8: "Agustus", 9: "September", 10: "Oktober", 11: "November", 12: "Desember"}
    now = datetime.datetime.now()
    return f"{now.day} {bulan_indo[now.month]} {now.year}"

def format_word_pro(doc, tag, content):
    """
    Formatting rapi: Titik dua lurus, Times New Roman, Justify.
    """
    for paragraph in doc.paragraphs:
        if tag in paragraph.text:
            paragraph.text = paragraph.text.replace(tag, "")
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            
            # Setting Tabulasi di 1.5 inci
            tab_stops = paragraph.paragraph_format.tab_stops
            tab_stops.add_tab_stop(Inches(1.5), WD_TAB_ALIGNMENT.LEFT)
            
            lines = content.split('\n')
            for i, line in enumerate(lines):
                # Bersihkan sisa-sisa karakter markdown AI
                clean_line = re.sub(r'[*#_]', '', line).strip()
                if not clean_line or tag in clean_line: continue 
                
                if ":" in clean_line:
                    label, detail = clean_line.split(":", 1)
                    # title() untuk cegah CAPSLOCK berlebih
                    run = paragraph.add_run(f"{label.strip().title()}\t: {detail.strip()}")
                else:
                    run = paragraph.add_run(clean_line)
                
                run.font.name = 'Times New Roman'
                run.font.size = Pt(11)
                if i < len(lines) - 1:
                    paragraph.add_run("\n")

# --- UI DASHBOARD ---
st.title("🛡️ Generator Surat PSH Tegal")
st.info("Status: Online. AI akan mendeteksi model terbaik secara otomatis.")

with st.container():
    col1, col2 = st.columns(2)
    with col1:
        nomor = st.text_input("Nomor Surat", placeholder="Contoh: 001/PSH/2026")
        hal = st.text_input("Perihal", placeholder="Contoh: Undangan Rapat")
    with col2:
        yth = st.text_input("Kepada Yth", placeholder="Contoh: Seluruh Anggota PSH")
    
    agenda_input = st.text_area("Rincian Agenda / Inti Surat", placeholder="Tulis jadwal, tempat, atau inti pesan di sini...")
    
    if st.button("✨ Susun Draf dengan AI"):
        if not agenda_input:
            st.error("Isi dulu detail agendanya, Bro!")
        else:
            with st.spinner("AI sedang berpikir..."):
                try:
                    prompt = (f"Buat isi surat resmi PSH Tegal berdasarkan: {agenda_input}. "
                              "Gunakan Bahasa Indonesia formal. Jangan gunakan huruf kapital semua. "
                              "Format: NAMA ATRIBUT : ISI. Tanpa salam pembuka/penutup, "
                              "langsung rincian intinya saja. Jangan tulis lagi nomor surat.")
                    
                    response = model.generate_content(prompt)
                    st.session_state['draf_final'] = response.text.strip()
                    st.success("Draf berhasil disusun!")
                except Exception as e:
                    st.error(f"Error AI: {str(e)}")

# --- AREA EDIT & DOWNLOAD ---
if 'draf_final' in st.session_state:
    st.subheader("📝 Edit & Download")
    isi_edit = st.text_area("Edit manual hasil AI di sini:", 
                            value=st.session_state['draf_final'], height=300)
    st.session_state['draf_final'] = isi_edit

    if st.button("💾 Proses File Word"):
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
            st.session_state['word_file'] = output.getvalue()
            st.success("File Word sudah jadi!")
        except Exception as e:
            st.error(f"Gagal cetak Word: {e}")

    if 'word_file' in st.session_state:
        st.download_button(
            label="📥 Download Hasil (Word)",
            data=st.session_state['word_file'],
            file_name=f"Surat_PSH_{nomor.replace('/', '-')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        st.warning("Tips: Buka file Word lalu 'Save as PDF' untuk hasil cetak terbaik.")

    if st.button("🔄 Reset"):
        for key in ['draf_final', 'word_file']:
            if key in st.session_state: del st.session_state[key]
        st.rerun()
