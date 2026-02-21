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
except Exception:
    st.error("Waduh, API Key belum disetting di Secrets Streamlit!")
    st.stop()

genai.configure(api_key=GEMINI_API_KEY)

@st.cache_resource
def load_ai_model():
    # Menggunakan model paling stabil agar tidak 404
    return genai.GenerativeModel('gemini-1.5-flash')

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
    Membersihkan tag dan mengisi konten dengan format rapi (Titik dua lurus).
    """
    for paragraph in doc.paragraphs:
        if tag in paragraph.text:
            paragraph.text = paragraph.text.replace(tag, "")
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            
            # Setting Tabulasi agar titik dua (:) lurus di 1.5 inci
            tab_stops = paragraph.paragraph_format.tab_stops
            tab_stops.add_tab_stop(Inches(1.5), WD_TAB_ALIGNMENT.LEFT)
            
            lines = content.split('\n')
            for i, line in enumerate(lines):
                # Menghilangkan teks dobel jika ada tag yang tidak sengaja tertulis lagi
                if tag in line: continue 
                
                if ":" in line:
                    label, detail = line.split(":", 1)
                    # Mengatur agar tidak CAPSLOCK berlebihan (Capitalize per kata)
                    run = paragraph.add_run(f"{label.strip().title()}\t: {detail.strip()}")
                else:
                    run = paragraph.add_run(line)
                
                run.font.name = 'Times New Roman'
                run.font.size = Pt(11)
                if i < len(lines) - 1:
                    paragraph.add_run("\n")

# --- UI DASHBOARD ---
st.title("🛡️ Generator Surat PSH Tegal")
st.info("Status: Online. Hasil akan diformat rapi (Times New Roman 11pt).")

# Form Input Utama
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
            with st.spinner("AI sedang menyusun kata-kata yang rapi..."):
                try:
                    # Prompt diperketat agar tidak dobel dan tidak capslock
                    prompt = (f"Buat isi surat resmi PSH Tegal dari: {agenda_input}. "
                              "Gunakan bahasa formal Indonesia. Jangan pakai CAPSLOCK semua. "
                              "Format: LABEL : ISI (Contoh: Acara : Rapat). "
                              "Tanpa salam pembuka, langsung inti saja. Jangan tulis lagi nomor surat/hal di dalam isi.")
                    
                    response = model.generate_content(prompt)
                    # Bersihkan karakter aneh
                    draf_hasil = re.sub(r'[*#_]', '', response.text).strip()
                    st.session_state['draf_final'] = draf_hasil
                except Exception as e:
                    st.error(f"Error AI: {str(e)}")

# --- AREA EDIT & EXPORT ---
if 'draf_final' in st.session_state:
    st.subheader("📝 Edit & Download")
    # Opsi Edit Ulang
    isi_edit = st.text_area("Cek kembali draf di bawah ini. Kamu bisa ubah manual jika ada yang salah:", 
                            value=st.session_state['draf_final'], height=300)
    st.session_state['draf_final'] = isi_edit

    col_btn1, col_btn2 = st.columns(2)
    
    with col_btn1:
        if st.button("💾 Proses File Word"):
            try:
                doc = Document("template_psh.docx")
                mapping = {
                    "{{nomor}}": nomor, "{{hal}}": hal, "{{yth}}": yth,
                    "{{pembuka}}": "Assalamu’alaikum Warahmatullahi Wabarakatuh,\nSalam Persaudaraan,",
                    "{{tanggal}}": get_tanggal_indo()
                }
                
                # Replace data header
                for old, new in mapping.items():
                    for p in doc.paragraphs:
                        if old in p.text:
                            for run in p.runs:
                                if old in run.text:
                                    run.text = run.text.replace(old, str(new))
                                    run.font.name = 'Times New Roman'

                # Isi bagian agenda dengan format rapi
                format_word_pro(doc, "{{agenda}}", st.session_state['draf_final'])
                format_word_pro(doc, "{{isi}}", st.session_state['draf_final'])

                output = io.BytesIO()
                doc.save(output)
                st.session_state['word_file'] = output.getvalue()
                st.success("File Word siap di-download!")
            except Exception as e:
                st.error(f"Gagal memproses Word: {e}")

    if 'word_file' in st.session_state:
        st.download_button(
            label="📥 Download Hasil (Word)",
            data=st.session_state['word_file'],
            file_name=f"Surat_PSH_{nomor.replace('/', '-')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        st.warning("Catatan: Untuk PDF, silakan gunakan fitur 'Save as PDF' di HP/Komputer kamu setelah membuka file Word ini agar format Kop Surat tidak berantakan.")

    if st.button("🔄 Reset / Buat Ulang"):
        del st.session_state['draf_final']
        if 'word_file' in st.session_state: del st.session_state['word_file']
        st.rerun()
