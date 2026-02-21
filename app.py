import streamlit as st
import datetime
import io
import re
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
import google.generativeai as genai

# --- KONFIGURASI API (PASTIKAN DI SECRETS) ---
try:
    GEMINI_API_KEY = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=GEMINI_API_KEY)
except:
    st.error("API Key bocor atau belum diisi di Secrets!")
    st.stop()

@st.cache_resource
def load_ai_model():
    available = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
    return genai.GenerativeModel('gemini-1.5-flash' if 'models/gemini-1.5-flash' in available else available[0])

model = load_ai_model()

# --- FUNGSI PENGISIAN KHUSUS (TIDAK MERUSAK TATA LETAK) ---
def ganti_tag_rapi(doc, tag, text_content, is_agenda=False):
    for p in doc.paragraphs:
        if tag in p.text:
            # Hapus tag, sisakan paragrafnya
            p.text = p.text.replace(tag, "")
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            
            lines = text_content.split('\n')
            for i, line in enumerate(lines):
                clean_line = re.sub(r'[*#_]', '', line).strip()
                if not clean_line: continue
                
                if i == 0:
                    run = p.add_run(clean_line)
                else:
                    new_p = p.insert_paragraph_before("") # Sisipkan baris baru di posisi tag
                    new_p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                    
                    if is_agenda and ":" in clean_line:
                        # Logika Titik Dua Lurus (2.0 inci)
                        new_p.paragraph_format.left_indent = Inches(0.5)
                        new_p.paragraph_format.tab_stops.add_tab_stop(Inches(2.0), WD_TAB_ALIGNMENT.LEFT)
                        run = new_p.add_run(f"{clean_line.split(':', 1)[0].strip()}\t: {clean_line.split(':', 1)[1].strip()}")
                    else:
                        # Paragraf Pembuka/Penutup (Menjorok)
                        new_p.paragraph_format.first_line_indent = Inches(0.5)
                        run = new_p.add_run(clean_line)
                
                run.font.name = 'Times New Roman'
                run.font.size = Pt(11)

# --- UI DASHBOARD ---
st.title("🛡️ Sekretaris Digital PSH Tegal")

with st.form("input_surat"):
    col1, col2 = st.columns(2)
    with col1:
        nomor = st.text_input("Nomor Surat", value="005/PSH/II/2026")
        hal = st.text_input("Perihal", value="Undangan Halal Bi Halal")
    with col2:
        yth = st.text_input("Kepada Yth", value="Seluruh Warga PSH Tegal")
    
    arahan = st.text_area("Inti Pesan:", placeholder="Contoh: Rapat tgl 25 jam 8 malam di TC...")
    submit = st.form_submit_button("✨ Susun & Rakit Surat")

if submit:
    with st.spinner("AI sedang merangkai kalimat..."):
        prompt = (f"Buat draf surat PSH Tegal dari arahan: {arahan}. "
                  "PISAHKAN JADI 2 BAGIAN dengan tanda '==='. "
                  "Bagian 1: Kalimat pembuka formal menanyakan kabar dan maksud surat. "
                  "Bagian 2: Rincian agenda dengan format 'Label : Isi'. "
                  "Jangan pakai salam pembuka/nomor lagi.")
        res = model.generate_content(prompt).text
        st.session_state['draf_raw'] = res

if 'draf_raw' in st.session_state:
    st.subheader("📝 Review Draf")
    edit_draf = st.text_area("Edit di sini (Gunakan === sebagai pemisah):", value=st.session_state['draf_raw'], height=300)
    
    if st.button("💾 Download File Word"):
        try:
            doc = Document("template_psh.docx")
            pembuka_ai, agenda_ai = edit_draf.split("===")

            # 1. Ganti Header (Nomor, Hal, Yth)
            header_map = {"{{nomor}}": nomor, "{{hal}}": hal, "{{yth}}": yth, "{{tanggal}}": "21 Februari 2026"}
            for p in doc.paragraphs:
                for k, v in header_map.items():
                    if k in p.text: p.text = p.text.replace(k, v)

            # 2. Masukkan Konten ke Tag (Tanpa merusak TTD di bawah)
            ganti_tag_rapi(doc, "{{pembuka}}", pembuka_ai.strip())
            ganti_tag_rapi(doc, "{{agenda}}", agenda_ai.strip(), is_agenda=True)

            out = io.BytesIO()
            doc.save(out)
            st.download_button("📥 Ambil File Surat", data=out.getvalue(), file_name="Surat_PSH_Jadi.docx")
        except Exception as e:
            st.error(f"Gagal! Pastikan file 'template_psh.docx' sudah di GitHub. Error: {e}")
