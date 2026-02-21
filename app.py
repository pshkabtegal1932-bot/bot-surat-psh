import streamlit as st
import datetime
import io
import re
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
import google.generativeai as genai

# --- KONEKSI AMAN (MENGAMBIL DARI SECRETS) ---
try:
    # Memanggil API Key secara aman dari Secrets Streamlit Cloud
    GEMINI_API_KEY = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=GEMINI_API_KEY)
except Exception:
    st.error("Error: API Key belum diisi di menu Secrets Streamlit!")
    st.stop()

@st.cache_resource
def load_ai_model():
    try:
        # Mencari model yang tersedia secara otomatis
        available = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
        name = 'models/gemini-1.5-flash' if 'models/gemini-1.5-flash' in available else available[0]
        return genai.GenerativeModel(name)
    except:
        st.error("Gagal koneksi ke server AI Google.")
        st.stop()

model = load_ai_model()

st.set_page_config(page_title="Sekretaris PSH Tegal", page_icon="📝")

# --- LOGIKA PENULISAN SURAT SEKRETARIS ---
def format_surat_sekretaris(doc, tag, content):
    """
    Mengatur paragraf menjorok dan poin-poin dengan titik dua sejajar.
    """
    for paragraph in doc.paragraphs:
        if tag in paragraph.text:
            paragraph.text = paragraph.text.replace(tag, "")
            
            lines = content.split('\n')
            for i, line in enumerate(lines):
                clean_line = re.sub(r'[*#_]', '', line).strip()
                if not clean_line: continue
                
                new_p = doc.add_paragraph()
                new_p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                
                # FORMAT POIN (Contoh: Acara : Halal Bihalal)
                if ":" in clean_line and len(clean_line.split(":")[0]) < 25:
                    label, detail = clean_line.split(":", 1)
                    new_p.paragraph_format.left_indent = Inches(0.5)
                    # Titik dua lurus di 2.0 inci
                    new_p.paragraph_format.tab_stops.add_tab_stop(Inches(2.0), WD_TAB_ALIGNMENT.LEFT)
                    run = new_p.add_run(f"{label.strip()}\t: {detail.strip()}")
                else:
                    # FORMAT PARAGRAF NARASI
                    new_p.paragraph_format.first_line_indent = Inches(0.5)
                    run = new_p.add_run(clean_line)
                
                run.font.name = 'Times New Roman'
                run.font.size = Pt(12)

# --- UI DASHBOARD ---
st.title("🛡️ Sekretaris Digital PSH Tegal")
st.info("Status: Online. AI akan menyusun kalimat sesuai pakem surat resmi PSH.")

with st.form("form_surat"):
    col1, col2 = st.columns(2)
    with col1:
        nomor = st.text_input("Nomor Surat", placeholder="Contoh: 003/PENGKAB.PSH/II/2026")
        hal = st.text_input("Perihal", placeholder="Contoh: Undangan Kegiatan")
    with col2:
        yth = st.text_input("Kepada Yth", placeholder="Contoh: Seluruh Warga PSH Tegal")
    
    arahan = st.text_area("Inti Pesan:", placeholder="Tulis instruksi di sini (Contoh: Rapat tgl 25 jam 8 malam di TC)...")
    submit = st.form_submit_button("✨ Susun Kalimat Resmi")

if submit:
    with st.spinner("Sekretaris sedang merangkai kata..."):
        try:
            prompt = (f"Bertindaklah sebagai Sekretaris PSH Tegal. Susun isi surat dari arahan ini: {arahan}. "
                      "INSTRUKSI KHUSUS: "
                      "1. Pakai kalimat pembuka 'Sehubungan dengan...' yang luwes. "
                      "2. Gunakan gaya bahasa persaudaraan yang sopan. "
                      "3. Poin rincian (Waktu, Tempat, dll) harus berformat 'Label : Isi'. "
                      "4. Tambahkan paragraf penutup 'Demikian surat ini kami sampaikan...'. "
                      "5. JANGAN menulis salam 'Assalamu'alaikum' (Sudah ada di template). "
                      "6. Gunakan huruf kapital normal (Bukan CAPSLOCK).")
            
            response = model.generate_content(prompt)
            st.session_state['draf_isi'] = response.text.strip()
            st.success("Draf Selesai!")
        except Exception as e:
            st.error(f"Gagal generate: {e}")

if 'draf_isi' in st.session_state:
    st.subheader("📝 Edit & Cetak")
    isi_final = st.text_area("Review draf sekretaris:", value=st.session_state['draf_isi'], height=300)
    st.session_state['draf_isi'] = isi_final

    if st.button("💾 Siapkan File Word"):
        try:
            doc = Document("template_psh.docx")
            tgl_now = f"Tegal, {datetime.datetime.now().day} Februari 2026"
            
            mapping = {"{{nomor}}": nomor, "{{hal}}": hal, "{{yth}}": yth, "{{tanggal}}": tgl_now}
            
            for old, new in mapping.items():
                for p in doc.paragraphs:
                    if old in p.text:
                        for run in p.runs:
                            if old in run.text:
                                run.text = run.text.replace(old, str(new))
                                run.font.name = 'Times New Roman'

            format_surat_sekretaris(doc, "{{isi}}", st.session_state['draf_isi'])
            
            out = io.BytesIO()
            doc.save(out)
            st.session_state['download_file'] = out.getvalue()
            st.success("Surat siap didownload!")
        except Exception as e:
            st.error(f"Error Word: {e}")

    if 'download_file' in st.session_state:
        st.download_button("📥 Download Surat", data=st.session_state['download_file'], 
                           file_name=f"Surat_PSH_{nomor.replace('/','_')}.docx")
