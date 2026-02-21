import streamlit as st
import datetime
import io
import re
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
import google.generativeai as genai

# --- KONEKSI AMAN (AMBIL DARI SECRETS) ---
try:
    GEMINI_API_KEY = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=GEMINI_API_KEY)
except Exception:
    st.error("Waduh, API Key belum disetting di Secrets Streamlit!")
    st.stop()

@st.cache_resource
def load_ai_model():
    try:
        available = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
        name = 'models/gemini-1.5-flash' if 'models/gemini-1.5-flash' in available else available[0]
        return genai.GenerativeModel(name)
    except:
        st.error("Koneksi AI Bermasalah!")
        st.stop()

model = load_ai_model()

# --- FUNGSI FORMATTING SESUAI KEINGINAN ---
def isi_ke_template(doc, tag, content):
    """
    Menghapus tag {{isi}} dan menggantinya dengan ketikan Sekretaris:
    - Paragraf Pembuka (Menjorok)
    - Poin-poin (Titik dua lurus di 2 inci)
    - Paragraf Penutup (Menjorok)
    """
    for paragraph in doc.paragraphs:
        if tag in paragraph.text:
            # Hapus tag {{isi}} dari paragraf tersebut
            paragraph.text = paragraph.text.replace(tag, "")
            
            lines = content.split('\n')
            for i, line in enumerate(lines):
                clean_line = re.sub(r'[*#_]', '', line).strip()
                if not clean_line: continue
                
                # Buat baris baru sebagai paragraf resmi
                new_p = doc.add_paragraph()
                new_p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                
                # JIKA BARIS ADALAH POIN ACARA (Ada titik dua)
                if ":" in clean_line and len(clean_line.split(":")[0]) < 25:
                    label, detail = clean_line.split(":", 1)
                    new_p.paragraph_format.left_indent = Inches(0.5) # Geser ke kanan
                    # Meluruskan titik dua di 2.0 inci
                    new_p.paragraph_format.tab_stops.add_tab_stop(Inches(2.0), WD_TAB_ALIGNMENT.LEFT)
                    run = new_p.add_run(f"{label.strip()}\t: {detail.strip()}")
                else:
                    # JIKA PARAGRAF NARASI (PEMBUKA/PENUTUP)
                    new_p.paragraph_format.first_line_indent = Inches(0.5) # Baris pertama menjorok
                    run = new_p.add_run(clean_line)
                
                run.font.name = 'Times New Roman'
                run.font.size = Pt(12)

# --- UI DASHBOARD ---
st.title("🛡️ Sekretaris Digital PSH Tegal")

with st.form("input_sekretaris"):
    col1, col2 = st.columns(2)
    with col1:
        nomor = st.text_input("Nomor Surat", placeholder="003/PENGKAB.PSH/II/2026")
        hal = st.text_input("Perihal", placeholder="Undangan Rapat")
    with col2:
        yth = st.text_input("Kepada Yth", placeholder="Seluruh Warga PSH")
    
    arahan = st.text_area("Instruksi Surat:", placeholder="Contoh: Rapat tgl 25 jam 8 malam di TC...")
    submit = st.form_submit_button("✨ Susun & Masukkan ke Template")

if submit:
    with st.spinner("Sekretaris sedang mengetik di template..."):
        try:
            prompt = (f"Bertindaklah sebagai Sekretaris PSH Tegal. Susun isi surat dari arahan: {arahan}. "
                      "Wajib ada paragraf pembuka 'Sehubungan dengan...', poin-poin acara yang rapi, "
                      "dan paragraf penutup 'Demikian surat ini...'. Jangan tulis salam dan nomor surat lagi.")
            response = model.generate_content(prompt)
            st.session_state['draf_isi'] = response.text.strip()
        except Exception as e:
            st.error(f"Error AI: {e}")

if 'draf_isi' in st.session_state:
    st.subheader("📝 Review Draf")
    isi_final = st.text_area("Edit manual jika perlu:", value=st.session_state['draf_isi'], height=300)
    
    if st.button("💾 Download Hasil (.docx)"):
        try:
            # MEMANGGIL FILE TEMPLATE KAMU
            doc = Document("template_psh.docx")
            
            # Isi Header otomatis
            mapping = {"{{nomor}}": nomor, "{{hal}}": hal, "{{yth}}": yth, 
                       "{{tanggal}}": f"Tegal, {datetime.datetime.now().day} Februari 2026"}
            
            for old, new in mapping.items():
                for p in doc.paragraphs:
                    if old in p.text:
                        for run in p.runs:
                            if old in run.text:
                                run.text = run.text.replace(old, str(new))

            # Memasukkan draf AI ke posisi {{isi}} di template Word
            isi_ke_template(doc, "{{isi}}", isi_final)
            
            # Export ke memori
            out = io.BytesIO()
            doc.save(out)
            st.download_button("📥 Klik untuk Download", data=out.getvalue(), file_name="Surat_PSH_Jadi.docx")
        except Exception as e:
            st.error(f"Gagal: Pastikan file 'template_psh.docx' sudah kamu upload ke GitHub! Error: {e}")
