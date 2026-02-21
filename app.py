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
    try:
        # Mencari daftar model otomatis agar tidak 404
        available_models = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
        model_name = 'models/gemini-1.5-flash' if 'models/gemini-1.5-flash' in available_models else available_models[0]
        return genai.GenerativeModel(model_name)
    except:
        st.error("Koneksi AI Bermasalah!")
        st.stop()

model = load_ai_model()

st.set_page_config(page_title="PSH Tegal Dashboard", page_icon="📝")

# --- FUNGSI TANGGAL INDO ---
def get_tanggal_indo():
    bulan_indo = {1: "Januari", 2: "Februari", 3: "Maret", 4: "April", 5: "Mei", 6: "Juni",
                  7: "Juli", 8: "Agustus", 9: "September", 10: "Oktober", 11: "November", 12: "Desember"}
    now = datetime.datetime.now()
    return f"{now.day} {bulan_indo[now.month]} {now.year}"

# --- FUNGSI FORMATTING SEKRETARIS ---
def format_surat_sekretaris(doc, tag, content):
    """
    Mengisi konten ke template dengan gaya surat resmi:
    Paragraf rapi (Justify) dan Poin-poin dengan titik dua lurus.
    """
    for paragraph in doc.paragraphs:
        if tag in paragraph.text:
            paragraph.text = paragraph.text.replace(tag, "")
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            
            lines = content.split('\n')
            for i, line in enumerate(lines):
                clean_line = re.sub(r'[*#_]', '', line).strip()
                if not clean_line: continue
                
                run = paragraph.add_run(clean_line)
                run.font.name = 'Times New Roman'
                run.font.size = Pt(12)
                
                # Jika baris mengandung titik dua, buat tabulasi lurus
                if ":" in clean_line:
                    paragraph.paragraph_format.tab_stops.add_tab_stop(Inches(1.5), WD_TAB_ALIGNMENT.LEFT)
                
                if i < len(lines) - 1:
                    paragraph.add_run("\n")

# --- UI DASHBOARD ---
st.title("🛡️ Sekretaris Digital PSH Tegal")
st.info("Ketik poin-poinnya saja, AI yang akan merangkai kalimat resminya.")

with st.form("input_surat"):
    col1, col2 = st.columns(2)
    with col1:
        nomor = st.text_input("Nomor Surat", placeholder="Contoh: 001/PSH/2026")
        hal = st.text_input("Perihal", placeholder="Contoh: Undangan Halal Bi Halal")
    with col2:
        yth = st.text_input("Kepada Yth", placeholder="Contoh: Seluruh Warga PSH Tegal")
    
    arahan = st.text_area("Apa inti suratnya?", placeholder="Contoh: Halal bihalal tanggal 24 maret di TC PSH jam 9 pagi, baju silat...")
    submit = st.form_submit_button("✨ Susun Surat Resmi")

if submit:
    if not arahan:
        st.error("Kasih arahan dulu ke Sekretarisnya, Bro!")
    else:
        with st.spinner("Sekretaris sedang mengetik..."):
            try:
                # PROMPT BARU: AI dipaksa jadi Sekretaris yang sopan
                prompt = (f"Bertindaklah sebagai Sekretaris PSH Tegal. Buatlah isi surat resmi berdasarkan arahan ini: {arahan}. "
                          "Gunakan struktur berikut: "
                          "1. Kalimat pembuka: 'Sehubungan dengan...' atau 'Dalam rangka...' yang relevan. "
                          "2. Kalimat pengantar sebelum poin. "
                          "3. Rincian acara dengan format 'Label : Isi' (Gunakan Huruf Kapital di awal kata saja). "
                          "4. Kalimat penutup yang formal dan sopan. "
                          "HANYA TULIS ISINYA SAJA. Jangan tulis salam Assalamuallaikum karena sudah ada di template.")
                
                response = model.generate_content(prompt)
                st.session_state['draf_final'] = response.text.strip()
            except Exception as e:
                st.error(f"Error AI: {str(e)}")

# --- EDITOR & DOWNLOAD ---
if 'draf_final' in st.session_state:
    st.subheader("📝 Review Draf Sekretaris")
    isi_final = st.text_area("Kamu bisa perbaiki kalimatnya di sini:", value=st.session_state['draf_final'], height=350)
    st.session_state['draf_final'] = isi_final

    if st.button("💾 Cetak ke Word"):
        try:
            doc = Document("template_psh.docx")
            # Isi Header
            mapping = {"{{nomor}}": nomor, "{{hal}}": hal, "{{yth}}": yth, "{{tanggal}}": get_tanggal_indo()}
            
            for old, new in mapping.items():
                for p in doc.paragraphs:
                    if old in p.text:
                        for run in p.runs:
                            if old in run.text:
                                run.text = run.text.replace(old, str(new))
                                run.font.name = 'Times New Roman'

            # Isi Konten (Mengganti {{isi}} atau {{agenda}})
            format_surat_sekretaris(doc, "{{isi}}", st.session_state['draf_final'])
            format_surat_sekretaris(doc, "{{agenda}}", st.session_state['draf_final'])

            output = io.BytesIO()
            doc.save(output)
            st.session_state['file_jadi'] = output.getvalue()
            st.success("Surat sudah rapi dan siap didownload!")
        except Exception as e:
            st.error(f"Gagal: {e}")

    if 'file_jadi' in st.session_state:
        st.download_button("📥 Download Surat (Word)", data=st.session_state['file_jadi'], file_name=f"Surat_PSH_{nomor.replace('/','-')}.docx")
