import streamlit as st
import datetime
import io
import re
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
import google.generativeai as genai

# --- KONEKSI AMAN VIA SECRETS ---
try:
    # Memanggil dari Secrets (Kunci tidak akan bocor lagi)
    GEMINI_API_KEY = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=GEMINI_API_KEY)
except Exception:
    st.error("Waduh, API Key belum disetting di Secrets Streamlit!")
    st.stop()

@st.cache_resource
def load_ai_model():
    try:
        # Deteksi otomatis model yang aktif di akun kamu
        available_models = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
        model_name = 'models/gemini-1.5-flash' if 'models/gemini-1.5-flash' in available_models else available_models[0]
        return genai.GenerativeModel(model_name)
    except Exception as e:
        st.error(f"Gagal koneksi AI: {e}")
        st.stop()

model = load_ai_model()

st.set_page_config(page_title="Sekretaris PSH Tegal", page_icon="📝")

# --- LOGIKA PENULISAN SURAT RESMI ---
def format_surat_sekretaris(doc, tag, content):
    """
    Mengatur tata letak surat agar poin-poin lurus dan paragraf rapi.
    """
    for paragraph in doc.paragraphs:
        if tag in paragraph.text:
            paragraph.text = paragraph.text.replace(tag, "")
            
            lines = content.split('\n')
            for i, line in enumerate(lines):
                clean_line = re.sub(r'[*#_]', '', line).strip()
                if not clean_line: continue
                
                # Buat paragraf baru untuk tiap baris agar spasi terjaga
                new_p = doc.add_paragraph()
                new_p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                
                # JIKA BARIS ADALAH DAFTAR POIN (Misal: Tanggal : ...)
                if ":" in clean_line and len(clean_line.split(":")[0]) < 25:
                    label, detail = clean_line.split(":", 1)
                    # Beri jarak ke kanan (Indent)
                    new_p.paragraph_format.left_indent = Inches(0.5)
                    # Atur titik dua agar sejajar lurus di posisi 2.0 inci
                    tab_stops = new_p.paragraph_format.tab_stops
                    tab_stops.add_tab_stop(Inches(2.0), WD_TAB_ALIGNMENT.LEFT)
                    
                    run = new_p.add_run(f"{label.strip()}\t: {detail.strip()}")
                else:
                    # JIKA BARIS ADALAH PARAGRAPH BIASA
                    new_p.paragraph_format.first_line_indent = Inches(0.5) # Baris pertama masuk
                    run = new_p.add_run(clean_line)
                
                run.font.name = 'Times New Roman'
                run.font.size = Pt(12)

# --- UI DASHBOARD ---
st.title("🛡️ Sekretaris Digital PSH Tegal")
st.info("Ketik instruksi, AI akan menyusun kalimat resmi dengan format paragraf & poin yang rapi.")

with st.form("input_surat"):
    col1, col2 = st.columns(2)
    with col1:
        nomor = st.text_input("Nomor Surat", placeholder="Contoh: 003/PENGKAB.PSH/II/2026")
        hal = st.text_input("Perihal", placeholder="Contoh: Edaran Agenda PSH")
    with col2:
        yth = st.text_input("Kepada Yth", placeholder="Contoh: Seluruh Warga PSH Tegal")
    
    arahan = st.text_area("Tulis Inti Pesan (Sekretaris akan merangkainya):", 
                          placeholder="Contoh: Halal bihalal tgl 29 maret jam 10 pagi, tempat di TC PSH...")
    submit = st.form_submit_button("✨ Susun Surat Resmi")

if submit:
    with st.spinner("Sekretaris sedang mengetik..."):
        try:
            # PROMPT AGAR AI BERPERILAKU SEBAGAI SEKRETARIS PROFESIONAL
            prompt = (f"Bertindaklah sebagai Sekretaris Organisasi PSH Tegal. "
                      f"Buat isi surat resmi berdasarkan instruksi ini: {arahan}. "
                      "FORMAT WAJIB: "
                      "1. Awali dengan paragraf pembuka yang luwes (Contoh: Sehubungan dengan agenda PSH Tegal...). "
                      "2. Gunakan bahasa formal tapi tetap penuh rasa persaudaraan. "
                      "3. Jika ada rincian waktu/tempat, buat dalam daftar 'Label : Isi' (Misal: Tanggal : 20 Februari 2026). "
                      "4. Akhiri dengan paragraf penutup yang sopan. "
                      "5. JANGAN menulis salam 'Assalamu'alaikum' karena sudah ada di template kertas surat. "
                      "6. Gunakan huruf kapital di awal kata saja (Title Case).")
            
            response = model.generate_content(prompt)
            st.session_state['draf_final'] = response.text.strip()
            st.success("Draf berhasil disusun!")
        except Exception as e:
            st.error(f"Gagal generate: {e}")

if 'draf_final' in st.session_state:
    st.subheader("📝 Review & Edit Sekretaris")
    isi_edit = st.text_area("Edit draf jika ada kalimat yang kurang pas:", value=st.session_state['draf_final'], height=300)
    st.session_state['draf_final'] = isi_edit

    if st.button("💾 Generate File Word"):
        try:
            doc = Document("template_psh.docx")
            # Pengisian Header
            mapping = {"{{nomor}}": nomor, "{{hal}}": hal, "{{yth}}": yth, 
                       "{{tanggal}}": f"Tegal, {datetime.datetime.now().day} Februari 2026"}
            
            for old, new in mapping.items():
                for p in doc.paragraphs:
                    if old in p.text:
                        for run in p.runs:
                            if old in run.text:
                                run.text = run.text.replace(old, str(new))
                                run.font.name = 'Times New Roman'

            # Pengisian Isi Utama dengan sistem Tabulasi & Indent
            format_surat_sekretaris(doc, "{{isi}}", st.session_state['draf_final'])
            
            output = io.BytesIO()
            doc.save(output)
            st.session_state['file_ok'] = output.getvalue()
            st.success("Surat siap didownload!")
        except Exception as e:
            st.error(f"Error proses Word: {e}")

    if 'file_ok' in st.session_state:
        st.download_button("📥 Download Surat", data=st.session_state['file_ok'], 
                           file_name=f"Surat_PSH_{nomor.replace('/','_')}.docx")
