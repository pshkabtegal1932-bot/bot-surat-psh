import streamlit as st
import datetime
import io
import re
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT

# --- KONEKSI AMAN VIA SECRETS ---
try:
    import google.generativeai as genai
    GEMINI_API_KEY = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=GEMINI_API_KEY)
except:
    st.error("Setting dulu API Key di Secrets Streamlit, Bro!")
    st.stop()

st.set_page_config(page_title="Sekretaris PSH Tegal", page_icon="📝")

# --- FUNGSI SEARCH & REPLACE (TIMES NEW ROMAN MUTLAK) ---
def rakit_isi_surat(doc, tag, content):
    """
    Mengisi tag {{isi}} dengan Times New Roman, poin ringkas, dan paragraf rapi.
    """
    for paragraph in doc.paragraphs:
        if tag in paragraph.text:
            paragraph.text = paragraph.text.replace(tag, "")
            
            lines = content.split('\n')
            for line in lines:
                clean_line = re.sub(r'[*#_]', '', line).strip()
                if not clean_line: continue
                
                new_p = doc.add_paragraph()
                new_p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                new_p.paragraph_format.line_spacing = 1.15
                
                # LOGIKA POIN RINGKAS (Acara, Waktu, Tempat)
                if ":" in clean_line and len(clean_line.split(":")[0]) < 20:
                    label, detail = clean_line.split(":", 1)
                    new_p.paragraph_format.left_indent = Inches(1.0)
                    # Titik dua lurus di 2.5 inci (sesuai image_6aaa44)
                    new_p.paragraph_format.tab_stops.add_tab_stop(Inches(2.5), WD_TAB_ALIGNMENT.LEFT)
                    run = new_p.add_run(f"{label.strip()}\t: {detail.strip()}")
                else:
                    # NARASI (Pembuka, Pakaian, Penutup)
                    new_p.paragraph_format.first_line_indent = Inches(0.5)
                    run = new_p.add_run(clean_line)
                
                # KUNCI TIMES NEW ROMAN 12PT
                run.font.name = 'Times New Roman'
                run.font.size = Pt(12)

# --- UI DASHBOARD ---
st.title("🛡️ Sekretaris Digital PSH Tegal")

with st.form("input_psh"):
    col1, col2 = st.columns(2)
    with col1:
        nomor = st.text_input("Nomor Surat", value="005/PSH/II/2026")
        hal = st.text_input("Perihal", value="Undangan Halal Bi Halal")
    with col2:
        yth = st.text_input("Kepada Yth", value="Seluruh Warga PSH Tegal")
    
    arahan = st.text_area("Instruksi (Contoh: Rapat tgl 25 jam 8 malam di TC, baju silat):")
    submit = st.form_submit_button("✨ Susun Surat")

if submit:
    with st.spinner("AI sedang merangkai kalimat..."):
        try:
            # FIX ERROR NOTFOUND: Gunakan model 'gemini-1.5-flash-latest' atau 'gemini-pro'
            model = genai.GenerativeModel('gemini-1.5-flash')
            prompt = (f"Bertindaklah sebagai Sekretaris PSH Tegal. Susun surat resmi dari: {arahan}. "
                      "ATURAN KERAS: "
                      "1. Paragraf Pembuka: Formal & Persaudaraan. "
                      "2. Agenda: Hanya POIN RINGKAS (Acara, Tanggal, Waktu, Tempat). "
                      "3. Pakaian: JANGAN JADI POIN. Tulis dalam satu paragraf setelah agenda. "
                      "4. Paragraf Penutup: Harus ada (Demikian surat ini...). "
                      "5. Tanpa salam pembuka/nomor.")
            response = model.generate_content(prompt)
            st.session_state['draf_raw'] = response.text
        except Exception as e:
            st.error(f"Gagal koneksi AI: {e}. Coba ganti model di kode.")

# --- REVIEW & DOWNLOAD ---
if 'draf_raw' in st.session_state:
    st.subheader("📝 Review & Edit")
    draf_edit = st.text_area("Edit draf sebelum dicetak:", value=st.session_state['draf_raw'], height=300)
    st.session_state['draf_raw'] = draf_edit

    c1, c2 = st.columns(2)
    with c1:
        if st.button("💾 Cetak ke Word"):
            try:
                doc = Document("template_psh.docx")
                
                # Update Header & Tanggal (Times New Roman)
                header_map = {"{{nomor}}": nomor, "{{hal}}": hal, "{{yth}}": yth, "{{tanggal}}": "Tegal, 21 Februari 2026"}
                for p in doc.paragraphs:
                    for k, v in header_map.items():
                        if k in p.text:
                            p.text = p.text.replace(k, v)
                            for run in p.runs:
                                run.font.name = 'Times New Roman'

                # Masukkan Isi Utama
                rakit_isi_surat(doc, "{{isi}}", draf_edit)
                
                out = io.BytesIO()
                doc.save(out)
                st.session_state['file_ok'] = out.getvalue()
                st.success("Surat Ready!")
            except:
                st.error("Gagal! Pastikan file 'template_psh.docx' ada di GitHub.")
    
    with c2:
        if st.button("🗑️ Reset"):
            for k in ['draf_raw', 'file_ok']:
                if k in st.session_state: del st.session_state[k]
            st.rerun()

    if 'file_ok' in st.session_state:
        st.download_button("📥 Download Surat", data=st.session_state['file_ok'], file_name="Surat_PSH.docx")
