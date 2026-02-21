import streamlit as st
import io
import re
import time
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT

# --- KONEKSI AI (DIPERBAIKI) ---
try:
    import google.generativeai as genai
    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
    
    def panggil_ai_pintar(prompt):
        try:
            # Cari model yang tersedia secara dinamis
            models = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
            m_name = models[0] if models else "models/gemini-1.5-flash"
            
            model = genai.GenerativeModel(m_name)
            response = model.generate_content(prompt)
            return response.text
        except Exception as e:
            # Jika error beneran (bukan limit), kita tampilin apa adanya biar lo tau
            return f"ERROR_SISTEM: {str(e)}"
except:
    st.error("API Key di Secrets belum bener, Bro!")
    st.stop()

st.set_page_config(page_title="Sekretaris PSH Tegal", page_icon="📝")

# --- FUNGSI RAKIT DOCX (KUNCI FORMAT 11PT & SINGLE SPACING) ---
def rakit_isi_surat(doc, tag, text, is_agenda=False):
    for paragraph in doc.paragraphs:
        if tag in paragraph.text:
            paragraph.text = paragraph.text.replace(tag, "")
            lines = text.split('\n')
            for line in lines:
                raw_line = line.strip()
                if not raw_line: continue
                
                new_p = paragraph.insert_paragraph_before()
                new_p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                new_p.paragraph_format.line_spacing = 1.0
                new_p.paragraph_format.space_after = Pt(0)
                new_p.paragraph_format.space_before = Pt(0)
                
                # Fitur *** untuk indentasi awal paragraf
                clean_text = re.sub(r'[*#_]', '', raw_line).strip()
                if raw_line.startswith("***"):
                    new_p.paragraph_format.first_line_indent = Inches(0.5)
                
                if is_agenda and ":" in clean_text:
                    label, detail = clean_text.split(":", 1)
                    new_p.paragraph_format.left_indent = Inches(1.0)
                    new_p.paragraph_format.tab_stops.add_tab_stop(Inches(2.5), WD_TAB_ALIGNMENT.LEFT)
                    run = new_p.add_run(f"{label.strip()}\t: {detail.strip()}")
                else:
                    run = new_p.add_run(clean_text)
                
                run.font.name = 'Times New Roman'
                run.font.size = Pt(11)

# --- UI DASHBOARD ---
st.title("🛡️ Sekretaris Digital PSH Tegal")

# Input Header
with st.container():
    col1, col2 = st.columns(2)
    with col1:
        nomor = st.text_input("Nomor Surat", value="005/PSH/II/2026")
        hal = st.text_input("Perihal (Hal)", value="Undangan Rapat")
        tgl_surat = st.text_input("Tanggal Surat", value="21 Februari 2026")
    with col2:
        yth = st.text_input("Kepada Yth", value="Seluruh Warga PSH Tegal")
        lamp = st.text_input("Lampiran", value="-")
        tempat = st.text_input("Di (Tempat)", value="Tempat")

arahan = st.text_area("Instruksi Agenda (AI cuma jalan kalau tombol diklik):")

# TOMBOL GENERATE (Hanya jalan sekali per klik)
if st.button("✨ Susun Surat"):
    if arahan:
        with st.spinner("AI lagi kerja..."):
            prompt = (f"Buat isi surat resmi PSH Tegal dari instruksi: {arahan}. "
                      "Hanya tulis narasi dan agenda. Pisahkan dengan '---'. TANPA SALAM.")
            st.session_state['draf_psh'] = panggil_ai_pintar(prompt)
    else:
        st.warning("Isi dulu instruksinya, Bro!")

# AREA EDIT & CETAK (Terpisah dari Generate)
if 'draf_psh' in st.session_state:
    st.subheader("📝 Review & Edit Draf")
    
    # Kunci input manual biar nggak panggil AI lagi
    draf_final = st.text_area("Edit manual (*** = paragraf masuk):", 
                              value=st.session_state['draf_psh'], 
                              height=300)
    
    if st.button("💾 Cetak & Download"):
        try:
            doc = Document("template_psh.docx")
            parts = draf_final.split("---")
            
            # Update Header
            h_map = {
                "{{nomor}}": nomor, "{{hal}}": hal, "{{yth}}": yth,
                "{{lamp}}": lamp, "{{tempat}}": tempat,
                "{{tanggal}}": tgl_surat 
            }
            
            for p in doc.paragraphs:
                p.paragraph_format.line_spacing = 1.0
                for k, v in h_map.items():
                    if k in p.text:
                        p.text = p.text.replace(k, v)
                        for run in p.runs:
                            run.font.name = 'Times New Roman'
                            run.font.size = Pt(11)

            # Isi Konten ke {{pembuka}} dan {{agenda}}
            rakit_isi_surat(doc, "{{pembuka}}", parts[0].strip())
            if len(parts) > 1:
                rakit_isi_surat(doc, "{{agenda}}", parts[1].strip(), is_agenda=True)

            # Bersihkan tag
            for p in doc.paragraphs:
                if "{{pembuka}}" in p.text or "{{agenda}}" in p.text: p.text = ""

            out = io.BytesIO()
            doc.save(out)
            st.download_button("📥 Download Surat", data=out.getvalue(), file_name="Surat_PSH.docx")
        except Exception as e:
            st.error(f"Gagal Cetak: {e}")
