import streamlit as st
import datetime
import io
import re
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT

# --- KONEKSI AI (AUTO-SCAN MODEL) ---
try:
    import google.generativeai as genai
    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])

    def get_model():
        for m in genai.list_models():
            if 'generateContent' in m.supported_generation_methods:
                return m.name
        return "models/gemini-1.5-flash"
except:
    st.error("API Key belum diset di Secrets!")
    st.stop()

# --- FUNGSI RAKIT ISI (TIMES NEW ROMAN MUTLAK) ---
def rakit_isi_surat(doc, tag, content):
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
                
                # JIKA POIN AGENDA (Titik dua lurus 2.5 inci)
                if ":" in clean_line and len(clean_line.split(":")[0]) < 20:
                    label, detail = clean_line.split(":", 1)
                    new_p.paragraph_format.left_indent = Inches(1.0)
                    new_p.paragraph_format.tab_stops.add_tab_stop(Inches(2.5), WD_TAB_ALIGNMENT.LEFT)
                    run = new_p.add_run(f"{label.strip()}\t: {detail.strip()}")
                else:
                    # NARASI (Pembuka, Pakaian, Penutup)
                    new_p.paragraph_format.first_line_indent = Inches(0.5)
                    run = new_p.add_run(clean_line)
                
                run.font.name = 'Times New Roman'
                run.font.size = Pt(12)

# --- UI DASHBOARD LENGKAP ---
st.title("🛡️ Sekretaris Digital PSH Tegal")

with st.form("input_psh"):
    st.subheader("📌 Informasi Surat")
    col1, col2 = st.columns(2)
    with col1:
        nomor = st.text_input("Nomor Surat", value="005/PSH/II/2026")
        hal = st.text_input("Perihal (Hal)", value="Undangan Halal Bi Halal")
        lamp = st.text_input("Lampiran", value="-")
    with col2:
        yth = st.text_input("Kepada Yth", value="Seluruh Warga PSH Tegal")
        tempat = st.text_input("Tempat", value="Tempat")
        # Input Tanggal Surat Sesuai Konteks image_6b22a2
        tgl_surat = st.text_input("Tanggal Surat", value="21 Februari 2026")
    
    st.subheader("📝 Konten & Agenda")
    arahan = st.text_area("Instruksi Detail (Contoh: Rapat tgl 25 jam 8 malam di TC, baju silat lengkap):")
    
    submit = st.form_submit_button("✨ Susun Surat Resmi")

if submit:
    with st.spinner("AI sedang merakit draf..."):
        try:
            model = genai.GenerativeModel(get_model())
            prompt = (f"Buat draf surat PSH Tegal dari instruksi: {arahan}. Pisahkan dengan '==='. "
                      "1. Pembuka formal (menanyakan kabar). "
                      "2. Agenda RINGKAS (Acara, Waktu, Tempat). "
                      "3. Narasi Pakaian & Penutup yang sopan. "
                      "Jangan tulis salam pembuka, nomor, atau perihal lagi.")
            st.session_state['draf_raw'] = model.generate_content(prompt).text
        except Exception as e:
            st.error(f"Gagal koneksi AI: {e}")

# --- EDIT, RESET, & DOWNLOAD ---
if 'draf_raw' in st.session_state:
    st.subheader("📝 Review & Edit Sekretaris")
    draf_final = st.text_area("Edit draf (Jaga pemisah ===):", value=st.session_state['draf_raw'], height=300)
    st.session_state['draf_raw'] = draf_final

    c1, c2 = st.columns(2)
    with c1:
        if st.button("💾 Cetak ke Word"):
            try:
                doc = Document("template_psh.docx")
                pembuka, agenda, penutup = draf_final.split("===")

                # 1. Update Header & Tanggal (Times New Roman)
                h_map = {
                    "{{nomor}}": nomor, "{{hal}}": hal, "{{yth}}": yth, 
                    "{{lamp}}": lamp, "{{tempat}}": tempat,
                    "{{tanggal}}": f"Tegal, {tgl_surat}"
                }
                
                for p in doc.paragraphs:
                    for k, v in h_map.items():
                        if k in p.text:
                            p.text = p.text.replace(k, v)
                            for run in p.runs:
                                run.font.name = 'Times New Roman'
                                run.font.size = Pt(11)

                # 2. Masukkan Konten ke Tag image_6c077c
                rakit_isi_surat(doc, "{{pembuka}}", pembuka.strip())
                rakit_isi_surat(doc, "{{agenda}}", f"{agenda.strip()}\n{penutup.strip()}")

                out = io.BytesIO()
                doc.save(out)
                st.session_state['file_jadi'] = out.getvalue()
                st.success("Surat siap di-download!")
            except:
                st.error("Gagal! Cek apakah tag {{pembuka}} dan {{agenda}} ada di template.")
    
    with c2:
        if st.button("🗑️ Hapus / Reset"):
            for k in ['draf_raw', 'file_jadi']:
                if k in st.session_state: del st.session_state[k]
            st.rerun()

    if 'file_jadi' in st.session_state:
        st.download_button("📥 Download Surat", data=st.session_state['file_jadi'], 
                           file_name=f"Surat_PSH_{nomor.replace('/','_')}.docx")
