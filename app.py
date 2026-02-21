import streamlit as st
import io
import re
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT

# --- KONEKSI AI AMAN ---
try:
    import google.generativeai as genai
    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])

    def cari_model():
        # Scan model yang bisa dipake biar nggak 404
        for m in genai.list_models():
            if 'generateContent' in m.supported_generation_methods:
                return m.name
        return "models/gemini-1.5-flash"
except:
    st.error("API Key belum diset di Secrets!")
    st.stop()

# --- FUNGSI PENGISI TEMPLATE (KUNCI: TIMES NEW ROMAN) ---
def isi_surat_rapi(doc, tag, konten, is_agenda=False):
    for p in doc.paragraphs:
        if tag in p.text:
            p.text = p.text.replace(tag, "")
            lines = konten.split('\n')
            for line in lines:
                clean = re.sub(r'[*#_]', '', line).strip()
                if not clean: continue
                
                new_p = doc.add_paragraph()
                new_p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                
                if is_agenda and ":" in clean:
                    # Titik dua lurus sesuai image_6aaa44
                    label, detail = clean.split(":", 1)
                    new_p.paragraph_format.left_indent = Inches(1.0)
                    new_p.paragraph_format.tab_stops.add_tab_stop(Inches(2.5), WD_TAB_ALIGNMENT.LEFT)
                    run = new_p.add_run(f"{label.strip()}\t: {detail.strip()}")
                else:
                    # Narasi (Pembuka/Baju/Penutup) menjorok
                    new_p.paragraph_format.first_line_indent = Inches(0.5)
                    run = new_p.add_run(clean)
                
                run.font.name = 'Times New Roman'
                run.font.size = Pt(12)

# --- UI STREAMLIT ---
st.title("🛡️ Sekretaris Digital PSH Tegal")

with st.form("input_form"):
    nomor = st.text_input("Nomor Surat", value="005/PSH/II/2026")
    yth = st.text_input("Kepada Yth", value="Seluruh Warga PSH Tegal")
    arahan = st.text_area("Instruksi (Contoh: Rapat tgl 25 jam 8 malam, baju silat):")
    submit = st.form_submit_button("✨ Susun Surat")

if submit:
    try:
        model = genai.GenerativeModel(cari_model())
        prompt = (f"Buat draf surat PSH Tegal dari: {arahan}. Pisahkan dengan '==='. "
                  "Bagian 1: Pembuka formal. Bagian 2: Agenda (Acara, Tanggal, Waktu, Tempat). "
                  "Bagian 3: Narasi baju & Penutup. JANGAN tulis salam/nomor.")
        res = model.generate_content(prompt).text
        st.session_state['draf'] = res
    except Exception as e:
        st.error(f"Kuota habis atau error: {e}. Tunggu 1 menit, Kontol!")

if 'draf' in st.session_state:
    edit_draf = st.text_area("Edit draf (Gunakan === sebagai pemisah):", value=st.session_state['draf'], height=300)
    
    if st.button("💾 Download Word"):
        doc = Document("template_psh.docx")
        pembuka, agenda, penutup = edit_draf.split("===")
        
        # Isi Header
        h_map = {"{{nomor}}": nomor, "{{yth}}": yth, "{{tanggal}}": "21 Februari 2026"}
        for p in doc.paragraphs:
            for k, v in h_map.items():
                if k in p.text: p.text = p.text.replace(k, v)

        # Isi Konten ke Tag Mentah (image_6c077c)
        isi_surat_rapi(doc, "{{pembuka}}", pembuka.strip())
        isi_surat_rapi(doc, "{{agenda}}", f"{agenda.strip()}\n{penutup.strip()}", is_agenda=True)

        out = io.BytesIO()
        doc.save(out)
        st.download_button("📥 Ambil File", data=out.getvalue(), file_name="Surat_PSH.docx")
