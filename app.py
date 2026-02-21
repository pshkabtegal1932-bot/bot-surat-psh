import streamlit as st
import datetime
import io
import re
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT

# --- KONEKSI AMAN & HANDLING QUOTA ---
try:
    import google.generativeai as genai
    GEMINI_API_KEY = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=GEMINI_API_KEY)
except:
    st.error("Waduh, cek API Key di Secrets Streamlit lo!")
    st.stop()

st.set_page_config(page_title="Sekretaris PSH Tegal", page_icon="📝")

# --- FUNGSI RAKIT SURAT (TIMES NEW ROMAN MUTLAK) ---
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
                
                # LOGIKA POIN RINGKAS (Acara, Tanggal, Waktu, Tempat)
                if ":" in clean_line and len(clean_line.split(":")[0]) < 20:
                    label, detail = clean_line.split(":", 1)
                    new_p.paragraph_format.left_indent = Inches(1.0)
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
        hal = st.text_input("Perihal", value="Undangan Kegiatan")
    with col2:
        yth = st.text_input("Kepada Yth", value="Seluruh Warga PSH Tegal")
    
    arahan = st.text_area("Instruksi (Contoh: Rapat tgl 25 jam 8 malam di TC, baju silat):")
    submit = st.form_submit_button("✨ Susun Surat")

if submit:
    with st.spinner("AI sedang merakit kalimat..."):
        try:
            # Gunakan model Flash yang lebih ringan kuotanya
            model = genai.GenerativeModel('gemini-1.5-flash')
            prompt = (f"Susun surat resmi PSH Tegal dari: {arahan}. "
                      "ATURAN KERAS: "
                      "1. Paragraf Pembuka: Formal. "
                      "2. Agenda: POIN RINGKAS (Acara, Tanggal, Waktu, Tempat). "
                      "3. Pakaian: JANGAN JADI POIN. Tulis dalam narasi paragraf setelah agenda. "
                      "4. Paragraf Penutup: Harus ada. "
                      "5. Tanpa salam pembuka/nomor.")
            response = model.generate_content(prompt)
            st.session_state['draf_raw'] = response.text
        except Exception as e:
            if "429" in str(e):
                st.error("Kuota Gratisan Habis! Tunggu 1 menit atau ganti API Key baru, Kontol!")
            else:
                st.error(f"Error: {e}")

# --- EDIT & DOWNLOAD ---
if 'draf_raw' in st.session_state:
    st.subheader("📝 Review & Edit")
    draf_edit = st.text_area("Sesuaikan draf:", value=st.session_state['draf_raw'], height=300)
    st.session_state['draf_raw'] = draf_edit

    c1, c2 = st.columns(2)
    with c1:
        if st.button("💾 Cetak ke Word"):
            try:
                doc = Document("template_psh.docx")
                # Update Header (Times New Roman)
                h_map = {"{{nomor}}": nomor, "{{hal}}": hal, "{{yth}}": yth, "{{tanggal}}": "Tegal, 21 Februari 2026"}
                for p in doc.paragraphs:
                    for k, v in h_map.items():
                        if k in p.text:
                            p.text = p.text.replace(k, v)
                            for run in p.runs: run.font.name = 'Times New Roman'

                rakit_isi_surat(doc, "{{isi}}", draf_edit)
                
                out = io.BytesIO()
                doc.save(out)
                st.session_state['file_ok'] = out.getvalue()
                st.success("Surat Ready!")
            except:
                st.error("Gagal! Pastikan 'template_psh.docx' ada di GitHub.")
    
    with c2:
        if st.button("🗑️ Reset"):
            for k in ['draf_raw', 'file_ok']:
                if k in st.session_state: del st.session_state[k]
            st.rerun()

    if 'file_ok' in st.session_state:
        st.download_button("📥 Download Surat", data=st.session_state['file_ok'], file_name="Surat_PSH.docx")
