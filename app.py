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
    def get_active_model():
        for m in genai.list_models():
            if 'generateContent' in m.supported_generation_methods:
                return m.name
        return "models/gemini-1.5-flash"
except:
    st.error("Setting API Key di Secrets dulu, Bangsat!")
    st.stop()

st.set_page_config(page_title="Sekretaris PSH Tegal", page_icon="📝")

# --- FUNGSI RAKIT ISI (LINE SPACING SINGLE / RAPAT) ---
def rakit_isi_surat(doc, tag, text, is_agenda=False):
    for paragraph in doc.paragraphs:
        if tag in paragraph.text:
            paragraph.text = paragraph.text.replace(tag, "")
            lines = text.split('\n')
            for line in lines:
                clean = re.sub(r'[*#_]', '', line).strip()
                if not clean: continue
                
                # Insert paragraph baru dengan settingan ketat
                new_p = paragraph.insert_paragraph_before()
                new_p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                
                # KUNCI KERAPATAN (Line Spacing Single & No Spacing After/Before)
                new_p.paragraph_format.line_spacing = 1.0
                new_p.paragraph_format.space_after = Pt(0)
                new_p.paragraph_format.space_before = Pt(0)
                
                if is_agenda and ":" in clean:
                    # Poin Agenda Sejajar
                    label, detail = clean.split(":", 1)
                    new_p.paragraph_format.left_indent = Inches(1.0)
                    new_p.paragraph_format.tab_stops.add_tab_stop(Inches(2.5), WD_TAB_ALIGNMENT.LEFT)
                    run = new_p.add_run(f"{label.strip()}\t: {detail.strip()}")
                else:
                    # Narasi (Pembuka/Baju/Penutup)
                    new_p.paragraph_format.first_line_indent = Inches(0.5)
                    run = new_p.add_run(clean)
                
                run.font.name = 'Times New Roman'
                run.font.size = Pt(12)

# --- UI DASHBOARD LENGKAP ---
st.title("🛡️ Sekretaris Digital PSH Tegal")

with st.form("input_form"):
    col1, col2 = st.columns(2)
    with col1:
        nomor = st.text_input("Nomor Surat", value="005/PSH/II/2026")
        hal = st.text_input("Perihal (Hal)", value="Undangan Halal Bi Halal")
        tgl_surat = st.text_input("Tanggal Surat", value="21 Februari 2026")
    with col2:
        yth = st.text_input("Kepada Yth", value="Seluruh Warga PSH Tegal")
        lamp = st.text_input("Lampiran", value="-")
        tempat = st.text_input("Di (Tempat)", value="Tempat")
    
    arahan = st.text_area("Instruksi (Contoh: Rapat tgl 25 jam 8 malam, baju silat):")
    submit = st.form_submit_button("✨ Susun Surat")

if submit:
    try:
        model = genai.GenerativeModel(get_active_model())
        prompt = (f"Susun surat resmi PSH Tegal dari: {arahan}. Pisahkan dengan '---'. "
                  "1. Pembuka formal. 2. Agenda RINGKAS (Acara, Waktu, Tempat). "
                  "3. Narasi Baju & Penutup. Jangan tulis nomor/hal.")
        st.session_state['draf'] = model.generate_content(prompt).text
    except Exception as e:
        st.error(f"Error AI: {e}")

if 'draf' in st.session_state:
    draf_edit = st.text_area("Review draf (Jaga pemisah ---):", value=st.session_state['draf'], height=250)
    
    if st.button("💾 Cetak & Download"):
        try:
            doc = Document("template_psh.docx")
            parts = draf_edit.split("---")
            
            # 1. Update Header (Kunci Times New Roman)
            h_map = {
                "{{nomor}}": nomor, "{{hal}}": hal, "{{yth}}": yth,
                "{{lamp}}": lamp, "{{tempat}}": tempat,
                "{{tanggal}}": f"Tegal, {tgl_surat}"
            }
            for p in doc.paragraphs:
                p.paragraph_format.line_spacing = 1.0 # Rapatkan juga header
                for k, v in h_map.items():
                    if k in p.text:
                        p.text = p.text.replace(k, v)
                        for run in p.runs: run.font.name = 'Times New Roman'

            # 2. Isi Konten (Ketik ke Tag {{pembuka}} dan {{agenda}})
            rakit_isi_surat(doc, "{{pembuka}}", parts[0].strip())
            rakit_isi_surat(doc, "{{agenda}}", f"{parts[1].strip()}\n{parts[2].strip()}", is_agenda=True)

            # Hapus tag sisa
            for p in doc.paragraphs:
                if "{{pembuka}}" in p.text or "{{agenda}}" in p.text:
                    p.text = ""

            out = io.BytesIO()
            doc.save(out)
            st.download_button("📥 Download Surat Rapat", data=out.getvalue(), file_name="Surat_PSH_Rapi.docx")
        except:
            st.error("Gagal! Cek tag {{pembuka}} dan {{agenda}} di template.docx lo!")
