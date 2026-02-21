import streamlit as st
import io
import re
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT

# --- KONEKSI AI ---
try:
    import google.generativeai as genai
    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
    def get_active_model():
        for m in genai.list_models():
            if 'generateContent' in m.supported_generation_methods:
                return m.name
        return "models/gemini-1.5-flash"
except:
    st.error("Setting API Key di Secrets dulu, Bro!")
    st.stop()

st.set_page_config(page_title="Sekretaris PSH Tegal", page_icon="📝")

# --- FUNGSI RAKIT ISI (RAPAT & INDENTASI KHUSUS) ---
def rakit_isi_surat(doc, tag, text, is_agenda=False):
    for paragraph in doc.paragraphs:
        if tag in paragraph.text:
            paragraph.text = paragraph.text.replace(tag, "")
            lines = text.split('\n')
            for line in lines:
                raw_line = line.strip()
                if not raw_line: continue
                
                new_p = paragraph.insert_paragraph_before()
                
                # 1. DEFAULT RATA KANAN KIRI (JUSTIFY)
                new_p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                
                # 2. PERKETAT JARAK (SINGLE SPACING)
                new_p.paragraph_format.line_spacing = 1.0
                new_p.paragraph_format.space_after = Pt(0)
                new_p.paragraph_format.space_before = Pt(0)
                
                # 3. LOGIKA INDENTASI AWAL PARAGRAF (***)
                clean_text = re.sub(r'[*#_]', '', raw_line).strip()
                if raw_line.startswith("***"):
                    # Maju ke kanan cuma di awal baris ini
                    new_p.paragraph_format.first_line_indent = Inches(0.5)
                
                if is_agenda and ":" in clean_text:
                    # 4. POIN AGENDA SEJAJAR
                    label, detail = clean_text.split(":", 1)
                    new_p.paragraph_format.left_indent = Inches(1.0)
                    new_p.paragraph_format.tab_stops.add_tab_stop(Inches(2.5), WD_TAB_ALIGNMENT.LEFT)
                    run = new_p.add_run(f"{label.strip()}\t: {detail.strip()}")
                else:
                    run = new_p.add_run(clean_text)
                
                # 5. TIMES NEW ROMAN 11pt
                run.font.name = 'Times New Roman'
                run.font.size = Pt(11)

# --- UI DASHBOARD ---
st.title("🛡️ Sekretaris Digital PSH Tegal")

with st.form("input_form"):
    col1, col2 = st.columns(2)
    with col1:
        nomor = st.text_input("Nomor Surat", value="005/PSH/II/2026")
        hal = st.text_input("Perihal (Hal)", value="Undangan Halal Bi Halal")
        # Default tgl langsung isi (Tegal, 21 Februari 2026)
        tgl_surat = st.text_input("Tanggal Surat", value="21 Februari 2026")
    with col2:
        yth = st.text_input("Kepada Yth", value="Seluruh Warga PSH Tegal")
        lamp = st.text_input("Lampiran", value="-")
        tempat = st.text_input("Di (Tempat)", value="Tempat")
    
    arahan = st.text_area("Instruksi Agenda:")
    submit = st.form_submit_button("✨ Susun Surat")

if submit:
    with st.spinner("Lagi dibikinin sekretaris pintar..."):
        try:
            model = genai.GenerativeModel(get_active_model())
            prompt = (f"Olah instruksi: {arahan}. DILARANG buat salam/yth/alamat. "
                      "Fokus ke inti informasi. Gunakan '---' sebagai pemisah bagian.")
            st.session_state['draf_psh'] = model.generate_content(prompt).text
        except Exception as e:
            st.error(f"Error AI: {e}")

if 'draf_psh' in st.session_state:
    st.subheader("📝 Review & Edit Draf")
    draf_final = st.text_area("Edit manual (Gunakan *** untuk paragraf menjorok):", 
                              value=st.session_state['draf_psh'], height=300, key="editor_surat")
    st.session_state['draf_psh'] = draf_final

    if st.button("💾 Cetak & Download"):
        try:
            doc = Document("template_psh.docx")
            parts = draf_final.split("---")
            
            # 1. Update Header (Kunci: Jangan double 'Tegal')
            # Mapping langsung isi value tanpa tambahan kata statis
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

            # 2. Isi Konten (Ketik ke Tag {{pembuka}} dan {{agenda}})
            rakit_isi_surat(doc, "{{pembuka}}", parts[0].strip())
            if len(parts) > 1:
                rakit_isi_surat(doc, "{{agenda}}", parts[1].strip(), is_agenda=True)

            # Hapus tag sisa
            for p in doc.paragraphs:
                if "{{pembuka}}" in p.text or "{{agenda}}" in p.text: p.text = ""

            out = io.BytesIO()
            doc.save(out)
            st.download_button("📥 Download Surat", data=out.getvalue(), file_name="Surat_PSH_Tegal.docx")
        except Exception as e:
            st.error(f"Gagal Cetak: {e}")
