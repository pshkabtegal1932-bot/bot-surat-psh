import streamlit as st
import io
import re
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT

# --- KONEKSI AI (DENGAN HANDLER MODEL) ---
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

# --- FUNGSI PENGISI KONTEN (LINE KETAT & FITUR INDENTASI) ---
def rakit_isi_surat(doc, tag, text, is_agenda=False):
    for paragraph in doc.paragraphs:
        if tag in paragraph.text:
            paragraph.text = paragraph.text.replace(tag, "")
            lines = text.split('\n')
            for line in lines:
                clean = re.sub(r'[*#_]', '', line).strip()
                if not clean: continue
                
                new_p = paragraph.insert_paragraph_before()
                
                # 1. DEFAULT RATA KANAN KIRI (JUSTIFY)
                new_p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                
                # 2. PERKETAT JARAK (SINGLE SPACING & NO SPACE AFTER)
                new_p.paragraph_format.line_spacing = 1.0
                new_p.paragraph_format.space_after = Pt(0)
                new_p.paragraph_format.space_before = Pt(0)
                
                # 3. FITUR GESER KANAN JIKA ADA ***
                if line.startswith("***"):
                    new_p.paragraph_format.left_indent = Inches(0.5)
                    clean = clean.replace("***", "").strip()

                if is_agenda and ":" in clean:
                    # 4. POIN : HARUS SAMA SEJAJAR
                    label, detail = clean.split(":", 1)
                    new_p.paragraph_format.left_indent = Inches(1.0)
                    new_p.paragraph_format.tab_stops.add_tab_stop(Inches(2.5), WD_TAB_ALIGNMENT.LEFT)
                    run = new_p.add_run(f"{label.strip()}\t: {detail.strip()}")
                else:
                    # NARASI DEFAULT
                    run = new_p.add_run(clean)
                
                # 5. HURUF DIPERKECIL (11pt) & TIMES NEW ROMAN
                run.font.name = 'Times New Roman'
                run.font.size = Pt(11)

# --- UI DASHBOARD ---
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
    
    arahan = st.text_area("Instruksi Agenda (Contoh: Rapat tgl 25 jam 8 malam, baju silat):")
    submit = st.form_submit_button("✨ Susun Surat")

if submit:
    with st.spinner("Lagi dibikinin sekretaris pintar..."):
        try:
            model = genai.GenerativeModel(get_active_model())
            prompt = (f"Olah instruksi ini: {arahan}. "
                      "ATURAN: Jangan buat salam/yth. Fokus ke informasi agenda. "
                      "Gunakan '---' untuk memisahkan paragraf pembuka dan list agenda.")
            st.session_state['draf'] = model.generate_content(prompt).text
        except Exception as e:
            st.error(f"Error AI: {e}")

if 'draf' in st.session_state:
    st.subheader("📝 Review Draf")
    draf_final = st.text_area("Edit draf (Gunakan *** untuk geser ke kanan):", value=st.session_state['draf'], height=250)
    
    if st.button("💾 Cetak & Download"):
        try:
            doc = Document("template_psh.docx")
            parts = draf_final.split("---")
            
            # Update Header (Times New Roman 11pt)
            h_map = {
                "{{nomor}}": nomor, "{{hal}}": hal, "{{yth}}": yth,
                "{{lamp}}": lamp, "{{tempat}}": tempat,
                "{{tanggal}}": f"Tegal, {tgl_surat}"
            }
            for p in doc.paragraphs:
                p.paragraph_format.line_spacing = 1.0
                for k, v in h_map.items():
                    if k in p.text:
                        p.text = p.text.replace(k, v)
                        for run in p.runs: 
                            run.font.name = 'Times New Roman'
                            run.font.size = Pt(11)

            # Isi Konten (Ketik ke Tag)
            rakit_isi_surat(doc, "{{pembuka}}", parts[0].strip())
            if len(parts) > 1:
                rakit_isi_surat(doc, "{{agenda}}", parts[1].strip(), is_agenda=True)

            # Hapus tag sisa
            for p in doc.paragraphs:
                if "{{pembuka}}" in p.text or "{{agenda}}" in p.text: p.text = ""

            out = io.BytesIO()
            doc.save(out)
            st.download_button("📥 Download Surat", data=out.getvalue(), file_name="Surat_PSH_Pintar.docx")
        except:
            st.error("Cek tag {{pembuka}} dan {{agenda}} di template lo!")
