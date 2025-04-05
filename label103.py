import streamlit as st
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL

# Fungsi untuk membuat file label undangan
def create_label_docx(nama_file_txt, template, nama_output="label_undangan.docx"):
    with open(nama_file_txt, encoding='utf-8') as f:
        daftar_nama = [line.strip() for line in f if line.strip()]

    doc = Document()
    section = doc.sections[0]
    section.top_margin = Cm(0)
    section.bottom_margin = Cm(0)
    section.left_margin = Cm(0)
    section.right_margin = Cm(0)

    label_per_halaman = 12
    jumlah_halaman = (len(daftar_nama) + label_per_halaman - 1) // label_per_halaman

    index = 0
    for _ in range(jumlah_halaman):
        table = doc.add_table(rows=4, cols=3)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False

        for row in table.rows:
            row.height = Cm(3.2)
            row.height_rule = True

        for r in range(4):
            for c in range(3):
                if index >= len(daftar_nama):
                    continue

                cell = table.cell(r, c)
                cell.width = Cm(6.4)
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                p.paragraph_format.left_indent = Cm(1.2)
                p.paragraph_format.line_spacing = Pt(12)

                # Atur margin atas berdasarkan baris
                if r == 0:
                    space_top = Pt(40)
                elif r == 1:
                    space_top = Pt(65)
                elif r == 2:
                    space_top = Pt(70)
                elif r == 3:
                    space_top = Pt(60)  # Naikkan dari sebelumnya

                p.paragraph_format.space_before = space_top

                # Gunakan template yang dimasukkan pengguna
                run = p.add_run(template.format(nama=daftar_nama[index]))
                run.font.name = "Calibri"
                run.font.size = Pt(11)

                index += 1

        if index < len(daftar_nama):
            doc.add_page_break()

    doc.save(nama_output)
    print(f"✅ File '{nama_output}' berhasil dibuat!")

# Streamlit interface
st.title("Generator Label Undangan")

# Input template kata-kata
template = st.text_input(
    "Masukkan template kata-kata (gunakan {nama} untuk nama):",
    "Kepada Yth,\n{nama}\nDi Tempat."
)

# Input file daftar nama
uploaded_file = st.file_uploader("Upload file daftar nama (.txt)", type="txt")

# Menangani tombol generate
if uploaded_file is not None:
    with open("daftar_nama.txt", "wb") as f:
        f.write(uploaded_file.getbuffer())
    
    if st.button("Generate Label"):
        create_label_docx("daftar_nama.txt", template)
        st.success("✅ Label berhasil dibuat!")
