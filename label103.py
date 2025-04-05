import streamlit as st
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
import os

def create_label_docx(daftar_nama, template_awal, nama_output="label_undangan.docx"):
    # Membuat dokumen baru
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Cm(0)
    section.bottom_margin = Cm(0)
    section.left_margin = Cm(0)
    section.right_margin = Cm(0)

    # Mengatur jumlah label per halaman (12 label per halaman)
    label_per_halaman = 12
    jumlah_halaman = (len(daftar_nama) + label_per_halaman - 1) // label_per_halaman

    index = 0
    for _ in range(jumlah_halaman):
        # Membuat tabel 4x3 untuk 12 label
        table = doc.add_table(rows=4, cols=3)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False

        # Menyesuaikan tinggi setiap baris pada tabel
        for row in table.rows:
            row.height = Cm(3.2)
            row.height_rule = True

        # Membuat label pada setiap cell tabel
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

                if r == 0:
                    space_top = Pt(40)
                elif r == 1:
                    space_top = Pt(65)
                elif r == 2:
                    space_top = Pt(70)
                elif r == 3:
                    space_top = Pt(60)

                p.paragraph_format.space_before = space_top

                # Menambahkan nama dengan template pesan yang sudah ditentukan oleh pengguna
                run = p.add_run(f"{template_awal[0]}\n{daftar_nama[index]}\n{template_awal[1]}")
                run.font.name = "Calibri"
                run.font.size = Pt(11)

                index += 1

        # Jika masih ada nama, tambahkan halaman baru
        if index < len(daftar_nama):
            doc.add_page_break()

    # Simpan hasil file
    doc.save(nama_output)
    return nama_output

# Streamlit Web App
st.title("Generator Label Undangan")
st.write("Pilih opsi di bawah untuk membuat label undangan:")

# Menambahkan input untuk template yang dapat diedit oleh pengguna
st.markdown("""
**Catatan:** Anda dapat mengedit bagian template, seperti "Kepada Yth," dan "Di Tempat."
Bagian yang dapat diedit adalah bagian pertama dan terakhir dalam format template.
""")

# Input untuk template kata-kata
template_awal = [
    st.text_input("Template Awal (misalnya: 'Kepada Yth,')", "Kepada Yth,"),
    st.text_input("Template Akhir (misalnya: 'Di Tempat.')", "Di Tempat.")
]

# Pilih opsi input manual atau upload file
input_option = st.radio("Pilih cara input daftar nama", ("Input Manual", "Upload File .txt"))

if input_option == "Input Manual":
    # Input manual daftar nama
    daftar_nama_input = st.text_area("Masukkan daftar nama (pisahkan setiap nama dengan baris baru)")
    if st.button("Generate Label") and daftar_nama_input:
        daftar_nama = daftar_nama_input.splitlines()
        output_file = "label_undangan.docx"
        result = create_label_docx(daftar_nama, template_awal, output_file)
        st.success(f"✅ Label berhasil dibuat! File disimpan di: {output_file}")

        # Offer untuk download file
        with open(output_file, "rb") as f:
            st.download_button("Download Label", f, file_name=output_file, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

elif input_option == "Upload File .txt":
    # Upload file .txt
    uploaded_file = st.file_uploader("Pilih file daftar nama", type="txt")

    if uploaded_file is not None:
        # Simpan file sementara
        temp_filename = "uploaded_daftar_nama.txt"
        with open(temp_filename, "wb") as f:
            f.write(uploaded_file.getbuffer())

        st.write(f"File {uploaded_file.name} berhasil diupload!")

        # Button untuk generate label
        if st.button("Generate Label"):
            with open(temp_filename, encoding='utf-8') as f:
                daftar_nama = [line.strip() for line in f if line.strip()]
            output_file = "label_undangan.docx"
            result = create_label_docx(daftar_nama, template_awal, output_file)
            st.success(f"✅ Label berhasil dibuat! File disimpan di: {output_file}")

            # Offer untuk download file
            with open(output_file, "rb") as f:
                st.download_button("Download Label", f, file_name=output_file, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

        # Hapus file sementara setelah proses selesai
        os.remove(temp_filename)
