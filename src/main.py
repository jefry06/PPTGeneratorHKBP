# from converter.GeneratePPT import convert_with_cover
#
# def main():
#     docx_path = "resources/29Desember.docx"  # Path ke file DOCX
#     pptx_path = "resources/29Desember.pptx"  # Path untuk menyimpan PPTX hasil konversi
#
#     print("Starting DOCX to PPT conversion...")
#     convert_with_cover(docx_path, pptx_path)
#     print(f"Conversion complete! File saved at {pptx_path}")
#
# if __name__ == "__main__":
#     main()


from converter.GeneratePPT import convert_with_cover

def main():
    docx_path = input("Masukkan nama file dan pathnya dalam format DOCX: ")
    if not docx_path.lower().endswith('.docx'):
        print("Error: File harus dalam format DOCX!")
        return

    pptx_path = docx_path.replace('.docx', '.pptx')

    print("Mulai membuat PPT...")
    convert_with_cover(docx_path, pptx_path)
    print(f"PPT Selesai! Silahkan cek di {pptx_path}")
    input("\nTekan Enter untuk keluar...")

if __name__ == "__main__":
    main()