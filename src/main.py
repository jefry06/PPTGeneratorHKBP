from converter.GeneratePPT import convert_with_cover

config = {
    'steps': [
        {'action': 'singing', 'number': 1, 'label': 'votum'},
        {'action': 'session', 'label': 'votum'},
        {'action': 'singing', 'number': 2, 'label': 'p a t i k'},
        {'action': 'patik', 'label': None},
        {'action': 'singing', 'number': 3, 'label': 'manopoti dosa'},
        {'action': 'session', 'label': 'manopoti dosa'},
        {'action': 'singing', 'number': 4, 'label': 'e p i s t e l'},
        {'action': 'session', 'label': 'e p i s t e l'},
        {'action': 'epistel', 'label': None},
        {'action': 'singing', 'number': 5, 'label': 'manghatindanghon haporseaon'},
        {'action': 'session', 'label': 'manghatindanghon haporseaon'},
        {'action': 'session', 'label': 'koor'},
        {'action': 'session', 'label': 'tingting'},
        {'action': 'session', 'label': 'sunggul'},
        {'action': 'singing', 'number': 6, 'label': 'j a m i t a'},
        {'action': 'session', 'label': 'j a m i t a'},
        {'action': 'singing', 'number': 7, 'label': 'tangiang'},
        {'action': 'session', 'label': 'tangiang pelean'},
        {'action': 'session', 'label': 'pangujungi'}
    ]
}

def main():
    docx_path = input("Masukkan nama file dan pathnya dalam format DOCX: ")
    if not docx_path.lower().endswith('.docx'):
        print("Error: File harus dalam format DOCX!")
        return

    pptx_path = docx_path.replace('.docx', '.pptx')

    print("Mulai membuat PPT...")
    convert_with_cover(docx_path, pptx_path, config)
    print(f"PPT Selesai! Silahkan cek di {pptx_path}")
    input("\nTekan Enter untuk keluar...")

if __name__ == "__main__":
    main()