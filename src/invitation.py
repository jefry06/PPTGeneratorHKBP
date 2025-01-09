import json

def create_invitations(data):
    template = u"""*Syalom {Salutation} {Nama},*

Selamat Tahun Baru 2025! Semoga {Salutation} dan keluarga selalu dalam keadaan sehat.  
Kami ingin memberitahukan bahwa hasil rapat ruas sektor 7 merekomendasikan {Salutation} sebagai *{seksi}* dalam kepanitiaan Pesta Gotilon 2025 HKBP Depok II.

Mohon konfirmasi apakah {Salutation} bersedia menerima rekomendasi tersebut.

Kami juga mengundang {Salutation} untuk hadir pada *Rapat Perdana Calon Panitia* yang akan diadakan:  
📅 *Rabu, 8 Januari 2025*  
🕒 *19.30 WIB*  
📍 *via Zoom (online) https://bit.ly/gotilon-8-januari*

Kehadiran {Salutation} sangat diharapkan untuk membahas rencana program Pesta Gotilon 2025.

Pada *Jumat, 10 Januari 2025, pukul 19.30 di Gereja HKBP Depok II*, kami mengundang kembali {Salutation} untuk acara Ramah Tamah Panitia dan finalisasi materi Pesta Gotilon yang akan dibawa ke *Rapat Huria*. Mohon konfirmasi apakah {Salutation} bisa hadir.

Terima kasih atas perhatian {Salutation}.

*Horas,*  
Jefry Pasaribu  
(Calon Panitia)"""

    for person in data:
        content = template.format(Salutation=person['salutation'], Nama=person['nama'], seksi=person['seksi'])
        hp_cleaned = person['hp'].replace('+62', '')
        filename = f"f:/HKBP/Gotilon 2025/invitation/result/{hp_cleaned}.txt"
        with open(filename, 'w', encoding='utf-8') as file:
            file.write(content)

# Example usage
with open('f:/HKBP/Gotilon 2025/invitation/list-piloting.json', 'r', encoding='utf-8-sig') as json_file:  # Handle BOM by using utf-8-sig
    data = json.load(json_file)

create_invitations(data)
