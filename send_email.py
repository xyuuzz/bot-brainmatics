import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import fitz  # PyMuPDF


def replace_text(page, search_text, replacement_text):
        text_instances = page.search_for(search_text)
        for inst in text_instances:
            area = fitz.Rect(inst[0], inst[1], inst[2], inst[3])
            page.add_redact_annot(area)

        # Terapkan redaksi
        page.apply_redactions()

        # Tambahkan teks pengganti
        for inst in text_instances:
            area = fitz.Rect(inst[0], inst[1], inst[2], inst[3])
            page.insert_text(area.tl, replacement_text, fontfile='calibri-regular.ttf', fontname='Calibri', fontsize=11, color=(0, 0, 0))

class SendEmail:
    email = None
    password = None
    smtp_server = None
    smtp_port = None

    def __init__(self, email, password):
        self.email = email
        self.password = password

    def setSmtpSettings(self, smtp_server, smtp_port):
        self.smtp_server = smtp_server
        self.smtp_port = smtp_port

    def getAttachmentPath(self, namaTraining):
        # Buka dokumen PDF
        pdf_document = fitz.open("./surat_konfirmasi.pdf")

        # Simpan salinan dari dokumen asli sebelum melakukan perubahan
        namaTrainingFormatted = namaTraining.replace(" ", "-").lower()
        pathFileBaru = f"bm-surat-konfirmasi-{namaTrainingFormatted}.pdf"
        pdf_document.save(pathFileBaru)

        # Tutup dokumen asli dan buka salinan untuk melakukan perubahan
        pdf_document.close()

        return pathFileBaru


    def getBodyEmail(self, namaAsisten, noHpAsisten, namaPeserta, namaTraining, tanggalTraining, waktuTraining, lokasiTraining, lokasiLinkGmaps, ruanganTraining, username, password):
        return f'''
            <p style="font-family: Helvetica, sans-serif; font-size: 9pt;">Yth Bapak/Ibu {namaPeserta},</p>
            
            <p style="font-family: Helvetica, sans-serif; font-size: 9pt;">Bersama email berikut kami kirimkan <strong>Konfirmasi Pelaksanaan {namaTraining}</strong> sebagai konfirmasi pelaksanaan training yang akan diselenggarakan pada:</p>
            
            <p style="font-family: Helvetica, sans-serif; font-size: 9pt;"><strong>Tanggal</strong>: {tanggalTraining}<br>
            <strong>Waktu</strong>: {waktuTraining}<br>
            <strong>Lokasi</strong>: <a href="{lokasiLinkGmaps}" target="_blank">{lokasiTraining}</a><br>
            <strong>Ruangan</strong>: {ruanganTraining}</p>
            
            <p style="font-family: Helvetica, sans-serif; font-size: 9pt;">Untuk mengakses materi, riwayat pelaksanaan training, dan sertifikat keikutsertaan training bisa diakses melalui Brainmatics Learning Management System (LMS), dengan cara sebagai berikut:</p>
            
            <ol style="font-family: Helvetica, sans-serif; font-size: 9pt;">
            <li>Silahkan mengakses situs <a href="https://brainmatics.com/" target="_blank">https://brainmatics.com/</a></li>
            <li>Pilih link login yang terdapat di pojok kiri atas</li>
            <li>Silahkan masukan username anda yaitu: <strong>{username}</strong></li>
            <li>Silahkan masukan password anda yaitu: <strong>{password}</strong></li>
            <li>Silahkan klik link View yang terdapat pada kolom Syllabus & Material, untuk mengakses materi training</li>
            <li>Silahkan klik link View yang terdapat pada kolom Certificate, untuk mengakses sertifikat keikutsertaan anda</li>
            </ol>
            
            <p style="font-family: Helvetica, sans-serif; font-size: 9pt;">Jika membutuhkan panduan penggunaan Brainmatics LMS dapat didownload melalui link berikut: <a href="https://1drv.ms/b/s!AiwttwOoOn4bipMM5RKOgVcl_nxQ6g?e=IcFjRg/" target="_blank">Panduan Penggunaan Brainmatics LMS</a>.</p>
            
            <p style="font-family: Helvetica, sans-serif; font-size: 9pt;">Semoga pelaksanaan training ini dapat berjalan dengan baik dan mendapatkan hasil yang memuaskan. Demikian yang dapat kami sampaikan, apabila ada informasi yang belum jelas, silahkan menghubungi kami kembali.</p>
            
            <p style="font-family: Helvetica, sans-serif; font-size: 9pt;">Salam,</p>
            
            <p style="font-family: Helvetica, sans-serif; font-size: 9pt;">--<br>
            {namaAsisten}<br>
            Training Staff<br>
            PT Brainmatics Indonesia Cendekia<br>
            Ph. +62 {noHpAsisten} | Phone/WA/Telegram +62 {noHpAsisten}<br>
            <a href="http://www.brainmatics.com" target="_blank">www.brainmatics.com</a></p>
        '''

    def send(self, subject, body, to_email, cc_emails, namaAsisten, attachment_path):
        # Membuat pesan MIMEMultipart
        msg = MIMEMultipart()
        msg['From'] = namaAsisten
        msg['To'] = to_email
        msg['CC'] = ", ".join(cc_emails)
        msg['Subject'] = subject

        # Menambahkan body ke email
        msg.attach(MIMEText(body, 'html'))

        # Menyiapkan attachment
        with open(attachment_path, "rb") as attachment_file:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment_file.read())
            encoders.encode_base64(part)
            part.add_header(
                "Content-Disposition",
                f"attachment; filename= {attachment_path}",
            )
            msg.attach(part)

        # Membuat sesi SMTP dan mengirim email
        try:
            with smtplib.SMTP(self.smtp_server, self.smtp_port) as server:
                server.starttls()  # Mengaktifkan mode TLS
                server.login(self.email, self.password)
                server.send_message(msg)
                print("Email berhasil dikirim!")
        except Exception as e:
            print(f"Error: {e}")
