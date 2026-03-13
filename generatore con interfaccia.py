import customtkinter as ctk
import random
import zipfile
import io
import threading
import os
import time
from datetime import timedelta, datetime

# --- LIBRERIE EMAIL CLASSICHE (MIME) ---
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.utils import formatdate
from tkinter import messagebox, filedialog

# --- CONFIGURAZIONE ESTETICA ---
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

# --- DATI MASTER ---
POOL_NOMI = [
    "Mario Rossi", "Luca Bianchi", "Giulia Verdi", "Francesca Neri", "Alessandro Gialli",
    "Elena Esposito", "Roberto Ruggiero", "Sofia Ferrari", "Matteo Ricci", "Chiara Romano",
    "Davide Colombo", "Sara Marino", "Antonio Greco", "Martina Bruno", "Lorenzo Gallo",
    "Valentina Conti", "Simone De Luca", "Giorgia Costa", "Federico Giordano", "Beatrice Rizzo",
    "Giacomo Mancini", "Alessia Lombardi", "Pietro Moretti", "Silvia Barbieri", "Marco Fontana"
]

DOMINI = ["azienda-finta.it", "corporate-demo.com", "training-example.org", "business-lab.net"]

OGGETTI_BASE = [
    "Budget 2025 preliminare", "Report vendite Q3", "Contratto fornitura IT", 
    "Aggiornamento policy Smart Working", "Invito webinar sicurezza", 
    "Problema accesso VPN", "Feedback presentazione cliente", 
    "Organizzazione team building", "Scadenza corsi obbligatori", "Piano ferie natalizie"
]

NOMI_ALLEGATI = [
    ("Report_Dati.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
    ("Contratto_Firmato.pdf", "application/pdf"),
    ("Note_Meeting.txt", "text/plain"),
    ("Slide_Project.pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation"),
    ("Screenshot_Errore.png", "image/png"),
    ("Database_Export.csv", "text/csv")
]

CORPI_INBOX = [
    "Ciao,\n\nCome d'accordo ti giro il file in allegato.\nDacci un'occhiata appena puoi.\n\n{sender}",
    "Buongiorno,\n\nIn allegato la documentazione richiesta.\nAttendo riscontro.\n\nSaluti,\n{sender}",
    "Attenzione:\n\nControlla questo documento urgentemente.\n\n{sender}",
    "Ciao a tutti,\n\nEcco il recap della riunione.\n\n{sender}",
    "Gentile collega,\n\nTi inoltro quanto ricevuto dal cliente.\n\nCordialmente,\n{sender}"
]

CORPI_SENT = [
    "Ciao {recipient},\n\nTi allego la versione corretta.\n\nSaluti,\n{me}",
    "Buongiorno,\n\nHo verificato i dati, tutto ok.\n\n{me}",
    "Ciao,\n\nGrazie per l'aggiornamento. Procedo all'archiviazione.\n\n{me}",
]

class EmailGeneratorApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Outlook Training Generator v9.0 (CRLF FIX)")
        self.geometry("900x800")
        self.save_path = os.getcwd()

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # Main Container
        self.main_frame = ctk.CTkScrollableFrame(self) 
        self.main_frame.grid(row=0, column=0, sticky="nsew", padx=0, pady=0)
        self.main_frame.grid_columnconfigure(0, weight=1)

        # Header
        self.header_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.header_frame.grid(row=0, column=0, pady=(20, 20), sticky="ew")
        
        self.label_title = ctk.CTkLabel(self.header_frame, text="Configurazione Scenario", font=ctk.CTkFont(size=24, weight="bold"))
        self.label_title.pack(side="left", padx=20)

        self.switch_fullscreen = ctk.CTkSwitch(self.header_frame, text="Schermo Intero", command=self.toggle_fullscreen)
        self.switch_fullscreen.pack(side="right", padx=20)

        # User Data
        self.frame_user = ctk.CTkFrame(self.main_frame)
        self.frame_user.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        
        ctk.CTkLabel(self.frame_user, text="Il tuo Nome:", width=100).grid(row=0, column=0, padx=10, pady=10)
        self.entry_name = ctk.CTkEntry(self.frame_user, width=200)
        self.entry_name.grid(row=0, column=1, padx=10, pady=10)
        self.entry_name.insert(0, "Alessandro Corsi")

        ctk.CTkLabel(self.frame_user, text="La tua Email:", width=100).grid(row=0, column=2, padx=10, pady=10)
        self.entry_email = ctk.CTkEntry(self.frame_user, width=200)
        self.entry_email.grid(row=0, column=3, padx=10, pady=10)
        self.entry_email.insert(0, "alessandro.corsi@azienda-finta.it")

        # Params
        self.frame_params = ctk.CTkFrame(self.main_frame)
        self.frame_params.grid(row=2, column=0, padx=20, pady=10, sticky="ew")
        self.frame_params.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(self.frame_params, text="Numero Totale Email:").grid(row=0, column=0, padx=10, pady=(10,0), sticky="w")
        self.slider_emails = ctk.CTkSlider(self.frame_params, from_=10, to=500, number_of_steps=49)
        self.slider_emails.set(50)
        self.slider_emails.grid(row=1, column=0, columnspan=3, padx=10, pady=(0,10), sticky="ew")

        ctk.CTkLabel(self.frame_params, text="Numero Interlocutori:").grid(row=2, column=0, padx=10, pady=(10,0), sticky="w")
        self.slider_colleghi = ctk.CTkSlider(self.frame_params, from_=2, to=20, number_of_steps=18)
        self.slider_colleghi.set(8)
        self.slider_colleghi.grid(row=3, column=0, columnspan=3, padx=10, pady=(0,10), sticky="ew")

        ctk.CTkLabel(self.frame_params, text="Percentuale Allegati (%):").grid(row=4, column=0, padx=10, pady=(10,0), sticky="w")
        self.slider_attach = ctk.CTkSlider(self.frame_params, from_=0, to=100, number_of_steps=100)
        self.slider_attach.set(30)
        self.slider_attach.grid(row=5, column=0, columnspan=3, padx=10, pady=(0,20), sticky="ew")

        # Date
        self.frame_date = ctk.CTkFrame(self.main_frame)
        self.frame_date.grid(row=3, column=0, padx=20, pady=10, sticky="ew")

        ctk.CTkLabel(self.frame_date, text="Data Inizio (GG/MM/AAAA):").pack(side="left", padx=10, pady=10)
        self.entry_start = ctk.CTkEntry(self.frame_date, width=100)
        self.entry_start.insert(0, "01/01/2024")
        self.entry_start.pack(side="left", padx=5)

        ctk.CTkLabel(self.frame_date, text="Data Fine (GG/MM/AAAA):").pack(side="left", padx=10, pady=10)
        self.entry_end = ctk.CTkEntry(self.frame_date, width=100)
        self.entry_end.insert(0, "31/12/2025")
        self.entry_end.pack(side="left", padx=5)

        # Path
        self.frame_path = ctk.CTkFrame(self.main_frame)
        self.frame_path.grid(row=4, column=0, padx=20, pady=10, sticky="ew")
        self.btn_path = ctk.CTkButton(self.frame_path, text="Scegli Cartella...", command=self.choose_directory, fg_color="#3B8ED0")
        self.btn_path.pack(side="left", padx=10, pady=10)
        self.lbl_path_display = ctk.CTkLabel(self.frame_path, text=f"Salva in: {self.save_path}")
        self.lbl_path_display.pack(side="left", padx=10)

        # Button
        self.btn_generate = ctk.CTkButton(self.main_frame, text="GENERA ARCHIVIO .ZIP", height=50, 
                                          font=("Arial", 16, "bold"), fg_color="#2CC985", hover_color="#20A065",
                                          command=self.start_generation_thread)
        self.btn_generate.grid(row=5, column=0, padx=20, pady=20, sticky="ew")
        
        self.progressbar = ctk.CTkProgressBar(self.main_frame)
        self.progressbar.grid(row=6, column=0, padx=20, pady=(0, 10), sticky="ew")
        self.progressbar.set(0)
        self.lbl_status = ctk.CTkLabel(self.main_frame, text="Pronto.")
        self.lbl_status.grid(row=7, column=0, pady=(0, 20))

    def toggle_fullscreen(self):
        state = self.switch_fullscreen.get()
        self.attributes("-fullscreen", bool(state))
        if state: self.bind("<Escape>", lambda e: self.exit_fullscreen())

    def exit_fullscreen(self):
        self.switch_fullscreen.deselect()
        self.attributes("-fullscreen", False)
        self.unbind("<Escape>")

    def choose_directory(self):
        directory = filedialog.askdirectory()
        if directory:
            self.save_path = directory
            self.lbl_path_display.configure(text=f"Salva in: {directory}")

    def start_generation_thread(self):
        self.btn_generate.configure(state="disabled", text="Generazione in corso...")
        self.progressbar.set(0)
        threading.Thread(target=self.run_logic, daemon=True).start()

    def run_logic(self):
        try:
            my_name = self.entry_name.get()
            my_email = self.entry_email.get()
            total_emails = int(self.slider_emails.get())
            num_interlocutori = int(self.slider_colleghi.get())
            attach_prob = int(self.slider_attach.get()) / 100.0
            
            try:
                d_start = datetime.strptime(self.entry_start.get(), "%d/%m/%Y").date()
                d_end = datetime.strptime(self.entry_end.get(), "%d/%m/%Y").date()
            except ValueError:
                raise ValueError("Formato data errato. Usa GG/MM/AAAA")

            pool_temp = list(POOL_NOMI)
            random.shuffle(pool_temp)
            selected_names = pool_temp[:num_interlocutori]
            
            colleghi = []
            for name in selected_names:
                email = f"{name.lower().replace(' ', '.')}@{random.choice(DOMINI)}"
                colleghi.append((name, email))

            zip_buffer = io.BytesIO()
            zip_filename = f"Archivio_Training_{datetime.now().strftime('%Y%m%d_%H%M')}.zip"

            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                for i in range(total_emails):
                    progress = (i + 1) / total_emails
                    self.progressbar.set(progress)
                    self.lbl_status.configure(text=f"Generazione mail {i+1}/{total_emails}...")
                    
                    f_type = 'inbox' if i < (total_emails * 0.7) else 'sent'
                    fname, msg_bytes = self.generate_single_email(i, f_type, my_name, my_email, colleghi, d_start, d_end, attach_prob)
                    
                    # Salviamo i bytes diretti
                    zip_file.writestr(fname, msg_bytes)

            full_save_path = os.path.join(self.save_path, zip_filename)
            with open(full_save_path, "wb") as f:
                f.write(zip_buffer.getvalue())

            self.lbl_status.configure(text=f"Completato! File salvato.")
            messagebox.showinfo("Successo", f"Archivio creato con successo in:\n{full_save_path}")

        except Exception as e:
            messagebox.showerror("Errore", str(e))
            self.lbl_status.configure(text="Errore.")
            print(e)
        finally:
            self.btn_generate.configure(state="normal", text="GENERA ARCHIVIO .ZIP")

    def generate_single_email(self, index, folder_type, my_name, my_email, colleghi, d_start, d_end, attach_prob):
        collega_nome, collega_email = random.choice(colleghi)
        
        delta = d_end - d_start
        int_delta = (delta.days * 24 * 60 * 60) + 86399
        random_second = random.randrange(max(1, int_delta))
        fake_date = datetime.combine(d_start, datetime.min.time()) + timedelta(seconds=random_second)
        
        # 1. Creiamo un contenitore 'mixed' (che può contenere allegati e testo)
        msg = MIMEMultipart('mixed')
        
        # Header essenziali
        msg['Date'] = formatdate(timeval=fake_date.timestamp(), localtime=True)
        
        if folder_type == 'inbox':
            sender_str = f"{collega_nome} <{collega_email}>"
            recipient_str = f"{my_name} <{my_email}>"
            body_tmpl = random.choice(CORPI_INBOX)
            folder = "Posta in Arrivo"
            subject = random.choice(OGGETTI_BASE)
            real_sender = collega_nome.split()[0]
            real_recip = my_name.split()[0]
        else:
            sender_str = f"{my_name} <{my_email}>"
            recipient_str = f"{collega_nome} <{collega_email}>"
            body_tmpl = random.choice(CORPI_SENT)
            folder = "Posta Inviata"
            subject = "RE: " + random.choice(OGGETTI_BASE)
            real_sender = my_name.split()[0]
            real_recip = collega_nome.split()[0]

        msg['Subject'] = subject
        msg['From'] = sender_str
        msg['To'] = recipient_str
        msg['Message-ID'] = f"<{random.randint(100000,999999)}.{index}.{fake_date.strftime('%Y%m%d')}@mailserver.fake>"
        msg['MIME-Version'] = "1.0"
        
        # Header Magici per Outlook
        msg['X-Unsent'] = '0' 
        msg['X-Mailer'] = 'Microsoft Outlook 16.0'
        msg['Content-Class'] = 'urn:content-classes:message'
        msg['X-MimeOLE'] = 'Produced By Microsoft MimeOLE V6.00.2900.2180'

        # Corpo del testo
        body_text = body_tmpl.format(sender=real_sender, recipient=real_recip, me=my_name.split()[0])
        body_html = f"""<html>
<body>
<p style="font-family: Calibri, sans-serif;">{body_text.replace(chr(10), '<br>')}</p>
</body>
</html>"""

        # 2. Creiamo la parte 'alternative' per Testo vs HTML
        msg_alternative = MIMEMultipart('alternative')
        msg_alternative.attach(MIMEText(body_text, "plain", "utf-8"))
        msg_alternative.attach(MIMEText(body_html, "html", "utf-8"))
        
        # Alleghiamo la parte testo al contenitore principale
        msg.attach(msg_alternative)

        # 3. Gestione Allegati
        if random.random() < attach_prob and folder_type == 'inbox':
            fname_att, mime_type = random.choice(NOMI_ALLEGATI)
            fname_final = f"{index}_{fname_att}"
            
            if hasattr(random, 'randbytes'):
                file_data = b'%FAKE-BINARY' + random.randbytes(2000)
            else:
                file_data = b'%FAKE-BINARY' + os.urandom(2000)

            maintype, subtype = mime_type.split('/', 1)
            
            part = MIMEBase(maintype, subtype)
            part.set_payload(file_data)
            encoders.encode_base64(part)
            
            part.add_header('Content-Disposition', f'attachment; filename="{fname_final}"')
            msg.attach(part)

        safe_sub = "".join([c if c.isalnum() else "_" for c in subject])[:30]
        filename = f"{folder}/{fake_date.strftime('%Y-%m-%d')}_{safe_sub}_{index}.eml"

        # --- FIX SUPREMO: FORZATURA CRLF ---
        # Otteniamo la stringa intera con i ritorni a capo standard di Python (\n)
        raw_string = msg.as_string()
        
        # Forziamo la sostituzione con \r\n (Carriage Return + Line Feed)
        # Questo è l'unico modo per cui Outlook riconosce gli header correttamente
        if '\r\n' not in raw_string:
            raw_string = raw_string.replace('\n', '\r\n')
            
        return filename, raw_string.encode('utf-8')

if __name__ == "__main__":
    app = EmailGeneratorApp()
    app.mainloop()