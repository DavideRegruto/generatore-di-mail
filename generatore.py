import customtkinter as ctk
import random
import zipfile
import io
import threading
import os
from datetime import timedelta, datetime
from email.message import EmailMessage
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

# --- LOGICA DI GENERAZIONE ---
class EmailGeneratorApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Configurazione Finestra
        self.title("Outlook Training Generator v3.0")
        self.geometry("900x800")
        
        # Variabile per il percorso di salvataggio (Default: cartella corrente)
        self.save_path = os.getcwd()

        # Layout a griglia
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # --- CONTAINER PRINCIPALE ---
        self.main_frame = ctk.CTkScrollableFrame(self) 
        self.main_frame.grid(row=0, column=0, sticky="nsew", padx=0, pady=0)
        self.main_frame.grid_columnconfigure(0, weight=1)

        # --- HEADER (TITOLO + FULLSCREEN) ---
        self.header_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.header_frame.grid(row=0, column=0, pady=(20, 20), sticky="ew")
        
        self.label_title = ctk.CTkLabel(self.header_frame, text="Configurazione Scenario", font=ctk.CTkFont(size=24, weight="bold"))
        self.label_title.pack(side="left", padx=20)

        # Switch Schermo Intero
        self.switch_fullscreen = ctk.CTkSwitch(self.header_frame, text="Schermo Intero", command=self.toggle_fullscreen)
        self.switch_fullscreen.pack(side="right", padx=20)

        # --- SEZIONE 1: I TUOI DATI ---
        self.frame_user = ctk.CTkFrame(self.main_frame)
        self.frame_user.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        
        ctk.CTkLabel(self.frame_user, text="Il tuo Nome:", width=100).grid(row=0, column=0, padx=10, pady=10)
        self.entry_name = ctk.CTkEntry(self.frame_user, placeholder_text="Es. Mario Rossi", width=200)
        self.entry_name.grid(row=0, column=1, padx=10, pady=10)
        self.entry_name.insert(0, "Alessandro Corsi")

        ctk.CTkLabel(self.frame_user, text="La tua Email:", width=100).grid(row=0, column=2, padx=10, pady=10)
        self.entry_email = ctk.CTkEntry(self.frame_user, placeholder_text="mario@azienda.it", width=200)
        self.entry_email.grid(row=0, column=3, padx=10, pady=10)
        self.entry_email.insert(0, "alessandro.corsi@azienda-finta.it")

        # --- SEZIONE 2: PARAMETRI ---
        self.frame_params = ctk.CTkFrame(self.main_frame)
        self.frame_params.grid(row=2, column=0, padx=20, pady=10, sticky="ew")
        self.frame_params.grid_columnconfigure(1, weight=1) # Per centrare meglio

        # Slider Numero Mail
        ctk.CTkLabel(self.frame_params, text="Numero Totale Email:").grid(row=0, column=0, padx=10, pady=(10,0), sticky="w")
        self.lbl_num_mail_val = ctk.CTkLabel(self.frame_params, text="100", font=("Arial", 12, "bold"))
        self.lbl_num_mail_val.grid(row=0, column=2, padx=10, pady=(10,0))
        
        self.slider_emails = ctk.CTkSlider(self.frame_params, from_=10, to=500, number_of_steps=49, command=self.update_email_label)
        self.slider_emails.set(100)
        self.slider_emails.grid(row=1, column=0, columnspan=3, padx=10, pady=(0,10), sticky="ew")

        # Slider Interlocutori
        ctk.CTkLabel(self.frame_params, text="Numero Interlocutori (Colleghi):").grid(row=2, column=0, padx=10, pady=(10,0), sticky="w")
        self.lbl_colleghi_val = ctk.CTkLabel(self.frame_params, text="8", font=("Arial", 12, "bold"))
        self.lbl_colleghi_val.grid(row=2, column=2, padx=10, pady=(10,0))

        self.slider_colleghi = ctk.CTkSlider(self.frame_params, from_=2, to=20, number_of_steps=18, command=self.update_colleghi_label)
        self.slider_colleghi.set(8)
        self.slider_colleghi.grid(row=3, column=0, columnspan=3, padx=10, pady=(0,10), sticky="ew")

        # Slider Allegati
        ctk.CTkLabel(self.frame_params, text="Percentuale Allegati (%):").grid(row=4, column=0, padx=10, pady=(10,0), sticky="w")
        self.lbl_allegati_val = ctk.CTkLabel(self.frame_params, text="40%", font=("Arial", 12, "bold"))
        self.lbl_allegati_val.grid(row=4, column=2, padx=10, pady=(10,0))

        self.slider_attach = ctk.CTkSlider(self.frame_params, from_=0, to=100, number_of_steps=100, command=self.update_attach_label)
        self.slider_attach.set(40)
        self.slider_attach.grid(row=5, column=0, columnspan=3, padx=10, pady=(0,20), sticky="ew")

        # --- SEZIONE 3: PERIODO ---
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

        # --- SEZIONE 4: PERCORSO SALVATAGGIO ---
        self.frame_path = ctk.CTkFrame(self.main_frame)
        self.frame_path.grid(row=4, column=0, padx=20, pady=10, sticky="ew")

        self.btn_path = ctk.CTkButton(self.frame_path, text="Scegli Cartella...", command=self.choose_directory, fg_color="#3B8ED0")
        self.btn_path.pack(side="left", padx=10, pady=10)

        self.lbl_path_display = ctk.CTkLabel(self.frame_path, text=f"Salva in: {self.save_path}", text_color="gray")
        self.lbl_path_display.pack(side="left", padx=10)

        # --- BOTTONE E PROGRESSO ---
        self.btn_generate = ctk.CTkButton(self.main_frame, text="GENERA ARCHIVIO .ZIP", height=50, 
                                          font=("Arial", 16, "bold"), fg_color="#2CC985", hover_color="#20A065",
                                          command=self.start_generation_thread)
        self.btn_generate.grid(row=5, column=0, padx=20, pady=20, sticky="ew")

        self.progressbar = ctk.CTkProgressBar(self.main_frame)
        self.progressbar.grid(row=6, column=0, padx=20, pady=(0, 10), sticky="ew")
        self.progressbar.set(0)

        self.lbl_status = ctk.CTkLabel(self.main_frame, text="Pronto.")
        self.lbl_status.grid(row=7, column=0, pady=(0, 20))

    # --- METODI GUI ---
    def toggle_fullscreen(self):
        state = self.switch_fullscreen.get()
        self.attributes("-fullscreen", bool(state))
        if state:
            self.bind("<Escape>", lambda e: self.exit_fullscreen())

    def exit_fullscreen(self):
        self.switch_fullscreen.deselect()
        self.attributes("-fullscreen", False)
        self.unbind("<Escape>")

    def choose_directory(self):
        directory = filedialog.askdirectory()
        if directory:
            self.save_path = directory
            display_text = directory if len(directory) < 50 else "..."+directory[-45:]
            self.lbl_path_display.configure(text=f"Salva in: {display_text}")

    def update_email_label(self, value):
        self.lbl_num_mail_val.configure(text=f"{int(value)}")
    
    def update_colleghi_label(self, value):
        self.lbl_colleghi_val.configure(text=f"{int(value)}")

    def update_attach_label(self, value):
        self.lbl_allegati_val.configure(text=f"{int(value)}%")

    def start_generation_thread(self):
        self.btn_generate.configure(state="disabled", text="Generazione in corso...")
        self.progressbar.set(0)
        threading.Thread(target=self.run_logic, daemon=True).start()

    # --- LOGICA CORE ---
    def run_logic(self):
        try:
            # 1. Recupero Dati
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

            # Creazione lista colleghi dinamica
            # Clona la lista per non modificare quella globale
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
                    # Update Progress
                    progress = (i + 1) / total_emails
                    self.progressbar.set(progress)
                    self.lbl_status.configure(text=f"Generazione mail {i+1}/{total_emails}...")
                    
                    # Genera singola mail
                    f_type = 'inbox' if i < (total_emails * 0.7) else 'sent'
                    fname, msg = self.generate_single_email(i, f_type, my_name, my_email, colleghi, d_start, d_end, attach_prob)
                    
                    # Writestr accetta lo slash come separatore di cartelle universale nello zip
                    zip_file.writestr(fname, msg.as_bytes())

            # Salvataggio nel percorso scelto
            full_save_path = os.path.join(self.save_path, zip_filename)

            with open(full_save_path, "wb") as f:
                f.write(zip_buffer.getvalue())

            self.lbl_status.configure(text=f"Completato! File salvato.")
            messagebox.showinfo("Successo", f"Archivio creato con successo in:\n{full_save_path}")

        except Exception as e:
            messagebox.showerror("Errore", str(e))
            self.lbl_status.configure(text="Errore.")
        
        finally:
            self.btn_generate.configure(state="normal", text="GENERA ARCHIVIO .ZIP")

    def generate_single_email(self, index, folder_type, my_name, my_email, colleghi, d_start, d_end, attach_prob):
        collega_nome, collega_email = random.choice(colleghi)
        
        # Data casuale
        delta = d_end - d_start
        int_delta = (delta.days * 24 * 60 * 60) + 86399
        random_second = random.randrange(max(1, int_delta)) # Protezione contro date uguali
        fake_date = datetime.combine(d_start, datetime.min.time()) + timedelta(seconds=random_second)
        # Formato data standard email
        date_str = fake_date.strftime("%a, %d %b %Y %H:%M:%S +0100")

        msg = EmailMessage()
        msg['Date'] = date_str

        # Setup Mittente/Destinatario
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

        # Corpo
        body_text = body_tmpl.format(sender=real_sender, recipient=real_recip, me=my_name.split()[0])

        # Allegati (Logica)
        has_attachment = random.random() < attach_prob
        if has_attachment and folder_type == 'inbox':
            body_text += "\n\n(Vedi allegato)"

        msg.set_content(body_text)
        msg['Subject'] = subject
        msg['From'] = sender_str
        msg['To'] = recipient_str
        msg['Message-ID'] = f"<{random.randint(10000,99999)}.{index}.{fake_date.strftime('%Y%m%d')}@mailserver.local>"

        # Priorità
        if random.random() > 0.9:
            msg['Importance'] = 'high'
            msg['X-Priority'] = '1'

        # Inserimento Allegato Fisico (Simulato)
        if has_attachment:
            fname_att, mime_type = random.choice(NOMI_ALLEGATI)
            fname_final = f"{index}_{fname_att}"
            
            if fname_att.endswith(".txt") or fname_att.endswith(".csv"):
                file_data = b"Dati di prova leggibili per il training."
            else:
                # Genera byte casuali (fake binary)
                # Fallback per versioni python vecchie
                if hasattr(random, 'randbytes'):
                    file_data = b'%FAKE-BINARY' + random.randbytes(2000)
                else:
                    file_data = b'%FAKE-BINARY' + os.urandom(2000)

            maintype, subtype = mime_type.split('/', 1)
            msg.add_attachment(file_data, maintype=maintype, subtype=subtype, filename=fname_final)

        # Nome file eml nel ZIP (Folder/Data_Soggetto.eml)
        safe_sub = "".join([c if c.isalnum() else "_" for c in subject])[:30]
        filename = f"{folder}/{fake_date.strftime('%Y-%m-%d')}_{safe_sub}_{index}.eml"

        return filename, msg

if __name__ == "__main__":
    app = EmailGeneratorApp()
    app.mainloop()
