# -*- coding: utf-8 -*-
"""ArubaMailer - invio programmato a blocchi di email da lista Excel tramite casella Aruba."""
import datetime as dt
import json
import os
import queue
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import mailer_engine as me

APP = "ArubaMailer"
CONFIG = os.path.join(me.app_dir(), "config.json")


def load_config():
    try:
        with open(CONFIG, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_config(cfg):
    try:
        with open(CONFIG, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2, ensure_ascii=False)
    except Exception as e:
        messagebox.showwarning(APP, "Impossibile salvare le impostazioni: %s" % e)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("ArubaMailer - invio a blocchi da casella Aruba")
        self.geometry("900x680")
        self.minsize(820, 600)
        self.q = queue.Queue()
        self.job = None
        self.campaign = None
        self.excel_path = ""
        self.cfg = load_config()
        try:
            ttk.Style().theme_use("vista")
        except Exception:
            pass

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=8, pady=8)
        self.tab_account = ttk.Frame(nb, padding=12)
        self.tab_list = ttk.Frame(nb, padding=12)
        self.tab_msg = ttk.Frame(nb, padding=12)
        self.tab_send = ttk.Frame(nb, padding=12)
        nb.add(self.tab_account, text=" 1. Account Aruba ")
        nb.add(self.tab_list, text=" 2. Lista Excel ")
        nb.add(self.tab_msg, text=" 3. Messaggio ")
        nb.add(self.tab_send, text=" 4. Programmazione e invio ")
        self.build_account()
        self.build_list()
        self.build_msg()
        self.build_send()
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.after(200, self.poll)

    # ------------------------------------------------------------ helpers
    def row(self, parent, r, label, var, width=40, show=None, hint=""):
        ttk.Label(parent, text=label).grid(row=r, column=0, sticky="w", pady=4)
        e = ttk.Entry(parent, textvariable=var, width=width, show=show)
        e.grid(row=r, column=1, sticky="w", pady=4)
        if hint:
            ttk.Label(parent, text=hint, foreground="#666").grid(row=r, column=2, sticky="w", padx=8)
        return e

    def account(self):
        return me.Account(self.v_email.get(), self.v_pass.get(), self.v_name.get(), self.v_reply.get(),
                          self.v_smtp.get(), self.v_smtp_port.get(), self.v_imap.get(), self.v_imap_port.get(),
                          use_ssl=True)

    def store_cfg(self):
        c = self.cfg
        c.update(email=self.v_email.get(), name=self.v_name.get(), reply=self.v_reply.get(),
                 smtp=self.v_smtp.get(), smtp_port=self.v_smtp_port.get(),
                 imap=self.v_imap.get(), imap_port=self.v_imap_port.get(),
                 remember=self.v_remember.get(),
                 batch=self.v_batch.get(), interval=self.v_interval.get(), period=self.v_period.get(),
                 cap=self.v_cap.get(), win_on=self.v_win_on.get(), win_from=self.v_win_from.get(),
                 win_to=self.v_win_to.get(), column=self.v_col.get(),
                 subject=self.v_subject.get(), body=self.txt_body.get("1.0", "end-1c"))
        c["pw"] = me.protect_password(self.v_pass.get()) if self.v_remember.get() else ""
        save_config(c)

    # ------------------------------------------------------------ tab 1
    def build_account(self):
        f = self.tab_account
        c = self.cfg
        self.v_name = tk.StringVar(value=c.get("name", ""))
        self.v_email = tk.StringVar(value=c.get("email", ""))
        self.v_pass = tk.StringVar(value=me.unprotect_password(c.get("pw", "")))
        self.v_remember = tk.BooleanVar(value=c.get("remember", False))
        self.v_reply = tk.StringVar(value=c.get("reply", ""))
        self.v_smtp = tk.StringVar(value=c.get("smtp", "smtps.aruba.it"))
        self.v_smtp_port = tk.StringVar(value=c.get("smtp_port", "465"))
        self.v_imap = tk.StringVar(value=c.get("imap", "imaps.aruba.it"))
        self.v_imap_port = tk.StringVar(value=c.get("imap_port", "993"))

        self.row(f, 0, "Nome mittente", self.v_name, hint="es. Mario Rossi (facoltativo)")
        self.row(f, 1, "Email mittente Aruba", self.v_email, hint="anche nome utente di accesso")
        self.row(f, 2, "Password", self.v_pass, show="•")
        ttk.Checkbutton(f, text="Ricorda password (cifrata con il tuo utente Windows)",
                        variable=self.v_remember).grid(row=3, column=1, sticky="w")
        self.row(f, 4, "Rispondi a (Reply-To)", self.v_reply, hint="facoltativo")
        ttk.Separator(f).grid(row=5, column=0, columnspan=3, sticky="ew", pady=10)
        self.row(f, 6, "Server SMTP (uscita)", self.v_smtp, 28)
        self.row(f, 7, "Porta SMTP (SSL)", self.v_smtp_port, 8)
        self.row(f, 8, "Server IMAP (rimbalzi)", self.v_imap, 28)
        self.row(f, 9, "Porta IMAP (SSL)", self.v_imap_port, 8)
        b = ttk.Frame(f)
        b.grid(row=10, column=0, columnspan=3, sticky="w", pady=14)
        ttk.Button(b, text="Prova connessione", command=self.test_conn).pack(side="left")
        ttk.Button(b, text="Invia una mail di prova a me stesso", command=self.test_mail).pack(side="left", padx=8)
        ttk.Button(b, text="Salva impostazioni", command=self.store_cfg).pack(side="left")
        info = ("Limiti Aruba applicati automaticamente:\n"
                "• una email per destinatario (Aruba ammette max 150 destinatari per invio)\n"
                "• max 250 email ogni 20 minuti (indicazione dell'assistenza Aruba)\n"
                "• max 8000 destinatari ogni 24 ore (policy Aruba)\n"
                "• in caso di 'policy violation' il programma si mette in pausa e, se si ripete, si ferma.")
        ttk.Label(f, text=info, foreground="#335", justify="left").grid(row=11, column=0, columnspan=3, sticky="w")

    def _bg(self, fn, ok_msg):
        def work():
            try:
                fn()
                self.q.put(("info", ok_msg))
            except Exception as e:
                self.q.put(("error", str(e) or repr(e)))
        threading.Thread(target=work, daemon=True).start()

    def test_conn(self):
        a = self.account()

        def fn():
            s = a.connect_smtp(timeout=30)
            s.quit()
        self._bg(fn, "Connessione e login SMTP riusciti.")

    def test_mail(self):
        a = self.account()
        subj = self.v_subject.get() or "Prova ArubaMailer"
        body = self.txt_body.get("1.0", "end-1c") or "Messaggio di prova."

        def fn():
            s = a.connect_smtp(timeout=30)
            try:
                st, resp = me.send_one(s, a, a.sender_email, "[PROVA] " + subj, body)
                if st != "sent":
                    raise RuntimeError(resp)
            finally:
                s.quit()
        self._bg(fn, "Mail di prova inviata a %s. Controlla la casella." % a.sender_email)

    # ------------------------------------------------------------ tab 2
    def build_list(self):
        f = self.tab_list
        top = ttk.Frame(f)
        top.pack(fill="x")
        ttk.Button(top, text="Scegli file Excel / CSV...", command=self.pick_excel).pack(side="left")
        ttk.Label(top, text="   Colonna email:").pack(side="left")
        self.v_col = tk.StringVar(value=self.cfg.get("column", ""))
        ttk.Entry(top, textvariable=self.v_col, width=14).pack(side="left")
        ttk.Label(top, text="(vuoto = automatico; oppure lettera es. B, o intestazione es. Email)",
                  foreground="#666").pack(side="left", padx=6)
        self.lbl_file = ttk.Label(f, text="Nessun file selezionato", font=("Segoe UI", 10, "bold"))
        self.lbl_file.pack(anchor="w", pady=(10, 2))
        self.lbl_stats = ttk.Label(f, text="", justify="left")
        self.lbl_stats.pack(anchor="w")
        ttk.Label(f, text="Indirizzi scartati (non validi) - correggili nell'Excel se vuoi includerli:").pack(anchor="w", pady=(10, 2))
        fr = ttk.Frame(f)
        fr.pack(fill="both", expand=True)
        self.lst_invalid = tk.Listbox(fr, height=10)
        sb = ttk.Scrollbar(fr, command=self.lst_invalid.yview)
        self.lst_invalid.config(yscrollcommand=sb.set)
        self.lst_invalid.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        b = ttk.Frame(f)
        b.pack(fill="x", pady=8)
        ttk.Button(b, text="Ricarica lista", command=lambda: self.load_list(self.excel_path)).pack(side="left")
        ttk.Button(b, text="Apri campagna esistente (.db)...", command=self.open_db).pack(side="left", padx=8)
        ttk.Button(b, text="Nuova campagna da zero (azzera esiti)", command=self.reset_campaign).pack(side="left")
        ttk.Label(f, text=("Gli esiti vengono salvati nel file <nome excel>_campagna.db accanto all'Excel: "
                           "se chiudi il programma o il PC si spegne, riaprendo lo stesso Excel riparti "
                           "esattamente da dove eri rimasto, senza doppi invii."),
                  foreground="#335", wraplength=840, justify="left").pack(anchor="w")

    def pick_excel(self):
        p = filedialog.askopenfilename(title="Lista indirizzi",
                                       filetypes=[("Excel o CSV", "*.xlsx *.xlsm *.csv *.txt"), ("Tutti", "*.*")])
        if p:
            self.load_list(p)

    def db_path_for(self, p):
        return os.path.splitext(p)[0] + "_campagna.db"

    def load_list(self, p):
        if not p:
            return
        if self.job and self.job.is_alive():
            messagebox.showwarning(APP, "Ferma prima l'invio in corso.")
            return
        try:
            valid, invalid, dups = me.load_addresses(p, self.v_col.get())
        except Exception as e:
            messagebox.showerror(APP, "Errore lettura file:\n%s" % e)
            return
        self.excel_path = p
        if self.campaign:
            self.campaign.close()
        dbp = self.db_path_for(p)
        existed = os.path.exists(dbp)
        self.campaign = me.Campaign(dbp)
        added = self.campaign.import_addresses(valid)
        self.lst_invalid.delete(0, "end")
        for r, v in invalid:
            self.lst_invalid.insert("end", "riga %d:  %s" % (r, v))
        self.lbl_file.config(text=os.path.basename(p))
        txt = "Indirizzi validi: %d    Duplicati rimossi: %d    Scartati: %d" % (len(valid), dups, len(invalid))
        if existed:
            txt += "\nCampagna esistente ripresa (%d nuovi indirizzi aggiunti)." % added
            s = self.campaign.get_meta("subject")
            if s and not self.v_subject.get():
                self.v_subject.set(s)
                self.txt_body.delete("1.0", "end")
                self.txt_body.insert("1.0", self.campaign.get_meta("body", ""))
        self.lbl_stats.config(text=txt)
        self.refresh_counts()
        self.update_estimate()

    def open_db(self):
        p = filedialog.askopenfilename(title="Campagna", filetypes=[("Campagna", "*_campagna.db *.db")])
        if not p:
            return
        if self.campaign:
            self.campaign.close()
        self.campaign = me.Campaign(p)
        self.excel_path = ""
        self.lbl_file.config(text=os.path.basename(p))
        self.lbl_stats.config(text="Campagna aperta.")
        s = self.campaign.get_meta("subject")
        if s:
            self.v_subject.set(s)
            self.txt_body.delete("1.0", "end")
            self.txt_body.insert("1.0", self.campaign.get_meta("body", ""))
        self.refresh_counts()
        self.update_estimate()

    def reset_campaign(self):
        if not self.excel_path:
            messagebox.showinfo(APP, "Seleziona prima un file Excel.")
            return
        if self.job and self.job.is_alive():
            messagebox.showwarning(APP, "Ferma prima l'invio in corso.")
            return
        if not messagebox.askyesno(APP, "Azzerare tutti gli esiti di questa lista?\n"
                                        "Gli indirizzi già inviati riceverebbero di nuovo la mail."):
            return
        dbp = self.db_path_for(self.excel_path)
        if self.campaign:
            self.campaign.close()
            self.campaign = None
        for ext in ("", "-wal", "-shm"):
            try:
                os.remove(dbp + ext)
            except FileNotFoundError:
                pass
        self.load_list(self.excel_path)

    # ------------------------------------------------------------ tab 3
    def build_msg(self):
        f = self.tab_msg
        self.v_subject = tk.StringVar(value=self.cfg.get("subject", ""))
        ttk.Label(f, text="Oggetto").pack(anchor="w")
        ttk.Entry(f, textvariable=self.v_subject).pack(fill="x", pady=(0, 10))
        ttk.Label(f, text="Testo (solo testo semplice)").pack(anchor="w")
        fr = ttk.Frame(f)
        fr.pack(fill="both", expand=True)
        self.txt_body = tk.Text(fr, wrap="word", font=("Segoe UI", 10), undo=True)
        sb = ttk.Scrollbar(fr, command=self.txt_body.yview)
        self.txt_body.config(yscrollcommand=sb.set)
        self.txt_body.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.txt_body.insert("1.0", self.cfg.get("body", ""))
        ttk.Label(f, text=("Ogni destinatario riceve la mail con il proprio indirizzo nel campo A: "
                           "nessuno vede gli altri indirizzi. Usa 'Invia una mail di prova' nella scheda 1 "
                           "per vedere come arriva."), foreground="#335", wraplength=840).pack(anchor="w", pady=6)

    # ------------------------------------------------------------ tab 4
    def build_send(self):
        f = self.tab_send
        c = self.cfg
        p = ttk.LabelFrame(f, text="Programmazione", padding=10)
        p.pack(fill="x")
        self.v_batch = tk.StringVar(value=c.get("batch", "250"))
        self.v_interval = tk.StringVar(value=c.get("interval", "6"))
        self.v_period = tk.StringVar(value=c.get("period", "35"))
        self.v_cap = tk.StringVar(value=c.get("cap", str(me.DEFAULT_DAILY_CAP)))
        self.v_start_mode = tk.StringVar(value="now")
        tomorrow = (dt.datetime.now() + dt.timedelta(days=1)).replace(hour=9, minute=0)
        self.v_start = tk.StringVar(value=tomorrow.strftime("%d/%m/%Y %H:%M"))
        self.v_win_on = tk.BooleanVar(value=c.get("win_on", False))
        self.v_win_from = tk.StringVar(value=c.get("win_from", "08:00"))
        self.v_win_to = tk.StringVar(value=c.get("win_to", "20:00"))

        def spin(r, label, var, lo, hi, hint):
            ttk.Label(p, text=label).grid(row=r, column=0, sticky="w", pady=3)
            s = ttk.Spinbox(p, from_=lo, to=hi, textvariable=var, width=8, command=self.update_estimate)
            s.grid(row=r, column=1, sticky="w")
            s.bind("<KeyRelease>", lambda e: self.update_estimate())
            ttk.Label(p, text=hint, foreground="#666").grid(row=r, column=2, columnspan=3, sticky="w", padx=8)

        spin(0, "Email per blocco", self.v_batch, 1, me.MAX_BATCH, "massimo 250")
        spin(1, "Secondi tra una email e l'altra", self.v_interval, int(me.MIN_INTERVAL_S), 120,
             "minimo 4 (consigliato 6: un blocco da 250 dura ~25 min)")
        spin(2, "Minuti tra l'inizio di un blocco e il successivo", self.v_period, 1, 1440,
             "viene alzato automaticamente se troppo basso per i limiti Aruba")
        spin(3, "Massimo email ogni 24 ore", self.v_cap, 1, me.HARD_DAILY_CAP, "limite Aruba 8000")
        ttk.Label(p, text="Avvio").grid(row=4, column=0, sticky="w", pady=3)
        ttk.Radiobutton(p, text="Subito", value="now", variable=self.v_start_mode,
                        command=self.update_estimate).grid(row=4, column=1, sticky="w")
        ttk.Radiobutton(p, text="Il giorno/ora:", value="at", variable=self.v_start_mode,
                        command=self.update_estimate).grid(row=4, column=2, sticky="w", padx=8)
        e = ttk.Entry(p, textvariable=self.v_start, width=18)
        e.grid(row=4, column=3, sticky="w")
        e.bind("<KeyRelease>", lambda ev: self.update_estimate())
        ttk.Label(p, text="gg/mm/aaaa hh:mm", foreground="#666").grid(row=4, column=4, sticky="w", padx=6)
        ttk.Checkbutton(p, text="Invia solo nella fascia oraria dalle", variable=self.v_win_on,
                        command=self.update_estimate).grid(row=5, column=0, sticky="w", pady=3)
        ttk.Entry(p, textvariable=self.v_win_from, width=7).grid(row=5, column=1, sticky="w")
        ttk.Label(p, text="alle").grid(row=5, column=2, sticky="w", padx=8)
        ttk.Entry(p, textvariable=self.v_win_to, width=7).grid(row=5, column=3, sticky="w")
        self.lbl_est = ttk.Label(p, text="", foreground="#0a5", font=("Segoe UI", 9, "bold"))
        self.lbl_est.grid(row=6, column=0, columnspan=5, sticky="w", pady=(8, 0))

        b = ttk.Frame(f)
        b.pack(fill="x", pady=8)
        self.btn_start = ttk.Button(b, text="▶  Avvia / Riprendi invio", command=self.start)
        self.btn_start.pack(side="left")
        self.btn_stop = ttk.Button(b, text="■  Ferma", command=self.stop, state="disabled")
        self.btn_stop.pack(side="left", padx=6)
        ttk.Button(b, text="Controlla rimbalzi", command=self.bounces).pack(side="left", padx=6)
        ttk.Button(b, text="Riprova errori temporanei", command=self.retry_temp).pack(side="left")
        ttk.Button(b, text="Esporta esiti in Excel...", command=self.export).pack(side="left", padx=6)

        s = ttk.Frame(f)
        s.pack(fill="x")
        self.pb = ttk.Progressbar(s, maximum=100)
        self.pb.pack(fill="x")
        self.lbl_counts = ttk.Label(s, text="Nessuna lista caricata", font=("Segoe UI", 10))
        self.lbl_counts.pack(anchor="w", pady=4)
        self.lbl_status = ttk.Label(s, text="", foreground="#a50")
        self.lbl_status.pack(anchor="w")

        fr = ttk.Frame(f)
        fr.pack(fill="both", expand=True, pady=(6, 0))
        self.txt_log = tk.Text(fr, height=10, font=("Consolas", 9), state="disabled", bg="#f7f7f7")
        sb = ttk.Scrollbar(fr, command=self.txt_log.yview)
        self.txt_log.config(yscrollcommand=sb.set)
        self.txt_log.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.update_estimate()

    def parse_start(self):
        if self.v_start_mode.get() == "now":
            return None
        return dt.datetime.strptime(self.v_start.get().strip(), "%d/%m/%Y %H:%M")

    def update_estimate(self):
        try:
            n = self.campaign.remaining() if self.campaign else 0
            b, iv, per, cap = int(self.v_batch.get()), float(self.v_interval.get()), float(self.v_period.get()), int(self.v_cap.get())
            try:
                start = self.parse_start()
            except ValueError:
                self.lbl_est.config(text="Data di avvio non valida (usa gg/mm/aaaa hh:mm)", foreground="#c00")
                return
            blocks, end, eff = me.estimate(n, b, iv, per, start, cap)
            dur = min(b, me.MAX_BATCH) * max(iv, me.MIN_INTERVAL_S) / 60
            txt = ("Da inviare: %d  →  %d blocchi da max %d, ogni blocco dura ~%.0f min, un blocco ogni %.0f min"
                   % (n, blocks, min(b, me.MAX_BATCH), dur, eff))
            if n:
                txt += ".  Fine prevista: %s" % end.strftime("%d/%m/%Y %H:%M")
                if self.v_win_on.get():
                    txt += " (più tardi se la fascia oraria interrompe)"
            if eff > per + 0.01:
                txt += "\nI minuti tra blocchi sono stati portati a %.0f per rispettare i limiti Aruba." % eff
            self.lbl_est.config(text=txt, foreground="#0a5")
        except (ValueError, tk.TclError):
            self.lbl_est.config(text="Valori non validi", foreground="#c00")

    def refresh_counts(self, c=None):
        if not self.campaign:
            return
        c = c or self.campaign.counts()
        done = c["sent"] + c["failed"] + c["bounced"]
        self.pb["value"] = 100.0 * done / c["total"] if c["total"] else 0
        self.lbl_counts.config(text=("Totale %d   |   Inviate %d   |   Da inviare %d   |   Errori temporanei %d   |   "
                                     "Rifiutate %d   |   Rimbalzate %d")
                               % (c["total"], c["sent"], c["pending"], c["temp_error"], c["failed"], c["bounced"]))

    def log(self, m):
        self.txt_log.config(state="normal")
        self.txt_log.insert("end", m + "\n")
        self.txt_log.see("end")
        self.txt_log.config(state="disabled")

    def start(self):
        if not self.campaign:
            messagebox.showinfo(APP, "Carica prima la lista (scheda 2).")
            return
        if not self.v_email.get() or not self.v_pass.get():
            messagebox.showinfo(APP, "Inserisci email e password Aruba (scheda 1).")
            return
        rem = self.campaign.remaining()
        if rem == 0:
            messagebox.showinfo(APP, "Non ci sono indirizzi da inviare in questa campagna.")
            return
        try:
            start = self.parse_start()
            settings = dict(subject=self.v_subject.get(), body=self.txt_body.get("1.0", "end-1c"),
                            batch_size=self.v_batch.get(), interval_s=self.v_interval.get(),
                            period_min=self.v_period.get(), daily_cap=self.v_cap.get(), start_at=start,
                            window_on=self.v_win_on.get(), window_from=self.v_win_from.get(),
                            window_to=self.v_win_to.get())
            if self.v_win_on.get():
                dt.datetime.strptime(self.v_win_from.get(), "%H:%M")
                dt.datetime.strptime(self.v_win_to.get(), "%H:%M")
            job = me.SendJob(self.campaign, self.account(), settings, lambda k, d: self.q.put((k, d)))
        except ValueError as e:
            messagebox.showerror(APP, "Controlla i dati: %s" % e)
            return
        old = self.campaign.get_meta("subject")
        if old and old != settings["subject"] and self.campaign.counts()["sent"] > 0:
            if not messagebox.askyesno(APP, "L'oggetto è diverso da quello usato per gli invii già fatti. Continuare?"):
                return
        when = "subito" if start is None else start.strftime("il %d/%m/%Y alle %H:%M")
        if not messagebox.askyesno(APP, "Avviare l'invio a %d indirizzi %s?\n\n"
                                        "Lascia il programma aperto e il PC acceso: la sospensione "
                                        "automatica di Windows viene bloccata durante la campagna." % (rem, when)):
            return
        self.store_cfg()
        self.job = job
        self.job.start()
        self.btn_start.config(state="disabled")
        self.btn_stop.config(state="normal")

    def stop(self):
        if self.job:
            self.job.stop()
            self.lbl_status.config(text="Arresto in corso (termino l'email corrente)...")

    def retry_temp(self):
        if not self.campaign:
            return
        n = self.campaign.reset_failed_to_pending()
        self.log("Rimessi in coda %d indirizzi con errore temporaneo." % n)
        self.refresh_counts()
        self.update_estimate()

    def bounces(self):
        if not self.campaign:
            return
        a = self.account()
        camp = self.campaign

        def work():
            try:
                me.check_bounces(a, camp, log=lambda m: self.q.put(("log", m)))
                self.q.put(("progress", camp.counts()))
            except Exception as e:
                self.q.put(("error", "Controllo rimbalzi non riuscito: %s" % e))
        threading.Thread(target=work, daemon=True).start()

    def export(self):
        if not self.campaign:
            return
        base = os.path.splitext(self.excel_path)[0] + "_esiti.xlsx" if self.excel_path else "esiti.xlsx"
        p = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=os.path.basename(base),
                                         initialdir=os.path.dirname(base) or None,
                                         filetypes=[("Excel", "*.xlsx")])
        if p:
            try:
                self.campaign.export_xlsx(p)
                messagebox.showinfo(APP, "Esiti salvati in:\n%s" % p)
            except Exception as e:
                messagebox.showerror(APP, "Errore esportazione: %s" % e)

    # ------------------------------------------------------------ loop eventi
    def poll(self):
        try:
            while True:
                k, d = self.q.get_nowait()
                if k == "log":
                    self.log(d)
                elif k == "progress":
                    self.refresh_counts(d)
                    self.update_estimate()
                elif k == "status":
                    self.lbl_status.config(text=d)
                elif k == "next":
                    if d is None:
                        self.lbl_status.config(text="Invio in corso...")
                elif k == "done":
                    self.lbl_status.config(text=d)
                    self.btn_start.config(state="normal")
                    self.btn_stop.config(state="disabled")
                    messagebox.showinfo(APP, d)
                elif k == "info":
                    messagebox.showinfo(APP, d)
                elif k == "error":
                    messagebox.showerror(APP, d)
        except queue.Empty:
            pass
        self.after(250, self.poll)

    def on_close(self):
        if self.job and self.job.is_alive():
            if not messagebox.askyesno(APP, "C'è un invio in corso o programmato. Chiudendo si ferma "
                                            "(potrai riprenderlo riaprendo lo stesso Excel). Chiudere?"):
                return
            self.job.stop()
            self.job.join(10)
        self.store_cfg()
        if self.campaign:
            self.campaign.close()
        self.destroy()


if __name__ == "__main__":
    App().mainloop()
