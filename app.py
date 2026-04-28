import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import logging
import queue
import os
from pathlib import Path

from generate_certificates import (
    CertificateConfig,
    create_directories,
    read_csv_participants,
    generate_all_certificates,
)


class QueueHandler(logging.Handler):
    """Encaminha registros de log para uma fila thread-safe."""

    def __init__(self, log_queue: queue.Queue):
        super().__init__()
        self.log_queue = log_queue

    def emit(self, record: logging.LogRecord) -> None:
        self.log_queue.put(record)


class App(tk.Tk):

    def __init__(self):
        super().__init__()
        self._build_window()
        self._build_widgets()
        self._wire_logging()
        self._poll_log_queue()

    # ── janela ────────────────────────────────────────────────────────────────

    def _build_window(self) -> None:
        self.title("Gerador de Certificados")
        self.resizable(False, False)
        w, h = 720, 600
        x = (self.winfo_screenwidth() - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")
        self.configure(bg="#F3F3F3")

    # ── widgets ───────────────────────────────────────────────────────────────

    def _build_widgets(self) -> None:
        pad = {"padx": 16, "pady": 6}

        # Título
        tk.Label(
            self, text="Gerador de Certificados",
            font=("Segoe UI", 16, "bold"),
            bg="#F3F3F3"
        ).pack(pady=(18, 2))
        tk.Label(
            self,
            text="Selecione os arquivos abaixo e clique em Gerar",
            font=("Segoe UI", 9), fg="#666666", bg="#F3F3F3"
        ).pack(pady=(0, 10))

        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=16)

        # Seleção de arquivos
        file_frame = tk.Frame(self, bg="#F3F3F3")
        file_frame.pack(fill="x", **pad)

        self._csv_var = tk.StringVar()
        self._template_var = tk.StringVar()
        self._output_var = tk.StringVar()

        self._add_file_row(file_frame, 0, "Relatório CSV:",
                           self._csv_var, self._pick_csv)
        self._add_file_row(file_frame, 1, "Modelo PPTX:",
                           self._template_var, self._pick_template)
        self._add_file_row(file_frame, 2, "Pasta de saída:",
                           self._output_var, self._pick_output, is_dir=True)

        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=16, pady=(4, 0))

        # Botão gerar
        self._btn_generate = tk.Button(
            self,
            text="Gerar Certificados",
            font=("Segoe UI", 11, "bold"),
            bg="#0078D4", fg="white",
            activebackground="#005A9E", activeforeground="white",
            relief="flat", padx=24, pady=10,
            cursor="hand2",
            command=self._on_generate
        )
        self._btn_generate.pack(pady=12)

        # Barra de progresso
        progress_frame = tk.Frame(self, bg="#F3F3F3")
        progress_frame.pack(fill="x", padx=16, pady=(0, 4))
        tk.Label(
            progress_frame, text="Progresso:", font=("Segoe UI", 9),
            bg="#F3F3F3"
        ).pack(side="left")
        self._progress_var = tk.DoubleVar(value=0)
        self._progress_bar = ttk.Progressbar(
            progress_frame, variable=self._progress_var,
            maximum=100, length=500, mode="determinate"
        )
        self._progress_bar.pack(side="left", padx=8)
        self._progress_label = tk.Label(
            progress_frame, text="–", font=("Segoe UI", 9),
            bg="#F3F3F3", width=8
        )
        self._progress_label.pack(side="left")

        # Log
        log_frame = tk.Frame(self, bd=0, relief="flat", bg="#1E1E1E")
        log_frame.pack(fill="both", expand=True, padx=16, pady=(0, 6))

        self._log_text = tk.Text(
            log_frame,
            state="disabled", wrap="word",
            font=("Consolas", 8),
            bg="#1E1E1E", fg="#D4D4D4",
            relief="flat", padx=6, pady=4
        )
        scrollbar = ttk.Scrollbar(log_frame, command=self._log_text.yview)
        self._log_text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self._log_text.pack(side="left", fill="both", expand=True)

        self._log_text.tag_config("INFO",    foreground="#D4D4D4")
        self._log_text.tag_config("WARNING", foreground="#DCDCAA")
        self._log_text.tag_config("ERROR",   foreground="#F44747")
        self._log_text.tag_config("DEBUG",   foreground="#6A9955")

        # Rodapé: resumo + botão abrir pasta
        bottom_frame = tk.Frame(self, bg="#F3F3F3")
        bottom_frame.pack(fill="x", padx=16, pady=(0, 12))

        self._summary_var = tk.StringVar(value="")
        tk.Label(
            bottom_frame, textvariable=self._summary_var,
            font=("Segoe UI", 9, "bold"), anchor="w", bg="#F3F3F3"
        ).pack(side="left")

        self._btn_open = tk.Button(
            bottom_frame,
            text="Abrir Pasta de Certificados",
            font=("Segoe UI", 9),
            bg="#107C10", fg="white",
            activebackground="#0B5A0B", activeforeground="white",
            relief="flat", padx=10, pady=4,
            cursor="hand2",
            command=self._open_output_folder,
            state="disabled"
        )
        self._btn_open.pack(side="right")

    def _add_file_row(self, parent, row, label_text, var, command,
                      is_dir=False) -> None:
        tk.Label(
            parent, text=label_text, font=("Segoe UI", 9),
            width=14, anchor="e", bg="#F3F3F3"
        ).grid(row=row, column=0, sticky="e", pady=5)

        entry = tk.Entry(
            parent, textvariable=var,
            state="readonly", font=("Segoe UI", 9), width=56,
            readonlybackground="white"
        )
        entry.grid(row=row, column=1, padx=(6, 4), pady=5, sticky="ew")

        tk.Button(
            parent, text="Alterar..." if is_dir else "Selecionar...",
            font=("Segoe UI", 9), command=command,
            relief="groove", cursor="hand2"
        ).grid(row=row, column=2, pady=5)

        parent.grid_columnconfigure(1, weight=1)

    # ── logging bridge ────────────────────────────────────────────────────────

    def _wire_logging(self) -> None:
        self._log_queue: queue.Queue = queue.Queue()
        self._queue_handler = QueueHandler(self._log_queue)
        self._queue_handler.setFormatter(
            logging.Formatter(
                "%(asctime)s - %(levelname)s - %(message)s",
                datefmt="%H:%M:%S"
            )
        )
        logging.getLogger().addHandler(self._queue_handler)

    def _poll_log_queue(self) -> None:
        try:
            while True:
                record = self._log_queue.get_nowait()
                self._append_log(record)
        except queue.Empty:
            pass
        finally:
            self.after(100, self._poll_log_queue)

    def _append_log(self, record: logging.LogRecord) -> None:
        msg = self._queue_handler.format(record)
        self._log_text.configure(state="normal")
        self._log_text.insert("end", msg + "\n", record.levelname)
        self._log_text.see("end")
        self._log_text.configure(state="disabled")

    # ── seleção de arquivos ───────────────────────────────────────────────────

    def _pick_csv(self) -> None:
        path = filedialog.askopenfilename(
            title="Selecionar relatório CSV",
            filetypes=[
                ("CSV / Texto", "*.csv *.txt"),
                ("Todos os arquivos", "*.*")
            ]
        )
        if not path:
            return
        self._csv_var.set(path)
        default_out = str(Path(path).parent / "certificados_gerados")
        self._output_var.set(default_out)

    def _pick_template(self) -> None:
        path = filedialog.askopenfilename(
            title="Selecionar modelo PPTX",
            filetypes=[
                ("PowerPoint", "*.pptx *.ppt"),
                ("Todos os arquivos", "*.*")
            ]
        )
        if path:
            self._template_var.set(path)

    def _pick_output(self) -> None:
        path = filedialog.askdirectory(title="Selecionar pasta de saída")
        if path:
            self._output_var.set(path)

    # ── ação principal ────────────────────────────────────────────────────────

    def _on_generate(self) -> None:
        csv_path = self._csv_var.get().strip()
        tpl_path = self._template_var.get().strip()
        out_path = self._output_var.get().strip()

        if not csv_path:
            messagebox.showwarning("Arquivo faltando",
                                   "Por favor, selecione o relatório CSV.")
            return
        if not tpl_path:
            messagebox.showwarning("Arquivo faltando",
                                   "Por favor, selecione o modelo PPTX.")
            return
        if not out_path:
            messagebox.showwarning("Pasta faltando",
                                   "Por favor, escolha a pasta de saída.")
            return

        self._btn_generate.config(state="disabled", text="Gerando...")
        self._btn_open.config(state="disabled")
        self._summary_var.set("")
        self._progress_var.set(0)
        self._progress_label.config(text="–")

        self._log_text.configure(state="normal")
        self._log_text.delete("1.0", "end")
        self._log_text.configure(state="disabled")

        base = Path(csv_path).parent
        config = CertificateConfig(
            csv_file=Path(csv_path),
            template_pptx=Path(tpl_path),
            output_dir=Path(out_path),
            temp_dir=base / ".temp_pptx",
            progress_callback=self._on_progress,
        )

        thread = threading.Thread(
            target=self._run_generation,
            args=(config,),
            daemon=True
        )
        thread.start()

    def _on_progress(self, current: int, total: int, name: str) -> None:
        pct = (current / total * 100) if total > 0 else 0
        self.after(0, self._update_progress, current, total, pct)

    def _update_progress(self, current: int, total: int, pct: float) -> None:
        self._progress_var.set(pct)
        self._progress_label.config(text=f"{current}/{total}")

    def _run_generation(self, config: CertificateConfig) -> None:
        report = None
        error = None
        try:
            create_directories(config)
            participants = read_csv_participants(config)

            if not participants:
                raise RuntimeError(
                    "Nenhum participante encontrado no CSV.\n\n"
                    "Verifique se o arquivo é o relatório de presença correto do Teams."
                )

            report = generate_all_certificates(participants, config)

        except Exception as exc:
            error = str(exc)

        self.after(0, self._on_generation_done, report, error)

    def _on_generation_done(self, report, error) -> None:
        self._btn_generate.config(state="normal", text="Gerar Certificados")

        if error:
            messagebox.showerror("Erro", f"Ocorreu um erro:\n\n{error}")
            self._summary_var.set("Falha ao gerar certificados.")
            return

        summary = (
            f"Total: {report['total']}  |  "
            f"Sucesso: {report['success']}  |  "
            f"Falhas: {report['failed']}  |  "
            f"Taxa: {report['success_rate']:.1f}%"
        )
        self._summary_var.set(summary)
        self._btn_open.config(state="normal")

        if report['failed'] > 0:
            failed_list = "\n".join(f"  - {n}" for n in report['failed_names'])
            messagebox.showwarning(
                "Concluído com falhas",
                f"{report['success']} de {report['total']} certificados gerados.\n\n"
                f"Participantes com falha:\n{failed_list}"
            )
        else:
            messagebox.showinfo(
                "Concluído!",
                f"Todos os {report['total']} certificados foram gerados com sucesso!\n\n"
                f"Salvos em:\n{self._output_var.get()}"
            )

    def _open_output_folder(self) -> None:
        path = self._output_var.get()
        if path and os.path.isdir(path):
            os.startfile(path)
        else:
            messagebox.showwarning(
                "Pasta não encontrada",
                "A pasta de saída ainda não existe ou não foi localizada."
            )


if __name__ == "__main__":
    app = App()
    app.mainloop()
