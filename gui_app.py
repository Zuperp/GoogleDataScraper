import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os

from file_utils import (
    detect_header_row_and_columns,
    read_keywords,
    read_mock_hits,
    update_hits_column,
)
from google_scraper import search_keyword

from dotenv import load_dotenv, set_key

DEFAULT_DOMAIN = "google.dk"
DEFAULT_LANG = "da"
DEFAULT_COUNTRY = "dk"
ENV_FILE = ".env"
load_dotenv(ENV_FILE)

class BatchKeywordHitsApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Batch Google Hits Inserter")
        self.geometry("670x420")
        self.resizable(False, False)

        self.main_file = None
        self.mock_file = None

        # Advanced settings state
        self.google_domain = tk.StringVar(value=os.getenv("GOOGLE_DOMAIN", DEFAULT_DOMAIN))
        self.language = tk.StringVar(value=os.getenv("HL", DEFAULT_LANG))
        self.country = tk.StringVar(value=os.getenv("GL", DEFAULT_COUNTRY))
        self.api_key = tk.StringVar(value=os.getenv("SERPAPI_API_KEY", ""))
        self.mock_mode = tk.BooleanVar(value=False)  # Default: real API mode

        # --- Main layout ---
        file_frame = ttk.LabelFrame(self, text="1. Vælg filer")
        file_frame.pack(padx=20, pady=20, fill="x")

        ttk.Label(file_frame, text="Excel med søgeord (Keywords):").grid(row=0, column=0, sticky="e", padx=4, pady=8)
        self.main_file_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.main_file_var, width=45, state="readonly").grid(row=0, column=1, padx=4)
        ttk.Button(file_frame, text="Vælg fil", command=self.choose_main_file).grid(row=0, column=2, padx=4)

        # The mock file input will be packed/unpacked based on mock_mode
        self.mock_row_widgets = []
        self.build_mock_row(file_frame)

        # --- Advanced Settings Button ---
        ttk.Button(self, text="Avanceret indstillinger", command=self.open_advanced_settings).pack(padx=20, pady=(0,10), anchor="e")

        # --- Output options ---
        option_frame = ttk.LabelFrame(self, text="2. Vælg output")
        option_frame.pack(padx=20, pady=(0, 16), fill="x")

        self.overwrite_var = tk.BooleanVar(value=False)
        ttk.Radiobutton(option_frame, text="Gem som NY fil (sikker)", variable=self.overwrite_var, value=False).pack(anchor="w", padx=8, pady=2)
        ttk.Radiobutton(option_frame, text="Overskriv den valgte fil (forsigtigt!)", variable=self.overwrite_var, value=True).pack(anchor="w", padx=8, pady=2)

        # --- Batch controls ---
        batch_frame = ttk.LabelFrame(self, text="3. Kør batch")
        batch_frame.pack(padx=20, pady=(0, 10), fill="x")

        self.progress_var = tk.StringVar(value="Ikke startet")
        self.progress_label = ttk.Label(batch_frame, textvariable=self.progress_var)
        self.progress_label.pack(anchor="w", padx=8, pady=4)

        self.run_button = ttk.Button(batch_frame, text="Start batch", command=self.run_batch)
        self.run_button.pack(pady=10)

        # Show/hide mock row at start
        self.update_mock_row()

    def build_mock_row(self, parent):
        # Row for "Excel med HITS (mock data):"
        self.mock_row_label = ttk.Label(parent, text="Excel med HITS (mock data):")
        self.mock_row_var = tk.StringVar()
        self.mock_row_entry = ttk.Entry(parent, textvariable=self.mock_row_var, width=45, state="readonly")
        self.mock_row_button = ttk.Button(parent, text="Vælg fil", command=self.choose_mock_file)
        self.mock_row_widgets = [self.mock_row_label, self.mock_row_entry, self.mock_row_button]

    def show_mock_row(self):
        for idx, widget in enumerate(self.mock_row_widgets):
            widget.grid(row=1, column=idx, padx=4, pady=8, sticky="e" if idx==0 else "")
        if self.mock_file:
            self.mock_row_var.set(os.path.basename(self.mock_file))

    def hide_mock_row(self):
        for widget in self.mock_row_widgets:
            widget.grid_forget()

    def update_mock_row(self):
        if self.mock_mode.get():
            self.show_mock_row()
        else:
            self.hide_mock_row()

    def choose_main_file(self):
        filename = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx")],
            title="Vælg hoved-Excel med søgeord"
        )
        if filename:
            self.main_file = filename
            self.main_file_var.set(os.path.basename(filename))

    def choose_mock_file(self):
        filename = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx")],
            title="Vælg Excel med HITS (mock data)"
        )
        if filename:
            self.mock_file = filename
            self.mock_row_var.set(os.path.basename(filename))

    def open_advanced_settings(self):
        win = tk.Toplevel(self)
        win.title("Avanceret indstillinger")
        win.geometry("390x320")
        win.resizable(False, False)
        row = 0

        # Test mode toggle (mock)
        mock_cb = ttk.Checkbutton(win, text="Testtilstand (MOCK)", variable=self.mock_mode, command=self.on_mock_mode_change)
        mock_cb.grid(row=row, column=0, columnspan=2, sticky="w", padx=8, pady=(16, 12))
        row += 1

        ttk.Label(win, text="Google domain (fx google.dk):").grid(row=row, column=0, sticky="e", padx=8, pady=12)
        domain_entry = ttk.Entry(win, textvariable=self.google_domain, width=22)
        domain_entry.grid(row=row, column=1, padx=8)
        row += 1

        ttk.Label(win, text="Sprogkode (hl):").grid(row=row, column=0, sticky="e", padx=8, pady=8)
        lang_entry = ttk.Entry(win, textvariable=self.language, width=22)
        lang_entry.grid(row=row, column=1, padx=8)
        row += 1

        ttk.Label(win, text="Landekode (gl):").grid(row=row, column=0, sticky="e", padx=8, pady=8)
        country_entry = ttk.Entry(win, textvariable=self.country, width=22)
        country_entry.grid(row=row, column=1, padx=8)
        row += 1

        ttk.Label(win, text="API nøgle:").grid(row=row, column=0, sticky="e", padx=8, pady=8)
        api_entry = ttk.Entry(win, textvariable=self.api_key, width=22, show="*")
        api_entry.grid(row=row, column=1, padx=8)
        row += 1

        self.api_message = tk.StringVar(value="")
        ttk.Label(win, textvariable=self.api_message, foreground="green").grid(row=row, column=0, columnspan=2, pady=(4, 0))
        row += 1

        def save_api_key():
            api_val = self.api_key.get().strip()
            set_key(ENV_FILE, "SERPAPI_API_KEY", api_val)
            self.api_message.set("API nøgle gemt i .env")
            load_dotenv(ENV_FILE, override=True)

        ttk.Button(win, text="Gem API nøgle", command=save_api_key).grid(row=row, column=1, pady=8, sticky="e")
        ttk.Button(win, text="Luk", command=win.destroy).grid(row=row, column=0, pady=8, sticky="w")

    def on_mock_mode_change(self):
        self.update_mock_row()

    def run_batch(self):
        MOCK_MODE = self.mock_mode.get()
        if not self.main_file or (MOCK_MODE and not self.mock_file):
            messagebox.showwarning(
                "Mangler fil",
                "Vælg hoved-Excel og (kun hvis testtilstand, også mock data fil)."
            )
            return

        self.run_button.config(state="disabled")
        self.progress_var.set("Indlæser filer...")
        self.update_idletasks()

        # 1. Detect header row and columns
        header_row, found_cols = detect_header_row_and_columns(self.main_file, search_cols=("Keyword", "HITS"), search_rows=8)
        if header_row is None:
            messagebox.showerror(
                "Header ikke fundet",
                "Kunne ikke finde header-rækken med både 'Keyword' og 'HITS'.\nTjek at filen er korrekt."
            )
            self.run_button.config(state="normal")
            return

        # 2. Read keywords
        df_keywords = read_keywords(self.main_file, keyword_col="Keyword", header_row=header_row)
        indices_to_process = [idx for idx, kw in df_keywords]
        keywords = [kw for idx, kw in df_keywords]

        if not keywords:
            messagebox.showwarning("Ingen søgeord", "Ingen søgeord fundet at behandle.")
            self.run_button.config(state="normal")
            return

        # 3. Prepare hits list
        n_total = len(keywords)
        n_filled = 0

        if MOCK_MODE:
            mock_hits = read_mock_hits(self.mock_file)
        else:
            mock_hits = []

        # Prepare for writing results
        results_list = []
        error_happened = False

        # Collect advanced params
        api_key = self.api_key.get().strip()
        google_domain = self.google_domain.get().strip() or DEFAULT_DOMAIN
        hl = self.language.get().strip() or DEFAULT_LANG
        gl = self.country.get().strip() or DEFAULT_COUNTRY

        for i, (idx, query) in enumerate(df_keywords):
            if MOCK_MODE:
                # Use next non-null mock value
                while mock_hits and (mock_hits[0] is None or str(mock_hits[0]).strip() == ""):
                    mock_hits.pop(0)
                if not mock_hits:
                    hits, error = None, "Not enough mock data"
                else:
                    hits, error = search_keyword(
                        query, api_key=None, mock_mode=True, mock_value=mock_hits.pop(0)
                    )
            else:
                hits, error = search_keyword(
                    query, api_key, google_domain, hl, gl, mock_mode=False
                )
            results_list.append(hits if error is None else error)
            n_filled += 1
            self.progress_var.set(f"Behandler række {n_filled} af {n_total}...")
            self.update_idletasks()
            if error is not None:
                error_happened = True

        # Write to Excel
        hits_col_idx = found_cols["HITS"]
        output_file = None
        if self.overwrite_var.get():
            output_file = update_hits_column(
                self.main_file, indices_to_process, results_list, header_row, hits_col_idx, overwrite=True
            )
        else:
            save_as = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                title="Vælg hvor resultatet skal gemmes"
            )
            if not save_as:
                self.progress_var.set("Batch annulleret.")
                self.run_button.config(state="normal")
                return
            output_file = update_hits_column(
                self.main_file, indices_to_process, results_list, header_row, hits_col_idx, overwrite=False, save_as=save_as
            )

        self.progress_var.set(f"Færdig! Skrevet til: {os.path.basename(output_file)} ({n_filled} rækker)")
        messagebox.showinfo(
            "Batch færdig",
            f"Kun HITS-kolonnen er opdateret i:\n{output_file}"
            + ("\nNogle rækker fejlede. Se efter 'ERROR' i resultatet." if error_happened else "")
        )
        self.run_button.config(state="normal")

def main():
    app = BatchKeywordHitsApp()
    app.mainloop()

if __name__ == "__main__":
    main()
