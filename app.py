from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from balance_core import DISPLAY_COLUMNS, compare_balances, export_results_csv_file, export_results_excel_file, load_balance


APP_TITLE = "Comparateur de balances N / N-1"


class BalanceComparatorApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1180x720")
        self.root.minsize(980, 620)

        self.balance_n_path: Path | None = None
        self.balance_n1_path: Path | None = None
        self.results: list[dict[str, str]] = []

        self.status_var = tk.StringVar(value="Chargez les deux balances pour lancer la comparaison.")
        self.balance_n_var = tk.StringVar(value="Aucun fichier selectionne")
        self.balance_n1_var = tk.StringVar(value="Aucun fichier selectionne")

        self._build_ui()

    def _build_ui(self) -> None:
        container = ttk.Frame(self.root, padding=16)
        container.pack(fill="both", expand=True)

        title = ttk.Label(
            container,
            text=APP_TITLE,
            font=("Segoe UI", 18, "bold"),
        )
        title.pack(anchor="w")

        subtitle = ttk.Label(
            container,
            text="Importez deux balances et obtenez automatiquement les variations en valeur et en pourcentage.",
            font=("Segoe UI", 10),
        )
        subtitle.pack(anchor="w", pady=(4, 16))

        controls = ttk.LabelFrame(container, text="Fichiers", padding=12)
        controls.pack(fill="x")

        ttk.Button(controls, text="Choisir balance N", command=self.choose_balance_n).grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(controls, textvariable=self.balance_n_var).grid(
            row=0, column=1, sticky="w", padx=(12, 0)
        )

        ttk.Button(controls, text="Choisir balance N-1", command=self.choose_balance_n1).grid(
            row=1, column=0, sticky="w", pady=(10, 0)
        )
        ttk.Label(controls, textvariable=self.balance_n1_var).grid(
            row=1, column=1, sticky="w", padx=(12, 0), pady=(10, 0)
        )

        action_bar = ttk.Frame(container, padding=(0, 16, 0, 10))
        action_bar.pack(fill="x")

        ttk.Button(action_bar, text="Calculer les variations", command=self.run_comparison).pack(
            side="left"
        )
        ttk.Button(action_bar, text="Exporter le resultat", command=self.export_results).pack(
            side="left", padx=(10, 0)
        )

        columns = DISPLAY_COLUMNS
        self.tree = ttk.Treeview(container, columns=columns, show="headings", height=22)
        self.tree.pack(fill="both", expand=True)

        widths = {
            "COMPTE": 140,
            "LIBELLE": 320,
            "SOLDE N": 140,
            "SOLDE N-1": 140,
            "VARIATION (ABS)": 150,
            "VARIATION (%)": 130,
        }
        for column in columns:
            self.tree.heading(column, text=column)
            anchor = "w" if column in {"COMPTE", "LIBELLE"} else "e"
            self.tree.column(column, width=widths[column], anchor=anchor)

        scrollbar = ttk.Scrollbar(self.tree, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")

        status = ttk.Label(container, textvariable=self.status_var, foreground="#334155")
        status.pack(anchor="w", pady=(10, 0))

    def choose_balance_n(self) -> None:
        path = self._pick_file()
        if path:
            self.balance_n_path = path
            self.balance_n_var.set(str(path))

    def choose_balance_n1(self) -> None:
        path = self._pick_file()
        if path:
            self.balance_n1_path = path
            self.balance_n1_var.set(str(path))

    def _pick_file(self) -> Path | None:
        path = filedialog.askopenfilename(
            title="Choisir un fichier de balance",
            filetypes=[
                ("Fichiers compatibles", "*.csv *.txt *.xlsx *.xls"),
                ("CSV", "*.csv"),
                ("Texte", "*.txt"),
                ("Excel", "*.xlsx *.xls"),
                ("Tous les fichiers", "*.*"),
            ],
        )
        return Path(path) if path else None

    def run_comparison(self) -> None:
        if not self.balance_n_path or not self.balance_n1_path:
            messagebox.showwarning(APP_TITLE, "Veuillez charger les balances N et N-1.")
            return

        try:
            balance_n = load_balance(self.balance_n_path)
            balance_n1 = load_balance(self.balance_n1_path)
            self.results = compare_balances(balance_n, balance_n1)
            self._refresh_table()
            self.status_var.set(f"{len(self.results)} ligne(s) comparee(s) avec succes.")
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"Erreur lors du traitement :\n{exc}")
            self.status_var.set("Une erreur est survenue pendant le calcul.")

    def export_results(self) -> None:
        if not self.results:
            messagebox.showinfo(APP_TITLE, "Aucun resultat a exporter pour le moment.")
            return

        path = filedialog.asksaveasfilename(
            title="Exporter le resultat",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv")],
            initialfile="variations_balances.xlsx",
        )
        if not path:
            return

        output_path = Path(path)
        if output_path.suffix.lower() == ".xlsx":
            export_results_excel_file(self.results, output_path)
        else:
            export_results_csv_file(self.results, output_path)

        self.status_var.set(f"Resultat exporte vers : {path}")
        messagebox.showinfo(APP_TITLE, "Export termine avec succes.")

    def _refresh_table(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)

        for row in self.results:
            self.tree.insert("", "end", values=[row[column] for column in DISPLAY_COLUMNS])


def main() -> None:
    root = tk.Tk()
    style = ttk.Style(root)
    if "vista" in style.theme_names():
        style.theme_use("vista")
    app = BalanceComparatorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
