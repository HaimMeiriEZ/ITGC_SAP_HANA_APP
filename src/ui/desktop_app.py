import os
import subprocess
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from src.config import AppConfig
from src.pipeline import process_file


class ValidationDesktopApp:
    def __init__(self, root: tk.Tk, base_dir: Path | None = None) -> None:
        self.root = root
        self.config = AppConfig.default(base_dir or Path.cwd())
        self.config.input_dir.mkdir(parents=True, exist_ok=True)
        self.config.output_dir.mkdir(parents=True, exist_ok=True)
        self.report_path: Path | None = None

        self.file_path_var = tk.StringVar(value="")
        self.required_columns_var = tk.StringVar(value="user_id,name,email")
        self.summary_vars = {
            "total": tk.StringVar(value="0"),
            "valid": tk.StringVar(value="0"),
            "invalid": tk.StringVar(value="0"),
            "status": tk.StringVar(value="ממתין להרצה"),
        }

        self._configure_root()
        self._build_ui()

    def _configure_root(self) -> None:
        self.root.title("מערכת בדיקות קבצים - ITGC")
        self.root.geometry("1100x760")
        self.root.minsize(980, 680)
        self.root.configure(bg="#f5f7fb")

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Header.TLabel", font=("Segoe UI", 18, "bold"), foreground="#16325c")
        style.configure("Section.TLabelframe.Label", font=("Segoe UI", 11, "bold"))
        style.configure("Action.TButton", font=("Segoe UI", 10, "bold"))

    def _build_ui(self) -> None:
        container = ttk.Frame(self.root, padding=16)
        container.pack(fill=tk.BOTH, expand=True)

        header = ttk.Frame(container)
        header.pack(fill=tk.X, pady=(0, 12))
        ttk.Label(header, text="מסך בדיקת קבצי TXT ו-Excel", style="Header.TLabel").pack(side=tk.RIGHT)

        actions = ttk.Frame(container)
        actions.pack(fill=tk.X, pady=(0, 10))

        self.select_button = ttk.Button(actions, text="בחירת קובץ", style="Action.TButton", command=self.choose_file)
        self.select_button.pack(side=tk.RIGHT, padx=4)

        self.run_button = ttk.Button(actions, text="הרץ בדיקה", style="Action.TButton", command=self.run_validation)
        self.run_button.pack(side=tk.RIGHT, padx=4)

        self.report_button = ttk.Button(actions, text="פתח דוח אקסל", style="Action.TButton", command=self.open_report, state=tk.DISABLED)
        self.report_button.pack(side=tk.RIGHT, padx=4)

        self.output_button = ttk.Button(actions, text="פתח תיקיית פלט", style="Action.TButton", command=self.open_output_folder)
        self.output_button.pack(side=tk.RIGHT, padx=4)

        self.clear_button = ttk.Button(actions, text="נקה מסך", style="Action.TButton", command=self.clear_results)
        self.clear_button.pack(side=tk.RIGHT, padx=4)

        source_frame = ttk.LabelFrame(container, text="קובץ ופרמטרים", style="Section.TLabelframe")
        source_frame.pack(fill=tk.X, pady=(0, 10))
        source_frame.columnconfigure(0, weight=1)

        ttk.Label(source_frame, text="נתיב הקובץ הנבחר:").grid(row=0, column=1, sticky="e", padx=8, pady=8)
        ttk.Entry(source_frame, textvariable=self.file_path_var, justify="right").grid(row=0, column=0, sticky="ew", padx=8, pady=8)

        ttk.Label(source_frame, text="עמודות חובה:").grid(row=1, column=1, sticky="e", padx=8, pady=8)
        ttk.Entry(source_frame, textvariable=self.required_columns_var, justify="right").grid(row=1, column=0, sticky="ew", padx=8, pady=8)

        summary_frame = ttk.LabelFrame(container, text="סיכום בדיקה", style="Section.TLabelframe")
        summary_frame.pack(fill=tk.X, pady=(0, 10))

        for index, (label_text, key) in enumerate([
            ("שורות שנבדקו", "total"),
            ("שורות תקינות", "valid"),
            ("שורות שגויות", "invalid"),
            ("סטטוס", "status"),
        ]):
            card = ttk.Frame(summary_frame, padding=10)
            card.grid(row=0, column=index, padx=6, pady=6, sticky="nsew")
            ttk.Label(card, text=label_text, font=("Segoe UI", 10, "bold")).pack()
            ttk.Label(card, textvariable=self.summary_vars[key], font=("Segoe UI", 12)).pack()
            summary_frame.columnconfigure(index, weight=1)

        results_frame = ttk.LabelFrame(container, text="רשימת שגיאות", style="Section.TLabelframe")
        results_frame.pack(fill=tk.BOTH, expand=True)
        results_frame.rowconfigure(0, weight=1)
        results_frame.columnconfigure(0, weight=1)

        columns = ("row_number", "column_name", "message")
        self.issues_tree = ttk.Treeview(results_frame, columns=columns, show="headings")
        self.issues_tree.heading("row_number", text="מספר שורה")
        self.issues_tree.heading("column_name", text="שם עמודה")
        self.issues_tree.heading("message", text="הודעת שגיאה")
        self.issues_tree.column("row_number", width=110, anchor="center")
        self.issues_tree.column("column_name", width=180, anchor="e")
        self.issues_tree.column("message", width=520, anchor="e")
        self.issues_tree.grid(row=0, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(results_frame, orient=tk.VERTICAL, command=self.issues_tree.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.issues_tree.configure(yscrollcommand=scrollbar.set)

    def choose_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="בחירת קובץ לבדיקה",
            filetypes=[("Supported files", "*.txt *.csv *.xlsx *.xlsm"), ("All files", "*.*")],
        )
        if file_path:
            self.file_path_var.set(file_path)

    def _parse_required_columns(self) -> list[str]:
        raw_value = self.required_columns_var.get().replace(";", ",").replace("\n", ",")
        return [item.strip() for item in raw_value.split(",") if item.strip()]

    def run_validation(self) -> None:
        file_path = self.file_path_var.get().strip()
        if not file_path:
            messagebox.showwarning("חסר קובץ", "יש לבחור קובץ לפני הרצת הבדיקה.")
            return

        try:
            result = process_file(
                Path(file_path),
                required_columns=self._parse_required_columns(),
                output_dir=self.config.output_dir,
            )
        except Exception as error:
            messagebox.showerror("שגיאה", f"אירעה שגיאה במהלך העיבוד:\n{error}")
            return

        self.summary_vars["total"].set(str(result.summary.total_rows))
        self.summary_vars["valid"].set(str(result.summary.valid_rows))
        self.summary_vars["invalid"].set(str(result.summary.invalid_rows))
        self.summary_vars["status"].set("תקין" if result.summary.is_valid else "נמצאו שגיאות")

        for item in self.issues_tree.get_children():
            self.issues_tree.delete(item)

        if result.issues:
            for issue in result.issues:
                row_value = issue.row_number if issue.row_number > 0 else "מבנה"
                self.issues_tree.insert("", tk.END, values=(row_value, issue.column_name, issue.message))
        else:
            self.issues_tree.insert("", tk.END, values=("-", "-", "לא נמצאו שגיאות"))

        self.report_path = result.report_path
        self.report_button.config(state=tk.NORMAL if self.report_path else tk.DISABLED)

        messagebox.showinfo(
            "הבדיקה הושלמה",
            "הקובץ נבדק בהצלחה ודוח האקסל נוצר בתיקיית הפלט.",
        )

    def clear_results(self) -> None:
        self.file_path_var.set("")
        self.required_columns_var.set("user_id,name,email")
        self.summary_vars["total"].set("0")
        self.summary_vars["valid"].set("0")
        self.summary_vars["invalid"].set("0")
        self.summary_vars["status"].set("ממתין להרצה")
        self.report_path = None
        self.report_button.config(state=tk.DISABLED)
        for item in self.issues_tree.get_children():
            self.issues_tree.delete(item)

    def open_output_folder(self) -> None:
        self._open_path(self.config.output_dir)

    def open_report(self) -> None:
        if self.report_path and self.report_path.exists():
            self._open_path(self.report_path)
        else:
            messagebox.showwarning("דוח לא זמין", "טרם נוצר דוח אקסל לפתיחה.")

    @staticmethod
    def _open_path(path: Path) -> None:
        if sys.platform.startswith("win"):
            os.startfile(path)  # type: ignore[attr-defined]
            return
        if sys.platform == "darwin":
            subprocess.run(["open", str(path)], check=False)
            return
        subprocess.run(["xdg-open", str(path)], check=False)


def launch_desktop_app() -> None:
    root = tk.Tk()
    ValidationDesktopApp(root)
    root.mainloop()
