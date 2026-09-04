from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import threading
from pathlib import Path
from tkinter import filedialog, messagebox

try:
    import customtkinter as ctk
except ModuleNotFoundError as exc:
    raise SystemExit(
        "CustomTkinter nao esta instalado. Execute: python -m pip install -r requirements.txt"
    ) from exc

from separador_unificado.processors import banco_brasil, mercado_pago, stone
from separador_unificado.processors.detection import BANK_BB, BANK_MP, BANK_STONE, SUPPORTED_BANKS, detect_bank
from separador_unificado.processors.models import ProcessResult


APP_NAME = "Separador de Extratos"
ROOT_DIR = Path(__file__).resolve().parent
RESOURCE_DIR = Path(getattr(sys, "_MEIPASS", ROOT_DIR))
DEFAULT_CATEGORY_RESOURCE = RESOURCE_DIR / "config" / "categorias.json"


def app_data_dir() -> Path:
    if os.name == "nt":
        base = Path(os.environ.get("APPDATA", Path.home()))
    else:
        base = Path(os.environ.get("XDG_CONFIG_HOME", Path.home() / ".config"))
    return base / "SeparadorExtratos"


DATA_DIR = app_data_dir()
SETTINGS_FILE = DATA_DIR / "settings.json"
USER_CATEGORIES_FILE = DATA_DIR / "categorias.json"


def documents_output_dir() -> Path:
    docs = Path.home() / "Documents"
    if not docs.exists():
        docs = Path.home()
    return docs / "Extratos Organizados"


def read_json(path: Path, default):
    if not path.exists():
        return default
    try:
        with path.open(encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default


def write_json(path: Path, data) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = path.with_suffix(path.suffix + ".tmp")
    with temp_path.open("w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
        f.write("\n")
    temp_path.replace(path)


def ensure_user_categories() -> Path:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    if not USER_CATEGORIES_FILE.exists():
        if DEFAULT_CATEGORY_RESOURCE.exists():
            shutil.copyfile(DEFAULT_CATEGORY_RESOURCE, USER_CATEGORIES_FILE)
        else:
            write_json(USER_CATEGORIES_FILE, {})
    return USER_CATEGORIES_FILE


def open_in_system(path: Path) -> None:
    path = path.resolve()
    if sys.platform.startswith("win"):
        os.startfile(str(path))  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])


class SettingsStore:
    def __init__(self) -> None:
        default_categories = ensure_user_categories()
        self.data = {
            "bank": BANK_BB,
            "output_dir": str(documents_output_dir()),
            "categories_path": str(default_categories),
            "recent_files": [],
        }
        self.data.update(read_json(SETTINGS_FILE, {}))

    def save(self) -> None:
        write_json(SETTINGS_FILE, self.data)

    def add_recent_file(self, path: str | Path) -> None:
        file_path = str(Path(path).resolve())
        recent = [p for p in self.data.get("recent_files", []) if p != file_path]
        recent.insert(0, file_path)
        self.data["recent_files"] = recent[:8]
        self.save()


class ConditionRow:
    def __init__(self, parent, data: dict | None = None, on_remove=None) -> None:
        self.frame = ctk.CTkFrame(parent, fg_color="#F6F7F9", corner_radius=6)
        self.on_remove = on_remove

        data = data or {}
        field = data.get("campo", "lancamento")
        if field not in ("lancamento", "detalhes", "ambos"):
            field = "ambos"
        mode = data.get("modo", "contem")
        if mode not in ("contem", "exato"):
            mode = "contem"
        terms = ", ".join(str(term) for term in data.get("termos", []))

        self.field_var = ctk.StringVar(value=field)
        self.mode_var = ctk.StringVar(value=mode)

        self.field_menu = ctk.CTkOptionMenu(
            self.frame,
            values=["lancamento", "detalhes", "ambos"],
            variable=self.field_var,
            width=120,
        )
        self.field_menu.grid(row=0, column=0, padx=(8, 6), pady=8, sticky="w")

        self.terms_entry = ctk.CTkEntry(self.frame, placeholder_text="termo 1, termo 2")
        self.terms_entry.insert(0, terms)
        self.terms_entry.grid(row=0, column=1, padx=6, pady=8, sticky="ew")

        self.mode_menu = ctk.CTkOptionMenu(
            self.frame,
            values=["contem", "exato"],
            variable=self.mode_var,
            width=90,
        )
        self.mode_menu.grid(row=0, column=2, padx=6, pady=8)

        self.remove_button = ctk.CTkButton(
            self.frame,
            text="Remover",
            width=82,
            fg_color="#8C3B3B",
            hover_color="#743131",
            command=self.remove,
        )
        self.remove_button.grid(row=0, column=3, padx=(6, 8), pady=8)

        self.frame.grid_columnconfigure(1, weight=1)

    def grid(self, row: int) -> None:
        self.frame.grid(row=row, column=0, sticky="ew", padx=0, pady=(0, 8))

    def remove(self) -> None:
        self.frame.destroy()
        if self.on_remove:
            self.on_remove(self)

    def to_condition(self) -> dict | None:
        terms = [term.strip() for term in self.terms_entry.get().split(",") if term.strip()]
        if not terms:
            return None

        field = self.field_var.get()
        if field == "ambos":
            field = "ambos"
        return {"campo": field, "termos": terms, "modo": self.mode_var.get()}


class CategoryEditor(ctk.CTkToplevel):
    def __init__(self, master, categories_path: Path, on_saved=None) -> None:
        super().__init__(master)
        self.title("Editar categorias")
        self.geometry("900x610")
        self.minsize(800, 520)
        self.transient(master)

        self.categories_path = categories_path
        self.on_saved = on_saved
        self.categories = read_json(categories_path, {})
        if not isinstance(self.categories, dict):
            self.categories = {}

        self.selected_name: str | None = None
        self.condition_rows: list[ConditionRow] = []

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(1, weight=1)

        title = ctk.CTkLabel(
            self,
            text="Categorias do Banco do Brasil",
            font=ctk.CTkFont(size=20, weight="bold"),
            anchor="w",
        )
        title.grid(row=0, column=0, columnspan=2, padx=20, pady=(18, 8), sticky="ew")

        sidebar = ctk.CTkFrame(self, width=260, corner_radius=8)
        sidebar.grid(row=1, column=0, padx=(20, 10), pady=(0, 20), sticky="nsew")
        sidebar.grid_rowconfigure(1, weight=1)
        sidebar.grid_propagate(False)

        ctk.CTkLabel(sidebar, text="Categorias", anchor="w", font=ctk.CTkFont(weight="bold")).grid(
            row=0, column=0, padx=14, pady=(14, 8), sticky="ew"
        )

        self.category_list = ctk.CTkScrollableFrame(sidebar, fg_color="transparent")
        self.category_list.grid(row=1, column=0, padx=10, pady=(0, 10), sticky="nsew")
        sidebar.grid_columnconfigure(0, weight=1)

        new_button = ctk.CTkButton(sidebar, text="Nova categoria", command=self.new_category)
        new_button.grid(row=2, column=0, padx=10, pady=(0, 8), sticky="ew")

        delete_button = ctk.CTkButton(
            sidebar,
            text="Excluir categoria",
            fg_color="#8C3B3B",
            hover_color="#743131",
            command=self.delete_selected,
        )
        delete_button.grid(row=3, column=0, padx=10, pady=(0, 12), sticky="ew")

        editor = ctk.CTkFrame(self, corner_radius=8)
        editor.grid(row=1, column=1, padx=(10, 20), pady=(0, 20), sticky="nsew")
        editor.grid_columnconfigure(0, weight=1)
        editor.grid_rowconfigure(4, weight=1)

        ctk.CTkLabel(editor, text="Nome", anchor="w", font=ctk.CTkFont(weight="bold")).grid(
            row=0, column=0, padx=16, pady=(16, 6), sticky="ew"
        )
        self.name_entry = ctk.CTkEntry(editor, placeholder_text="Nome da categoria")
        self.name_entry.grid(row=1, column=0, padx=16, pady=(0, 14), sticky="ew")

        header = ctk.CTkFrame(editor, fg_color="transparent")
        header.grid(row=2, column=0, padx=16, pady=(0, 8), sticky="ew")
        header.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(header, text="Condições", anchor="w", font=ctk.CTkFont(weight="bold")).grid(
            row=0, column=0, sticky="w"
        )
        ctk.CTkButton(header, text="Adicionar condição", width=150, command=self.add_condition).grid(
            row=0, column=1, sticky="e"
        )

        labels = ctk.CTkFrame(editor, fg_color="transparent")
        labels.grid(row=3, column=0, padx=24, pady=(0, 4), sticky="ew")
        labels.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(labels, text="Campo", text_color="#5B6472").grid(row=0, column=0, sticky="w")
        ctk.CTkLabel(labels, text="Termos separados por vírgula", text_color="#5B6472").grid(
            row=0, column=1, sticky="w", padx=18
        )
        ctk.CTkLabel(labels, text="Modo", text_color="#5B6472").grid(row=0, column=2, sticky="w")

        self.conditions_frame = ctk.CTkScrollableFrame(editor, fg_color="transparent")
        self.conditions_frame.grid(row=4, column=0, padx=16, pady=(0, 12), sticky="nsew")
        self.conditions_frame.grid_columnconfigure(0, weight=1)

        actions = ctk.CTkFrame(editor, fg_color="transparent")
        actions.grid(row=5, column=0, padx=16, pady=(0, 16), sticky="ew")
        actions.grid_columnconfigure(0, weight=1)
        ctk.CTkButton(actions, text="Salvar alterações", width=150, command=self.save_selected).grid(
            row=0, column=1, padx=(8, 0), sticky="e"
        )

        self.refresh_category_list()
        first_name = next(iter(self.categories), None)
        if first_name:
            self.select_category(first_name)
        else:
            self.new_category()

    def refresh_category_list(self) -> None:
        for widget in self.category_list.winfo_children():
            widget.destroy()

        for row, name in enumerate(self.categories):
            selected = name == self.selected_name
            button = ctk.CTkButton(
                self.category_list,
                text=name,
                anchor="w",
                fg_color="#2F6F6D" if selected else "#E7EAEE",
                text_color="#FFFFFF" if selected else "#1D2430",
                hover_color="#275E5C" if selected else "#DCE1E7",
                command=lambda n=name: self.select_category(n),
            )
            button.grid(row=row, column=0, padx=2, pady=3, sticky="ew")
        self.category_list.grid_columnconfigure(0, weight=1)

    def select_category(self, name: str) -> None:
        self.selected_name = name
        self.name_entry.delete(0, "end")
        self.name_entry.insert(0, name)
        self.clear_conditions()

        condicoes = self.categories.get(name, {}).get("condicoes", [])
        for condition in condicoes:
            self.add_condition(condition)
        if not condicoes:
            self.add_condition()

        self.refresh_category_list()

    def clear_conditions(self) -> None:
        for row in self.condition_rows:
            row.frame.destroy()
        self.condition_rows.clear()

    def add_condition(self, data: dict | None = None) -> None:
        row = ConditionRow(self.conditions_frame, data, on_remove=self.remove_condition_row)
        self.condition_rows.append(row)
        self.layout_condition_rows()

    def remove_condition_row(self, row: ConditionRow) -> None:
        self.condition_rows = [item for item in self.condition_rows if item is not row]
        if not self.condition_rows:
            self.add_condition()
        self.layout_condition_rows()

    def layout_condition_rows(self) -> None:
        for idx, row in enumerate(self.condition_rows):
            row.grid(idx)

    def new_category(self) -> None:
        base_name = "Nova categoria"
        name = base_name
        suffix = 2
        while name in self.categories:
            name = f"{base_name} {suffix}"
            suffix += 1
        self.categories[name] = {"condicoes": []}
        self.selected_name = name
        self.refresh_category_list()
        self.select_category(name)
        self.name_entry.focus_set()
        self.name_entry.select_range(0, "end")

    def delete_selected(self) -> None:
        if not self.selected_name:
            return
        if not messagebox.askyesno("Excluir categoria", f"Excluir '{self.selected_name}'?"):
            return

        del self.categories[self.selected_name]
        self.selected_name = None
        write_json(self.categories_path, self.categories)
        if self.on_saved:
            self.on_saved()
        self.refresh_category_list()

        first_name = next(iter(self.categories), None)
        if first_name:
            self.select_category(first_name)
        else:
            self.new_category()

    def save_selected(self) -> None:
        if not self.selected_name:
            return

        new_name = self.name_entry.get().strip()
        if not new_name:
            messagebox.showwarning("Nome obrigatório", "Informe um nome para a categoria.")
            return

        if new_name != self.selected_name and new_name in self.categories:
            messagebox.showwarning("Nome duplicado", "Ja existe uma categoria com esse nome.")
            return

        conditions = []
        for row in self.condition_rows:
            condition = row.to_condition()
            if condition:
                conditions.append(condition)

        if not conditions:
            messagebox.showwarning("Condição obrigatória", "Adicione pelo menos um termo.")
            return

        if new_name != self.selected_name:
            del self.categories[self.selected_name]

        self.categories[new_name] = {"condicoes": conditions}
        self.selected_name = new_name
        write_json(self.categories_path, self.categories)
        if self.on_saved:
            self.on_saved()
        self.refresh_category_list()
        messagebox.showinfo("Categorias", "Categoria salva com sucesso.")


class SeparadorApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        self.settings = SettingsStore()
        self.last_output_path: Path | None = None

        self.title(APP_NAME)
        self.geometry("1020x620")
        self.minsize(900, 480)

        self.bank_var = ctk.StringVar(value=self.settings.data.get("bank", BANK_BB))
        self.input_path_var = ctk.StringVar(value="")
        self.output_dir_var = ctk.StringVar(value=self.settings.data.get("output_dir", str(documents_output_dir())))
        self.categories_path_var = ctk.StringVar(
            value=self.settings.data.get("categories_path", str(ensure_user_categories()))
        )
        self.status_var = ctk.StringVar(value="Pronto")
        self.result_var = ctk.StringVar(value="")

        self.configure(fg_color="#EEF1F4")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        self.build_header()
        self.build_body()
        self.build_status_bar()
        self.refresh_recent_files()
        self.update_category_visibility()

    def build_header(self) -> None:
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.grid(row=0, column=0, padx=24, pady=(22, 10), sticky="ew")
        header.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            header,
            text=APP_NAME,
            font=ctk.CTkFont(size=28, weight="bold"),
            text_color="#1C2530",
        ).grid(row=0, column=0, sticky="w")
        ctk.CTkLabel(
            header,
            text="Organize extratos do Banco do Brasil e Mercado Pago em poucos cliques.",
            font=ctk.CTkFont(size=13),
            text_color="#596575",
        ).grid(row=1, column=0, sticky="w", pady=(2, 0))

    def build_body(self) -> None:
        body = ctk.CTkFrame(self, fg_color="transparent")
        body.grid(row=1, column=0, padx=24, pady=8, sticky="nsew")
        body.grid_columnconfigure(1, weight=1)
        body.grid_rowconfigure(0, weight=1)

        controls = ctk.CTkScrollableFrame(body, width=430, corner_radius=8, fg_color="#FFFFFF")
        controls.grid(row=0, column=0, padx=(0, 14), sticky="nsew")
        controls.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(controls, text="Processamento", font=ctk.CTkFont(size=17, weight="bold")).grid(
            row=0, column=0, padx=18, pady=(18, 10), sticky="w"
        )

        ctk.CTkLabel(controls, text="Banco", text_color="#596575").grid(
            row=1, column=0, padx=18, pady=(0, 6), sticky="w"
        )
        self.bank_selector = ctk.CTkSegmentedButton(
            controls,
            values=list(SUPPORTED_BANKS),
            variable=self.bank_var,
            command=self.on_bank_changed,
            selected_color="#2F6F6D",
            selected_hover_color="#275E5C",
        )
        self.bank_selector.grid(row=2, column=0, padx=18, pady=(0, 16), sticky="ew")

        self.file_entry = self.path_field(
            controls,
            row=3,
            label="Arquivo do extrato",
            variable=self.input_path_var,
            button_text="Selecionar",
            command=self.browse_input,
        )

        self.output_entry = self.path_field(
            controls,
            row=6,
            label="Pasta de saída",
            variable=self.output_dir_var,
            button_text="Alterar",
            command=self.browse_output,
        )

        self.category_panel = ctk.CTkFrame(controls, fg_color="#F6F7F9", corner_radius=8)
        self.category_panel.grid(row=9, column=0, padx=18, pady=(0, 16), sticky="ew")
        self.category_panel.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            self.category_panel,
            text="Categorias do BB",
            font=ctk.CTkFont(weight="bold"),
            anchor="w",
        ).grid(row=0, column=0, padx=12, pady=(12, 4), sticky="ew")

        self.categories_entry = ctk.CTkEntry(self.category_panel, textvariable=self.categories_path_var)
        self.categories_entry.grid(row=1, column=0, columnspan=2, padx=12, pady=(0, 8), sticky="ew")
        ctk.CTkButton(
            self.category_panel,
            text="Escolher arquivo",
            width=130,
            command=self.browse_categories,
        ).grid(row=2, column=0, padx=(12, 6), pady=(0, 12), sticky="ew")
        ctk.CTkButton(
            self.category_panel,
            text="Editar categorias",
            width=130,
            fg_color="#4B5F76",
            hover_color="#3F5064",
            command=self.edit_categories,
        ).grid(row=2, column=1, padx=(6, 12), pady=(0, 12), sticky="ew")

        self.process_button = ctk.CTkButton(
            controls,
            text="Processar extrato",
            height=42,
            font=ctk.CTkFont(size=15, weight="bold"),
            fg_color="#2F6F6D",
            hover_color="#275E5C",
            command=self.process_current_file,
        )
        self.process_button.grid(row=10, column=0, padx=18, pady=(4, 14), sticky="ew")

        self.result_panel = ctk.CTkFrame(controls, fg_color="#EEF5F2", corner_radius=8)
        self.result_panel.grid(row=11, column=0, padx=18, pady=(0, 18), sticky="ew")
        self.result_panel.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            self.result_panel,
            textvariable=self.result_var,
            anchor="w",
            justify="left",
            text_color="#1F493E",
            wraplength=365,
        ).grid(row=0, column=0, columnspan=2, padx=12, pady=(12, 8), sticky="ew")
        ctk.CTkButton(
            self.result_panel,
            text="Abrir arquivo",
            width=130,
            command=self.open_last_output,
        ).grid(row=1, column=0, padx=(12, 6), pady=(0, 12), sticky="ew")
        ctk.CTkButton(
            self.result_panel,
            text="Abrir pasta",
            width=130,
            command=self.open_last_output_folder,
        ).grid(row=1, column=1, padx=(6, 12), pady=(0, 12), sticky="ew")
        self.result_panel.grid_remove()

        main_panel = ctk.CTkFrame(body, corner_radius=8, fg_color="#FFFFFF")
        main_panel.grid(row=0, column=1, sticky="nsew")
        main_panel.grid_columnconfigure(0, weight=1)
        main_panel.grid_rowconfigure(1, weight=1)

        top = ctk.CTkFrame(main_panel, fg_color="transparent")
        top.grid(row=0, column=0, padx=18, pady=(18, 8), sticky="ew")
        top.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(top, text="Arquivos recentes", font=ctk.CTkFont(size=17, weight="bold")).grid(
            row=0, column=0, sticky="w"
        )
        ctk.CTkButton(
            top,
            text="Limpar histórico",
            width=130,
            fg_color="#E7EAEE",
            hover_color="#DCE1E7",
            text_color="#1D2430",
            command=self.clear_history,
        ).grid(row=0, column=1, sticky="e")

        self.recent_frame = ctk.CTkScrollableFrame(main_panel, fg_color="transparent")
        self.recent_frame.grid(row=1, column=0, padx=18, pady=(0, 18), sticky="nsew")
        self.recent_frame.grid_columnconfigure(0, weight=1)

    def build_status_bar(self) -> None:
        status_bar = ctk.CTkFrame(self, fg_color="#DCE2E8", height=34, corner_radius=0)
        status_bar.grid(row=2, column=0, sticky="ew")
        status_bar.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(status_bar, textvariable=self.status_var, text_color="#334052", anchor="w").grid(
            row=0, column=0, padx=24, sticky="ew"
        )

    def path_field(self, parent, row: int, label: str, variable, button_text: str, command):
        ctk.CTkLabel(parent, text=label, text_color="#596575").grid(
            row=row, column=0, padx=18, pady=(0, 6), sticky="w"
        )
        group = ctk.CTkFrame(parent, fg_color="transparent")
        group.grid(row=row + 1, column=0, padx=18, pady=(0, 16), sticky="ew")
        group.grid_columnconfigure(0, weight=1)
        entry = ctk.CTkEntry(group, textvariable=variable)
        entry.grid(row=0, column=0, padx=(0, 8), sticky="ew")
        ctk.CTkButton(group, text=button_text, width=96, command=command).grid(row=0, column=1)
        return entry

    def on_bank_changed(self, _value=None) -> None:
        self.settings.data["bank"] = self.bank_var.get()
        self.settings.save()
        self.update_category_visibility()

    def update_category_visibility(self) -> None:
        if self.bank_var.get() == BANK_BB:
            self.category_panel.grid()
        else:
            self.category_panel.grid_remove()

    def browse_input(self) -> None:
        path = filedialog.askopenfilename(
            title="Selecionar extrato",
            filetypes=[("Planilhas Excel", "*.xlsx *.xlsm *.xls"), ("Todos os arquivos", "*.*")],
        )
        if path:
            self.set_input_path(path)

    def set_input_path(self, path: str | Path) -> None:
        file_path = Path(path)
        self.input_path_var.set(str(file_path))
        self.settings.add_recent_file(file_path)
        self.refresh_recent_files()

        detected = detect_bank(file_path)
        if detected:
            self.bank_var.set(detected)
            self.on_bank_changed()
            self.status_var.set(f"Banco detectado: {detected}")
        else:
            self.status_var.set("Nao consegui detectar o banco. Confira a seleção antes de processar.")

    def browse_output(self) -> None:
        path = filedialog.askdirectory(title="Selecionar pasta de saída")
        if path:
            self.output_dir_var.set(path)
            self.settings.data["output_dir"] = path
            self.settings.save()

    def browse_categories(self) -> None:
        path = filedialog.askopenfilename(
            title="Selecionar categorias",
            filetypes=[("Arquivos JSON", "*.json"), ("Todos os arquivos", "*.*")],
        )
        if path:
            self.categories_path_var.set(path)
            self.settings.data["categories_path"] = path
            self.settings.save()

    def edit_categories(self) -> None:
        path = Path(self.categories_path_var.get().strip() or ensure_user_categories())
        if not path.exists():
            write_json(path, {})
        self.categories_path_var.set(str(path))
        self.settings.data["categories_path"] = str(path)
        self.settings.save()
        CategoryEditor(self, path, on_saved=self.on_categories_saved)

    def on_categories_saved(self) -> None:
        self.status_var.set("Categorias salvas.")

    def refresh_recent_files(self) -> None:
        for widget in self.recent_frame.winfo_children():
            widget.destroy()

        recent = self.settings.data.get("recent_files", [])
        existing = [Path(path) for path in recent if Path(path).exists()]

        if not existing:
            ctk.CTkLabel(
                self.recent_frame,
                text="Nenhum arquivo recente ainda.",
                text_color="#6A7482",
                anchor="w",
            ).grid(row=0, column=0, padx=4, pady=8, sticky="ew")
            return

        for row_idx, path in enumerate(existing):
            row = ctk.CTkFrame(self.recent_frame, fg_color="#F6F7F9", corner_radius=8)
            row.grid(row=row_idx, column=0, padx=2, pady=6, sticky="ew")
            row.grid_columnconfigure(0, weight=1)

            detected = detect_bank(path) or "Banco nao identificado"
            ctk.CTkLabel(
                row,
                text=path.name,
                font=ctk.CTkFont(size=14, weight="bold"),
                anchor="w",
            ).grid(row=0, column=0, padx=12, pady=(10, 0), sticky="ew")
            ctk.CTkLabel(
                row,
                text=f"{detected} | {path.parent}",
                text_color="#5F6B79",
                anchor="w",
                wraplength=540,
            ).grid(row=1, column=0, padx=12, pady=(0, 10), sticky="ew")
            ctk.CTkButton(
                row,
                text="Usar",
                width=80,
                command=lambda p=path: self.set_input_path(p),
            ).grid(row=0, column=1, rowspan=2, padx=12, pady=10)

    def clear_history(self) -> None:
        self.settings.data["recent_files"] = []
        self.settings.save()
        self.refresh_recent_files()
        self.status_var.set("Histórico limpo.")

    def validate_before_processing(self) -> bool:
        input_path = Path(self.input_path_var.get().strip())
        if not input_path.exists():
            messagebox.showwarning("Arquivo obrigatório", "Selecione um arquivo de extrato válido.")
            return False

        if input_path.suffix.lower() not in {".xlsx", ".xlsm", ".xls"}:
            messagebox.showwarning("Formato não suportado", "Selecione um arquivo .xlsx, .xlsm ou .xls.")
            return False

        if input_path.suffix.lower() == ".xls" and self.bank_var.get() != BANK_STONE:
            messagebox.showwarning(
                "Formato não suportado",
                "Arquivos .xls são suportados apenas para extratos Stone. Para BB ou Mercado Pago, use .xlsx.",
            )
            return False

        output_dir_text = self.output_dir_var.get().strip()
        if not output_dir_text:
            messagebox.showwarning("Pasta obrigatória", "Selecione uma pasta de saída.")
            return False

        selected_bank = self.bank_var.get()
        detected = detect_bank(input_path)
        if detected and detected != selected_bank:
            return messagebox.askyesno(
                "Banco diferente",
                f"Esse arquivo parece ser de {detected}, mas a seleção atual é {selected_bank}.\n\n"
                "Deseja processar mesmo assim?",
            )

        if not detected:
            return messagebox.askyesno(
                "Banco não identificado",
                "Nao consegui identificar o banco automaticamente.\n\nDeseja processar com a seleção atual?",
            )

        if selected_bank == BANK_BB:
            categories_path = Path(self.categories_path_var.get().strip())
            if not categories_path.exists():
                return messagebox.askyesno(
                    "Categorias não encontradas",
                    "O arquivo de categorias não foi encontrado. O extrato será agrupado em Outros.\n\nContinuar?",
                )

        if selected_bank == BANK_STONE:
            try:
                header_match = stone.analyze_columns(input_path)
            except Exception as exc:
                messagebox.showerror("Erro ao ler o extrato", str(exc))
                return False

            if header_match.missing:
                messagebox.showerror(
                    "Não é possível separar os lançamentos",
                    stone.format_missing_columns_message(header_match.missing),
                )
                return False

            if not header_match.is_exact:
                detalhes = "\n".join(
                    f'- {stone.DISPLAY_HEADERS.get(campo, campo)}: usei a coluna "{valor}"'
                    for campo, valor in header_match.fuzzy_matches.items()
                )
                return messagebox.askyesno(
                    "Formato diferente do esperado",
                    "O formato do arquivo parece diferente do extrato Stone padrão, mas encontrei colunas "
                    f"equivalentes:\n\n{detalhes}\n\nDeseja continuar mesmo assim?",
                )

        return True

    def process_current_file(self) -> None:
        if not self.validate_before_processing():
            return

        input_path = Path(self.input_path_var.get().strip())
        output_dir = Path(self.output_dir_var.get().strip())
        categories_path = Path(self.categories_path_var.get().strip())
        selected_bank = self.bank_var.get()

        self.process_button.configure(state="disabled", text="Processando...")
        self.status_var.set("Processando extrato...")
        self.result_panel.grid_remove()

        thread = threading.Thread(
            target=self._process_worker,
            args=(input_path, output_dir, categories_path, selected_bank),
            daemon=True,
        )
        thread.start()

    def _process_worker(
        self,
        input_path: Path,
        output_dir: Path,
        categories_path: Path,
        selected_bank: str,
    ) -> None:
        try:
            if selected_bank == BANK_BB:
                result = banco_brasil.process_file(input_path, output_dir, categories_path)
            elif selected_bank == BANK_MP:
                result = mercado_pago.process_file(input_path, output_dir)
            elif selected_bank == BANK_STONE:
                result = stone.process_file(input_path, output_dir)
            else:
                raise ValueError(f"Banco nao suportado: {selected_bank}")
        except Exception as exc:
            self.after(0, lambda error=exc: self.on_processing_error(error))
            return

        self.after(0, lambda processed_result=result: self.on_processing_success(processed_result))

    def on_processing_success(self, result: ProcessResult) -> None:
        self.last_output_path = result.output_path
        self.settings.data["bank"] = self.bank_var.get()
        self.settings.data["output_dir"] = self.output_dir_var.get().strip()
        self.settings.data["categories_path"] = self.categories_path_var.get().strip()
        self.settings.add_recent_file(self.input_path_var.get().strip())

        details = f"Arquivo gerado:\n{result.output_path}"
        if result.banco == BANK_BB:
            details += f"\n\nEntradas: {result.entradas} | Saídas: {result.saidas}"
        elif result.banco == BANK_STONE:
            details += f"\n\nCréditos: {result.entradas} | Débitos: {result.saidas}"
        elif result.total_rows:
            details += f"\n\nLançamentos: {result.total_rows}"

        self.result_var.set(details)
        self.result_panel.grid()
        self.status_var.set("Arquivo gerado com sucesso.")
        self.process_button.configure(state="normal", text="Processar extrato")
        self.refresh_recent_files()

    def on_processing_error(self, error: Exception) -> None:
        self.status_var.set("Erro ao processar extrato.")
        self.process_button.configure(state="normal", text="Processar extrato")
        messagebox.showerror("Erro ao processar", str(error))

    def open_last_output(self) -> None:
        if self.last_output_path and self.last_output_path.exists():
            open_in_system(self.last_output_path)

    def open_last_output_folder(self) -> None:
        if self.last_output_path:
            open_in_system(self.last_output_path.parent)


def main() -> None:
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")
    if "--smoke-test" in sys.argv:
        app = SeparadorApp()
        app.update()
        app.destroy()
        return

    app = SeparadorApp()
    app.mainloop()


if __name__ == "__main__":
    main()
