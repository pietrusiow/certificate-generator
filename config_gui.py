from __future__ import annotations

import csv
import json
import logging
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser, scrolledtext
from typing import Dict, List, Optional, Tuple, Union

try:
    from PIL import Image, ImageDraw, ImageFont, ImageTk
except ImportError as exc:  # pragma: no cover - handled at runtime
    raise SystemExit("Pillow is required: pip install pillow") from exc


DEFAULT_CONFIG_PATH = Path("config/content_config.json")
PREVIEW_MAX_WIDTH = 900
PREVIEW_MAX_HEIGHT = 600
DEFAULT_VALUES = {
    "background_image": "./background/sample.png",
    "font_path": "./fonts/Lato/Lato-Black.ttf",
    "font_size": "32",
    "text_y": "160",
    "long_name_threshold": "",
    "long_name_font_size": "",
    "long_name_text_y": "",
    "split_name_threshold": "",
    "split_name_line_gap": "",
    "split_name_font_size": "",
    "split_name_text_y": "",
    "orientation": "L",
    "text_color": "#000000",
}

EMAIL_CONFIG_PATH = Path("config/email_config.json")
SMTP_CONFIG_PATH = Path("config/smtp_config.json")
DEBUG_CONFIG_PATH = Path("config/debug_mode.json")


class ConfigGUI:
    """Tkinter GUI for editing certificate content configuration files."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Certificate Config Editor")
        self.root.minsize(1100, 650)

        self.current_config_path: Optional[Path] = None
        self.preview_after_id: Optional[str] = None
        self.preview_image = None
        self.preview_photo = None

        self.vars = {key: tk.StringVar(value=value) for key, value in DEFAULT_VALUES.items()}
        self.preview_name_var = tk.StringVar(value="Sample Recipient")
        self.path_var = tk.StringVar(value="(unsaved)")
        self.status_var = tk.StringVar(value="")

        self.email_vars = {
            "subject": tk.StringVar(),
            "throttle_per_minute": tk.StringVar(),
        }
        self.email_status_var = tk.StringVar(value="")
        self.email_body_widget: Optional[scrolledtext.ScrolledText] = None

        self.smtp_vars = {
            "smtp_server": tk.StringVar(),
            "smtp_port": tk.StringVar(),
            "email_sender": tk.StringVar(),
            "email_password": tk.StringVar(),
        }
        self.smtp_status_var = tk.StringVar(value="")

        self.debug_var = tk.BooleanVar(value=False)
        self.debug_status_var = tk.StringVar(value="")

        self.participants_path = Path("participants.csv")
        self.participants_path_var = tk.StringVar(value=str(self.participants_path))
        self.participants_columns: List[str] = []
        self.participants_tree: Optional[ttk.Treeview] = None
        self.participants_status_var = tk.StringVar(value="")
        self.participant_entry_vars: Dict[str, tk.StringVar] = {}
        self.participants_form_frame: Optional[ttk.Frame] = None

        self._attach_var_traces()

        self._build_layout()
        self._load_initial_configs()
        self.schedule_preview_update()

    # ------------------------------------------------------------------ UI setup
    def _build_layout(self) -> None:
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill="both", expand=True)

        left_frame = ttk.Frame(main_frame)
        left_frame.pack(side="left", fill="both", expand=True)

        preview_frame = ttk.Frame(main_frame)
        preview_frame.pack(side="right", fill="both", expand=True, padx=(15, 0))

        self.notebook = ttk.Notebook(left_frame)
        self.notebook.pack(fill="both", expand=True)

        content_frame = ttk.Frame(self.notebook, padding=10)
        participants_frame = ttk.Frame(self.notebook, padding=10)
        email_frame = ttk.Frame(self.notebook, padding=10)
        smtp_frame = ttk.Frame(self.notebook, padding=10)
        debug_frame = ttk.Frame(self.notebook, padding=10)

        self.content_tab = content_frame
        self.notebook.add(content_frame, text="Certificate Layout")
        self.notebook.add(participants_frame, text="Participants")
        self.notebook.add(email_frame, text="Email Template")
        self.notebook.add(smtp_frame, text="SMTP Server")
        self.notebook.add(debug_frame, text="Debug Mode")

        self._build_content_tab(content_frame)
        self._build_participants_tab(participants_frame)
        self._build_email_tab(email_frame)
        self._build_smtp_tab(smtp_frame)
        self._build_debug_tab(debug_frame)

        # Preview canvas --------------------------------------------------------
        self.preview_frame = preview_frame
        ttk.Label(self.preview_frame, text="Live Preview").pack(anchor="w")
        self.preview_canvas = tk.Canvas(
            self.preview_frame,
            background="#e0e0e0",
            highlightthickness=1,
            highlightbackground="#cccccc",
            width=PREVIEW_MAX_WIDTH,
            height=PREVIEW_MAX_HEIGHT,
        )
        self.preview_canvas.pack(fill="both", expand=True, pady=(5, 0))

        self.notebook.bind("<<NotebookTabChanged>>", lambda *_: self._update_preview_visibility())
        self._update_preview_visibility()

    def _build_content_tab(self, frame: ttk.Frame) -> None:
        ttk.Label(frame, text="Configuration File").grid(row=0, column=0, sticky="w")
        ttk.Label(frame, textvariable=self.path_var, foreground="#555555", wraplength=260, justify="left").grid(
            row=1, column=0, columnspan=3, sticky="w", pady=(0, 10)
        )

        ttk.Button(frame, text="Load...", command=self.select_and_load_config).grid(
            row=2, column=0, pady=(0, 10), sticky="ew"
        )
        ttk.Button(frame, text="Save", command=self.save_config).grid(
            row=2, column=1, pady=(0, 10), sticky="ew", padx=5
        )
        ttk.Button(frame, text="Save As...", command=self.save_config_as).grid(
            row=2, column=2, pady=(0, 10), sticky="ew"
        )

        row = 3
        self._build_path_entry(
            frame,
            row,
            label="Background Image",
            key="background_image",
            dialog_title="Select Background Image",
            filetypes=[("Image files", "*.png;*.jpg;*.jpeg"), ("All files", "*.*")],
        )
        row += 2

        self._build_path_entry(
            frame,
            row,
            label="Font File",
            key="font_path",
            dialog_title="Select Font File",
            filetypes=[("TrueType font", "*.ttf;*.otf"), ("All files", "*.*")],
        )
        row += 2

        self._build_simple_entry(frame, row, "Font Size (pt)", "font_size")
        row += 1
        self._build_simple_entry(frame, row, "Text Baseline Y (mm)", "text_y")
        row += 1
        self._build_simple_entry(frame, row, "Long Name Threshold (chars)", "long_name_threshold")
        row += 1
        self._build_simple_entry(frame, row, "Long Name Font Size (pt)", "long_name_font_size")
        row += 1
        self._build_simple_entry(frame, row, "Long Name Baseline Y (mm)", "long_name_text_y")
        row += 1
        self._build_simple_entry(frame, row, "Split Name Threshold (chars)", "split_name_threshold")
        row += 1
        self._build_simple_entry(frame, row, "Split Line Gap (mm)", "split_name_line_gap")
        row += 1
        self._build_simple_entry(frame, row, "Split Name Font Size (pt)", "split_name_font_size")
        row += 1
        self._build_simple_entry(frame, row, "Split Name Baseline Y (mm)", "split_name_text_y")
        row += 1

        ttk.Label(frame, text="Orientation").grid(row=row, column=0, sticky="w", pady=(10, 0))
        orientation_box = ttk.Combobox(
            frame,
            textvariable=self.vars["orientation"],
            values=("L", "P"),
            state="readonly",
            width=5,
        )
        orientation_box.grid(row=row, column=1, sticky="w", pady=(10, 0))
        row += 1

        ttk.Label(frame, text="Text Color").grid(row=row, column=0, sticky="w", pady=(10, 0))
        color_frame = ttk.Frame(frame)
        color_frame.grid(row=row, column=1, columnspan=2, sticky="ew", pady=(10, 0))
        color_entry = ttk.Entry(color_frame, textvariable=self.vars["text_color"], width=12)
        color_entry.pack(side="left", fill="x", expand=True)
        ttk.Button(color_frame, text="Pick...", command=self.choose_text_color).pack(side="left", padx=(5, 0))
        row += 1

        ttk.Label(frame, text="Preview Name").grid(row=row, column=0, sticky="w", pady=(10, 0))
        ttk.Entry(frame, textvariable=self.preview_name_var).grid(
            row=row, column=1, columnspan=2, sticky="ew", pady=(10, 0)
        )
        row += 1

        ttk.Button(frame, text="Refresh Preview", command=self.update_preview).grid(
            row=row, column=0, columnspan=3, pady=(15, 0), sticky="ew"
        )
        row += 1

        ttk.Label(frame, textvariable=self.status_var, foreground="#444444", wraplength=260, justify="left").grid(
            row=row, column=0, columnspan=3, sticky="w", pady=(15, 0)
        )

        for col in range(3):
            frame.columnconfigure(col, weight=1)

    def _build_participants_tab(self, frame: ttk.Frame) -> None:
        frame.columnconfigure(0, weight=1)

        intro = (
            "Review and edit the participant list stored in CSV format. "
            "You can load another CSV, add new rows, remove selected entries, and save the file."
        )
        ttk.Label(frame, text=intro, wraplength=360, justify="left").grid(row=0, column=0, sticky="w")

        row = 1
        path_frame = ttk.Frame(frame)
        path_frame.grid(row=row, column=0, sticky="ew", pady=(10, 0))
        path_frame.columnconfigure(1, weight=1)
        ttk.Label(path_frame, text="Current file:").grid(row=0, column=0, sticky="w")
        ttk.Label(path_frame, textvariable=self.participants_path_var, foreground="#555555", wraplength=280).grid(
            row=0, column=1, sticky="w"
        )

        row += 1
        button_frame = ttk.Frame(frame)
        button_frame.grid(row=row, column=0, sticky="ew", pady=(10, 0))
        for idx in range(3):
            button_frame.columnconfigure(idx, weight=1)
        ttk.Button(
            button_frame,
            text="Load CSV...",
            command=self.load_participants_via_dialog,
        ).grid(row=0, column=0, sticky="ew", padx=(0, 5))
        ttk.Button(
            button_frame,
            text="Save",
            command=self.save_participants,
        ).grid(row=0, column=1, sticky="ew", padx=5)
        ttk.Button(
            button_frame,
            text="Save As...",
            command=self.save_participants_as,
        ).grid(row=0, column=2, sticky="ew", padx=(5, 0))

        row += 1
        tree_frame = ttk.Frame(frame)
        tree_frame.grid(row=row, column=0, sticky="nsew", pady=(10, 0))
        frame.rowconfigure(row, weight=1)
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)

        self.participants_tree = ttk.Treeview(tree_frame, show="headings", selectmode="browse")
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.participants_tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.participants_tree.xview)
        self.participants_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.participants_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        row += 1
        ttk.Button(
            frame,
            text="Remove Selected",
            command=self.remove_selected_participant,
        ).grid(row=row, column=0, sticky="ew", pady=(10, 0))

        row += 1
        self.participants_form_frame = ttk.LabelFrame(frame, text="Add Participant")
        self.participants_form_frame.grid(row=row, column=0, sticky="ew", pady=(15, 0))
        self.participants_form_frame.columnconfigure(1, weight=1)

        row += 1
        ttk.Label(
            frame,
            textvariable=self.participants_status_var,
            foreground="#444444",
            wraplength=360,
            justify="left",
        ).grid(row=row, column=0, sticky="w", pady=(12, 0))

        self._set_participants_columns(["FirstName", "LastName", "Email"])

    def _build_email_tab(self, frame: ttk.Frame) -> None:
        frame.columnconfigure(0, weight=1)

        intro = (
            "Adjust the email message that is sent with each certificate. "
            "Make it clear why the recipient is receiving the message and what to do next."
        )
        ttk.Label(frame, text=intro, wraplength=360, justify="left").grid(row=0, column=0, sticky="w")

        row = 1
        ttk.Label(frame, text="Subject").grid(row=row, column=0, sticky="w", pady=(10, 0))
        row += 1
        ttk.Entry(frame, textvariable=self.email_vars["subject"]).grid(row=row, column=0, sticky="ew")
        row += 1
        ttk.Label(
            frame,
            text="Appears in the recipient's inbox. Keep it short, friendly, and specific.",
            foreground="#555555",
            wraplength=360,
            justify="left",
        ).grid(row=row, column=0, sticky="w", pady=(4, 0))

        row += 1
        ttk.Label(frame, text="Body (HTML)").grid(row=row, column=0, sticky="w", pady=(12, 0))
        row += 1
        self.email_body_widget = scrolledtext.ScrolledText(frame, height=10, wrap="word")
        self.email_body_widget.grid(row=row, column=0, sticky="nsew")
        frame.rowconfigure(row, weight=1)
        row += 1
        ttk.Label(
            frame,
            text="Use HTML for formatting. Include {name} to insert the participant name automatically.",
            foreground="#555555",
            wraplength=360,
            justify="left",
        ).grid(row=row, column=0, sticky="w", pady=(4, 0))

        row += 1
        ttk.Label(frame, text="Throttle Per Minute").grid(row=row, column=0, sticky="w", pady=(12, 0))
        row += 1
        ttk.Entry(frame, textvariable=self.email_vars["throttle_per_minute"]).grid(row=row, column=0, sticky="ew")
        row += 1
        ttk.Label(
            frame,
            text="Limit how many emails are sent per minute so you stay within your provider's sending limits. "
            "Set 0 to send as fast as possible.",
            foreground="#555555",
            wraplength=360,
            justify="left",
        ).grid(row=row, column=0, sticky="w", pady=(4, 0))

        row += 1
        button_frame = ttk.Frame(frame)
        button_frame.grid(row=row, column=0, sticky="ew", pady=(15, 0))
        button_frame.columnconfigure(0, weight=1)
        button_frame.columnconfigure(1, weight=1)
        ttk.Button(
            button_frame,
            text="Reload from File",
            command=lambda: self.load_email_config(show_message=True),
        ).grid(row=0, column=0, sticky="ew", padx=(0, 5))
        ttk.Button(
            button_frame,
            text="Save Changes",
            command=self.save_email_config,
        ).grid(row=0, column=1, sticky="ew")

        row += 1
        ttk.Label(frame, textvariable=self.email_status_var, foreground="#444444", wraplength=360, justify="left").grid(
            row=row, column=0, sticky="w", pady=(12, 0)
        )

    def _build_smtp_tab(self, frame: ttk.Frame) -> None:
        frame.columnconfigure(0, weight=1)

        intro = (
            "Configure how the application connects to your email provider. "
            "These settings come from your provider's SMTP documentation or admin dashboard."
        )
        ttk.Label(frame, text=intro, wraplength=360, justify="left").grid(row=0, column=0, sticky="w")

        row = 1
        ttk.Label(frame, text="SMTP Server").grid(row=row, column=0, sticky="w", pady=(10, 0))
        row += 1
        ttk.Entry(frame, textvariable=self.smtp_vars["smtp_server"]).grid(row=row, column=0, sticky="ew")
        row += 1
        ttk.Label(
            frame,
            text="Example: smtp.gmail.com or smtp.yourdomain.com. Check your email provider's support page.",
            foreground="#555555",
            wraplength=360,
            justify="left",
        ).grid(row=row, column=0, sticky="w", pady=(4, 0))

        row += 1
        ttk.Label(frame, text="SMTP Port").grid(row=row, column=0, sticky="w", pady=(12, 0))
        row += 1
        ttk.Entry(frame, textvariable=self.smtp_vars["smtp_port"]).grid(row=row, column=0, sticky="ew")
        row += 1
        ttk.Label(
            frame,
            text="Common values are 587 for TLS or 465 for SSL. Use the value recommended by your provider.",
            foreground="#555555",
            wraplength=360,
            justify="left",
        ).grid(row=row, column=0, sticky="w", pady=(4, 0))

        row += 1
        ttk.Label(frame, text="Sender Email Address").grid(row=row, column=0, sticky="w", pady=(12, 0))
        row += 1
        ttk.Entry(frame, textvariable=self.smtp_vars["email_sender"]).grid(row=row, column=0, sticky="ew")
        row += 1
        ttk.Label(
            frame,
            text="Certificates will be sent from this address. Use a mailbox that your participants expect.",
            foreground="#555555",
            wraplength=360,
            justify="left",
        ).grid(row=row, column=0, sticky="w", pady=(4, 0))

        row += 1
        ttk.Label(frame, text="SMTP Password or App Password").grid(row=row, column=0, sticky="w", pady=(12, 0))
        row += 1
        ttk.Entry(frame, textvariable=self.smtp_vars["email_password"], show="*").grid(row=row, column=0, sticky="ew")
        row += 1
        ttk.Label(
            frame,
            text="Use the password or application-specific password provided by your email provider. "
            "Keep it secret and regenerate it if compromised.",
            foreground="#555555",
            wraplength=360,
            justify="left",
        ).grid(row=row, column=0, sticky="w", pady=(4, 0))

        row += 1
        button_frame = ttk.Frame(frame)
        button_frame.grid(row=row, column=0, sticky="ew", pady=(15, 0))
        button_frame.columnconfigure(0, weight=1)
        button_frame.columnconfigure(1, weight=1)
        ttk.Button(
            button_frame,
            text="Reload from File",
            command=lambda: self.load_smtp_config(show_message=True),
        ).grid(row=0, column=0, sticky="ew", padx=(0, 5))
        ttk.Button(
            button_frame,
            text="Save Changes",
            command=self.save_smtp_config,
        ).grid(row=0, column=1, sticky="ew")

        row += 1
        ttk.Label(frame, textvariable=self.smtp_status_var, foreground="#444444", wraplength=360, justify="left").grid(
            row=row, column=0, sticky="w", pady=(12, 0)
        )

    def _build_debug_tab(self, frame: ttk.Frame) -> None:
        frame.columnconfigure(0, weight=1)

        intro = (
            "Debug mode lets you test the process without sending real emails. "
            "Choose \"Test\" while rehearsing and switch to \"Full\" when you are ready to deliver certificates."
        )
        ttk.Label(frame, text=intro, wraplength=360, justify="left").grid(row=0, column=0, sticky="w")

        row = 1
        ttk.Checkbutton(
            frame,
            text="Enable test mode (do not send emails)",
            variable=self.debug_var,
            onvalue=True,
            offvalue=False,
        ).grid(row=row, column=0, sticky="w", pady=(12, 0))

        row += 1
        ttk.Label(
            frame,
            text="When checked, the generator keeps logging but skips the SMTP send step.",
            foreground="#555555",
            wraplength=360,
            justify="left",
        ).grid(row=row, column=0, sticky="w", pady=(4, 0))

        row += 1
        button_frame = ttk.Frame(frame)
        button_frame.grid(row=row, column=0, sticky="ew", pady=(15, 0))
        button_frame.columnconfigure(0, weight=1)
        button_frame.columnconfigure(1, weight=1)
        ttk.Button(
            button_frame,
            text="Reload from File",
            command=lambda: self.load_debug_config(show_message=True),
        ).grid(row=0, column=0, sticky="ew", padx=(0, 5))
        ttk.Button(
            button_frame,
            text="Save Changes",
            command=self.save_debug_config,
        ).grid(row=0, column=1, sticky="ew")

        row += 1
        ttk.Label(frame, textvariable=self.debug_status_var, foreground="#444444", wraplength=360, justify="left").grid(
            row=row, column=0, sticky="w", pady=(12, 0)
        )

    def _set_participants_columns(self, columns: List[str]) -> None:
        self.participants_columns = [col for col in columns if col]

        if self.participants_tree is not None:
            self.participants_tree["columns"] = self.participants_columns
            for col in self.participants_columns:
                self.participants_tree.heading(col, text=col)
                self.participants_tree.column(col, width=140, anchor="center")

        self._rebuild_participant_form()

    def _rebuild_participant_form(self) -> None:
        if self.participants_form_frame is None:
            return

        for child in self.participants_form_frame.winfo_children():
            child.destroy()

        self.participant_entry_vars = {}

        if not self.participants_columns:
            ttk.Label(
                self.participants_form_frame,
                text="Load a CSV file to add new participants.",
                foreground="#555555",
                wraplength=320,
                justify="left",
            ).grid(row=0, column=0, sticky="w", padx=5, pady=5)
            return

        ttk.Label(
            self.participants_form_frame,
            text="Enter values for each column and click Add Participant.",
            foreground="#555555",
            wraplength=320,
            justify="left",
        ).grid(row=0, column=0, columnspan=2, sticky="w", padx=5, pady=(5, 10))

        row = 1
        for column in self.participants_columns:
            pretty_label = column.replace("_", " ")
            var = tk.StringVar()
            self.participant_entry_vars[column] = var

            ttk.Label(self.participants_form_frame, text=pretty_label).grid(
                row=row, column=0, sticky="w", padx=5, pady=3
            )
            ttk.Entry(self.participants_form_frame, textvariable=var).grid(
                row=row, column=1, sticky="ew", padx=5, pady=3
            )
            row += 1

        ttk.Button(
            self.participants_form_frame,
            text="Add Participant",
            command=self.add_participant,
        ).grid(row=row, column=0, columnspan=2, sticky="ew", padx=5, pady=(10, 5))
        self.participants_form_frame.columnconfigure(1, weight=1)

    def _update_preview_visibility(self) -> None:
        if not hasattr(self, "notebook") or not hasattr(self, "preview_frame"):
            return

        current_tab = self.notebook.select()
        should_show = current_tab == str(self.content_tab)

        if should_show:
            if self.preview_frame.winfo_manager() != "pack":
                self.preview_frame.pack(side="right", fill="both", expand=True, padx=(15, 0))
        else:
            self.preview_frame.pack_forget()

    def _attach_var_traces(self) -> None:
        for var in self.vars.values():
            var.trace_add("write", lambda *_args: self.schedule_preview_update())
        self.preview_name_var.trace_add("write", lambda *_args: self.schedule_preview_update())

    def _build_simple_entry(self, parent: ttk.Frame, row: int, label: str, key: str) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=5)
        entry = ttk.Entry(parent, textvariable=self.vars[key])
        entry.grid(row=row, column=1, columnspan=2, sticky="ew", pady=5)

    def _build_path_entry(
        self,
        parent: ttk.Frame,
        row: int,
        *,
        label: str,
        key: str,
        dialog_title: str,
        filetypes: List[Tuple[str, str]],
    ) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=5)
        entry = ttk.Entry(parent, textvariable=self.vars[key])
        entry.grid(row=row, column=0, columnspan=2, sticky="ew", pady=5)
        ttk.Button(
            parent,
            text="Browse...",
            command=lambda: self._select_path_for_key(key, dialog_title, filetypes),
        ).grid(row=row, column=2, sticky="ew", padx=(5, 0), pady=5)

    # ------------------------------------------------------------------ Config IO
    def _load_initial_configs(self) -> None:
        self._load_initial_content_config()
        self.load_participants()
        self.load_email_config()
        self.load_smtp_config()
        self.load_debug_config()

    def _load_initial_content_config(self) -> None:
        if DEFAULT_CONFIG_PATH.exists():
            self.load_config(DEFAULT_CONFIG_PATH)
        else:
            self.status_var.set("Default config not found; using template values.")

    def select_and_load_config(self) -> None:
        selected = filedialog.askopenfilename(
            title="Load Content Configuration",
            initialdir=str(DEFAULT_CONFIG_PATH.parent),
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
        )
        if selected:
            self.load_config(Path(selected))

    def load_config(self, path: Path) -> None:
        try:
            with path.open("r", encoding="utf-8") as fh:
                data = json.load(fh)
        except (OSError, json.JSONDecodeError) as exc:
            messagebox.showerror("Load Failed", f"Could not load config:\n{exc}")
            logging.exception("Failed to load config from %s", path)
            return

        self.current_config_path = path
        self.path_var.set(str(path))

        for key, var in self.vars.items():
            if key in data:
                var.set(str(data[key]))
            elif key in DEFAULT_VALUES:
                var.set(DEFAULT_VALUES[key])
            else:
                var.set("")

        self.status_var.set(f"Loaded configuration from {path}")
        self.schedule_preview_update()

    def collect_config(self) -> Optional[Dict[str, Union[float, int, str]]]:
        config: Dict[str, Union[float, int, str]] = {}

        for key, var in self.vars.items():
            value = var.get().strip()
            if not value:
                if key == "orientation":
                    messagebox.showerror("Invalid Value", "Orientation cannot be empty.")
                    return None
                continue

            if key == "font_size":
                number = self._safe_int(value)
                if number is None:
                    messagebox.showerror("Invalid Value", "Font size must be an integer.")
                    return None
                config[key] = number
            elif key == "text_y":
                position = self._safe_float(value)
                if position is None:
                    messagebox.showerror("Invalid Value", "Text baseline (text_y) must be numeric.")
                    return None
                config[key] = int(position) if position.is_integer() else position
            elif key == "long_name_threshold":
                threshold = self._safe_int(value)
                if threshold is None:
                    messagebox.showerror("Invalid Value", "Long name threshold must be an integer.")
                    return None
                config[key] = threshold
            elif key == "long_name_font_size":
                long_font = self._safe_int(value)
                if long_font is None:
                    messagebox.showerror("Invalid Value", "Long name font size must be an integer.")
                    return None
                config[key] = long_font
            elif key == "long_name_text_y":
                long_baseline = self._safe_float(value)
                if long_baseline is None:
                    messagebox.showerror("Invalid Value", "Long name baseline must be numeric.")
                    return None
                config[key] = int(long_baseline) if long_baseline.is_integer() else long_baseline
            elif key == "split_name_threshold":
                split_threshold = self._safe_int(value)
                if split_threshold is None:
                    messagebox.showerror("Invalid Value", "Split name threshold must be an integer.")
                    return None
                config[key] = split_threshold
            elif key == "split_name_line_gap":
                gap_value = self._safe_float(value)
                if gap_value is None:
                    messagebox.showerror("Invalid Value", "Split line gap must be numeric.")
                    return None
                config[key] = int(gap_value) if gap_value.is_integer() else gap_value
            elif key == "split_name_font_size":
                split_font = self._safe_float(value)
                if split_font is None:
                    messagebox.showerror("Invalid Value", "Split name font size must be numeric.")
                    return None
                config[key] = int(split_font) if split_font.is_integer() else split_font
            elif key == "split_name_text_y":
                split_baseline = self._safe_float(value)
                if split_baseline is None:
                    messagebox.showerror("Invalid Value", "Split name baseline must be numeric.")
                    return None
                config[key] = int(split_baseline) if split_baseline.is_integer() else split_baseline
            elif key == "orientation":
                config[key] = value.upper() if value.upper() in {"L", "P"} else "L"
            else:
                config[key] = value

        return config

    def save_config(self) -> None:
        if not self.current_config_path:
            self.save_config_as()
            return

        config = self.collect_config()
        if config is None:
            return

        try:
            with self.current_config_path.open("w", encoding="utf-8") as fh:
                json.dump(config, fh, indent=2)
            self.status_var.set(f"Saved configuration to {self.current_config_path}")
        except OSError as exc:
            messagebox.showerror("Save Failed", f"Could not save config:\n{exc}")
            logging.exception("Failed to save config to %s", self.current_config_path)

    def save_config_as(self) -> None:
        selected = filedialog.asksaveasfilename(
            defaultextension=".json",
            title="Save Content Configuration",
            initialdir=str(DEFAULT_CONFIG_PATH.parent),
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
        )
        if not selected:
            return

        self.current_config_path = Path(selected)
        self.path_var.set(str(self.current_config_path))
        self.save_config()

    def load_participants_via_dialog(self) -> None:
        initial_dir = str(self.participants_path.parent) if self.participants_path else "."
        selected = filedialog.askopenfilename(
            title="Load Participants CSV",
            initialdir=initial_dir,
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if selected:
            self.load_participants(Path(selected), show_message=True)

    def load_participants(self, path: Optional[Path] = None, *, show_message: bool = False) -> None:
        target_path = Path(path) if path else self.participants_path
        count = 0

        try:
            with target_path.open("r", encoding="utf-8-sig", newline="") as fh:
                reader = csv.DictReader(fh)
                columns = reader.fieldnames
                if not columns:
                    raise ValueError("CSV file does not contain a header row.")

                self._set_participants_columns(columns)
                if self.participants_tree is not None:
                    self.participants_tree.delete(*self.participants_tree.get_children())
                    for row in reader:
                        values = [(row.get(col, "") or "").strip() for col in self.participants_columns]
                        self.participants_tree.insert("", "end", values=values)
                        count += 1
        except FileNotFoundError:
            message = (
                f"Participants file not found at {target_path}. "
                "Load an existing CSV or add entries and save them."
            )
            self.participants_status_var.set(message)
            if show_message:
                messagebox.showwarning("Participants Missing", message)
            return
        except Exception as exc:
            self.participants_status_var.set(f"Could not load participants: {exc}")
            logging.exception("Failed to load participants from %s", target_path)
            if show_message:
                messagebox.showerror("Load Failed", f"Could not load participants:\n{exc}")
            return

        self.participants_path = target_path
        self.participants_path_var.set(str(self.participants_path))
        summary = f"Loaded {count} participant{'s' if count != 1 else ''} from {target_path}"
        self.participants_status_var.set(summary)
        if show_message:
            messagebox.showinfo("Participants Loaded", summary)

    def save_participants(self) -> None:
        if not self.participants_columns:
            messagebox.showerror("No Columns", "Load a CSV file before saving participants.")
            return

        if self.participants_tree is None:
            return

        rows = []
        for item in self.participants_tree.get_children():
            values = self.participants_tree.item(item, "values")
            row = {}
            for idx, column in enumerate(self.participants_columns):
                value = values[idx] if idx < len(values) else ""
                row[column] = str(value).strip()
            rows.append(row)

        try:
            with self.participants_path.open("w", encoding="utf-8", newline="") as fh:
                writer = csv.DictWriter(fh, fieldnames=self.participants_columns)
                writer.writeheader()
                writer.writerows(rows)
        except OSError as exc:
            messagebox.showerror("Save Failed", f"Could not save participants:\n{exc}")
            logging.exception("Failed to save participants to %s", self.participants_path)
            return

        summary = f"Saved {len(rows)} participant{'s' if len(rows) != 1 else ''} to {self.participants_path}"
        self.participants_status_var.set(summary)

    def save_participants_as(self) -> None:
        selected = filedialog.asksaveasfilename(
            defaultextension=".csv",
            title="Save Participants CSV",
            initialdir=str(self.participants_path.parent),
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if not selected:
            return

        self.participants_path = Path(selected)
        self.participants_path_var.set(str(self.participants_path))
        self.save_participants()

    def add_participant(self) -> None:
        if not self.participants_columns:
            messagebox.showerror("No Columns", "Load a CSV file before adding participants.")
            return

        values = []
        for column in self.participants_columns:
            var = self.participant_entry_vars.get(column)
            values.append(var.get().strip() if var else "")

        if all(not value for value in values):
            messagebox.showerror("Missing Data", "Enter at least one value before adding a participant.")
            return

        if self.participants_tree is not None:
            self.participants_tree.insert("", "end", values=values)

        for var in self.participant_entry_vars.values():
            var.set("")

        self.participants_status_var.set(
            "Participant added to the table. Remember to save the CSV to persist the change."
        )

    def remove_selected_participant(self) -> None:
        if self.participants_tree is None:
            return

        selection = self.participants_tree.selection()
        if not selection:
            messagebox.showinfo("No Selection", "Select a participant row to remove.")
            return

        for item in selection:
            self.participants_tree.delete(item)

        count = len(selection)
        self.participants_status_var.set(
            f"Removed {count} participant{'s' if count != 1 else ''}. Save the CSV to confirm the change."
        )

    def load_email_config(self, *, show_message: bool = False) -> None:
        path = EMAIL_CONFIG_PATH
        try:
            with path.open("r", encoding="utf-8") as fh:
                data = json.load(fh)
        except FileNotFoundError:
            self.email_vars["subject"].set("")
            self.email_vars["throttle_per_minute"].set("")
            if self.email_body_widget:
                self.email_body_widget.delete("1.0", tk.END)
            message = f"Email config not found at {path}. Using empty defaults."
            self.email_status_var.set(message)
            if show_message:
                messagebox.showwarning("Email Config Missing", message)
            return
        except (OSError, json.JSONDecodeError) as exc:
            self.email_status_var.set(f"Could not load email config: {exc}")
            logging.exception("Failed to load email config from %s", path)
            if show_message:
                messagebox.showerror("Load Failed", f"Could not load email config:\n{exc}")
            return

        self.email_vars["subject"].set(str(data.get("subject", "")))
        throttle_value = data.get("throttle_per_minute", "")
        self.email_vars["throttle_per_minute"].set(str(throttle_value))
        if self.email_body_widget:
            self.email_body_widget.delete("1.0", tk.END)
            self.email_body_widget.insert("1.0", data.get("body", ""))
        self.email_status_var.set(f"Loaded email settings from {path}")

    def save_email_config(self) -> None:
        throttle_raw = self.email_vars["throttle_per_minute"].get().strip()
        if not throttle_raw:
            throttle_value = 0
        else:
            throttle_value = self._safe_int(throttle_raw)
            if throttle_value is None or throttle_value < 0:
                messagebox.showerror("Invalid Value", "Throttle per minute must be a non-negative integer.")
                return

        body = ""
        if self.email_body_widget:
            body = self.email_body_widget.get("1.0", tk.END).rstrip()

        data = {
            "subject": self.email_vars["subject"].get().strip(),
            "body": body,
            "throttle_per_minute": throttle_value,
        }

        try:
            with EMAIL_CONFIG_PATH.open("w", encoding="utf-8") as fh:
                json.dump(data, fh, indent=2, ensure_ascii=False)
            self.email_status_var.set(f"Saved email settings to {EMAIL_CONFIG_PATH}")
        except OSError as exc:
            messagebox.showerror("Save Failed", f"Could not save email config:\n{exc}")
            logging.exception("Failed to save email config to %s", EMAIL_CONFIG_PATH)

    def load_smtp_config(self, *, show_message: bool = False) -> None:
        path = SMTP_CONFIG_PATH
        try:
            with path.open("r", encoding="utf-8") as fh:
                data = json.load(fh)
        except FileNotFoundError:
            for var in self.smtp_vars.values():
                var.set("")
            message = (
                f"SMTP config not found at {path}. Enter the server details provided by your email provider."
            )
            self.smtp_status_var.set(message)
            if show_message:
                messagebox.showwarning("SMTP Config Missing", message)
            return
        except (OSError, json.JSONDecodeError) as exc:
            self.smtp_status_var.set(f"Could not load SMTP config: {exc}")
            logging.exception("Failed to load SMTP config from %s", path)
            if show_message:
                messagebox.showerror("Load Failed", f"Could not load SMTP config:\n{exc}")
            return

        self.smtp_vars["smtp_server"].set(str(data.get("smtp_server", "")))
        self.smtp_vars["smtp_port"].set(str(data.get("smtp_port", "")))
        self.smtp_vars["email_sender"].set(str(data.get("email_sender", "")))
        self.smtp_vars["email_password"].set(str(data.get("email_password", "")))
        self.smtp_status_var.set(f"Loaded SMTP settings from {path}")

    def save_smtp_config(self) -> None:
        server = self.smtp_vars["smtp_server"].get().strip()
        port_raw = self.smtp_vars["smtp_port"].get().strip()
        sender = self.smtp_vars["email_sender"].get().strip()
        password = self.smtp_vars["email_password"].get()

        if not server:
            messagebox.showerror("Invalid Value", "SMTP server is required.")
            return

        port = self._safe_int(port_raw)
        if port is None or port <= 0:
            messagebox.showerror("Invalid Value", "SMTP port must be a positive integer.")
            return

        if not sender:
            messagebox.showerror("Invalid Value", "Sender email address is required.")
            return

        data = {
            "smtp_server": server,
            "smtp_port": port,
            "email_sender": sender,
            "email_password": password,
        }

        try:
            with SMTP_CONFIG_PATH.open("w", encoding="utf-8") as fh:
                json.dump(data, fh, indent=2)
            self.smtp_status_var.set(f"Saved SMTP settings to {SMTP_CONFIG_PATH}")
        except OSError as exc:
            messagebox.showerror("Save Failed", f"Could not save SMTP config:\n{exc}")
            logging.exception("Failed to save SMTP config to %s", SMTP_CONFIG_PATH)

    def load_debug_config(self, *, show_message: bool = False) -> None:
        path = DEBUG_CONFIG_PATH
        try:
            with path.open("r", encoding="utf-8") as fh:
                data = json.load(fh)
        except FileNotFoundError:
            self.debug_var.set(True)
            message = f"Debug config not found at {path}. Defaulting to Test mode."
            self.debug_status_var.set(message)
            if show_message:
                messagebox.showwarning("Debug Config Missing", message)
            return
        except (OSError, json.JSONDecodeError) as exc:
            self.debug_status_var.set(f"Could not load debug config: {exc}")
            logging.exception("Failed to load debug config from %s", path)
            if show_message:
                messagebox.showerror("Load Failed", f"Could not load debug config:\n{exc}")
            return

        raw_value = data.get("debug_mode", "Test")
        normalized = str(raw_value).strip().lower()

        if normalized in {"test", "t", "false"}:
            self.debug_var.set(True)
            label = "Test"
        elif normalized in {"full", "f", "true"}:
            self.debug_var.set(False)
            label = "Full"
        else:
            message = (
                f"Unsupported debug_mode value '{raw_value}'. "
                "Expected 'Test'/'T' for test mode or 'Full'/'F' for production."
            )
            self.debug_var.set(True)
            self.debug_status_var.set(message)
            logging.warning(message)
            if show_message:
                messagebox.showwarning("Invalid Value", message)
            return

        self.debug_status_var.set(f"Loaded debug mode ({label}) from {path}")

    def save_debug_config(self) -> None:
        mode_value = "Test" if self.debug_var.get() else "Full"
        data = {"debug_mode": mode_value}

        try:
            with DEBUG_CONFIG_PATH.open("w", encoding="utf-8") as fh:
                json.dump(data, fh, indent=2)
            self.debug_status_var.set(f"Saved debug mode ({mode_value}) to {DEBUG_CONFIG_PATH}")
        except OSError as exc:
            messagebox.showerror("Save Failed", f"Could not save debug config:\n{exc}")
            logging.exception("Failed to save debug config to %s", DEBUG_CONFIG_PATH)

    # ------------------------------------------------------------------ Preview
    def schedule_preview_update(self, *_args) -> None:
        if self.preview_after_id:
            self.root.after_cancel(self.preview_after_id)
        self.preview_after_id = self.root.after(250, self.update_preview)

    def update_preview(self) -> None:
        self.preview_after_id = None
        background_path = Path(self.vars["background_image"].get())
        image = self._load_background_image(background_path)

        draw = ImageDraw.Draw(image)
        preview_text = self.preview_name_var.get().strip() or "Sample Recipient"

        first_line = second_line = ""
        initial_split = self._should_split_preview_name(preview_text)
        if initial_split:
            first_line, second_line = self._split_preview_lines(preview_text)
        should_split = initial_split and bool(first_line) and bool(second_line)

        page_width_mm, page_height_mm = self._page_dimensions_mm()
        font_size_pt, baseline_override = self._resolve_preview_style(preview_text, should_split)
        font = self._load_preview_font(page_height_mm, image.height, font_size_pt)
        color = self._parse_color(self.vars["text_color"].get())

        baseline_mm = self._resolve_baseline_mm(font_size_pt, baseline_override)
        baseline_px = self._mm_to_pixels(baseline_mm, page_height_mm, image.height)

        try:
            if should_split:
                gap_mm = self._resolve_split_gap_mm(font_size_pt)
                gap_px = max(self._mm_to_pixels(gap_mm, page_height_mm, image.height), 0.0)
                first_baseline_px = baseline_px - gap_px
                first_top_px = self._baseline_to_top_px(font, first_baseline_px)
                second_top_px = self._baseline_to_top_px(font, baseline_px)
                if first_line:
                    first_x = self._center_text_x(draw, image.width, first_line, font)
                    draw.text((first_x, first_top_px), first_line, font=font, fill=color)
                second_x = self._center_text_x(draw, image.width, second_line, font)
                draw.text((second_x, second_top_px), second_line, font=font, fill=color)
            else:
                top_px = self._baseline_to_top_px(font, baseline_px)
                text_x = self._center_text_x(draw, image.width, preview_text, font)
                draw.text((text_x, top_px), preview_text, font=font, fill=color)
        except OSError as exc:
            logging.exception("Failed to draw text onto preview: %s", exc)

        display_image = self._downscale_for_canvas(image)
        self.preview_photo = ImageTk.PhotoImage(display_image)
        self.preview_canvas.delete("all")
        self.preview_canvas.config(
            width=self.preview_photo.width(),
            height=self.preview_photo.height(),
        )
        self.preview_canvas.create_image(0, 0, anchor="nw", image=self.preview_photo)

    def _resolve_preview_style(self, full_name: str, split: bool) -> Tuple[float, Optional[float]]:
        base_font_size = self._safe_float(self.vars["font_size"].get()) or 32.0
        default_baseline = self._safe_float(self.vars["text_y"].get())

        threshold_value = self._safe_int(self.vars["long_name_threshold"].get())
        alt_font_size = self._safe_float(self.vars["long_name_font_size"].get())
        alt_baseline = self._safe_float(self.vars["long_name_text_y"].get())

        if threshold_value is not None and self._count_name_characters(full_name) > threshold_value:
            font_size = alt_font_size or base_font_size
            baseline = alt_baseline if alt_baseline is not None else default_baseline
        else:
            font_size = base_font_size
            baseline = default_baseline

        if split:
            font_size, baseline = self._apply_split_preview_overrides(font_size, baseline)

        return font_size, baseline

    def _apply_split_preview_overrides(
        self, font_size: float, baseline: Optional[float]
    ) -> Tuple[float, Optional[float]]:
        raw_font = self.vars["split_name_font_size"].get().strip()
        if raw_font:
            override_font = self._safe_float(raw_font)
            if override_font is not None:
                font_size = override_font
            else:
                logging.warning(
                    "Invalid split_name_font_size value '%s'; preview keeps font size %s.",
                    raw_font,
                    font_size,
                )

        raw_baseline = self.vars["split_name_text_y"].get().strip()
        if raw_baseline:
            override_baseline = self._safe_float(raw_baseline)
            if override_baseline is not None:
                baseline = override_baseline
            else:
                logging.warning(
                    "Invalid split_name_text_y value '%s'; preview keeps baseline %s.",
                    raw_baseline,
                    baseline,
                )

        return font_size, baseline

    def _should_split_preview_name(self, full_name: str) -> bool:
        raw_value = self.vars["split_name_threshold"].get().strip()
        if raw_value:
            threshold = self._safe_int(raw_value)
            if threshold is None:
                logging.warning(
                    "Invalid split_name_threshold value '%s'; preview will not split recipient name.",
                    raw_value,
                )
                return False
        else:
            threshold = 24
        return self._count_name_characters(full_name) > threshold

    def _resolve_split_gap_mm(self, font_size_pt: float) -> float:
        raw_value = self.vars["split_name_line_gap"].get().strip()
        if raw_value:
            gap = self._safe_float(raw_value)
            if gap is not None:
                return gap
            logging.warning(
                "Invalid split_name_line_gap value '%s'; using fallback spacing for preview.",
                raw_value,
            )
        return self._pt_to_mm(font_size_pt or 0.0) * 0.85

    @staticmethod
    def _split_preview_lines(full_name: str) -> Tuple[str, str]:
        parts = full_name.strip().split()
        if len(parts) <= 1:
            stripped = full_name.strip()
            return stripped, ""
        first_line = " ".join(parts[:-1]).strip()
        second_line = parts[-1].strip()
        return (first_line or full_name.strip(), second_line)

    def choose_text_color(self) -> None:
        initial = self.vars["text_color"].get() or "#000000"
        color_code = colorchooser.askcolor(color=initial)
        if color_code and color_code[1]:
            self.vars["text_color"].set(color_code[1])

    # ------------------------------------------------------------------ Helpers
    def _select_path_for_key(self, key: str, title: str, filetypes: List[Tuple[str, str]]) -> None:
        initial = self.vars[key].get()
        initialdir = Path(initial).parent if initial else DEFAULT_CONFIG_PATH.parent
        selected = filedialog.askopenfilename(title=title, initialdir=initialdir, filetypes=filetypes)
        if selected:
            self.vars[key].set(selected)

    def _load_background_image(self, path: Path) -> Image.Image:
        try:
            image = Image.open(path).convert("RGBA")
        except (FileNotFoundError, OSError):
            logging.warning("Background image not found or invalid: %s", path)
            image = Image.new("RGBA", (1600, 1131), "#dddddd")
            fallback_draw = ImageDraw.Draw(image)
            fallback_draw.text((20, 20), "Background preview unavailable", fill="#444444")
        return image

    def _load_preview_font(self, page_height_mm: float, image_height: int, font_size_pt: float) -> ImageFont.ImageFont:
        font_path = Path(self.vars["font_path"].get())
        font_size_pt = font_size_pt or 32.0
        font_size_mm = self._pt_to_mm(font_size_pt)
        font_size_px = max(
            int(round(self._mm_to_pixels(font_size_mm, page_height_mm, image_height))),
            1,
        )
        try:
            return ImageFont.truetype(str(font_path), font_size_px)
        except (OSError, ValueError):
            logging.warning("Falling back to default font for preview. Invalid font path: %s", font_path)
            return ImageFont.load_default()

    def _resolve_baseline_mm(self, font_size_pt: float, baseline_override: Optional[float]) -> float:
        if baseline_override is not None:
            return baseline_override

        font_size_pt = font_size_pt or 32.0
        return self._pt_to_mm(font_size_pt)

    def _page_dimensions_mm(self) -> Tuple[float, float]:
        orientation = (self.vars["orientation"].get() or "L").upper()
        if orientation == "P":
            return 210.0, 297.0
        return 297.0, 210.0

    @staticmethod
    def _count_name_characters(value: str) -> int:
        return sum(1 for char in value if not char.isspace())

    @staticmethod
    def _mm_to_pixels(value_mm: float, total_mm: float, total_px: int) -> float:
        if total_mm <= 0:
            return float(value_mm)
        return (value_mm / total_mm) * float(total_px)

    @staticmethod
    def _baseline_to_top_px(font: ImageFont.ImageFont, baseline_px: float) -> int:
        try:
            ascent, _ = font.getmetrics()
        except AttributeError:
            ascent = getattr(font, "size", 0)
        top_px = baseline_px - ascent
        return int(round(top_px))

    def _center_text_x(self, draw: ImageDraw.ImageDraw, width: int, text: str, font: ImageFont.ImageFont) -> int:
        try:
            bbox = draw.textbbox((0, 0), text, font=font)
            text_width = bbox[2] - bbox[0]
        except AttributeError:
            text_width, _ = draw.textsize(text, font=font)
        return max(int((width - text_width) / 2), 0)

    def _downscale_for_canvas(self, image: Image.Image) -> Image.Image:
        width, height = image.size
        scale = min(PREVIEW_MAX_WIDTH / width, PREVIEW_MAX_HEIGHT / height, 1.0)
        if scale < 1.0:
            new_size = (int(width * scale), int(height * scale))
            return image.resize(new_size, Image.LANCZOS)
        return image

    @staticmethod
    def _pt_to_mm(points: float) -> float:
        return points * 25.4 / 72.0

    @staticmethod
    def _safe_float(value: str) -> Optional[float]:
        try:
            return float(value)
        except (TypeError, ValueError):
            return None

    @staticmethod
    def _safe_int(value: str) -> Optional[int]:
        try:
            return int(value)
        except (TypeError, ValueError):
            return None

    @staticmethod
    def _parse_color(value: str) -> str:
        if isinstance(value, str) and len(value) in {4, 7} and value.startswith("#"):
            return value
        return "#000000"


def main() -> None:
    logging.basicConfig(level=logging.INFO)
    root = tk.Tk()
    ConfigGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
