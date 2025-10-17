from __future__ import annotations

import json
import logging
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser
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
    "orientation": "L",
    "text_color": "#000000",
}


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

        self._attach_var_traces()

        self._build_layout()
        self._load_initial_config()
        self.schedule_preview_update()

    # ------------------------------------------------------------------ UI setup
    def _build_layout(self) -> None:
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill="both", expand=True)

        form_frame = ttk.Frame(main_frame)
        form_frame.pack(side="left", fill="y")

        preview_frame = ttk.Frame(main_frame)
        preview_frame.pack(side="right", fill="both", expand=True, padx=(15, 0))

        # Config file controls --------------------------------------------------
        ttk.Label(form_frame, text="Configuration File").grid(row=0, column=0, sticky="w")
        ttk.Label(form_frame, textvariable=self.path_var, foreground="#555555").grid(
            row=1, column=0, columnspan=3, sticky="w", pady=(0, 10)
        )

        ttk.Button(form_frame, text="Load...", command=self.select_and_load_config).grid(
            row=2, column=0, pady=(0, 10), sticky="ew"
        )
        ttk.Button(form_frame, text="Save", command=self.save_config).grid(
            row=2, column=1, pady=(0, 10), sticky="ew", padx=5
        )
        ttk.Button(form_frame, text="Save As...", command=self.save_config_as).grid(
            row=2, column=2, pady=(0, 10), sticky="ew"
        )

        # Field controls --------------------------------------------------------
        row = 3
        self._build_path_entry(
            form_frame,
            row,
            label="Background Image",
            key="background_image",
            dialog_title="Select Background Image",
            filetypes=[("Image files", "*.png;*.jpg;*.jpeg"), ("All files", "*.*")],
        )
        row += 2

        self._build_path_entry(
            form_frame,
            row,
            label="Font File",
            key="font_path",
            dialog_title="Select Font File",
            filetypes=[("TrueType font", "*.ttf;*.otf"), ("All files", "*.*")],
        )
        row += 2

        self._build_simple_entry(form_frame, row, "Font Size (pt)", "font_size")
        row += 1
        self._build_simple_entry(form_frame, row, "Text Baseline Y (mm)", "text_y")
        row += 1

        ttk.Label(form_frame, text="Orientation").grid(row=row, column=0, sticky="w", pady=(10, 0))
        orientation_box = ttk.Combobox(
            form_frame,
            textvariable=self.vars["orientation"],
            values=("L", "P"),
            state="readonly",
            width=5,
        )
        orientation_box.grid(row=row, column=1, sticky="w", pady=(10, 0))
        row += 1

        ttk.Label(form_frame, text="Text Color").grid(row=row, column=0, sticky="w", pady=(10, 0))
        color_frame = ttk.Frame(form_frame)
        color_frame.grid(row=row, column=1, columnspan=2, sticky="ew", pady=(10, 0))
        color_entry = ttk.Entry(color_frame, textvariable=self.vars["text_color"], width=12)
        color_entry.pack(side="left", fill="x", expand=True)
        ttk.Button(color_frame, text="Pick...", command=self.choose_text_color).pack(side="left", padx=(5, 0))
        row += 1

        ttk.Label(form_frame, text="Preview Name").grid(row=row, column=0, sticky="w", pady=(10, 0))
        ttk.Entry(form_frame, textvariable=self.preview_name_var).grid(
            row=row, column=1, columnspan=2, sticky="ew", pady=(10, 0)
        )
        row += 1

        ttk.Button(form_frame, text="Refresh Preview", command=self.update_preview).grid(
            row=row, column=0, columnspan=3, pady=(15, 0), sticky="ew"
        )
        row += 1

        ttk.Label(form_frame, textvariable=self.status_var, foreground="#444444", wraplength=260, justify="left").grid(
            row=row, column=0, columnspan=3, sticky="w", pady=(15, 0)
        )

        for col in range(3):
            form_frame.columnconfigure(col, weight=1)

        # Preview canvas --------------------------------------------------------
        ttk.Label(preview_frame, text="Live Preview").pack(anchor="w")
        self.preview_canvas = tk.Canvas(
            preview_frame,
            background="#e0e0e0",
            highlightthickness=1,
            highlightbackground="#cccccc",
            width=PREVIEW_MAX_WIDTH,
            height=PREVIEW_MAX_HEIGHT,
        )
        self.preview_canvas.pack(fill="both", expand=True, pady=(5, 0))

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
    def _load_initial_config(self) -> None:
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
                config[key] = int(position) if float(position).is_integer() else position
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

        page_width_mm, page_height_mm = self._page_dimensions_mm()
        font = self._load_preview_font(page_height_mm, image.height)
        color = self._parse_color(self.vars["text_color"].get())

        baseline_mm = self._resolve_baseline_mm()
        baseline_px = self._mm_to_pixels(baseline_mm, page_height_mm, image.height)
        top_px = self._baseline_to_top_px(font, baseline_px)

        text_x = self._center_text_x(draw, image.width, preview_text, font)

        try:
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

    def _load_preview_font(self, page_height_mm: float, image_height: int) -> ImageFont.ImageFont:
        font_path = Path(self.vars["font_path"].get())
        font_size_pt = self._safe_float(self.vars["font_size"].get()) or 32.0
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

    def _resolve_baseline_mm(self) -> float:
        baseline = self._safe_float(self.vars["text_y"].get())
        if baseline is not None:
            return baseline

        font_size_pt = self._safe_float(self.vars["font_size"].get()) or 32.0
        return self._pt_to_mm(font_size_pt)

    def _page_dimensions_mm(self) -> Tuple[float, float]:
        orientation = (self.vars["orientation"].get() or "L").upper()
        if orientation == "P":
            return 210.0, 297.0
        return 297.0, 210.0

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
