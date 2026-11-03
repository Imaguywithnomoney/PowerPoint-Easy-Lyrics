from __future__ import annotations

import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from app.ppt.exporter import ExportConfig, export_lyrics_to_pptx


def get_asset_path(filename: str) -> Path:
    """Resolve asset paths whether running from source or frozen PyInstaller bundle."""
    if hasattr(sys, "_MEIPASS"):
        base = Path(sys._MEIPASS)
    else:
        base = Path(__file__).resolve().parent.parent
    return base / "assets" / filename


class PPTXWizardApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("PowerPoint Easy Lyrics")
        self.root.geometry("920x720")
        self.root.minsize(820, 600)

        # Set Window Icons
        self._set_window_icon()

        # Default Settings
        self.max_stanzas_var = tk.IntVar(value=30)
        self.max_lines_var = tk.IntVar(value=1)
        self.effect_var = tk.StringVar(value="Appear/Disappear")
        self.output_var = tk.StringVar(
            value=str(Path.home() / "Desktop" / "Enter the name of pptx here.pptx")
        )

        self._build_ui()

    def _set_window_icon(self) -> None:
        ico_file = get_asset_path("icon.ico")
        if ico_file.exists():
            try:
                self.root.iconbitmap(str(ico_file))
            except Exception:
                pass

        png_file = get_asset_path("icon.png")
        if png_file.exists():
            try:
                self._app_icon = tk.PhotoImage(file=str(png_file))
                self.root.iconphoto(True, self._app_icon)
            except Exception:
                pass

    def _build_ui(self) -> None:
        frame = ttk.Frame(self.root, padding=14)
        frame.pack(fill=tk.BOTH, expand=True)

        # Top Header Bar: Logo + Heading on left, Instructions button on top right
        header = ttk.Frame(frame)
        header.pack(fill=tk.X, pady=(0, 10))

        title_container = ttk.Frame(header)
        title_container.pack(side=tk.LEFT, fill=tk.Y)

        logo_file = get_asset_path("header_logo.png")
        if logo_file.exists():
            try:
                self._header_logo = tk.PhotoImage(file=str(logo_file))
                logo_label = ttk.Label(title_container, image=self._header_logo)
                logo_label.pack(side=tk.LEFT, padx=(0, 10))
            except Exception:
                pass

        title_label = ttk.Label(
            title_container,
            text="Lyrics to PowerPoint in a Second",
            font=("Segoe UI", 16, "bold"),
        )
        title_label.pack(side=tk.LEFT, anchor=tk.W)

        btn_instructions = tk.Button(
            header,
            text="📖 Instructions",
            font=("Segoe UI", 10, "bold"),
            bg="#2563eb",
            fg="#ffffff",
            activebackground="#1d4ed8",
            activeforeground="#ffffff",
            relief=tk.FLAT,
            cursor="hand2",
            padx=14,
            pady=6,
            command=self._show_instructions,
        )
        btn_instructions.pack(side=tk.RIGHT, anchor=tk.E)

        # Separator line
        ttk.Separator(frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=(0, 10))

        # Lyrics Input Section
        ttk.Label(
            frame,
            text="Paste Song Lyrics",
            font=("Segoe UI", 13, "bold"),
        ).pack(anchor=tk.W)

        self.lyrics_input = tk.Text(frame, wrap=tk.WORD, height=18)
        self.lyrics_input.pack(fill=tk.BOTH, expand=True, pady=(8, 12))

        # Options Row
        options = ttk.Frame(frame)
        options.pack(fill=tk.X)

        ttk.Label(options, text="Max stanzas per slide").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Spinbox(
            options,
            from_=1,
            to=100,
            textvariable=self.max_stanzas_var,
            width=6,
        ).grid(row=0, column=1, padx=8)

        ttk.Label(options, text="Lines per stanza").grid(
            row=0, column=2, sticky="w"
        )
        ttk.Spinbox(
            options,
            from_=1,
            to=20,
            textvariable=self.max_lines_var,
            width=6,
        ).grid(row=0, column=3, padx=8)

        ttk.Label(options, text="Effect").grid(
            row=0, column=4, sticky="e", padx=(16, 0)
        )
        ttk.Combobox(
            options,
            textvariable=self.effect_var,
            values=["Appear/Disappear", "Fade/Fade", "Wipe/Fade"],
            width=18,
            state="readonly",
        ).grid(row=0, column=5, padx=8)

        # Output Path Row
        output_row = ttk.Frame(frame)
        output_row.pack(fill=tk.X, pady=(12, 0))
        ttk.Label(output_row, text="Output PPTX").pack(side=tk.LEFT)
        ttk.Entry(output_row, textvariable=self.output_var).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=8
        )
        ttk.Button(output_row, text="Browse", command=self._pick_output).pack(
            side=tk.LEFT
        )

        # Action Row
        action_row = ttk.Frame(frame)
        action_row.pack(fill=tk.X, pady=(14, 0))
        ttk.Button(
            action_row,
            text="Generate PPTX",
            command=self._generate,
        ).pack(side=tk.RIGHT)

    def _show_instructions(self) -> None:
        dialog = tk.Toplevel(self.root)
        dialog.title("Instructions & App Info - PowerPoint Easy Lyrics")
        dialog.geometry("740x620")
        dialog.minsize(620, 480)
        dialog.transient(self.root)
        dialog.grab_set()

        ico_file = get_asset_path("icon.ico")
        if ico_file.exists():
            try:
                dialog.iconbitmap(str(ico_file))
            except Exception:
                pass

        container = ttk.Frame(dialog, padding=16)
        container.pack(fill=tk.BOTH, expand=True)

        header_frame = ttk.Frame(container)
        header_frame.pack(fill=tk.X, pady=(0, 10))

        logo_file = get_asset_path("header_logo.png")
        if logo_file.exists():
            try:
                self._dialog_logo = tk.PhotoImage(file=str(logo_file))
                ttk.Label(header_frame, image=self._dialog_logo).pack(side=tk.LEFT, padx=(0, 10))
            except Exception:
                pass

        ttk.Label(
            header_frame,
            text="PowerPoint Easy Lyrics — Guide & Instructions",
            font=("Segoe UI", 14, "bold"),
        ).pack(side=tk.LEFT)

        # Styled Scrolled Text Area
        text_frame = ttk.Frame(container)
        text_frame.pack(fill=tk.BOTH, expand=True)

        text_area = tk.Text(
            text_frame,
            wrap=tk.WORD,
            font=("Segoe UI", 10),
            bg="#f8fafc",
            fg="#0f172a",
            padx=16,
            pady=16,
            relief=tk.SOLID,
            borderwidth=1,
        )
        scroll = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=text_area.yview)
        text_area.configure(yscrollcommand=scroll.set)

        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        text_area.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Configure stylish tags for formatting
        text_area.tag_configure("h1", font=("Segoe UI", 12, "bold"), foreground="#1e3a8a", spacing1=10, spacing3=4)
        text_area.tag_configure("body", font=("Segoe UI", 10), foreground="#334155", spacing1=2, spacing3=3)
        text_area.tag_configure("bold", font=("Segoe UI", 10, "bold"), foreground="#0f172a")
        text_area.tag_configure("bullet", font=("Segoe UI", 10), foreground="#1e293b", lmargin1=16, lmargin2=32, spacing1=3, spacing3=3)
        text_area.tag_configure("important_box", background="#fef2f2", foreground="#991b1b", font=("Segoe UI", 10, "bold"), lmargin1=14, lmargin2=14, rmargin=14, spacing1=8, spacing3=8)
        text_area.tag_configure("tip_box", background="#eff6ff", foreground="#1e40af", font=("Segoe UI", 9, "italic"), lmargin1=14, lmargin2=14, rmargin=14, spacing1=6, spacing3=6)

        # Populate styled content
        text_area.insert(tk.END, "💡 ABOUT THIS APPLICATION\n", "h1")
        text_area.insert(
            tk.END,
            "This app is designed to help people who have difficulty creating PowerPoint lyric presentations. "
            "It is especially useful for churches, worship teams, events, and other occasions where lyrics need "
            "to be displayed cleanly on a screen.\n\n",
            "body",
        )

        text_area.insert(tk.END, "⚠️ IMPORTANT NOTE ON BACKGROUNDS & FORMATTING\n", "h1")
        text_area.insert(
            tk.END,
            "IMPORTANT: You will need to manually add your background after the PowerPoint presentation has been generated. "
            "I apologize for not mentioning this earlier. Once the presentation is created, you can freely modify it to your liking, "
            "including changing the fonts, font sizes, text colors, and other formatting options.\n\n",
            "important_box",
        )

        text_area.insert(tk.END, "📋 INSTRUCTIONS & USAGE\n", "h1")
        text_area.insert(
            tk.END,
            "Hello there!\n"
            "This app allows you to create PowerPoint slides according to your preferences. You can make the background run seamlessly "
            "without interruptions or cuts. For example, if you want all the lyrics on one slide, stacked on top of each other with "
            "animations and a looping background, you can easily do that here.\n\n"
            "You can also choose to display one stanza per slide. Here is how the settings work:\n\n",
            "body",
        )

        text_area.insert(tk.END, "• Max Stanza per Slide: ", "bold")
        text_area.insert(tk.END, "Determines how many stanzas will appear on one slide. These stanzas will be stacked on top of each other with animations. When you click the Next button during slideshow, the previous stanza will disappear, and the next one will appear.\n", "bullet")

        text_area.insert(tk.END, "• Lines per Stanza: ", "bold")
        text_area.insert(tk.END, "Determines how many lines of lyrics will appear in each text box.\n\n", "bullet")

        text_area.insert(
            tk.END,
            "🌟 Recommendation: I recommend using the one-slide method (higher Max Stanza count, e.g. 30) if you want to have a continuously looping motion background without slide transitions interrupting it.\n\n",
            "tip_box",
        )

        text_area.insert(tk.END, "📌 IMPORTANT NOTES\n", "h1")
        text_area.insert(tk.END, "• Normal PowerPoint Pop-up: ", "bold")
        text_area.insert(tk.END, "If a PowerPoint window suddenly pops up during the process, this is completely normal. The app uses Microsoft PowerPoint to generate and sequence the presentation animations, so this is part of how it works.\n", "bullet")

        text_area.insert(tk.END, "• Dependencies & Safety: ", "bold")
        text_area.insert(tk.END, "Before running the app from source, make sure to check requirements.txt and install the dependencies. All dependencies (python-pptx, pywin32, etc.) are safe, legitimate open-source libraries standard in the Python ecosystem.\n", "bullet")

        text_area.config(state=tk.DISABLED)

        btn_close = ttk.Button(dialog, text="Close", command=dialog.destroy)
        btn_close.pack(pady=(12, 4))

    def _pick_output(self) -> None:
        chosen = filedialog.asksaveasfilename(
            title="Save PowerPoint Lyrics",
            defaultextension=".pptx",
            initialfile="Enter the name of pptx here.pptx",
            filetypes=[("PowerPoint Presentation", "*.pptx")],
        )
        if chosen:
            self.output_var.set(chosen)

    def _generate(self) -> None:
        lyrics = self.lyrics_input.get("1.0", tk.END).strip()
        if not lyrics:
            messagebox.showwarning(
                "PowerPoint Easy Lyrics", "Please paste lyrics first."
            )
            return

        out_path = self.output_var.get().strip()
        if not out_path:
            messagebox.showwarning(
                "PowerPoint Easy Lyrics", "Please specify an output PPTX path."
            )
            return

        effect_in, effect_out = self._map_effect(self.effect_var.get())
        cfg = ExportConfig(
            max_stanzas_per_slide=self.max_stanzas_var.get(),
            use_local_ai=False,
            animation_in=effect_in,
            animation_out=effect_out,
            max_lines_per_stanza=self.max_lines_var.get(),
            use_powerpoint_com=True,
            require_powerpoint_com=False,
        )
        try:
            output = export_lyrics_to_pptx(lyrics, out_path, cfg)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror(
                "PowerPoint Easy Lyrics", f"Generation failed:\n{exc}"
            )
            return

        messagebox.showinfo(
            "PowerPoint Easy Lyrics",
            f"Presentation created successfully:\n{output}",
        )

    @staticmethod
    def _map_effect(choice: str) -> tuple[str, str]:
        if choice == "Wipe/Fade":
            return ("wipe", "fade")
        if choice == "Appear/Disappear":
            return ("appear", "disappear")
        return ("fade", "fade")


def main() -> None:
    root = tk.Tk()
    PPTXWizardApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
