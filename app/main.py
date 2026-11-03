from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from app.ppt.exporter import ExportConfig, export_lyrics_to_pptx


INSTRUCTIONS_TEXT = """INFO:
This app is designed to help people who have difficulty creating PowerPoint lyric presentations. It is especially useful for churches, events, and other occasions where lyrics need to be displayed on a screen.

============================================================
INSTRUCTIONS:
============================================================
Hello there!

This app allows you to create PowerPoint slides according to your preferences. You can make the background run seamlessly without interruptions or cuts. For example, if you want all the lyrics on one slide, stacked on top of each other with animations and a looping background, you can easily do that here.

You can also choose to display one stanza per slide. Here is how the settings work:

• Max Stanza per Slide: Determines how many stanzas will appear on one slide. These stanzas will be stacked on top of each other with animations. When you click the Next button, the previous stanza will disappear, and the next one will appear.

• Lines per Stanza: Determines how many lines of lyrics will appear in each text box.

Recommendation:
I recommend using the one-slide method if you want to have a continuously looping background without interruptions.

============================================================
IMPORTANT NOTES:
============================================================
• If a PowerPoint window suddenly pops up during the process, this is completely normal. The app uses PowerPoint to generate the presentation, so this is part of how it works.

• Before running the app, make sure to check the requirements.txt file. You need to install the required dependencies using Command Prompt, Terminal, or PowerShell.

• The dependencies are safe and legitimate. You can also search for them online if you want to learn more about them.
"""


class PPTXWizardApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("PowerPoint Easy Lyrics")
        self.root.geometry("920x700")
        self.root.minsize(800, 580)

        # Default Settings requested:
        # Max Stanza per Slide = 30
        # Lines per Stanza = 1
        # Default filename = Enter the name of pptx here.pptx
        self.max_stanzas_var = tk.IntVar(value=30)
        self.max_lines_var = tk.IntVar(value=1)
        self.effect_var = tk.StringVar(value="Appear/Disappear")
        self.output_var = tk.StringVar(
            value=str(Path.home() / "Desktop" / "Enter the name of pptx here.pptx")
        )

        self._build_ui()

    def _build_ui(self) -> None:
        frame = ttk.Frame(self.root, padding=14)
        frame.pack(fill=tk.BOTH, expand=True)

        # Top Header Bar: Heading on left, Instructions button on top right
        header = ttk.Frame(frame)
        header.pack(fill=tk.X, pady=(0, 10))

        title_label = ttk.Label(
            header,
            text="Lyrics to PowerPoint in a Second",
            font=("Segoe UI", 16, "bold"),
        )
        title_label.pack(side=tk.LEFT, anchor=tk.W)

        btn_instructions = tk.Button(
            header,
            text="Instructions",
            font=("Segoe UI", 10, "bold"),
            bg="#2563eb",
            fg="#ffffff",
            activebackground="#1d4ed8",
            activeforeground="#ffffff",
            relief=tk.FLAT,
            cursor="hand2",
            padx=14,
            pady=5,
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
        dialog.geometry("680x520")
        dialog.minsize(550, 400)
        dialog.transient(self.root)
        dialog.grab_set()

        container = ttk.Frame(dialog, padding=16)
        container.pack(fill=tk.BOTH, expand=True)

        header_lbl = ttk.Label(
            container,
            text="Instructions & App Information",
            font=("Segoe UI", 14, "bold"),
        )
        header_lbl.pack(anchor=tk.W, pady=(0, 10))

        text_area = tk.Text(
            container,
            wrap=tk.WORD,
            font=("Segoe UI", 10),
            bg="#f8fafc",
            fg="#0f172a",
            padx=12,
            pady=12,
            relief=tk.SOLID,
            borderwidth=1,
        )
        scroll = ttk.Scrollbar(container, orient=tk.VERTICAL, command=text_area.yview)
        text_area.configure(yscrollcommand=scroll.set)

        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        text_area.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        text_area.insert(tk.END, INSTRUCTIONS_TEXT)
        text_area.config(state=tk.DISABLED)

        btn_close = ttk.Button(dialog, text="Close", command=dialog.destroy)
        btn_close.pack(pady=(0, 12))

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
