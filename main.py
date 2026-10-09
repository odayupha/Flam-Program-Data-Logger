"""
main.py – Flap Sensor Data Logger
──────────────────────────────────
Full-featured desktop data logger for the API 8810A Digital Protractor.
Built for GMF AeroAsia flap rigging inspections.

Author : Senior Python Dev – GMF AeroAsia Tooling
Date   : 2026-10-08
"""

# ── stdlib ───────────────────────────────────────────────────────────
import os
import sys
import math
import time
import threading
import winsound
from pathlib import Path
from tkinter import filedialog, messagebox, StringVar, IntVar, BooleanVar

# ── third-party ──────────────────────────────────────────────────────
import customtkinter as ctk
from PIL import Image
import openpyxl

# ── local ────────────────────────────────────────────────────────────
from serial_handler import SerialReader, list_com_ports


# ═════════════════════════════════════════════════════════════════════
#  CONSTANTS
# ═════════════════════════════════════════════════════════════════════
APP_TITLE = "Flap Sensor Data Logger — API 8810A"
APP_SIZE = "1100x780"
DEFAULT_START_ROW = 11
RESULT_COLUMN = "G"           # Column G = "RESULT"
DEFAULT_COMMAND = "READ?"
POLL_INTERVAL_MS = 1000      
LOGO_FILENAME = "logo.png"

# Colour palette (industrial-light theme accents)
CLR_BG          = "#F0F2F5"
CLR_CARD        = "#FFFFFF"
CLR_BORDER      = "#D1D9E6"
CLR_PRIMARY     = "#1A73E8"
CLR_PRIMARY_HVR = "#1557B0"
CLR_SUCCESS     = "#0F9D58"
CLR_DANGER      = "#D93025"
CLR_TEXT_DARK   = "#202124"
CLR_TEXT_MID    = "#5F6368"
CLR_TEXT_LIGHT  = "#9AA0A6"
CLR_LIVE_BG     = "#E8F0FE"
CLR_ACCENT_GOLD = "#F9AB00"


# ═════════════════════════════════════════════════════════════════════
#  HELPER – Rounded-corner card frame
# ═════════════════════════════════════════════════════════════════════
class CardFrame(ctk.CTkFrame):
    """A styled card container with subtle shadow-like border."""

    def __init__(self, master, title: str = "", **kw):
        super().__init__(
            master,
            corner_radius=12,
            fg_color=CLR_CARD,
            border_width=1,
            border_color=CLR_BORDER,
            **kw,
        )
        if title:
            lbl = ctk.CTkLabel(
                self,
                text=title,
                font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
                text_color=CLR_TEXT_MID,
                anchor="w",
            )
            lbl.pack(fill="x", padx=16, pady=(12, 4))
            sep = ctk.CTkFrame(self, height=1, fg_color=CLR_BORDER)
            sep.pack(fill="x", padx=16, pady=(0, 8))


# ═════════════════════════════════════════════════════════════════════
#  StatusDot – tiny animated status indicator
# ═════════════════════════════════════════════════════════════════════
class StatusDot(ctk.CTkLabel):
    """A small coloured dot that can pulse."""

    def __init__(self, master, **kw):
        super().__init__(master, text="●", font=ctk.CTkFont(size=16), width=20, **kw)
        self.set_state("off")

    def set_state(self, state: str):
        colour_map = {"on": CLR_SUCCESS, "off": CLR_TEXT_LIGHT, "error": CLR_DANGER}
        self.configure(text_color=colour_map.get(state, CLR_TEXT_LIGHT))


# ═════════════════════════════════════════════════════════════════════
#  MAIN APPLICATION
# ═════════════════════════════════════════════════════════════════════
class FlapSensorApp(ctk.CTk):
    """Root window for the Flap Sensor Data Logger."""

    def __init__(self):
        super().__init__()

        # ── Appearance ───────────────────────────────────────────────
        ctk.set_appearance_mode("Light")
        ctk.set_default_color_theme("blue")

        self.title(APP_TITLE)
        self.geometry(APP_SIZE)
        self.minsize(980, 700)
        self.configure(fg_color=CLR_BG)
        self._set_icon()

        # ── Application state ────────────────────────────────────────
        self._serial_reader: SerialReader | None = None
        self._live_value: float | None = None
        self._current_row: int = DEFAULT_START_ROW
        self._workbook: openpyxl.Workbook | None = None
        self._excel_path: str | None = None
        self._active_sheet_name: str | None = None
        self._save_count: int = 0

        # Tkinter variables
        self._var_port = StringVar(value="— select —")
        self._var_baud = StringVar(value="9600")
        self._var_cmd = StringVar(value=DEFAULT_COMMAND)
        self._var_sim = BooleanVar(value=False)
        self._var_sheet = StringVar(value="— load Excel first —")

        # ── Build UI ─────────────────────────────────────────────────
        self._build_header()
        self._build_body()
        self._build_footer()

        # ── Key bindings ─────────────────────────────────────────────
        self.bind("<Return>", self._on_enter_pressed)
        self.bind("<KP_Enter>", self._on_enter_pressed)

        # ── Graceful shutdown ────────────────────────────────────────
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ─────────────────────────────────────────────────────────────────
    #  ICON
    # ─────────────────────────────────────────────────────────────────
    def _set_icon(self):
        """Set the window icon if logo exists."""
        logo = self._find_logo()
        if logo:
            try:
                self.iconbitmap(default="")
            except Exception:
                pass

    # ─────────────────────────────────────────────────────────────────
    #  HEADER  (Logo + App Title)
    # ─────────────────────────────────────────────────────────────────
    def _build_header(self):
        header = ctk.CTkFrame(self, fg_color=CLR_CARD, corner_radius=0, height=72,
                              border_width=0)
        header.pack(fill="x", side="top")
        header.pack_propagate(False)

        # Inner container for centering
        inner = ctk.CTkFrame(header, fg_color="transparent")
        inner.pack(fill="both", expand=True, padx=20)

        # ── Logo (left) ─────────────────────────────────────────────
        logo_path = self._find_logo()
        if logo_path:
            try:
                pil_img = Image.open(logo_path)
                # Scale to header height keeping aspect ratio
                target_h = 50
                ratio = target_h / pil_img.height
                target_w = int(pil_img.width * ratio)
                logo_ctk = ctk.CTkImage(light_image=pil_img, size=(target_w, target_h))
                logo_label = ctk.CTkLabel(inner, image=logo_ctk, text="")
                logo_label.pack(side="left", padx=(0, 12))
            except Exception:
                self._logo_placeholder(inner, side="left")
        else:
            self._logo_placeholder(inner, side="left")

        # ── Title (centre-left) ─────────────────────────────────────
        title_frame = ctk.CTkFrame(inner, fg_color="transparent")
        title_frame.pack(side="left", fill="y", padx=4)

        ctk.CTkLabel(
            title_frame,
            text="FLAP SENSOR DATA LOGGER",
            font=ctk.CTkFont(family="Segoe UI", size=20, weight="bold"),
            text_color=CLR_TEXT_DARK,
        ).pack(anchor="w", pady=(10, 0))

        ctk.CTkLabel(
            title_frame,
            text="API 8810A Digital Protractor  ·  GMF AeroAsia",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color=CLR_TEXT_MID,
        ).pack(anchor="w")

        # ── Logo (right – duplicate placeholder) ────────────────────
        if logo_path:
            try:
                pil_img2 = Image.open(logo_path)
                target_h2 = 50
                ratio2 = target_h2 / pil_img2.height
                target_w2 = int(pil_img2.width * ratio2)
                logo_ctk2 = ctk.CTkImage(light_image=pil_img2, size=(target_w2, target_h2))
                logo_label2 = ctk.CTkLabel(inner, image=logo_ctk2, text="")
                logo_label2.pack(side="right", padx=(12, 0))
            except Exception:
                self._logo_placeholder(inner, side="right")
        else:
            self._logo_placeholder(inner, side="right")

        # Bottom accent line
        accent = ctk.CTkFrame(self, fg_color=CLR_PRIMARY, height=3, corner_radius=0)
        accent.pack(fill="x", side="top")

    @staticmethod
    def _logo_placeholder(parent, side="left"):
        """Small placeholder when logo.png is not found."""
        ph = ctk.CTkFrame(parent, width=100, height=50, fg_color=CLR_LIVE_BG,
                          corner_radius=8, border_width=1, border_color=CLR_BORDER)
        ph.pack(side=side, padx=4)
        ph.pack_propagate(False)
        ctk.CTkLabel(ph, text="LOGO", font=ctk.CTkFont(size=11),
                     text_color=CLR_TEXT_LIGHT).place(relx=0.5, rely=0.5, anchor="center")

    # ─────────────────────────────────────────────────────────────────
    #  BODY  (3 columns: Serial | Live Display | Excel)
    # ─────────────────────────────────────────────────────────────────
    def _build_body(self):
        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=16, pady=(12, 8))

        body.columnconfigure(0, weight=1, minsize=310)
        body.columnconfigure(1, weight=2, minsize=340)
        body.columnconfigure(2, weight=1, minsize=310)

        # ── LEFT COLUMN ─────────────────────────────────────────────
        left = ctk.CTkFrame(body, fg_color="transparent")
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 6))

        self._build_serial_card(left)
        self._build_excel_card(left)

        # ── CENTRE COLUMN ───────────────────────────────────────────
        centre = ctk.CTkFrame(body, fg_color="transparent")
        centre.grid(row=0, column=1, sticky="nsew", padx=6)

        self._build_live_display(centre)
        self._build_control_card(centre)

        # ── RIGHT COLUMN ────────────────────────────────────────────
        right = ctk.CTkFrame(body, fg_color="transparent")
        right.grid(row=0, column=2, sticky="nsew", padx=(6, 0))

        self._build_log_card(right)

    # ── Serial Card ──────────────────────────────────────────────────
    def _build_serial_card(self, parent):
        card = CardFrame(parent, title="🔌  SERIAL CONNECTION")
        card.pack(fill="x", pady=(0, 8))

        pad = dict(padx=16, pady=(2, 2))

        # COM Port
        ctk.CTkLabel(card, text="COM Port", font=ctk.CTkFont(size=12),
                     text_color=CLR_TEXT_MID).pack(anchor="w", **pad)
        port_frame = ctk.CTkFrame(card, fg_color="transparent")
        port_frame.pack(fill="x", padx=16, pady=(0, 6))

        ports = list_com_ports() or ["— none detected —"]
        self._combo_port = ctk.CTkOptionMenu(
            port_frame, variable=self._var_port, values=ports,
            width=160, fg_color=CLR_LIVE_BG, button_color=CLR_PRIMARY,
            text_color=CLR_TEXT_DARK, dropdown_fg_color=CLR_CARD,
        )
        self._combo_port.pack(side="left")

        btn_refresh = ctk.CTkButton(
            port_frame, text="⟳", width=36, height=32,
            fg_color=CLR_LIVE_BG, hover_color=CLR_BORDER,
            text_color=CLR_TEXT_DARK,
            command=self._refresh_ports,
        )
        btn_refresh.pack(side="left", padx=(6, 0))

        # Baud Rate
        ctk.CTkLabel(card, text="Baud Rate", font=ctk.CTkFont(size=12),
                     text_color=CLR_TEXT_MID).pack(anchor="w", **pad)
        self._combo_baud = ctk.CTkOptionMenu(
            card, variable=self._var_baud, values=["9600", "19200", "115200"],
            width=160, fg_color=CLR_LIVE_BG, button_color=CLR_PRIMARY,
            text_color=CLR_TEXT_DARK, dropdown_fg_color=CLR_CARD,
        )
        self._combo_baud.pack(anchor="w", padx=16, pady=(0, 6))

        # Command
        ctk.CTkLabel(card, text="Query Command", font=ctk.CTkFont(size=12),
                     text_color=CLR_TEXT_MID).pack(anchor="w", **pad)
        self._entry_cmd = ctk.CTkEntry(
            card, textvariable=self._var_cmd, width=260,
            border_color=CLR_BORDER, fg_color=CLR_LIVE_BG,
            text_color=CLR_TEXT_DARK, placeholder_text="READ?",
        )
        self._entry_cmd.pack(anchor="w", padx=16, pady=(0, 8))

        # Simulator checkbox
        self._chk_sim = ctk.CTkCheckBox(
            card, text="Simulator Mode (no hardware)",
            variable=self._var_sim,
            font=ctk.CTkFont(size=12), text_color=CLR_TEXT_MID,
            fg_color=CLR_PRIMARY, hover_color=CLR_PRIMARY_HVR,
            border_color=CLR_BORDER,
        )
        self._chk_sim.pack(anchor="w", padx=16, pady=(0, 8))

        # Connect / Disconnect
        btn_frame = ctk.CTkFrame(card, fg_color="transparent")
        btn_frame.pack(fill="x", padx=16, pady=(0, 14))

        self._btn_connect = ctk.CTkButton(
            btn_frame, text="▶  Connect", width=120, height=36,
            fg_color=CLR_SUCCESS, hover_color="#0C8C4E",
            font=ctk.CTkFont(size=13, weight="bold"),
            command=self._connect_serial,
        )
        self._btn_connect.pack(side="left")

        self._btn_disconnect = ctk.CTkButton(
            btn_frame, text="■  Disconnect", width=120, height=36,
            fg_color=CLR_DANGER, hover_color="#B7271D",
            font=ctk.CTkFont(size=13, weight="bold"),
            command=self._disconnect_serial, state="disabled",
        )
        self._btn_disconnect.pack(side="left", padx=(8, 0))

        # Status dot
        status_frame = ctk.CTkFrame(card, fg_color="transparent")
        status_frame.pack(fill="x", padx=16, pady=(0, 12))
        self._status_dot = StatusDot(status_frame)
        self._status_dot.pack(side="left")
        self._lbl_status = ctk.CTkLabel(
            status_frame, text="Disconnected",
            font=ctk.CTkFont(size=12), text_color=CLR_TEXT_LIGHT,
        )
        self._lbl_status.pack(side="left", padx=(2, 0))

    # ── Excel Card ───────────────────────────────────────────────────
    def _build_excel_card(self, parent):
        card = CardFrame(parent, title="📊  EXCEL FILE")
        card.pack(fill="both", expand=True, pady=(0, 0))

        pad = dict(padx=16, pady=(2, 2))

        # Browse
        ctk.CTkLabel(card, text="Master Workbook", font=ctk.CTkFont(size=12),
                     text_color=CLR_TEXT_MID).pack(anchor="w", **pad)

        browse_frame = ctk.CTkFrame(card, fg_color="transparent")
        browse_frame.pack(fill="x", padx=16, pady=(0, 6))

        self._lbl_file = ctk.CTkLabel(
            browse_frame, text="No file loaded",
            font=ctk.CTkFont(size=11), text_color=CLR_TEXT_LIGHT,
            anchor="w", width=180,
        )
        self._lbl_file.pack(side="left", fill="x", expand=True)

        self._btn_browse = ctk.CTkButton(
            browse_frame, text="📂 Browse", width=90, height=32,
            fg_color=CLR_PRIMARY, hover_color=CLR_PRIMARY_HVR,
            font=ctk.CTkFont(size=12),
            command=self._browse_excel,
        )
        self._btn_browse.pack(side="right")

        # Sheet selector
        ctk.CTkLabel(card, text="Active Sheet", font=ctk.CTkFont(size=12),
                     text_color=CLR_TEXT_MID).pack(anchor="w", **pad)
        self._combo_sheet = ctk.CTkOptionMenu(
            card, variable=self._var_sheet,
            values=["— load Excel first —"],
            width=240, fg_color=CLR_LIVE_BG, button_color=CLR_PRIMARY,
            text_color=CLR_TEXT_DARK, dropdown_fg_color=CLR_CARD,
            command=self._on_sheet_changed,
        )
        self._combo_sheet.pack(anchor="w", padx=16, pady=(0, 8))

        # Current row indicator
        row_frame = ctk.CTkFrame(card, fg_color=CLR_LIVE_BG, corner_radius=8,
                                 border_width=1, border_color=CLR_BORDER)
        row_frame.pack(fill="x", padx=16, pady=(4, 14))

        ctk.CTkLabel(
            row_frame, text="Current Test Point Row:",
            font=ctk.CTkFont(size=12), text_color=CLR_TEXT_MID,
        ).pack(side="left", padx=(12, 4), pady=8)

        self._lbl_row = ctk.CTkLabel(
            row_frame,
            text=str(DEFAULT_START_ROW),
            font=ctk.CTkFont(family="Consolas", size=18, weight="bold"),
            text_color=CLR_PRIMARY,
        )
        self._lbl_row.pack(side="left", padx=(0, 12), pady=8)

        # Save count badge
        self._lbl_saves = ctk.CTkLabel(
            row_frame,
            text="0 saved",
            font=ctk.CTkFont(size=11),
            text_color=CLR_SUCCESS,
        )
        self._lbl_saves.pack(side="right", padx=12, pady=8)

    # ── Live Display ─────────────────────────────────────────────────
    def _build_live_display(self, parent):
        card = CardFrame(parent, title="📐  LIVE ANGLE READING")
        card.pack(fill="both", expand=True, pady=(0, 8))

        # Large reading
        display_frame = ctk.CTkFrame(card, fg_color=CLR_LIVE_BG, corner_radius=16,
                                     border_width=2, border_color=CLR_PRIMARY)
        display_frame.pack(fill="both", expand=True, padx=20, pady=(8, 4))

        self._lbl_angle = ctk.CTkLabel(
            display_frame,
            text="— — —",
            font=ctk.CTkFont(family="Consolas", size=72, weight="bold"),
            text_color=CLR_TEXT_DARK,
        )
        self._lbl_angle.pack(expand=True)

        self._lbl_unit = ctk.CTkLabel(
            display_frame,
            text="degrees (°)",
            font=ctk.CTkFont(family="Segoe UI", size=16),
            text_color=CLR_TEXT_MID,
        )
        self._lbl_unit.pack(pady=(0, 12))

        # Timestamp of last update
        self._lbl_ts = ctk.CTkLabel(
            card,
            text="Last update: —",
            font=ctk.CTkFont(size=11),
            text_color=CLR_TEXT_LIGHT,
        )
        self._lbl_ts.pack(pady=(0, 10))

    # ── Control Card ─────────────────────────────────────────────────
    def _build_control_card(self, parent):
        card = CardFrame(parent, title="⌨  CONTROLS")
        card.pack(fill="x", pady=(0, 0))

        # Instruction
        instr = ctk.CTkLabel(
            card,
            text="Tekan  ENTER  untuk menyimpan data ke Excel",
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            text_color=CLR_PRIMARY,
        )
        instr.pack(pady=(6, 4))

        hint = ctk.CTkLabel(
            card,
            text=f"Nilai akan ditulis ke kolom {RESULT_COLUMN} pada baris aktif",
            font=ctk.CTkFont(size=12),
            text_color=CLR_TEXT_MID,
        )
        hint.pack(pady=(0, 8))

        # Buttons row
        btn_row = ctk.CTkFrame(card, fg_color="transparent")
        btn_row.pack(pady=(0, 14))

        self._btn_save_enter = ctk.CTkButton(
            btn_row, text="⏎  Save Data (Enter)", width=170, height=40,
            fg_color=CLR_PRIMARY, hover_color=CLR_PRIMARY_HVR,
            font=ctk.CTkFont(size=13, weight="bold"),
            command=lambda: self._on_enter_pressed(None),
        )
        self._btn_save_enter.pack(side="left", padx=4)

        self._btn_export = ctk.CTkButton(
            btn_row, text="💾  Export / Save As", width=170, height=40,
            fg_color=CLR_ACCENT_GOLD, hover_color="#E09D00",
            text_color=CLR_TEXT_DARK,
            font=ctk.CTkFont(size=13, weight="bold"),
            command=self._export_saveas,
        )
        self._btn_export.pack(side="left", padx=4)

    # ── Log Card (right column) ──────────────────────────────────────
    def _build_log_card(self, parent):
        card = CardFrame(parent, title="📝  EVENT LOG")
        card.pack(fill="both", expand=True)

        self._log_text = ctk.CTkTextbox(
            card, font=ctk.CTkFont(family="Consolas", size=11),
            fg_color=CLR_LIVE_BG, text_color=CLR_TEXT_DARK,
            border_width=1, border_color=CLR_BORDER,
            wrap="word", state="disabled", corner_radius=8,
        )
        self._log_text.pack(fill="both", expand=True, padx=16, pady=(4, 14))

    # ─────────────────────────────────────────────────────────────────
    #  FOOTER / STATUS BAR
    # ─────────────────────────────────────────────────────────────────
    def _build_footer(self):
        footer = ctk.CTkFrame(self, fg_color=CLR_CARD, corner_radius=0, height=30)
        footer.pack(fill="x", side="bottom")
        footer.pack_propagate(False)

        ctk.CTkLabel(
            footer,
            text="Flap Sensor Data Logger v1.0  |  GMF AeroAsia  |  API 8810A",
            font=ctk.CTkFont(size=11), text_color=CLR_TEXT_LIGHT,
        ).pack(side="left", padx=16)

        self._lbl_footer_status = ctk.CTkLabel(
            footer, text="Ready",
            font=ctk.CTkFont(size=11), text_color=CLR_TEXT_LIGHT,
        )
        self._lbl_footer_status.pack(side="right", padx=16)

    # ═════════════════════════════════════════════════════════════════
    #  CALLBACKS — Serial
    # ═════════════════════════════════════════════════════════════════
    def _refresh_ports(self):
        ports = list_com_ports() or ["— none detected —"]
        self._combo_port.configure(values=ports)
        if ports:
            self._var_port.set(ports[0])
        self._log("Ports refreshed: " + ", ".join(ports))

    def _connect_serial(self):
        port = self._var_port.get()
        baud = int(self._var_baud.get())
        cmd = self._var_cmd.get()  # sanitised inside SerialReader
        sim = self._var_sim.get()

        if not sim and port.startswith("—"):
            messagebox.showwarning("No Port", "Pilih COM Port terlebih dahulu\natau aktifkan Simulator Mode.")
            return

        self._serial_reader = SerialReader(
            port=port,
            baudrate=baud,
            command=cmd,
            interval=POLL_INTERVAL_MS / 1000.0,
            on_data=self._on_serial_data,
            on_error=self._on_serial_error,
            on_log=self._on_serial_log,
            simulator=sim,
        )
        self._serial_reader.connect()

        # UI state
        self._btn_connect.configure(state="disabled")
        self._btn_disconnect.configure(state="normal")
        self._status_dot.set_state("on")
        mode = "SIMULATOR" if sim else port
        self._lbl_status.configure(text=f"Connected ({mode})", text_color=CLR_SUCCESS)
        self._lbl_footer_status.configure(text=f"Connected — {mode} @ {baud} baud")
        self._log(f"Connected: {mode} @ {baud} baud")
        self._log(f"Query command: '{cmd}' (terminator \\r\\n added automatically)")

    def _disconnect_serial(self):
        if self._serial_reader:
            self._serial_reader.disconnect()
            self._serial_reader = None

        self._btn_connect.configure(state="normal")
        self._btn_disconnect.configure(state="disabled")
        self._status_dot.set_state("off")
        self._lbl_status.configure(text="Disconnected", text_color=CLR_TEXT_LIGHT)
        self._lbl_footer_status.configure(text="Disconnected")
        self._lbl_angle.configure(text="— — —")
        self._live_value = None
        self._log("Disconnected")

    # ── Serial data callback (called from worker thread!) ────────────
    def _on_serial_data(self, value: float):
        """Thread-safe update via `after()`. Schedules UI refresh on main thread."""
        self.after(0, self._update_live_display, value)

    def _on_serial_error(self, msg: str):
        self.after(0, self._log, f"⚠ {msg}")

    def _on_serial_log(self, msg: str):
        """Verbose serial debug messages piped to event log."""
        self.after(0, self._log, f"🔧 {msg}")

    def _update_live_display(self, value: float):
        """Update the big angle readout (main thread)."""
        self._live_value = value
        # Format: fixed 2 decimals, at least 6 chars wide for consistent look
        display = f"{value:>8.2f}"
        self._lbl_angle.configure(text=display.strip())
        now = time.strftime("%H:%M:%S")
        self._lbl_ts.configure(text=f"Last update: {now}")

    # ═════════════════════════════════════════════════════════════════
    #  CALLBACKS — Excel
    # ═════════════════════════════════════════════════════════════════
    def _browse_excel(self):
        path = filedialog.askopenfilename(
            title="Select Master Excel File",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if not path:
            return

        try:
            wb = openpyxl.load_workbook(path)
        except Exception as exc:
            messagebox.showerror("Error", f"Gagal membuka Excel:\n{exc}")
            return

        self._workbook = wb
        self._excel_path = path
        self._current_row = DEFAULT_START_ROW
        self._save_count = 0

        # Update UI
        short = os.path.basename(path)
        self._lbl_file.configure(text=short, text_color=CLR_TEXT_DARK)
        self._lbl_row.configure(text=str(self._current_row))
        self._lbl_saves.configure(text="0 saved")

        sheets = wb.sheetnames
        self._combo_sheet.configure(values=sheets)
        if sheets:
            self._var_sheet.set(sheets[0])
            self._active_sheet_name = sheets[0]

        self._log(f"Loaded: {short}  (sheets: {', '.join(sheets)})")

    def _on_sheet_changed(self, sheet_name: str):
        self._active_sheet_name = sheet_name
        self._current_row = DEFAULT_START_ROW
        self._save_count = 0
        self._lbl_row.configure(text=str(self._current_row))
        self._lbl_saves.configure(text="0 saved")
        self._log(f"Active sheet → {sheet_name}")

    # ═════════════════════════════════════════════════════════════════
    #  ENTER — Save current angle to Excel
    # ═════════════════════════════════════════════════════════════════
    def _on_enter_pressed(self, event):
        """Write the current live value to the Excel sheet at G{row}."""
        # Validations
        if self._live_value is None:
            self._log("⚠ No live data — connect to sensor first.")
            self._flash_display(CLR_DANGER)
            return

        if self._workbook is None or self._excel_path is None:
            self._log("⚠ No Excel file loaded — browse for a file first.")
            self._flash_display(CLR_DANGER)
            return

        if self._active_sheet_name is None:
            self._log("⚠ No active sheet selected.")
            return

        ws = self._workbook[self._active_sheet_name]
        cell_ref = f"{RESULT_COLUMN}{self._current_row}"
        value = self._live_value

        try:
            ws[cell_ref] = value
            self._workbook.save(self._excel_path)
        except PermissionError:
            messagebox.showerror(
                "File Locked",
                f"File Excel sedang dibuka program lain.\n"
                f"Tutup file terlebih dahulu, lalu coba lagi.",
            )
            self._log(f"✗ PermissionError saving {cell_ref}")
            return
        except Exception as exc:
            messagebox.showerror("Save Error", f"Gagal menyimpan:\n{exc}")
            self._log(f"✗ Error: {exc}")
            return

        # Success!
        self._save_count += 1
        self._log(f"✓ Saved {value:.2f} → {cell_ref}  (#{self._save_count})")
        self._lbl_saves.configure(text=f"{self._save_count} saved")
        self._flash_display(CLR_SUCCESS)

        # Beep confirmation
        threading.Thread(
            target=lambda: winsound.Beep(1000, 200), daemon=True
        ).start()

        # Advance row
        self._current_row += 1
        self._lbl_row.configure(text=str(self._current_row))

    def _flash_display(self, colour: str):
        """Quick flash the live display border to give visual feedback."""
        original = CLR_PRIMARY
        try:
            parent = self._lbl_angle.master  # display_frame
            parent.configure(border_color=colour)
            self.after(350, lambda: parent.configure(border_color=original))
        except Exception:
            pass

    # ═════════════════════════════════════════════════════════════════
    #  EXPORT / SAVE AS
    # ═════════════════════════════════════════════════════════════════
    def _export_saveas(self):
        if self._workbook is None:
            messagebox.showinfo("Info", "Load file Excel terlebih dahulu.")
            return

        dest = filedialog.asksaveasfilename(
            title="Export / Save As",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile="RESULT_export.xlsx",
        )
        if not dest:
            return

        try:
            self._workbook.save(dest)
            self._log(f"📦 Exported → {os.path.basename(dest)}")
            messagebox.showinfo("Export Berhasil", f"File tersimpan:\n{dest}")
        except Exception as exc:
            messagebox.showerror("Export Error", str(exc))

    # ═════════════════════════════════════════════════════════════════
    #  LOGGING
    # ═════════════════════════════════════════════════════════════════
    def _log(self, message: str):
        """Append a timestamped message to the event log textbox."""
        ts = time.strftime("%H:%M:%S")
        line = f"[{ts}]  {message}\n"
        self._log_text.configure(state="normal")
        self._log_text.insert("end", line)
        self._log_text.see("end")
        self._log_text.configure(state="disabled")

    # ═════════════════════════════════════════════════════════════════
    #  HELPERS
    # ═════════════════════════════════════════════════════════════════
    def _find_logo(self) -> str | None:
        """Search for logo.png next to the executable / script."""
        candidates = [
            Path(__file__).parent / LOGO_FILENAME,
            Path(sys.executable).parent / LOGO_FILENAME,
            Path.cwd() / LOGO_FILENAME,
        ]
        for p in candidates:
            if p.exists():
                return str(p)
        return None

    def _on_close(self):
        """Clean shutdown."""
        if self._serial_reader and self._serial_reader.is_connected:
            self._serial_reader.disconnect()
        self.destroy()


# ═════════════════════════════════════════════════════════════════════
#  ENTRY POINT
# ═════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    app = FlapSensorApp()
    app.mainloop()
