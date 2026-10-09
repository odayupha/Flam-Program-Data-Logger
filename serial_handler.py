"""
serial_handler.py
─────────────────
Background serial communication handler for the API 8810A Angle Position
Indicator (North Atlantic Industries).  Runs in a dedicated daemon thread
so the CustomTkinter GUI never blocks.

Key protocol details (NAI 8810A over USB / Virtual COM / FTDI):
  • Every command MUST be terminated with \r\n (CR+LF).
  • The instrument replies with an ASCII float (e.g. "2.427\r\n").
  • The input buffer must be flushed BEFORE each query to avoid
    reading stale data from a previous cycle.
  • Supported query commands (SCPI-style):
        MEASure:ANGle?     – read angle (degrees)
        FETC:ANG? CH1      – fetch angle, channel 1
        FETC:ANG? CH2      – fetch angle, channel 2
        FETCH?  /  READ?   – read active channel
  • Read timeout should be 1.0 – 2.0 s to tolerate slow responses.

Author : Senior Python Dev – GMF AeroAsia Tooling
Rev    : 2.0  (2026-10-09) – full refactor for reliable NAI 8810A comms
"""

import re
import threading
import time
import random
from typing import Callable, Optional

import serial
import serial.tools.list_ports


# Compiled regex: matches an optional minus sign, digits, optional
# decimal point + more digits.  Covers "2.427", "-0.15", "274.926000".
_RE_FLOAT = re.compile(r"-?\d+\.?\d*")


# ──────────────────────────────────────────────────────────────────────
#  Utility: list available COM ports
# ──────────────────────────────────────────────────────────────────────
def list_com_ports() -> list[str]:
    """Return a list of available COM port names (e.g. ['COM3', 'COM5'])."""
    ports = serial.tools.list_ports.comports()
    return [p.device for p in sorted(ports)]


# ──────────────────────────────────────────────────────────────────────
#  SerialReader – threaded reader for the API 8810A
# ──────────────────────────────────────────────────────────────────────
class SerialReader:
    """
    Manages a serial connection to the API 8810A.

    Parameters
    ----------
    port : str          – COM port name (e.g. "COM3").
    baudrate : int      – Baud rate (typically 9600).
    command : str       – Primary query command (e.g. "MEASure:ANGle?").
    interval : float    – Seconds between queries (default 0.5).
    on_data : callable  – Callback ``on_data(value: float)`` invoked
                          from the worker thread; the GUI should schedule
                          the UI update via ``root.after()``.
    on_error : callable – Callback ``on_error(msg: str)`` for warnings.
    on_log : callable   – Callback ``on_log(msg: str)`` for verbose logs.
    simulator : bool    – If True, generate fake angle readings.
    """

    def __init__(
        self,
        port: str,
        baudrate: int,
        command: str,
        interval: float = 0.5,
        on_data: Optional[Callable[[float], None]] = None,
        on_error: Optional[Callable[[str], None]] = None,
        on_log: Optional[Callable[[str], None]] = None,
        simulator: bool = False,
    ):
        self.port = port
        self.baudrate = baudrate
        self.command = self._sanitise_command(command)
        self.interval = interval
        self.on_data = on_data
        self.on_error = on_error
        self.on_log = on_log
        self.simulator = simulator

        self._serial: Optional[serial.Serial] = None
        self._thread: Optional[threading.Thread] = None
        self._running = threading.Event()
        self._lock = threading.Lock()

    # ── connect ──────────────────────────────────────────────────────
    def connect(self) -> None:
        """Open the serial port (or start the simulator) and begin polling."""
        if self._running.is_set():
            return  # already running

        if not self.simulator:
            try:
                self._serial = serial.Serial(
                    port=self.port,
                    baudrate=self.baudrate,
                    bytesize=serial.EIGHTBITS,
                    parity=serial.PARITY_NONE,
                    stopbits=serial.STOPBITS_ONE,
                    timeout=2.0,           # read timeout 2.0 s
                    write_timeout=1.0,     # write timeout 1.0 s
                    xonxoff=False,
                    rtscts=False,
                    dsrdtr=False,
                )
                
                # CRITICAL: Activate DTR & RTS to wake up FTDI chip / NAI 8810A
                self._serial.dtr = True
                self._serial.rts = True
                
                # Allow the FTDI / USB-serial adapter to settle after activating signals
                time.sleep(0.2)

                # Drain any boot-up garbage from the device
                self._serial.reset_input_buffer()
                self._serial.reset_output_buffer()

                self._log_info(
                    f"Port {self.port} opened: {self.baudrate} 8N1, "
                    f"timeout=1.5 s"
                )
            except serial.SerialException as exc:
                if self.on_error:
                    self.on_error(f"Serial open error: {exc}")
                return

        self._running.set()
        self._thread = threading.Thread(target=self._poll_loop, daemon=True)
        self._thread.start()

    # ── disconnect ───────────────────────────────────────────────────
    def disconnect(self) -> None:
        """Stop polling and close the serial port."""
        self._running.clear()
        if self._thread is not None:
            self._thread.join(timeout=3)
            self._thread = None

        with self._lock:
            if self._serial is not None and self._serial.is_open:
                try:
                    # SAFE CLOSE: reset buffers before closing to prevent PermissionError on reconnect
                    self._serial.reset_input_buffer()
                    self._serial.reset_output_buffer()
                    self._serial.close()
                except Exception:
                    pass
                self._serial = None

    # ══════════════════════════════════════════════════════════════════
    #  INTERNAL POLLING LOOP
    # ══════════════════════════════════════════════════════════════════
    def _poll_loop(self) -> None:
        """Continuously send the query and parse responses."""
        sim_angle = random.uniform(0.0, 360.0)

        while self._running.is_set():
            try:
                if self.simulator:
                    value = self._simulate(sim_angle)
                    sim_angle = value
                else:
                    value = self._query_device()

                if value is not None:
                    if self.on_data:
                        self.on_data(value)
                else:
                    # No valid data received
                    self._log_info("Empty response (timeout) or invalid format")

            except serial.SerialException as exc:
                # Port-level error (cable unplugged, etc.)
                if self.on_error:
                    self.on_error(f"Serial I/O error: {exc}")

            except Exception as exc:
                if self.on_error:
                    self.on_error(f"Unexpected: {exc}")

            # ── Interruptible sleep (check every 50 ms) ─────────────
            waited = 0.0
            while waited < self.interval and self._running.is_set():
                time.sleep(0.05)
                waited += 0.05

    # ══════════════════════════════════════════════════════════════════
    #  QUERY THE REAL DEVICE
    # ══════════════════════════════════════════════════════════════════
    def _query_device(self) -> Optional[float]:
        """
        Send the active command to the API 8810A and parse the reply.

        Protocol sequence:
          1. Flush input buffer  → discard stale / buffered data
          2. Write command + \\r\\n terminator
          3. readline()          → read until \\n or timeout
          4. Parse numeric float from response
        """
        with self._lock:
            if self._serial is None or not self._serial.is_open:
                return None

            try:
                # ① Flush stale data from the receive buffer
                self._serial.reset_input_buffer()

                # ② Build and send the command with proper terminator
                cmd_clean = self._sanitise_command(self.command)
                wire_bytes = f"{cmd_clean}\r\n".encode("ascii", errors="ignore")
                self._serial.write(wire_bytes)

                # ③ Read the response line (blocks up to timeout)
                raw_bytes = self._serial.readline()

            except serial.SerialException:
                raise  # let poll_loop handle it
            except Exception as exc:
                self._log_info(f"Write/Read error: {exc}")
                return None

        # ── Decode and strip outside the lock ────────────────────────
        if not raw_bytes:
            self._log_info("Empty response (timeout)")
            return None

        raw_str = raw_bytes.decode("ascii", errors="ignore").strip()

        if not raw_str:
            self._log_info("Blank response after strip")
            return None

        # ④ Parse the numeric value
        value = self._parse_angle(raw_str)

        if value is None:
            self._log_info(f"Unparseable response: '{raw_str}'")

        return value

    # ══════════════════════════════════════════════════════════════════
    #  PARSING
    # ══════════════════════════════════════════════════════════════════
    @staticmethod
    def _parse_angle(text: str) -> Optional[float]:
        """
        Safely extract a float from the device response.

        The API 8810A typically replies with a plain ASCII float:
            "2.427"   "274.926000"   "-0.05"

        Returns the value rounded to 2 decimal places, or None.
        """
        # First try: direct float conversion (fastest path)
        try:
            return round(float(text), 2)
        except (ValueError, TypeError):
            pass

        # Second try: regex extraction (handles prefixed/suffixed data)
        match = _RE_FLOAT.search(text)
        if match:
            try:
                return round(float(match.group()), 2)
            except (ValueError, TypeError):
                pass

        return None

    # ══════════════════════════════════════════════════════════════════
    #  SIMULATOR
    # ══════════════════════════════════════════════════════════════════
    @staticmethod
    def _simulate(current: float) -> float:
        """Generate a slowly drifting simulated angle value."""
        drift = random.gauss(0, 0.15)
        noisy = (current + drift) % 360.0
        return round(noisy, 2)

    # ══════════════════════════════════════════════════════════════════
    #  HELPERS
    # ══════════════════════════════════════════════════════════════════
    @staticmethod
    def _sanitise_command(cmd: str) -> str:
        """
        Clean up a command string from UI input.

        - Strips leading/trailing whitespace
        - Removes literal '\\r' / '\\n' escape sequences that may have
          been typed into the textbox (the terminator is added at send time)
        """
        if not cmd:
            return ""
        cleaned = cmd.strip()
        # Remove any literal backslash-r / backslash-n typed by the user
        cleaned = cleaned.replace("\\r", "").replace("\\n", "")
        # Also remove actual CR/LF if someone pasted them
        cleaned = cleaned.replace("\r", "").replace("\n", "")
        return cleaned

    def _log_info(self, msg: str) -> None:
        """Send a verbose log message via the on_log callback."""
        if self.on_log:
            self.on_log(msg)

    # ── properties ───────────────────────────────────────────────────
    @property
    def is_connected(self) -> bool:
        return self._running.is_set()
