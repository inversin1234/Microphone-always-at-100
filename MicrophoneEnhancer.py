import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
from typing import Dict, List, Optional

import comtypes
from pycaw.pycaw import AudioUtilities, IMMDeviceEnumerator, EDataFlow
from pycaw.constants import CLSID_MMDeviceEnumerator, DEVICE_STATE


def get_audio_devices(direction: str = "in", state: int = DEVICE_STATE.ACTIVE.value) -> List:
    """Return a list of audio endpoint devices for the requested direction."""
    devices: List = []
    flow = EDataFlow.eCapture.value if direction == "in" else EDataFlow.eRender.value
    device_enumerator = comtypes.CoCreateInstance(
        CLSID_MMDeviceEnumerator,
        IMMDeviceEnumerator,
        comtypes.CLSCTX_INPROC_SERVER,
    )
    if device_enumerator is None:
        return devices
    collection = device_enumerator.EnumAudioEndpoints(flow, state)
    if collection is None:
        return devices
    for index in range(collection.GetCount()):
        imm_device = collection.Item(index)
        devices.append(AudioUtilities.CreateDevice(imm_device))
    return devices


class MicrophoneApp:
    """Interactive Tk application that keeps the microphone at the target volume."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Microphone volume control")
        self.root.geometry("420x520")
        self.root.minsize(420, 520)
        self.root.configure(bg="#f5f7fb")
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        self.style = ttk.Style(self.root)
        self._configure_styles()

        self.devices: List = []
        self.device_map: Dict[str, object] = {}
        self.current_device = None
        self.endpoint_volume = None

        self.monitoring = False
        self.after_id: Optional[str] = None
        self.frequency_seconds = 5

        self.device_name_var = tk.StringVar()
        self.frequency_var = tk.StringVar(value="5")
        self.target_volume_var = tk.IntVar(value=100)
        self.status_message_var = tk.StringVar(value="Select a microphone to begin.")
        self.device_details_var = tk.StringVar(value="No device selected.")
        self.last_applied_var = tk.StringVar(value="Last applied: —")
        self.next_check_var = tk.StringVar(value="Next check in: —")
        self.target_display_var = tk.StringVar(value="Target: 100%")

        self._build_layout()
        self.refresh_devices(initial=True)

    def _configure_styles(self) -> None:
        try:
            self.style.theme_use("clam")
        except tk.TclError:
            pass
        palette = {
            "background": "#f5f7fb",
            "card": "#ffffff",
            "accent": "#2962ff",
            "accent_hover": "#1c43b5",
            "danger": "#d32f2f",
            "text": "#1a1c2d",
            "muted": "#5c6270",
        }
        self.style.configure("TFrame", background=palette["background"])
        self.style.configure("Card.TFrame", background=palette["card"], relief="flat")
        self.style.configure(
            "Title.TLabel",
            background=palette["card"],
            foreground=palette["text"],
            font=("Segoe UI", 16, "bold"),
        )
        self.style.configure(
            "Body.TLabel",
            background=palette["card"],
            foreground=palette["muted"],
            font=("Segoe UI", 10),
            wraplength=360,
        )
        self.style.configure(
            "Section.TLabel",
            background=palette["card"],
            foreground=palette["text"],
            font=("Segoe UI", 11, "bold"),
        )
        self.style.configure(
            "TButton",
            font=("Segoe UI", 10, "bold"),
            padding=(14, 8),
        )
        self.style.configure(
            "Accent.TButton",
            background=palette["accent"],
            foreground="#ffffff",
        )
        self.style.map(
            "Accent.TButton",
            background=[("active", palette["accent_hover"]), ("disabled", "#aeb8fb")],
            foreground=[("disabled", "#f5f7fb")],
        )
        self.style.configure(
            "Danger.TButton",
            background=palette["danger"],
            foreground="#ffffff",
        )
        self.style.map(
            "Danger.TButton",
            background=[("active", "#9a2424"), ("disabled", "#f0b8b8")],
            foreground=[("disabled", "#f5f7fb")],
        )
        self.style.configure(
            "Info.TLabel",
            background=palette["card"],
            foreground=palette["muted"],
            font=("Segoe UI", 9),
        )
        self.palette = palette

    def _build_layout(self) -> None:
        main = ttk.Frame(self.root, style="Card.TFrame", padding=24)
        main.pack(fill=tk.BOTH, expand=True, padx=24, pady=24)

        header = ttk.Frame(main, style="Card.TFrame")
        header.pack(fill=tk.X)

        ttk.Label(header, text="Microphone guardian", style="Title.TLabel").pack(anchor=tk.W)
        ttk.Label(
            header,
            text=(
                "Keep your input level locked where you want it. Refresh the device list,"
                " set a custom enforcement interval, and monitor live status updates."
            ),
            style="Body.TLabel",
        ).pack(anchor=tk.W, pady=(6, 18))

        device_section = ttk.Frame(main, style="Card.TFrame")
        device_section.pack(fill=tk.X, pady=(0, 18))

        ttk.Label(device_section, text="Recording device", style="Section.TLabel").pack(anchor=tk.W)

        combo_row = ttk.Frame(device_section, style="Card.TFrame")
        combo_row.pack(fill=tk.X, pady=(8, 0))

        self.device_combo = ttk.Combobox(combo_row, textvariable=self.device_name_var, state="readonly", width=34)
        self.device_combo.pack(side=tk.LEFT, expand=True, fill=tk.X)
        self.device_combo.bind("<<ComboboxSelected>>", lambda _event: self._on_device_selected())

        ttk.Button(
            combo_row,
            text="Refresh",
            command=self.refresh_devices,
            style="TButton",
        ).pack(side=tk.LEFT, padx=(8, 0))

        ttk.Label(device_section, textvariable=self.device_details_var, style="Info.TLabel").pack(
            anchor=tk.W, pady=(6, 0)
        )

        controls_section = ttk.Frame(main, style="Card.TFrame")
        controls_section.pack(fill=tk.X, pady=(0, 18))

        ttk.Label(controls_section, text="Enforcement settings", style="Section.TLabel").pack(anchor=tk.W)

        freq_frame = ttk.Frame(controls_section, style="Card.TFrame")
        freq_frame.pack(fill=tk.X, pady=(10, 0))

        ttk.Label(freq_frame, text="Check every", style="Body.TLabel").pack(side=tk.LEFT)
        freq_entry = ttk.Entry(freq_frame, textvariable=self.frequency_var, width=5, justify="center")
        freq_entry.pack(side=tk.LEFT, padx=(8, 6))
        ttk.Label(freq_frame, text="seconds", style="Body.TLabel").pack(side=tk.LEFT)
        self.frequency_entry = freq_entry

        volume_frame = ttk.Frame(controls_section, style="Card.TFrame")
        volume_frame.pack(fill=tk.X, pady=(14, 0))

        ttk.Label(volume_frame, textvariable=self.target_display_var, style="Body.TLabel").pack(anchor=tk.W)
        self.volume_slider = ttk.Scale(
            volume_frame,
            from_=0,
            to=100,
            orient=tk.HORIZONTAL,
            variable=self.target_volume_var,
            command=self._on_volume_change,
        )
        self.volume_slider.pack(fill=tk.X, pady=(6, 10))

        self.target_progress = ttk.Progressbar(volume_frame, maximum=100, value=100)
        self.target_progress.pack(fill=tk.X)

        actions = ttk.Frame(main, style="Card.TFrame")
        actions.pack(fill=tk.X, pady=(0, 18))

        self.start_button = ttk.Button(actions, text="Start monitoring", style="Accent.TButton", command=self.start_monitoring)
        self.start_button.pack(side=tk.LEFT, expand=True, fill=tk.X)

        self.stop_button = ttk.Button(
            actions,
            text="Stop",
            style="Danger.TButton",
            command=self.stop_monitoring,
            state=tk.DISABLED,
        )
        self.stop_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(12, 0))

        status_section = ttk.Frame(main, style="Card.TFrame")
        status_section.pack(fill=tk.BOTH, expand=True)

        ttk.Label(status_section, text="Live status", style="Section.TLabel").pack(anchor=tk.W)

        status_row = ttk.Frame(status_section, style="Card.TFrame")
        status_row.pack(fill=tk.X, pady=(10, 6))

        self.status_indicator = tk.Canvas(
            status_row,
            width=14,
            height=14,
            highlightthickness=0,
            bg=self.palette["card"],
            bd=0,
        )
        self.status_indicator.pack(side=tk.LEFT, pady=2)
        self.status_indicator_circle = self.status_indicator.create_oval(2, 2, 12, 12, fill="#b0b7c3", outline="")

        ttk.Label(status_row, textvariable=self.status_message_var, style="Body.TLabel").pack(side=tk.LEFT, padx=(10, 0))

        ttk.Label(status_section, textvariable=self.last_applied_var, style="Info.TLabel").pack(anchor=tk.W, pady=(6, 0))
        ttk.Label(status_section, textvariable=self.next_check_var, style="Info.TLabel").pack(anchor=tk.W)

    def refresh_devices(self, initial: bool = False) -> None:
        """Fetch the list of active recording devices and update the combo box."""
        selected = self.device_name_var.get()
        self.update_status("Refreshing recording devices…", level="info")
        self.devices = get_audio_devices("in")
        mic_names = [device.FriendlyName for device in self.devices]
        self.device_map = {device.FriendlyName: device for device in self.devices}

        self.device_combo["values"] = mic_names
        if selected in self.device_map:
            self.device_combo.set(selected)
        elif mic_names:
            self.device_combo.set(mic_names[0])
        else:
            self.device_combo.set("")

        if not mic_names:
            self.device_details_var.set("No active recording devices detected.")
            self.current_device = None
            self.endpoint_volume = None
            self.update_status("Connect or enable a microphone, then refresh.", level="warning")
            self.stop_monitoring()
            return

        if not initial or mic_names:
            self.update_status("Select a microphone to start monitoring.", level="info")
        self._on_device_selected()

    def _on_device_selected(self) -> None:
        name = self.device_name_var.get()
        self.current_device = self.device_map.get(name)
        if not self.current_device:
            self.device_details_var.set("No device selected.")
            self.endpoint_volume = None
            return
        self.endpoint_volume = self.current_device.EndpointVolume
        device_id = getattr(self.current_device, "id", "Unknown ID")
        self.device_details_var.set(f"Friendly name: {name}\nDevice ID: {device_id}")
        if self.monitoring:
            self.update_status(f"Monitoring '{name}'.", level="success")

    def _on_volume_change(self, value: str) -> None:
        try:
            numeric = int(float(value))
        except ValueError:
            numeric = self.target_volume_var.get()
        numeric = max(0, min(100, numeric))
        self.target_volume_var.set(numeric)
        self.target_display_var.set(f"Target: {numeric}%")
        self.target_progress["value"] = numeric

    def get_frequency_seconds(self) -> Optional[int]:
        raw = self.frequency_var.get().strip()
        try:
            value = int(raw)
            if value <= 0:
                raise ValueError
            return value
        except ValueError:
            messagebox.showerror("Invalid frequency", "Please enter a positive number of seconds (e.g. 5).")
            self.frequency_var.set(str(self.frequency_seconds))
            return None

    def start_monitoring(self) -> None:
        if self.monitoring:
            return
        selected = self.device_name_var.get()
        if not selected:
            messagebox.showinfo("No device selected", "Please choose a microphone before starting monitoring.")
            return
        device = self.device_map.get(selected)
        if not device:
            messagebox.showwarning("Device unavailable", "The selected microphone is no longer available. Refresh the list.")
            return
        frequency = self.get_frequency_seconds()
        if frequency is None:
            return
        self.frequency_seconds = frequency
        self.current_device = device
        self.endpoint_volume = device.EndpointVolume
        self.monitoring = True
        self.start_button.state(["disabled"])
        self.stop_button.state(["!disabled"])
        self.device_combo.configure(state="disabled")
        self.update_status(f"Monitoring '{selected}'.", level="success")
        self.enforce_target_volume()

    def enforce_target_volume(self) -> None:
        if not self.monitoring or not self.endpoint_volume:
            return
        target = max(0, min(100, self.target_volume_var.get())) / 100.0
        try:
            self.endpoint_volume.SetMasterVolumeLevelScalar(target, None)
        except Exception as exc:  # noqa: BLE001
            self.update_status(f"Failed to set volume: {exc}", level="error")
            self.stop_monitoring()
            return
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.last_applied_var.set(f"Last applied: {timestamp}")
        self.update_status(
            f"Enforcing {int(target * 100)}% on '{self.device_name_var.get()}'.",
            level="success",
        )
        self.schedule_next_enforcement()

    def schedule_next_enforcement(self) -> None:
        if not self.monitoring:
            return
        if self.after_id:
            self.root.after_cancel(self.after_id)
            self.after_id = None
        delay = self.get_frequency_seconds()
        if delay is None:
            self.stop_monitoring()
            return
        self.frequency_seconds = delay
        self.next_check_var.set(f"Next check in: {delay} s")
        self.after_id = self.root.after(delay * 1000, self.enforce_target_volume)

    def stop_monitoring(self) -> None:
        if not self.monitoring and not self.after_id:
            return
        if self.after_id:
            self.root.after_cancel(self.after_id)
            self.after_id = None
        self.monitoring = False
        self.start_button.state(["!disabled"])
        self.stop_button.state(["disabled"])
        self.device_combo.configure(state="readonly")
        self.next_check_var.set("Next check in: —")
        if self.current_device:
            self.update_status("Monitoring paused.", level="info")
        else:
            self.update_status("Select a microphone to begin.", level="info")

    def update_status(self, message: str, level: str = "info") -> None:
        colors = {
            "info": "#90a4ae",
            "success": "#1faa00",
            "warning": "#ffb300",
            "error": "#d32f2f",
        }
        color = colors.get(level, colors["info"])
        self.status_message_var.set(message)
        self.status_indicator.itemconfig(self.status_indicator_circle, fill=color)

    def on_close(self) -> None:
        if self.after_id:
            self.root.after_cancel(self.after_id)
        self.root.destroy()


def main() -> None:
    root = tk.Tk()
    app = MicrophoneApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
