import tkinter as tk
from tkinter import ttk
import comtypes
from pycaw.pycaw import AudioUtilities, IMMDeviceEnumerator, EDataFlow
from pycaw.constants import CLSID_MMDeviceEnumerator, DEVICE_STATE

def MyGetAudioDevices(direction="in", State=DEVICE_STATE.ACTIVE.value):
    devices = []
    if direction == "in":
        Flow = EDataFlow.eCapture.value
    else:
        Flow = EDataFlow.eRender.value
    deviceEnumerator = comtypes.CoCreateInstance(
        CLSID_MMDeviceEnumerator,
        IMMDeviceEnumerator,
        comtypes.CLSCTX_INPROC_SERVER)
    if deviceEnumerator is None:
        return devices
    collection = deviceEnumerator.EnumAudioEndpoints(Flow, State)
    if collection is None:
        return devices
    num_devices = collection.GetCount()
    for i in range(num_devices):
        imm_device = collection.Item(i)
        devices.append(AudioUtilities.CreateDevice(imm_device))
    return devices

root = tk.Tk()
root.title("Microphone volume control")
root.configure(bg="#f0f0f0")
root.geometry("340x360")


main_frame = tk.Frame(root, bg="#f0f0f0", padx=20, pady=20)
main_frame.pack(expand=True)

label = tk.Label(main_frame, text="Select a mic:", bg="#f0f0f0", font=("Arial", 10))
label.pack(pady=(0, 10))

mics = []
mic_dict = {}


def refresh_devices():
    """Refresh the list of available recording devices."""
    global mics, mic_dict
    mics = MyGetAudioDevices("in")
    mic_names = [mic.FriendlyName for mic in mics]
    mic_dict = {mic.FriendlyName: mic for mic in mics}
    current_selection = combo.get()
    combo["values"] = mic_names
    if current_selection in mic_dict:
        combo.set(current_selection)
    elif mic_names:
        combo.set(mic_names[0])
    else:
        combo.set("")


combo = ttk.Combobox(main_frame, state="readonly", width=30)
combo.pack(pady=(0, 10))

refresh_button = tk.Button(
    main_frame,
    text="Refresh devices",
    command=refresh_devices,
    bg="#2196F3",
    fg="white",
    padx=10,
    pady=5,
)
refresh_button.pack(pady=(0, 15))


freq_label = tk.Label(main_frame, text="Check frequency (seconds):", bg="#f0f0f0", font=("Arial", 10))
freq_label.pack(pady=(0, 5))
freq_entry = tk.Entry(main_frame, width=10, justify="center")
freq_entry.insert(0, "5")
freq_entry.pack()


volume_label = tk.Label(main_frame, text="Target volume (%)", bg="#f0f0f0", font=("Arial", 10))
volume_label.pack(pady=(15, 5))

volume_var = tk.IntVar(value=100)
volume_slider = tk.Scale(
    main_frame,
    from_=0,
    to=100,
    orient=tk.HORIZONTAL,
    variable=volume_var,
    resolution=1,
    length=220,
    bg="#f0f0f0",
    highlightthickness=0,
)
volume_slider.pack()


status_label = tk.Label(
    main_frame,
    text="Select a microphone to begin.",
    bg="#f0f0f0",
    fg="#333333",
    wraplength=260,
    justify="center",
)
status_label.pack(pady=(15, 0))


current_volume = None
after_id = None


def start():
    global current_volume, after_id
    selected = combo.get()
    if not selected:
        status_label.config(text="No microphone selected. Please choose a device.", fg="#b00020")
        return
    device = mic_dict.get(selected)
    if not device:
        status_label.config(
            text="Selected microphone is unavailable. Refresh the device list.",
            fg="#b00020",
        )
        return

    current_volume = device.EndpointVolume

    def set_target_volume():
        global after_id
        if current_volume:
            target_level = max(0, min(100, volume_var.get())) / 100.0
            current_volume.SetMasterVolumeLevelScalar(target_level, None)
            status_label.config(
                text=f"Enforcing {int(target_level * 100)}% on '{selected}'.",
                fg="#2e7d32",
            )
        try:
            freq = int(freq_entry.get())
            if freq <= 0:
                raise ValueError
        except ValueError:
            freq = 5
            freq_entry.delete(0, tk.END)
            freq_entry.insert(0, str(freq))
        after_id = root.after(freq * 1000, set_target_volume)

    set_target_volume()


def stop():
    global after_id
    if after_id:
        root.after_cancel(after_id)
        after_id = None
        status_label.config(text="Monitoring paused.", fg="#333333")


button_frame = tk.Frame(main_frame, bg="#f0f0f0")
button_frame.pack(pady=20)

button_start = tk.Button(
    button_frame,
    text="Start",
    command=start,
    bg="#4CAF50",
    fg="white",
    padx=10,
    pady=5,
)
button_start.pack(side=tk.LEFT, padx=10)

button_stop = tk.Button(
    button_frame,
    text="Stop",
    command=stop,
    bg="#f44336",
    fg="white",
    padx=10,
    pady=5,
)
button_stop.pack(side=tk.RIGHT, padx=10)


refresh_devices()

root.mainloop()