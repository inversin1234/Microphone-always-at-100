import tkinter as tk
from tkinter import ttk
import comtypes
from ctypes import cast, POINTER
from pycaw.pycaw import AudioUtilities, IMMDeviceEnumerator, EDataFlow, IAudioEndpointVolume
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
root.configure(bg='#f0f0f0')
root.geometry("300x250")


main_frame = tk.Frame(root, bg='#f0f0f0', padx=20, pady=20)
main_frame.pack(expand=True)

label = tk.Label(main_frame, text="Select a mic:", bg='#f0f0f0', font=('Arial', 10))
label.pack(pady=(0, 10))

mics = MyGetAudioDevices("in")
mic_names = [mic.FriendlyName for mic in mics]
mic_dict = {mic.FriendlyName: mic for mic in mics}

combo = ttk.Combobox(main_frame, values=mic_names, state="readonly", width=30)
combo.pack(pady=(0, 15))


freq_label = tk.Label(main_frame, text="Frequency (seconds):", bg='#f0f0f0', font=('Arial', 10))
freq_label.pack(pady=(0, 5))
freq_entry = tk.Entry(main_frame, width=10)
freq_entry.insert(0, "5")
freq_entry.pack()

current_volume = None
after_id = None

def start():
    global current_volume, after_id
    selected = combo.get()
    if not selected:
        return
    device = mic_dict[selected]
    current_volume = device.EndpointVolume
    
    def set_max_volume():
        global after_id
        if current_volume:
            current_volume.SetMasterVolumeLevelScalar(1.0, None)
        try:
            freq = int(freq_entry.get())
            after_id = root.after(freq * 1000, set_max_volume)
        except ValueError:
            after_id = root.after(5000, set_max_volume)
    
    set_max_volume()

def stop():
    global after_id
    if after_id:
        root.after_cancel(after_id)
        after_id = None


button_frame = tk.Frame(main_frame, bg='#f0f0f0')
button_frame.pack(pady=20)

button_start = tk.Button(button_frame, text="Start", command=start, bg='#4CAF50', fg='white', padx=10, pady=5)
button_start.pack(side=tk.LEFT, padx=10)

button_stop = tk.Button(button_frame, text="Stop", command=stop, bg='#f44336', fg='white', padx=10, pady=5)
button_stop.pack(side=tk.RIGHT, padx=10)

root.mainloop()