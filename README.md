# Microphone-always-at-100

Microphone-always-at-100 is a tiny Windows utility that keeps your microphone input level pinned at 100% so you never sound quiet in calls again. The script watches your preferred recording device at a configurable interval and immediately pushes it back to full volume whenever an app or driver turns it down.

## Features
- **Auto-reset microphone input level** – Select any active capture device and the script continuously enforces your preferred volume.
- **Adjustable target volume** – Choose the exact percentage you want enforced instead of being limited to 100%.
- **Configurable refresh rate** – Choose how often (in seconds) the check runs to balance responsiveness and resource use.

- **Live device management** – Refresh the microphone list without restarting the app, with clear availability messaging.
- **Polished monitoring dashboard** – Modern Tkinter styling, visual target gauge, and live status indicators show when the level was last applied and when it will be checked again.
=======
- **Live device management** – Refresh the microphone list without restarting the app and see real-time status updates about enforcement.
- **Simple graphical interface** – Built with Tkinter for an easy, dependency-light UI.


## Requirements
- Windows 10/11 (the [`pycaw`](https://github.com/AndreMiras/pycaw) audio APIs only work on Windows through WASAPI).
- Python 3.8 or newer installed and on your `PATH`.
- The following Python packages:
  - `pycaw`
  - `comtypes`
  - `tkinter` (ships with the standard Python installer on Windows).

Install the missing Python packages with:

```bash
pip install pycaw comtypes
```

## Usage
1. Clone or download this repository.
2. Launch the script:
   ```bash
   python MicrophoneEnhancer.py
   ```
3. Pick your microphone from the drop-down list.
4. Set how frequently (in seconds) you want the script to re-apply the target volume level.
5. Use the slider to choose the exact volume percentage to enforce (default: 100%) and verify the target via the progress gauge.
6. Click **Start monitoring** to begin. The volume will be forced to the selected level on the specified cadence while the status panel confirms each enforcement.
5. Use the slider to choose the exact volume percentage to enforce (default: 100%).
6. Click **Start** to begin monitoring. The volume will be forced to the selected level on the specified cadence.
7. Click **Stop** to pause monitoring when you no longer need it.

> **Tip:** Start the utility before joining meetings that tend to lower your microphone. Leaving it running in the background is usually sufficient, since the volume enforcement only happens on the chosen interval.

## Troubleshooting
- **I do not see my microphone in the list.** Make sure the device is enabled in Windows and appears in *Sound Settings → Recording*. The script only lists active capture devices.
- **The script crashes on start-up.** Confirm you are running it on Windows and that `pycaw`/`comtypes` are installed for the same Python interpreter you use to run the script.
- **I want a different target volume.** The current implementation forces 100% volume. Modify the call to `SetMasterVolumeLevelScalar` in `MicrophoneEnhancer.py` if you want to enforce a different level.

## Contributing
Bug reports, feature requests, and pull requests are welcome. Please open an issue to discuss major changes before submitting a PR.

## License
This project is released under the [MIT License](LICENSE).
