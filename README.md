# XJ Audio Latency Tester

A Windows desktop tool for measuring audio round-trip latency across multiple output devices simultaneously.

## What it does

Plays a known test signal (chirp sweep) through a selected audio output, records it back through a microphone input, then computes the exact delay using **normalized cross-correlation** — giving ~0.02 ms (1 sample) precision.

Useful for comparing latency between:
- Multiple monitor outputs (HDMI / DisplayPort audio)
- Bluetooth speakers vs wired
- Different audio interfaces
- Voicemeeter virtual inputs vs physical hardware

## Features

- **Multi-device testing** — add any number of output devices, test them sequentially in one run
- **Cross-correlation algorithm** — chirp signal (200 Hz → 8 kHz, 100 ms) + Stopwatch-based timing offset compensation; no artificial floor on measurable latency
- **Bluetooth warm-up** — optional 1-second low-amplitude init phase to wake BT devices before measurement
- **WASAPI shared mode output** with automatic format conversion (handles 48 kHz / 32-bit float devices); falls back to WinMM if WASAPI is blocked
- **Threshold calibration** — auto-calibrate detection threshold from ambient noise
- **Relative latency** — when testing multiple devices, shows offset from the highest-latency device
- **Save recordings** — export captured audio as WAV for offline analysis
- **Persistent settings** — device selection, threshold, BT warm-up saved between sessions

## Requirements

- Windows 10 / 11
- .NET Framework 4.8 (pre-installed on all modern Windows)
- A microphone or loopback input positioned to capture the output devices

## Usage

1. **Output Devices** — select one or more devices to test (use **+** to add rows)
2. **Input Device** — select the microphone that will capture the playback
3. Adjust **Detection Threshold** or click **Calibrate Threshold** to auto-set from ambient noise
4. Enable **Bluetooth Warm-up** if testing BT speakers
5. Click **Run Latency Analysis**

Results show absolute latency per device and relative offsets between devices.

### Voicemeeter note

When Voicemeeter Potato is running, physical hardware outputs are locked by its kernel driver. To measure latency through Voicemeeter's routing, select one of the virtual inputs (**VoiceMeeter Input**, **VoiceMeeter Aux Input**, **VoiceMeeter VAIO3 Input**) as the output device, and configure routing inside Voicemeeter to the desired physical output.

## Build

```
dotnet build XJAudioLatencyTester.csproj -c Release
```

Or publish as a single EXE (all dependencies embedded via Costura.Fody):

```
dotnet publish XJAudioLatencyTester.csproj -c Release -o dist
```

## How the latency algorithm works

1. Generates a **linear chirp** (200 → 8000 Hz, 100 ms) preceded by an optional 1 s warm-up tone
2. Starts recording and playback simultaneously, capturing exact start times via `Stopwatch`
3. Compensates for the recording→playback start offset
4. Runs **normalized cross-correlation** between the reference chirp and the recording over a ±50 ms / +800 ms search window
5. Uses a **coarse pass** (0.25 ms step) followed by a **fine pass** (per-sample) around the peak
6. Reports `detected_position − expected_position` as the measured latency

## Code Signing

Code signing for release binaries is provided by the [SignPath Foundation](https://signpath.org).

## Dependencies

- [NAudio](https://github.com/naudio/NAudio) 2.2.1 — audio I/O (WASAPI, WinMM, WaveIn/WaveOut)
- [Costura.Fody](https://github.com/Fody/Costura) 5.7.0 — embeds all DLLs into the EXE at build time
