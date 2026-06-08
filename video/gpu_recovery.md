# GPU Recovery

Diagnose and recover an NVIDIA GPU that Windows has marked **disabled** (Device Manager **Code 22**). When the card drops into this state, Windows falls back to the *Microsoft Basic Render Driver*, `nvidia-smi` stops responding, and every NVENC-accelerated tool in this `video/` folder (re-encoder, screen recorder, etc.) silently loses GPU acceleration. This tool detects the condition and applies the one fix that actually works.

## The symptom

- NVIDIA app shows **"No driver installed"**.
- Windows → System → About shows **"No dedicated VRAM / No GPU installed"** and *Microsoft Basic Render Driver*.
- Device Manager still lists the card by its correct name, but with an error.
- `nvidia-smi` fails or reports no device.

The card still appears in Device Manager with its real name, which means **the GPU is healthy and present on the PCIe bus** — it is not dead and not a driver crash. Confirm with:

```powershell
Get-PnpDevice -PresentOnly -Class Display |
  Where-Object { $_.InstanceId -like 'PCI\VEN_10DE*' } |
  Select-Object FriendlyName, Status, ConfigManagerErrorCode
```

`ConfigManagerErrorCode = 22` is the giveaway: **Code 22 = "this device is disabled."** Not Code 43 (fell off the bus), not Code 10 (cannot start), not Code 31 (driver problem) — just *switched off*.

## Root cause and the non-obvious trap

The card gets **disabled** — by a manual Device Manager click, or by a utility (on this machine, **MSI Center** is the prime suspect; there is *no* crash/TDR in the event log when it happens). That part is mundane. The trap is the fix:

> **The PowerShell `Enable-PnpDevice` cmdlet reports success but does *not* actually start the device.**

It returns cleanly, but the card stays at Code 22 through:

- `Enable-PnpDevice` (clears the registry disable-flag, `ConfigFlags = 0`, but the devnode never starts)
- a PnP rescan (`pnputil /scan-devices`)
- a disable → enable cycle
- cycling the **parent PCIe root port**
- **a full reboot**

The native tool does the real enable in one shot:

```powershell
pnputil /enable-device "PCI\VEN_10DE&DEV_2D04&...&REV_A1\<serial>"
```

```
Enabling device: ...RTX 5060 Ti
Device enabled successfully.
```

→ immediately back to `Status = OK / Code 0`, driver loaded, `nvidia-smi` responding. **No reboot needed.** `gpu_recovery.py` wraps exactly this: discover the disabled NVIDIA adapter, run `pnputil /enable-device`, verify with `nvidia-smi`.

## Requirements

- Windows + an NVIDIA GPU.
- **Administrator rights** (enabling a device requires them — the `.bat` self-elevates via UAC).
- Python 3.x (standard library only — no extra packages).

## Running

**Recommended — double-click** `gpu_recovery.bat`. It requests admin (accept the UAC prompt), diagnoses the GPU, and fixes it if it finds it disabled.

**From a terminal:**

```bash
python gpu_recovery.py             # diagnose, then fix if disabled (needs admin to fix)
python gpu_recovery.py --diagnose  # report only, make no changes
```

The GPU's device-instance ID is **discovered dynamically** — nothing is hardcoded, so the tool works on any machine/card.

## What it does

1. Enumerates present NVIDIA display adapters and reads each one's `ConfigManagerErrorCode` and registry `ConfigFlags`.
2. Runs `nvidia-smi` to show whether the card is live.
3. If any GPU is at **Code 22**, runs `pnputil /enable-device <instance>` (admin) and re-verifies.
4. If the card *stays* Code 22 after the enable, prints recent System-log events touching the device — to help track down what is re-disabling it.

## Output (healthy result)

```
✅ Recovered. nvidia-smi: NVIDIA GeForce RTX 5060 Ti, 591.86, 16311 MiB, 33, 9 %
```

## Troubleshooting

| Issue | Solution |
|-------|----------|
| "Enabling the device needs administrator rights" | Use `gpu_recovery.bat` (self-elevates), or run an elevated terminal. |
| No NVIDIA adapter found | Card may be physically unseated or not enumerating — check the slot / power; this is then a hardware issue, not Code 22. |
| Still Code 22 after the fix | Something is actively re-disabling it. Check the printed System-log events, MSI Center, and Task Scheduler. |
| Windows shows "4 GB" VRAM | Cosmetic — the WMI `AdapterRAM` field overflows at 4 GB. Trust `nvidia-smi` (reports the true 16 GB). |
| `nvidia-smi` "insufficient permissions" | Run elevated, or simply confirm via Device Manager that the status is OK. |

## Related

- `video_reencoder.py`, `screen_recorder.py` — the NVENC-dependent tools this recovery protects.
