# CUDA / PyTorch GPU Setup

## System Info

| Component | Value |
|-----------|-------|
| GPU | NVIDIA GeForce RTX 5060 Ti (Blackwell) |
| VRAM | 16 GB |
| Driver | 581.57 |
| CUDA (driver max) | 13.0 |
| Python | 3.14.3 |
| PyTorch | 2.11.0 |

## Install Command

Use `cu128` — the latest stable PyTorch index compatible with Blackwell GPUs and CUDA 13.0 (backward-compatible):

```bash
pip install torch==2.11.0+cu128 torchvision==0.26.0+cu128 torchaudio==2.11.0+cu128 --index-url https://download.pytorch.org/whl/cu128
```

## Why not cu121 (or older)?

- The `cu121` index does not carry PyTorch 2.9+ or Python 3.14 wheels
- RTX 50-series (Blackwell) GPUs require CUDA 12.8+ for full compute support
- `cu128` is the correct channel for this GPU

## Verify GPU is Active

```python
import torch
print(torch.cuda.is_available())       # True
print(torch.cuda.get_device_name(0))   # NVIDIA GeForce RTX 5060 Ti
```

Verified output after install:

```
torch: 2.11.0+cu128
CUDA available: True
CUDA version: 12.8
Device name: NVIDIA GeForce RTX 5060 Ti
VRAM: 17.1 GB
```

## Whisper / faster-whisper Notes

- `faster-whisper` with `device="cuda"` will use the RTX 5060 Ti automatically
- Set `compute_type="float16"` for best performance on Blackwell
- 16 GB VRAM is sufficient for `large-v3` model
