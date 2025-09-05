# CUDA Setup and Activation Guide

## Overview
This guide will help you activate CUDA on your Windows system with NVIDIA GeForce GTX 1070 GPU and configure it for use with Python virtual environments.

## Prerequisites
- NVIDIA GeForce GTX 1070 GPU (8GB VRAM)
- Windows 10/11 with PowerShell
- Python virtual environment at `E:\automation\automation\.venv`

## Step 1: Verify GPU and Drivers

Check if your NVIDIA GPU is detected:

```powershell
nvidia-smi
```

Expected output should show your GTX 1070 with driver information.

## Step 2: Install CUDA Toolkit

### Download CUDA 12.1
1. Visit: https://developer.nvidia.com/cuda-12-1-0-download-archive
2. Select: Windows → Local → exe (local) → Download
3. Run the installer as Administrator
4. Choose "Custom" installation
5. Select these components:
   - CUDA Toolkit 12.1
   - CUDA Visual Studio Integration
   - CUDA Samples
   - CUDA Documentation

### Verify CUDA Installation
```powershell
nvcc --version
```

Expected output:
```
nvcc: NVIDIA (R) Cuda compiler driver
Copyright (c) 2005-2023 NVIDIA Corporation
Built on Mon_Apr__3_17:36:15_PDT_2023
Cuda compilation tools, release 12.1, V12.1.105
Build cuda_12.1.r12.1/compiler.32688072_0
```

## Step 3: Install cuDNN (Optional but Recommended)

1. Download cuDNN from: https://developer.nvidia.com/cudnn
2. Sign in to NVIDIA Developer account (create if needed)
3. Download cuDNN for CUDA 12.x
4. Extract files to: `C:\Program Files\NVIDIA GPU Computing Toolkit\CUDA\v12.1\`

## Step 4: Configure Python Virtual Environment

### Remove CPU-only PyTorch
```powershell
cd E:\automation\automation

# Remove current CPU-only PyTorch
& ".\.venv\Scripts\python.exe" -m pip uninstall torch torchvision torchaudio -y
```

### Install CUDA-enabled PyTorch
```powershell
# Install PyTorch with CUDA 12.1 support
& ".\.venv\Scripts\python.exe" -m pip install torch torchvision torchaudio --index-url https://download.pytorch.org/whl/cu121
```

## Step 5: Verify CUDA Activation

### Basic CUDA Test
```powershell
& ".\.venv\Scripts\python.exe" -c "
import torch
print('PyTorch version:', torch.__version__)
print('CUDA available:', torch.cuda.is_available())
print('CUDA version:', torch.version.cuda)
print('GPU count:', torch.cuda.device_count())
print('GPU name:', torch.cuda.get_device_name(0) if torch.cuda.is_available() else 'N/A')
"
```

Expected output:
```
PyTorch version: 2.8.0+cu121
CUDA available: True
CUDA version: 12.1
GPU count: 1
GPU name: NVIDIA GeForce GTX 1070
```

### GPU Memory Test
```powershell
& ".\.venv\Scripts\python.exe" -c "
import torch
if torch.cuda.is_available():
    device = torch.device('cuda')
    x = torch.randn(1000, 1000).to(device)
    y = torch.matmul(x, x)
    print('✅ CUDA working! Matrix multiplication successful')
    print('GPU Memory used:', torch.cuda.memory_allocated(device) / 1024**2, 'MB')
else:
    print('❌ CUDA not available')
"
```

## Step 6: Environment Variables (Optional)

Add these to your system environment variables for better CUDA detection:

```powershell
# Add to System Environment Variables
CUDA_PATH = C:\Program Files\NVIDIA GPU Computing Toolkit\CUDA\v12.1
CUDA_HOME = C:\Program Files\NVIDIA GPU Computing Toolkit\CUDA\v12.1
Path += %CUDA_PATH%\bin
Path += %CUDA_PATH%\libnvvp
```

## Troubleshooting

### Common Issues

1. **"CUDA not available" after installation**
   - Restart your computer
   - Check if virtual environment is activated correctly
   - Verify PyTorch installation: `pip list | findstr torch`

2. **Memory allocation errors**
   - Close other GPU-intensive applications
   - Check GPU memory usage: `nvidia-smi`

3. **Version compatibility issues**
   - Ensure CUDA toolkit version matches PyTorch CUDA version
   - GTX 1070 supports CUDA 12.x

### Check GPU Memory Usage
```powershell
nvidia-smi --query-gpu=memory.used,memory.total --format=csv
```

## Performance Tips

1. **GPU Memory Management**
   ```python
   # Clear GPU cache
   torch.cuda.empty_cache()

   # Check memory usage
   print(torch.cuda.memory_summary())
   ```

2. **Data Transfer Optimization**
   ```python
   # Use pinned memory for faster CPU-GPU transfers
   data = torch.randn(1000, 1000).pin_memory()

   # Use non_blocking transfers
   tensor = tensor.to(device, non_blocking=True)
   ```

3. **Multi-GPU (if applicable)**
   ```python
   # Use DataParallel for multiple GPUs
   model = torch.nn.DataParallel(model)
   ```

## Testing Your Setup

### Simple CUDA Benchmark
```python
import torch
import time

device = torch.device('cuda' if torch.cuda.is_available() else 'cpu')
print(f'Using device: {device}')

# Matrix multiplication benchmark
sizes = [1000, 2000, 5000]

for size in sizes:
    a = torch.randn(size, size).to(device)
    b = torch.randn(size, size).to(device)

    start_time = time.time()
    c = torch.matmul(a, b)
    torch.cuda.synchronize()  # Wait for GPU computation to complete
    end_time = time.time()

    print(f'Size {size}x{size}: {end_time - start_time:.4f} seconds')
```

## Next Steps

Once CUDA is activated, you can:
- Run GPU-accelerated machine learning models
- Use CUDA-enabled libraries (TensorFlow, PyTorch, etc.)
- Process large datasets faster
- Run computer vision and NLP models with GPU acceleration

## Support

If you encounter issues:
1. Check NVIDIA forums: https://forums.developer.nvidia.com/
2. PyTorch CUDA installation guide: https://pytorch.org/get-started/locally/
3. Verify your GPU compatibility: https://developer.nvidia.com/cuda-gpus

---

*Last updated: September 2024*
*Tested on: Windows 11, NVIDIA GeForce GTX 1070, CUDA 12.1*
