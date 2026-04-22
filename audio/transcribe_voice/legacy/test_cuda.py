#!/usr/bin/env python3
"""
CUDA Test Script
Run this script to verify CUDA installation and functionality.
"""

import sys
import torch
import time

def test_cuda_basic():
    """Test basic CUDA functionality"""
    print("=== CUDA Basic Test ===")
    print(f"PyTorch version: {torch.__version__}")
    print(f"CUDA available: {torch.cuda.is_available()}")

    if torch.cuda.is_available():
        print(f"CUDA version: {torch.version.cuda}")
        print(f"GPU count: {torch.cuda.device_count()}")
        print(f"GPU name: {torch.cuda.get_device_name(0)}")
        print(f"GPU memory: {torch.cuda.get_device_properties(0).total_memory / 1024**3:.1f} GB")
    else:
        print("❌ CUDA not available")
        return False

    return True

def test_cuda_computation():
    """Test CUDA computation with matrix multiplication"""
    print("\n=== CUDA Computation Test ===")

    if not torch.cuda.is_available():
        print("❌ CUDA not available for computation test")
        return False

    device = torch.device('cuda')
    print(f"Using device: {device}")

    # Test different matrix sizes
    sizes = [500, 1000, 2000]

    for size in sizes:
        print(f"\nTesting {size}x{size} matrix multiplication...")

        # Create random matrices on GPU
        a = torch.randn(size, size).to(device)
        b = torch.randn(size, size).to(device)

        # Time the computation
        start_time = time.time()
        c = torch.matmul(a, b)
        torch.cuda.synchronize()  # Ensure computation is complete
        end_time = time.time()

        elapsed = end_time - start_time
        print(f"Time: {elapsed:.4f} seconds")
        # Check memory usage
        memory_used = torch.cuda.memory_allocated(device) / 1024**2
        print(f"Memory used: {memory_used:.1f} MB")

    print("✅ CUDA computation test completed successfully")
    return True

def test_memory_management():
    """Test GPU memory management"""
    print("\n=== Memory Management Test ===")

    if not torch.cuda.is_available():
        print("❌ CUDA not available for memory test")
        return False

    device = torch.device('cuda')

    # Allocate some memory
    tensors = []
    for i in range(5):
        tensor = torch.randn(1000, 1000).to(device)
        tensors.append(tensor)
        memory_used = torch.cuda.memory_allocated(device) / 1024**2
        print(f"Memory after tensor {i+1}: {memory_used:.1f} MB")

    # Clear memory
    del tensors
    torch.cuda.empty_cache()
    memory_after = torch.cuda.memory_allocated(device) / 1024**2
    print(f"Memory after cleanup: {memory_after:.1f} MB")

    print("✅ Memory management test completed")
    return True

def main():
    """Main test function"""
    print("🚀 CUDA Test Script")
    print("=" * 50)

    try:
        # Run all tests
        basic_ok = test_cuda_basic()
        if basic_ok:
            comp_ok = test_cuda_computation()
            mem_ok = test_memory_management()

            if comp_ok and mem_ok:
                print("\n🎉 All CUDA tests passed! Your setup is working correctly.")
                print("You can now use CUDA acceleration in your Python projects.")
            else:
                print("\n⚠️ Some tests failed. Check your CUDA installation.")
        else:
            print("\n❌ Basic CUDA test failed. Please check your installation following CUDA.md")

    except Exception as e:
        print(f"\n❌ Test failed with error: {e}")
        print("Please check your CUDA installation and try again.")

    print("\n" + "=" * 50)

if __name__ == "__main__":
    main()
