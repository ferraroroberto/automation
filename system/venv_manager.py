import logging
import os
import subprocess
import sys
import ctypes
import platform

log = logging.getLogger(__name__)

def is_in_virtualenv():
    """Check if currently running in a virtual environment."""
    # Check standard virtualenv indicators
    in_venv = hasattr(sys, 'real_prefix') or (hasattr(sys, 'base_prefix') and sys.base_prefix != sys.prefix)
    
    # Additional check: look for VIRTUAL_ENV environment variable
    virtual_env = os.environ.get('VIRTUAL_ENV')
    
    # Check if the Python executable path contains typical VENV paths
    executable_path = sys.executable.lower()
    venv_indicators = ['.venv', 'virtualenv', 'venv', 'env']
    path_suggests_venv = any(indicator in executable_path for indicator in venv_indicators)
    
    return in_venv or bool(virtual_env) or path_suggests_venv

def get_venv_info():
    """Get information about the current virtual environment."""
    # Get the Python executable being used
    python_executable = sys.executable
    
    # Check VIRTUAL_ENV environment variable
    virtual_env = os.environ.get('VIRTUAL_ENV', '')
    
    if is_in_virtualenv():
        # Determine the actual VENV path
        venv_path = virtual_env if virtual_env else sys.prefix
        venv_name = os.path.basename(venv_path)
        return {
            "status": "active",
            "name": venv_name,
            "path": venv_path,
            "executable": python_executable
        }
    else:
        return {
            "status": "inactive",
            "name": "global Python",
            "path": sys.prefix,
            "executable": python_executable
        }

def deactivate_venv():
    """Provide instructions to deactivate the virtual environment."""
    log.info("To deactivate the virtual environment, you need to:")
    if os.name == 'nt':  # Windows
        log.info("1. Close this script")
        log.info("2. Run 'deactivate' in your command prompt")
        log.info("3. Make sure to run 'python' and not the VENV python when restarting")
        log.info("TIP: If deactivation doesn't work, try running this script with:")
        log.info("   %s %s", os.path.join(sys.base_prefix, 'python'), os.path.basename(__file__))
    else:  # Unix/Linux/Mac
        log.info("1. Close this script")
        log.info("2. Run 'deactivate' in your terminal")
        log.info("3. Make sure to run 'python' and not the VENV python when restarting")
        log.info("TIP: If deactivation doesn't work, try running this script with:")
        log.info("   %s %s", os.path.join(sys.base_prefix, 'bin/python'), os.path.basename(__file__))

    sys.exit("Please restart this script after deactivating the virtual environment.")

def get_installed_packages():
    """Get a list of installed packages."""
    result = subprocess.run([sys.executable, "-m", "pip", "list", "--format=freeze"], 
                           capture_output=True, text=True, check=True)
    return result.stdout.strip().split('\n')

def is_admin():
    """Check if the script is running with administrator privileges."""
    try:
        if platform.system() == 'Windows':
            return ctypes.windll.shell32.IsUserAnAdmin() != 0
        else:
            # On Unix systems, check if user is root (uid 0)
            return os.geteuid() == 0
    except (AttributeError, OSError):
        return False

def elevate_privileges():
    """Re-run the script with elevated privileges (Windows only)."""
    if platform.system() != 'Windows':
        log.warning("Privilege elevation is only supported on Windows.")
        return False

    try:
        script_path = os.path.abspath(__file__)
        log.info("Re-launching with administrator privileges...")

        # Use Python executable to run this script with admin rights
        ctypes.windll.shell32.ShellExecuteW(
            None,                          # Parent window handle
            "runas",                       # Operation ("runas" means run as admin)
            sys.executable,                # Program to execute (Python)
            f'"{script_path}"',            # Parameters (the script path)
            os.path.dirname(script_path),  # Working directory
            1                              # Show command window
        )

        # Exit the current non-elevated process
        sys.exit(0)

    except Exception as e:
        log.error("Failed to elevate privileges: %s", e)
        return False

def safe_uninstall_package(package):
    """Safely uninstall a package, handling permission errors."""
    try:
        # First try normal uninstall
        result = subprocess.run(
            [sys.executable, "-m", "pip", "uninstall", "-y", package],
            capture_output=True,
            text=True
        )
        
        # Check if there was a permission error
        if result.returncode != 0 and "PermissionError" in result.stderr:
            log.warning("Permission error when uninstalling %s. Trying with --user flag...", package)
            
            # Try with --user flag
            user_result = subprocess.run(
                [sys.executable, "-m", "pip", "uninstall", "--user", "-y", package],
                capture_output=True,
                text=True
            )
            
            if user_result.returncode == 0:
                return True, f"Successfully uninstalled {package} with --user flag"
            else:
                return False, f"Failed to uninstall {package}: {user_result.stderr}"
        
        elif result.returncode == 0:
            return True, f"Successfully uninstalled {package}"
        else:
            return False, f"Failed to uninstall {package}: {result.stderr}"
            
    except Exception as e:
        return False, f"Error uninstalling {package}: {str(e)}"

def clean_environment(packages_to_keep=None):
    """Clean the current environment by removing unnecessary packages."""
    if packages_to_keep is None:
        packages_to_keep = ['pip', 'setuptools', 'wheel']
    
    installed_packages = get_installed_packages()
    count = 0
    failed = 0

    log.info("Found %d installed packages", len(installed_packages))
    log.info("Will keep: %s", ', '.join(packages_to_keep))

    # Check if running as admin for global Python installations
    if not is_in_virtualenv() and not is_admin():
        log.warning("WARNING: You are not running with administrator privileges.")
        log.warning("You may encounter permission errors when uninstalling packages from a global Python installation.")
        log.info("Options:")
        log.info("  1. Restart this script with administrator privileges (recommended)")
        log.info("  2. Use a virtual environment instead")
        log.info("  3. Continue anyway (some uninstalls may fail)")

        choice = input("\nHow would you like to proceed? (1/2/3): ")

        if choice == '1':
            # Attempt to re-launch with elevated privileges
            elevate_privileges()
            # If we get here, elevation failed
            log.error("Failed to restart with administrator privileges.")
            proceed = input("Continue anyway? (y/n): ")
            if proceed.lower() != 'y':
                return
        elif choice == '2':
            log.info("Creating a virtual environment is recommended. Steps:")
            log.info("1. Run: python -m venv .venv")
            log.info("2. Activate it: .venv\\Scripts\\activate (Windows) or source .venv/bin/activate (Unix)")
            log.info("3. Install only the packages you need in the virtual environment")
            sys.exit("Please create a virtual environment and try again.")

    log.info("Packages to be uninstalled:")

    packages_to_uninstall = []

    for package_line in installed_packages:
        package_name = package_line.split('==')[0]
        if package_name.lower() not in [pkg.lower() for pkg in packages_to_keep]:
            log.info("  - %s", package_name)
            packages_to_uninstall.append(package_name)

    if not packages_to_uninstall:
        log.info("No packages to uninstall!")
        return

    # Ask for confirmation
    confirm = input("\nAre you sure you want to uninstall these packages? (y/n): ")
    if confirm.lower() != 'y':
        log.info("Operation cancelled.")
        return

    # Uninstall packages
    log.info("Uninstalling packages...")

    results = []
    for package in packages_to_uninstall:
        log.info("Uninstalling %s...", package)
        success, message = safe_uninstall_package(package)

        if success:
            count += 1
            log.info("  ✓ %s", message)
        else:
            failed += 1
            log.warning("  ✗ %s", message)

        results.append((package, success, message))

    log.info("Completed! Uninstalled %d packages successfully.", count)

    if failed > 0:
        log.warning("%d packages failed to uninstall. Summary of failures:", failed)
        for package, success, message in results:
            if not success:
                log.warning("  - %s: %s", package, message)

        log.info("Troubleshooting tips:")
        log.info("1. Run this script with administrator privileges")
        log.info("2. Try uninstalling problematic packages manually: pip uninstall <package>")
        log.info("3. Consider using a virtual environment instead of modifying your global Python")

    if count > 0:
        log.info("Your environment has been cleaned up!")

def main():
    """Main function to run the VENV manager."""
    # Check for --admin flag which could be used when script re-launches itself
    if "--admin" in sys.argv and not is_admin() and platform.system() == 'Windows':
        elevate_privileges()
    
    # Display script banner
    log.info("=" * 50)
    log.info("Python Virtual Environment Manager")
    log.info("=" * 50)

    venv_info = get_venv_info()

    log.info("=" * 50)
    log.info("Current Python Environment: %s", venv_info['name'])
    log.info("Status: %s", 'VIRTUAL ENVIRONMENT' if venv_info['status'] == 'active' else 'GLOBAL PYTHON')
    log.info("Path: %s", venv_info['path'])
    log.info("Python Executable: %s", venv_info['executable'])
    log.info("=" * 50)

    if venv_info['status'] == 'active':
        log.info("You are currently in a virtual environment.")
        choice = input("Do you want to deactivate it before proceeding (y) or force continue (f)? (y/f/n): ")
        if choice.lower() == 'y':
            deactivate_venv()
        elif choice.lower() == 'f':
            log.info("Forcing continuation despite being in a virtual environment...")
        else:
            log.info("Continuing in the virtual environment...")

    # Show installed packages
    packages = get_installed_packages()
    log.info("You have %d package(s) installed:", len(packages))

    # Display packages in columns (optional condensed view)
    row = []
    for i, pkg_line in enumerate(packages):
        pkg_name = pkg_line.split('==')[0]
        row.append(f"{pkg_name:25}")
        if (i + 1) % 3 == 0:
            log.info("\t".join(row))
            row = []
    if row:
        log.info("\t".join(row))

    # Ask if user wants to clean the environment
    choice = input("Do you want to clean this environment (remove unnecessary packages)? (y/n): ")
    if choice.lower() == 'y':
        # Define packages to keep
        essential_packages = ['pip', 'setuptools', 'wheel']

        # Ask for additional packages to keep
        user_kept_packages = input("\nEnter any additional packages to keep (comma-separated) or press Enter to continue: ")
        if user_kept_packages.strip():
            additional_packages = [pkg.strip() for pkg in user_kept_packages.split(',')]
            essential_packages.extend(additional_packages)

        clean_environment(essential_packages)

    log.info("Thank you for using the VENV Manager!")

if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    main()
