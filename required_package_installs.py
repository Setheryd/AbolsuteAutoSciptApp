import subprocess
import sys
import os

# Define the log file path
LOGFILE = "install_log.txt"

# List of packages to install
PACKAGES = [
    "altgraph",
    "attrs",
    "beautifulsoup4",
    "bs4",
    "certifi",
    "cffi",
    "charset-normalizer",
    "contourpy",
    "cycler",
    "docopt",
    "et-xmlfile",
    "fonttools",
    "h11",
    "idna",
    "kiwisolver",
    "matplotlib",
    "numpy",
    "openpyxl",
    "outcome",
    "packaging",
    "pandas",
    "pefile",
    "pillow",
    "pip",
    "pipreqs",
    "pycparser",
    "pyinstaller",
    "pyinstaller-hooks-contrib",
    "pyparsing",
    "PySide6",
    "PySide6_Addons",
    "PySide6_Essentials",
    "PySocks",
    "python-dateutil",
    "pytz",
    "pywin32",
    "pywin32-ctypes",
    "requests",
    "selenium",
    "setuptools",
    "shiboken6",
    "six",
    "sniffio",
    "sortedcontainers",
    "soupsieve",
    "trio",
    "trio-websocket",
    "typing_extensions",
    "tzdata",
    "urllib3",
    "websocket-client",
    "wsproto",
    "yarg",
]

def log(message):
    """
    Writes a message to both the log file and the console.
    """
    with open(LOGFILE, "a", encoding="utf-8") as log_file:
        log_file.write(message + "\n")
    print(message)

def run_command(command, success_message, error_message):
    """
    Runs a shell command, logs the output, and handles errors.
    """
    try:
        log(f"> {command}")
        result = subprocess.run(
            command,
            shell=True,
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True
        )
        # Log the standard output
        if result.stdout:
            log(result.stdout.strip())
        # Log the standard error
        if result.stderr:
            log(result.stderr.strip())
        log(success_message)
    except subprocess.CalledProcessError:
        log(error_message)
        display_log()
        input("Press Enter to exit...")
        sys.exit(1)

def display_log():
    """
    Displays the contents of the log file.
    """
    log("\n--- Installation Log ---")
    try:
        with open(LOGFILE, "r", encoding="utf-8") as log_file:
            print(log_file.read())
    except FileNotFoundError:
        print("Log file not found.")

def verify_package(package):
    """
    Verifies if a package is installed by using 'pip show'.
    """
    try:
        subprocess.run(
            [sys.executable, "-m", "pip", "show", package],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True
        )
        log(f"{package} is installed successfully.")
    except subprocess.CalledProcessError:
        log(f"Verification failed: {package} is not installed.")
        return False
    return True

def main():
    # Initialize the log file
    if os.path.exists(LOGFILE):
        os.remove(LOGFILE)  # Remove existing log file
    with open(LOGFILE, "w", encoding="utf-8") as log_file:
        log_file.write("============================================\n")
        log_file.write("         Installing Required Packages       \n")
        log_file.write("============================================\n")
    
    log("Checking if Python is installed...")

    # Check if Python is accessible
    try:
        python_version = subprocess.run(
            ["python", "--version"],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True
        )
        # Some Python versions output version info to stderr
        version_info = python_version.stdout.strip() or python_version.stderr.strip()
        log(f"Python is installed: {version_info}")
    except subprocess.CalledProcessError:
        log("Python is not installed or not added to PATH.")
        log("Please install Python from https://www.python.org/downloads/ and ensure it's added to your system's PATH.")
        display_log()
        input("Press Enter to exit...")
        sys.exit(1)

    log("Python check completed successfully. Continuing...")

    # Upgrade pip to the latest version
    log("\nUpgrading pip to the latest version...")
    run_command(
        f"python -m pip install --upgrade pip --user",
        "Pip upgrade completed.",
        "Failed to upgrade pip. Please check your internet connection and Python setup.",
    )

    log("Pip upgrade done. Proceeding to install packages...")

    # Iterate through the list of packages and install each
    for package in PACKAGES:
        log(f"\nInstalling {package}...")
        run_command(
            f"python -m pip install {package} --user",
            f"{package} installed successfully.",
            f"Failed to install {package}. Please check your internet connection and Python setup.",
        )

    
        # Keep the Command Prompt window open
        input("Press Enter to exit...")

if __name__ == "__main__":
    main()
