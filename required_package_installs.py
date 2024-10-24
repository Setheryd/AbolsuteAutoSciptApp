import subprocess
import sys
import os
import time

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


def display_progress_bar(current, total, bar_length=40):
    """
    Displays a user-friendly progress bar in the console.
    """
    progress = current / total
    block = int(bar_length * progress)
    bar = "#" * block + "-" * (bar_length - block)
    print(f"\rProgress: [{bar}] {current}/{total} packages", end="")


def run_command(command, success_message, error_message, package=None):
    """
    Runs a shell command, logs the output, and handles errors.
    Warnings and other output will be suppressed from the console.
    """
    try:
        result = subprocess.run(
            command,
            shell=True,
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )

        # Filter out warnings from stderr
        filtered_stderr = "\n".join(
            line for line in result.stderr.splitlines() if "WARNING" not in line
        )

        # Log the filtered standard error (if any) but don't display warnings to user
        if filtered_stderr:
            log(filtered_stderr.strip())
        
        log(success_message)
    except subprocess.CalledProcessError:
        # Only show package name on failure
        log(error_message)
        if package:
            log(f"Error occurred while installing: {package}")



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
            text=True,
        )
        return True
    except subprocess.CalledProcessError:
        return False


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
            text=True,
        )
        # Some Python versions output version info to stderr
        version_info = python_version.stdout.strip() or python_version.stderr.strip()
        log(f"Python is installed: {version_info}")
    except subprocess.CalledProcessError:
        log("Python is not installed or not added to PATH.")
        log(
            "Please install Python from https://www.python.org/downloads/ and ensure it's added to your system's PATH."
        )
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

    # Install packages with progress bar
    total_packages = len(PACKAGES)
    for index, package in enumerate(PACKAGES, start=1):
        display_progress_bar(index, total_packages)

        # Install the package
        run_command(
            f"python -m pip install {package} --user",
            f"",
            f"Failed to install {package}. Please check your internet connection and Python setup.",
            package,
        )
        # Short delay to simulate loading
        time.sleep(0.1)

    # Final progress completion
    display_progress_bar(total_packages, total_packages)
    print("\nInstallation completed.")


if __name__ == "__main__":
    main()
