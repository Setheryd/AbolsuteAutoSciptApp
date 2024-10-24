import subprocess
import sys
import os

# Define the log file path
LOGFILE = "install_log.txt"

def log(message):
    """
    Writes a message to both the log file and the console.
    """
    with open(LOGFILE, "a", encoding="utf-8") as log_file:
        log_file.write(message + "\n")
    print(message)

def check_error(error_message):
    """
    Checks the last subprocess return code.
    If an error occurred, logs the error message, displays the log,
    and exits the script.
    """
    if subprocess.call(["echo", "%ERRORLEVEL%"], shell=True) != 0:
        log(f"Error: {error_message}")
        log("Installation aborted.")
        display_log()
        input("Press Enter to exit...")
        sys.exit(1)

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

def main():
    # Initialize the log file
    if os.path.exists(LOGFILE):
        os.remove(LOGFILE)  # Remove existing log file
    with open(LOGFILE, "w", encoding="utf-8") as log_file:
        log_file.write("============================================\n")
        log_file.write("       Installing pywin32, bs4, and pandas  \n")
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
        f"pip install --upgrade pip --user",
        "Pip upgrade completed.",
        "Failed to upgrade pip. Please check your internet connection and Python setup.",
    )

    log("Pip upgrade done. Proceeding to install pywin32...")

    # Install pywin32
    log("\nInstalling pywin32...")
    run_command(
        f"pip install pywin32 --user",
        "pywin32 installed.",
        "Failed to install pywin32. Please check your internet connection and Python setup.",
    )

    log("pywin32 installed. Proceeding to install bs4...")

    # Install bs4
    log("\nInstalling bs4...")
    run_command(
        f"pip install bs4 --user",
        "bs4 installed.",
        "Failed to install bs4. Please check your internet connection and Python setup.",
    )

    log("bs4 installed. Proceeding to install pandas...")

    # Install pandas
    log("\nInstalling pandas...")
    run_command(
        f"pip install pandas --user",
        "pandas installed.",
        "Failed to install pandas. Please check your internet connection and Python setup.",
    )

    # Verify installations
    log("\nVerifying installations...")

    # Verify pywin32
    try:
        verification = subprocess.run(
            [sys.executable, "-c", "import win32; print('pywin32 installed successfully.')"],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True
        )
        log(verification.stdout.strip())
    except subprocess.CalledProcessError:
        log("pywin32 verification failed.")
        display_log()
        input("Press Enter to exit...")
        sys.exit(1)

    # Verify bs4
    try:
        verification = subprocess.run(
            [sys.executable, "-c", "from bs4 import BeautifulSoup; print('bs4 installed successfully.')"],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True
        )
        log(verification.stdout.strip())
    except subprocess.CalledProcessError:
        log("bs4 verification failed.")
        display_log()
        input("Press Enter to exit...")
        sys.exit(1)

    # Verify pandas
    try:
        verification = subprocess.run(
            [sys.executable, "-c", "import pandas as pd; print('pandas installed successfully.')"],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True
        )
        log(verification.stdout.strip())
    except subprocess.CalledProcessError:
        log("pandas verification failed.")
        display_log()
        input("Press Enter to exit...")
        sys.exit(1)

    # Final log entry
    log("\n============================================")
    log("   All packages installed successfully!      ")
    log("============================================")

    # Display the log file
    log("\nInstallation complete. Reviewing log...")
    display_log()

    # Keep the Command Prompt window open
    input("Press Enter to exit...")

if __name__ == "__main__":
    main()
