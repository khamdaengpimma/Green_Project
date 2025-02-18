import webbrowser
import subprocess
import os
import sys
import screen_brightness_control as sbc
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import shutil
from psutil import virtual_memory, disk_partitions, disk_usage
import platform
def check_admin_privileges():
    """Check if the script is running with administrative privileges."""
    try:
        # Try to run a simple command that requires admin privileges (net session)
        subprocess.check_call("net session", shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        return True
    except subprocess.CalledProcessError:
        return False

def run_as_admin():
    """Restart the script with administrative privileges."""
    script = os.path.abspath(__file__)  # Get the script path
    python_exe = sys.executable  # Get the path to the Python interpreter

    # Run the script with Python explicitly as an administrator
    subprocess.run(["powershell", "-Command", f"Start-Process '{python_exe}' -ArgumentList '{script}' -Verb runAs"], check=True)

def set_time_zone():
    """Set the time zone to SE Asia Standard Time."""
    subprocess.run(["tzutil", "/s", "SE Asia Standard Time"], check=True)
def openlink():
    urls = [
    "https://trinhanlaptop.vn/tools/test-ban-phim-online/",
    "https://www.loom.com/webcam-mic-test",
    "https://trinhanlaptop.vn/tools/test-man-hinh-online/",
    "https://youtu.be/gC9KG7wSCXU"
    ]

    for url in urls:
        webbrowser.open(url)
def open_app(app_path):
    """Opens an application using the given file path."""
    if os.path.exists(app_path):
        subprocess.Popen(app_path, shell=True)
        print(f"Opening: {app_path}")
    else:
        print(f"Error: File not found -> {app_path}")

def generate_battery_report():
    """Generates and opens the battery report."""
    report_path = os.path.expandvars(r"%USERPROFILE%\Desktop\battery_report.html")

    # Generate the battery report
    subprocess.run("powercfg /batteryreport /output {}".format(report_path), shell=True)
    print("Battery report generated at:", report_path)

    # Open the report
    if os.path.exists(report_path):
        subprocess.run(f'start "" "{report_path}"', shell=True)
    else:
        print("Error: Battery report not found.")


def set_brightness_and_volume():
    """Sets screen brightness to 100% and volume to maximum."""
    
    # Set screen brightness to 100%
    try:
        sbc.set_brightness(100)
        print("Brightness set to 100%")
    except Exception as e:
        print(f"Error setting brightness: {e}")

    # Set system volume to maximum
    try:
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = cast(interface, POINTER(IAudioEndpointVolume))
        volume.SetMasterVolumeLevelScalar(1.0, None)  # 1.0 = 100%
        print("Volume set to maximum")
    except Exception as e:
        print(f"Error setting volume: {e}")

def run_powershell_script():
    # Define the path to the PowerShell script
    script_path = os.path.join(os.getcwd(), "Program", "Remove.ps1")
    
    # Check if the PowerShell script exists
    if not os.path.exists(script_path):
        print(f"Error: PowerShell script not found at {script_path}")
        return

    # Run the PowerShell script using subprocess
    try:
        subprocess.run(["powershell", "-ExecutionPolicy", "Bypass", "-File", script_path], check=True)
        print("PowerShell script executed successfully.")
    except subprocess.CalledProcessError as e:
        print(f"Error running PowerShell script: {e}")

def execute_script():
    # Check if fltmc command succeeds
    try:
        subprocess.run(["fltmc"], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except subprocess.CalledProcessError:
        sys.exit()

    # Construct the path for MAS_AIO.cmd
    mas_aio_cmd = os.path.join("Program", "MAS_AIO.cmd")
    
    # Debug: Print the path to verify it is correct
    print(f"MAS_AIO.cmd path: {mas_aio_cmd}")
    
    # Check if the file exists before running the command
    if not os.path.exists(mas_aio_cmd):
        print(f"Error: {mas_aio_cmd} not found!")
        sys.exit()

    # Call MAS_AIO.cmd with specified arguments
    subprocess.run([mas_aio_cmd, "/KMS38", "/Ohook"])

    # Cleanup if running from Setup\Scripts
    if os.path.abspath(os.path.dirname(__file__)) == os.path.join(os.getenv("SystemRoot"), "Setup", "Scripts"):
        shutil.rmtree(os.path.dirname(__file__), ignore_errors=True)
def get_cpu_name():
    """Get CPU name"""
    try:
        if platform.system() == "Windows":
            cmd = "wmic cpu get Name"
            cpu_name = subprocess.check_output(cmd, shell=True).decode().split("\n")[1].strip()
        else:
            cpu_name = platform.processor()
    except Exception as e:
        cpu_name = f"Unable to retrieve CPU name ({e})"
    cpuid=str(cpu_name[19])+str(cpu_name[21])
    return cpuid

def get_total_disk_space():
    """Calculate total disk space (excluding removable disks)"""
    total_space = sum(disk_usage(partition.mountpoint).total for partition in disk_partitions() if "fixed" in partition.opts.lower()) // (1024 ** 3)

    # Apply custom conditions for disk space
    if total_space < 200:
        return 128
    elif total_space > 300:
        return 512
    elif total_space >600:
        return 1024
    return 256

def get_system_info():
    """Retrieve CPU, RAM, and disk information"""
    cpu_name = f"{get_cpu_name()}"
    ram_info = virtual_memory()
    ram_total = f"{round(ram_info.total / (1024 ** 3))}"  # Round RAM size
    disk_total = f"{get_total_disk_space()}"
    return f"{cpu_name}{ram_total}{disk_total}.bat"

def create_script(filename):
    try:
        # Get the current user's desktop path dynamically
        if sys.platform == "win32":
            desktop_path = os.path.join(os.environ["USERPROFILE"], "Desktop")  # Windows
        else:
            # For macOS or Linux
            desktop_path = os.path.join(os.environ["HOME"], "Desktop")
        
        # Construct the full file path to the desktop
        file_path = os.path.join(desktop_path, filename)
        
        # Open the file for writing
        with open(file_path, "w") as file:
            # Writing commands to the file
            file.write('@echo off\n')
            file.write('taskkill /im sysinfoPC.exe /f\n')
            file.write('rd /s /q "BaoDuyQB"\n')
            file.write('shutdown /s /f /t 0\n')
            file.write(f'del "{filename}"\n')

        print(f"Script {filename} created successfully on the Desktop!")

    except Exception as e:
        print(f"Error: {e}")

def main():
    if not check_admin_privileges():
        run_as_admin()
        return  # Exit after relaunching as admin
    # Run the function
    set_time_zone()
    open_app(r"Program\UpdateTime.exe")
    open_app(r"Program\sysinfoPC.exe")
    run_powershell_script()
    set_brightness_and_volume()
    openlink()
    generate_battery_report()
    create_script(get_system_info())
    execute_script()
    


# Run the main function
if __name__ == "__main__":
    main()