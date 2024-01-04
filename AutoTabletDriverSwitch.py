import subprocess
import time
import psutil
import pygetwindow as gw  # Import the pygetwindow library

def is_tablet_connected():
    tablet_process_names = ["TabTip.exe"]

    for process_name in tablet_process_names:
        command = f'tasklist /FI "IMAGENAME eq {process_name}" /NH'
        try:
            result = subprocess.run(command, shell=True, check=True, stdout=subprocess.PIPE, text=True)
            if process_name in result.stdout:
                return True
        except subprocess.CalledProcessError as e:
            print(f"Error checking process {process_name}: {e}")

    return False

def run_batch_file(file_path):
    subprocess.run(file_path, shell=True, check=True)

def minimize_opentabletdriver():
    opentabletdriver_window = gw.getWindowsWithTitle("OpenTabletDriver")[0]
    opentabletdriver_window.minimize()

def close_wacomcenter():
    process_name = "WacomCenterUI.exe"
    
    for process in psutil.process_iter(['pid', 'name']):
        if process.info['name'] == process_name:
            # Terminate the process
            process.terminate()
            print(f"Terminated {process_name} process with PID {process.info['pid']}")
            return
    
    print(f"{process_name} process not found.")

def check_and_run_opentabletdriver(opentabletdriver_path):
    try:
        # Check if "OpenTabletDriver" is running
        tasklist_output = subprocess.check_output(['tasklist'], text=True)
        if "OpenTabletDriver.UX.Wpf.exe" not in tasklist_output:
            # Run the daemon if not running
            daemon_path = r"C:\Users\Jason\Desktop\Stuff\OpenTabletDriver\OpenTabletDriver.Daemon.exe"
            subprocess.Popen([daemon_path])
            time.sleep(5)  # Adjust the delay as needed

            # Run OpenTabletDriver
            subprocess.Popen([opentabletdriver_path])
            # Wait for the main application to start (adjust the sleep time as needed)
            time.sleep(10)
            # Minimize OpenTabletDriver
            minimize_opentabletdriver()
    except subprocess.CalledProcessError as e:
        print("Error checking tasklist:", e)
    except Exception as e:
        print(f"Error running OpenTabletDriver: {e}")

# Paths to your batch file and opentabletdriver.exe
batch_file_path = r"C:\Users\Jason\Desktop\Scripts\Wacom Scripts\DisableWacomDrivers.bat"
opentabletdriver_path = r"C:\Users\Jason\Desktop\Stuff\OpenTabletDriver\OpenTabletDriver.UX.Wpf.exe"

# Flag to track tablet connection status
tablet_connected = False

while True:
    if is_tablet_connected():
        if not tablet_connected:
            tablet_connected = True
            print("Tablet connected. Proceeding with the script.")
            
            # Run the batch file, wait, close cmd window, and run it again
            run_batch_file(batch_file_path)
            
            time.sleep(20)
            
            # Close WacomCenter
            close_wacomcenter()
            
            # Check and run OpenTabletDriver if not running
            check_and_run_opentabletdriver(opentabletdriver_path)

            # Adjust the interval between iterations as needed
            time.sleep(10)  # 1 minute
        else:
            #print("Tablet still connected. Waiting...")
            time.sleep(10)  # 1 minute
    else:
        tablet_connected = False
        #print("Tablet not connected. Waiting...")
        time.sleep(10)  # 1 minute


