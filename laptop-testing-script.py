import platform
import psutil
import time
import multiprocessing
import os
from datetime import datetime
import re
import subprocess
import cv2  # For webcam detection
from screeninfo import get_monitors  # For screen resolution

def get_system_info():
    system_info = "System Information:\n"

    # Get basic system information
    system = platform.system()
    
    # Better Windows version detection
    try:
        reg_key = subprocess.check_output("reg query \"HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\" /v CurrentBuild", shell=True).decode()
        build_number = int(reg_key.split()[-1])
        
        # Windows 11 started from build 22000
        if system == "Windows" and build_number >= 22000:
            system_info += "Operating System: Windows 11\n"
        else:
            system_info += f"Operating System: {system}\n"
            
        system_info += f"OS Build Number: {build_number}\n"
        
    except:
        # Fallback to platform info if reg query fails
        system_info += f"Operating System: {system}\n"
        system_info += f"OS Release: {platform.release()}\n"
        system_info += f"OS Version: {platform.version()}\n"

    # Rest of the system info gathering remains the same
    try:
        manufacturer = subprocess.check_output("wmic computersystem get manufacturer", shell=True).decode()
        manufacturer = manufacturer.split('\n')[1].strip()
        
        model = subprocess.check_output("wmic computersystem get model", shell=True).decode()
        model = model.split('\n')[1].strip()
        
        serial = subprocess.check_output("wmic bios get serialnumber", shell=True).decode()
        serial = serial.split('\n')[1].strip()
        
        product_number = subprocess.check_output("wmic baseboard get product", shell=True).decode()
        product_number = product_number.split('\n')[1].strip()
    except:
        manufacturer = "Unknown"
        model = "Unknown"
        serial = "Unknown"
        product_number = "Unknown"

    system_info += f"Manufacturer: {manufacturer}\n"
    system_info += f"Model: {model}\n"
    system_info += f"Serial Number: {serial}\n"
    system_info += f"Product Number: {product_number}\n"

    return system_info

def get_cpu_info():
    cpu_info = "CPU Information:\n"
    
    # Get CPU details using wmic
    try:
        processor = subprocess.check_output("wmic cpu get name", shell=True).decode()
        processor = processor.split('\n')[1].strip()
        cpu_info += f"Processor: {processor}\n"
    except:
        cpu_info += "Processor: Unable to retrieve\n"
    
    # Get core information
    cpu_info += f"Physical cores: {psutil.cpu_count(logical=False)}\n"
    cpu_info += f"Total cores: {psutil.cpu_count(logical=True)}\n"
    
    # Get frequency information
    cpu_freq = psutil.cpu_freq()
    if cpu_freq:
        cpu_info += f"Max Frequency: {cpu_freq.max:.2f} MHz\n"
        cpu_info += f"Min Frequency: {cpu_freq.min:.2f} MHz\n"
        cpu_info += f"Current Frequency: {cpu_freq.current:.2f} MHz\n"
    
    return cpu_info

def get_memory_info():
    memory_info = "Memory Information:\n"

    # Total physical memory
    total_memory = psutil.virtual_memory().total / (1024**3)
    memory_info += f"Total Physical Memory: {total_memory:.2f} GB\n"

    # Get detailed memory information using wmic
    try:
        result = subprocess.check_output("wmic memorychip get capacity,speed,manufacturer", shell=True).decode()
        memory_info += "\nDetailed Memory Information:\n"
        memory_info += result
    except Exception as e:
        memory_info += f"\nError retrieving detailed memory information: {str(e)}\n"

    return memory_info

def get_gpu_info():
    gpu_info = "GPU Information:\n"
    try:
        result = subprocess.check_output("wmic path win32_videocontroller get caption,driverversion", shell=True).decode()
        gpu_info += result
    except Exception as e:
        gpu_info += f"Error retrieving GPU information: {str(e)}\n"
    return gpu_info

def get_detailed_io_info():
    io_info = "Detailed I/O Information:\n"

    # Hard Drives
    io_info += "\nHard Drives:\n"
    for partition in psutil.disk_partitions():
        try:
            usage = psutil.disk_usage(partition.mountpoint)
            io_info += f"Device: {partition.device}\n"
            io_info += f"  Mountpoint: {partition.mountpoint}\n"
            io_info += f"  File System: {partition.fstype}\n"
            io_info += f"  Total Size: {usage.total / (1024**3):.2f} GB\n"
            io_info += f"  Used: {usage.used / (1024**3):.2f} GB\n"
            io_info += f"  Free: {usage.free / (1024**3):.2f} GB\n\n"
        except PermissionError:
            io_info += f"Device: {partition.device} (Access Denied)\n\n"

    # Network Adapters
    io_info += "\nNetwork Adapters:\n"
    try:
        result = subprocess.check_output("wmic nic get name,manufacturer", shell=True).decode()
        io_info += result
    except Exception as e:
        io_info += f"Error retrieving network adapter information: {str(e)}\n"

    return io_info

def get_battery_health():
    battery_info = "Battery Health Information:\n"
    try:
        # Get basic battery information using psutil
        battery = psutil.sensors_battery()
        if battery:
            battery_info += f"Battery percentage: {battery.percent}%\n"
            battery_info += f"Power plugged in: {battery.power_plugged}\n"
            if battery.secsleft != -1:
                hours = battery.secsleft // 3600
                minutes = (battery.secsleft % 3600) // 60
                battery_info += f"Time left: {hours}h {minutes}m\n"
        
        # Try using powercfg for more detailed battery report
        try:
            # Generate battery report
            report_path = os.path.join(os.getenv('TEMP'), 'battery-report.xml')
            subprocess.run(['powercfg', '/batteryreport', '/xml', '/output', report_path], capture_output=True)
            
            if os.path.exists(report_path):
                with open(report_path, 'r') as f:
                    report_content = f.read()
                    
                # Parse design capacity and current capacity from the XML
                design_pattern = r'<DesignCapacity>(\d+)</DesignCapacity>'
                current_pattern = r'<FullChargeCapacity>(\d+)</FullChargeCapacity>'
                
                design_match = re.search(design_pattern, report_content)
                current_match = re.search(current_pattern, report_content)
                
                if design_match and current_match:
                    design_cap = int(design_match.group(1))
                    current_cap = int(current_match.group(1))
                    
                    wear_level = ((design_cap - current_cap) / design_cap) * 100
                    battery_info += f"\nBattery Design Capacity: {design_cap} mWh\n"
                    battery_info += f"Current Full Capacity: {current_cap} mWh\n"
                    battery_info += f"Battery Wear Level: {wear_level:.2f}%\n"
                    
                    # Add health status assessment
                    if wear_level < 20:
                        battery_info += "Battery Health Status: Good\n"
                    elif wear_level < 40:
                        battery_info += "Battery Health Status: Fair\n"
                    else:
                        battery_info += "Battery Health Status: Poor - Consider replacement\n"
                
                # Clean up the temporary file
                os.remove(report_path)
            
        except Exception as e:
            battery_info += f"\nCould not generate detailed battery report: {str(e)}\n"
            
    except Exception as e:
        battery_info += f"Error retrieving battery information: {str(e)}\n"
    
    return battery_info

def get_display_info():
    display_info = "Display Information:\n"
    try:
        monitors = get_monitors()
        for i, monitor in enumerate(monitors, 1):
            display_info += f"Display {i}:\n"
            display_info += f"Name: {monitor.name}\n"
            display_info += f"Resolution: {monitor.width}x{monitor.height}\n"
            if hasattr(monitor, 'refresh_rate') and monitor.refresh_rate:
                display_info += f"Refresh Rate: {monitor.refresh_rate}Hz\n"
    except Exception:
        try:
            result = subprocess.check_output("wmic path Win32_VideoController get VideoModeDescription,CurrentRefreshRate", shell=True).decode()
            if result.strip():
                display_info += result
        except:
            display_info += "Unable to retrieve display information\n"
    
    return display_info

def get_webcam_info():
    webcam_info = "Webcam Information:\n"
    try:
        webcam_list = []
        index = 0
        
        while True:
            cap = cv2.VideoCapture(index)
            if not cap.isOpened():
                break
            
            width = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
            height = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
            fps = int(cap.get(cv2.CAP_PROP_FPS))
            
            webcam_list.append({
                'index': index,
                'resolution': f"{width}x{height}",
                'fps': fps
            })
            
            cap.release()
            index += 1
        
        if webcam_list:
            for cam in webcam_list:
                webcam_info += f"Webcam {cam['index']}:\n"
                webcam_info += f"Resolution: {cam['resolution']}\n"
                webcam_info += f"FPS: {cam['fps']}\n"
        else:
            webcam_info += "No webcams detected\n"
            
        if platform.system() == 'Windows':
            try:
                result = subprocess.check_output("wmic path Win32_PnPEntity where (deviceid like '%USB%') get caption", shell=True).decode()
                webcam_devices = list(set([line.strip() for line in result.split('\n') if 'camera' in line.lower() or 'webcam' in line.lower()]))
                
                if webcam_devices:
                    webcam_info += "Detected Webcam Devices:\n"
                    for device in webcam_devices:
                        if device:
                            webcam_info += f"{device}\n"
            except:
                pass
                
    except Exception as e:
        webcam_info += f"Errors detecting webcams: {str(e)}\n"
    
    return webcam_info

def main():
    print("\nGathering laptop information...")
    
    info = "Laptop Testing Script Results\n"
    info += "==============================\n\n"
    info += get_system_info() + "\n"
    info += get_cpu_info() + "\n"
    info += get_memory_info() + "\n"
    info += get_gpu_info() + "\n"
    info += get_detailed_io_info() + "\n"
    info += get_battery_health() + "\n"
    info += get_display_info() + "\n"
    info += get_webcam_info() + "\n"
    
    # Get the path to the user's Documents folder
    documents_path = os.path.expanduser("~/Documents")
    
    # Create a filename with the current date and time
    filename = f"laptop_info_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    
    # Full path to the output file
    file_path = os.path.join(documents_path, filename)
    
    # Write the information to the file
    with open(file_path, "w") as f:
        f.write(info)
    
    print(f"Information gathered and saved to: {file_path}")
    print("Here's a summary of the information:")
    print(info)

if __name__ == "__main__":
    main()
