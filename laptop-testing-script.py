import platform
import psutil
import time
import multiprocessing
import wmi
import GPUtil
import os
from datetime import datetime
import re
import platform
import wmi

def get_system_info():
    c = wmi.WMI()
    system_info = "System Information:\n"

    # Get basic system information
    system = platform.system()
    release = platform.release()
    version = platform.version()

    system_info += f"Operating System: {system}\n"
    system_info += f"OS Release: {release}\n"
    system_info += f"OS Version: {version}\n"

    # Get detailed computer system information
    computer_info = c.Win32_ComputerSystem()[0]
    bios_info = c.Win32_BIOS()[0]

    # Manufacturer and Model
    system_info += f"Manufacturer: {computer_info.Manufacturer}\n"
    system_info += f"Model: {computer_info.Model}\n"

    # Serial Number
    if hasattr(bios_info, 'SerialNumber'):
        system_info += f"Serial Number: {bios_info.SerialNumber}\n"
    else:
        system_info += "Serial Number: Not available\n"

    # Product Code (SKU)
    if hasattr(computer_info, 'SystemSKUNumber'):
        system_info += f"Product Code (SKU): {computer_info.SystemSKUNumber}\n"
    else:
        system_info += "Product Code (SKU): Not available\n"

    return system_info

def get_cpu_info():
    c = wmi.WMI()
    cpu_info = "CPU Information:\n"
    
    # Get CPU details
    cpu_details = c.Win32_Processor()[0]
    
    # Get basic CPU information
    cpu_info += f"Processor: {cpu_details.Name}\n"
    cpu_info += f"Manufacturer: {cpu_details.Manufacturer}\n"
    cpu_info += f"Base Clock Speed: {cpu_details.MaxClockSpeed} MHz\n"
    
    # Determine CPU family and generation
    cpu_name = cpu_details.Name
    cpu_family = "Unknown"
    cpu_generation = "Unknown"
    
    # Regular expressions for different CPU families
    intel_core_regex = r'(i[3579])-(\d{4,5})'
    intel_pentium_regex = r'Pentium (\w+)?'
    intel_celeron_regex = r'Celeron (\w+)?'
    intel_xeon_regex = r'Xeon (\w+)?'
    
    if 'Intel' in cpu_name:
        core_match = re.search(intel_core_regex, cpu_name)
        if core_match:
            cpu_family = core_match.group(1)
            model_number = core_match.group(2)
            cpu_generation = int(model_number[0])
        elif re.search(intel_pentium_regex, cpu_name):
            cpu_family = "Pentium"
        elif re.search(intel_celeron_regex, cpu_name):
            cpu_family = "Celeron"
        elif re.search(intel_xeon_regex, cpu_name):
            cpu_family = "Xeon"
    elif 'AMD' in cpu_name:
        cpu_family = "AMD"
        # Add more specific AMD detection if needed
    
    cpu_info += f"CPU Family: {cpu_family}\n"
    if cpu_generation != "Unknown":
        cpu_info += f"Generation: {cpu_generation}\n"
    
    # Get core information
    cpu_info += f"Physical cores: {psutil.cpu_count(logical=False)}\n"
    cpu_info += f"Total cores: {psutil.cpu_count(logical=True)}\n"
    
    # Get frequency information
    cpu_freq = psutil.cpu_freq()
    cpu_info += f"Max Frequency: {cpu_freq.max:.2f} MHz\n"
    cpu_info += f"Min Frequency: {cpu_freq.min:.2f} MHz\n"
    cpu_info += f"Current Frequency: {cpu_freq.current:.2f} MHz\n"
    
    return cpu_info  # Add this return statement

import psutil
import wmi

def get_memory_info():
    memory_info = "Memory Information:\n"
    c = wmi.WMI()

    try:
        # Total physical memory
        total_memory = psutil.virtual_memory().total / (1024**3)
        memory_info += f"Total Physical Memory: {total_memory:.2f} GB\n\n"

        # Attempt to get detailed memory information
        physical_memory = c.Win32_PhysicalMemory()
        
        if physical_memory:
            for i, mem in enumerate(physical_memory, start=1):
                memory_info += f"RAM Slot #{i}:\n"
                
                # Capacity
                capacity_gb = mem.Capacity / (1024**3) if mem.Capacity else "Unknown"
                memory_info += f"  Capacity: {capacity_gb:.2f} GB\n" if isinstance(capacity_gb, float) else f"  Capacity: {capacity_gb}\n"
                
                # Type
                memory_type = mem.MemoryType
                type_dict = {0: "Unknown", 20: "DDR", 21: "DDR2", 22: "DDR2 FB-DIMM", 24: "DDR3", 26: "DDR4"}
                memory_info += f"  Type: {type_dict.get(memory_type, f'Unknown ({memory_type})')}\n"
                
                # Speed
                speed = mem.Speed or "Unknown"
                memory_info += f"  Speed: {speed} MHz\n"
                
                # Manufacturer (Brand)
                manufacturer = mem.Manufacturer or "Unknown"
                memory_info += f"  Manufacturer: {manufacturer}\n"
                
                memory_info += "\n"
        else:
            memory_info += "Detailed memory information not available.\n"

    except Exception as e:
        memory_info += f"Error retrieving detailed memory information: {str(e)}\n"
        memory_info += "Only basic memory information is available.\n"

    return memory_info

def get_gpu_info():
    gpus = GPUtil.getGPUs()
    gpu_info = "GPU Information:\n"
    if gpus:
        for gpu in gpus:
            gpu_info += f"GPU: {gpu.name}\nMemory Total: {gpu.memoryTotal} MB\nMemory Free: {gpu.memoryFree} MB\nMemory Used: {gpu.memoryUsed} MB\n\n"
    else:
        gpu_info += "No GPU detected\n"
    return gpu_info

import wmi
import psutil

def get_detailed_io_info():
    c = wmi.WMI()
    io_info = "Detailed I/O Information:\n"

    # USB Controllers (including Thunderbolt and Type-C)
    usb_controllers = c.Win32_USBController()
    usb_hubs = c.Win32_USBHub()
    io_info += f"USB Controllers and Ports:\n"
    
    type_c_count = 0
    for hub in usb_hubs:
        if hub.DeviceID.startswith("USB\\ROOT_HUB30"):  # USB 3.0 and above
            type_c_count += 1

    for controller in usb_controllers:
        io_info += f"  - {controller.Name}\n"
        if 'thunderbolt' in controller.Name.lower():
            io_info += "    Thunderbolt USB detected (likely Type-C)\n"
        elif '3.1' in controller.Name or '3.2' in controller.Name:
            io_info += "    USB 3.1/3.2 detected (possibly Type-C)\n"
        elif '3.0' in controller.Name:
            io_info += "    USB 3.0 (SuperSpeed)\n"
        elif '2.0' in controller.Name:
            io_info += "    USB 2.0 (High Speed)\n"
        else:
            io_info += "    USB version unknown\n"

    if type_c_count > 0:
        io_info += f"  Potential USB Type-C ports detected: {type_c_count}\n"
    else:
        io_info += "  No USB Type-C ports definitively detected\n"

    # Network Adapters
    network_adapters = c.Win32_NetworkAdapter(PhysicalAdapter=True)
    io_info += f"\nNetwork Adapters: {len(network_adapters)}\n"
    for adapter in network_adapters:
        io_info += f"  - {adapter.Name}\n"
        if adapter.Speed:
            try:
                speed = int(adapter.Speed) / 1_000_000
                io_info += f"    Speed: {speed:.2f} Mbps\n"
            except ValueError:
                io_info += f"    Speed: {adapter.Speed} (unable to convert to Mbps)\n"

    # Video Outputs
    video_outputs = c.Win32_VideoController()
    io_info += "\nVideo Outputs:\n"
    for output in video_outputs:
        io_info += f"  - {output.Name}\n"
        try:
            if output.VideoOutputTechnology:
                tech = output.VideoOutputTechnology
                if tech == 3:
                    io_info += "    VGA\n"
                elif tech == 4:
                    io_info += "    DVI\n"
                elif tech == 5:
                    io_info += "    HDMI\n"
                elif tech == 6:
                    io_info += "    SVIDEO\n"
                elif tech == 8:
                    io_info += "    DisplayPort\n"
                else:
                    io_info += f"    Other (Code: {tech})\n"
            else:
                io_info += "    Output technology information not available\n"
        except AttributeError:
            io_info += "    Output technology information not available\n"

    # Audio Devices
    sound_devices = c.Win32_SoundDevice()
    io_info += f"\nAudio Devices: {len(sound_devices)}\n"
    for device in sound_devices:
        io_info += f"  - {device.Name}\n"

    # Hard Drives (including external)
    io_info += "\nHard Drives:\n"
    for disk in psutil.disk_partitions():
        try:
            disk_usage = psutil.disk_usage(disk.mountpoint)
            io_info += f"  - Drive: {disk.device}\n"
            io_info += f"    Mountpoint: {disk.mountpoint}\n"
            io_info += f"    File System: {disk.fstype}\n"
            io_info += f"    Total Size: {disk_usage.total / (1024**3):.2f} GB\n"
            io_info += f"    Used: {disk_usage.used / (1024**3):.2f} GB\n"
            io_info += f"    Free: {disk_usage.free / (1024**3):.2f} GB\n"
        except PermissionError:
            io_info += f"  - Drive: {disk.device} (Access Denied)\n"

    return io_info

def get_battery_health():
    c = wmi.WMI()
    batteries = c.Win32_Battery()
    
    if not batteries:
        return "Battery Information:\nNo battery detected\n"
    
    battery_info = "Battery Health Information:\n"
    for battery in batteries:
        battery_info += f"Battery Name: {battery.Name}\n"
        
        try:
            # Get design capacity
            design_capacity = battery.DesignCapacity
            
            # Get full charge capacity
            full_charge_capacity = battery.FullChargeCapacity
            
            if design_capacity and full_charge_capacity:
                health_percentage = (full_charge_capacity / design_capacity) * 100
                battery_info += f"Battery Health: {health_percentage:.2f}%\n"
            else:
                battery_info += "Battery Health: Unable to calculate (missing capacity information)\n"
        
        except Exception as e:
            battery_info += f"Battery Health: Unable to retrieve (Error: {str(e)})\n"
        
        battery_info += "\n"  # Add a newline for readability if there are multiple batteries
    
    return battery_info

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
    
    # Get the path to the user's Documents folder
    documents_path = os.path.join(os.path.expanduser("~"), "Documents")
    
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