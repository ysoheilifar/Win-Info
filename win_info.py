import wmi
import psutil
import platform
import csv
import os

# Get Computer Name
computer_name = os.getenv("COMPUTERNAME")
csv_filename = f"{computer_name}_system_info.csv"

# Function to Print Section Header
def print_section(title):
    print("\n" + "-" * 50)
    print(title)
    print("-" * 50)

# Gather system info
system_info = []

# System Information
print_section("System Information")
computer = wmi.WMI()
os_info = computer.Win32_OperatingSystem()[0]
system_info.append(["Computer Name", computer_name])
system_info.append(["OS Name", os_info.Caption])
system_info.append(["OS Version", os_info.Version])
system_info.append(["Architecture", platform.architecture()[0]])
system_info.append(["System Type", os_info.OSArchitecture])
system_info.append(["System Manufacturer", computer.Win32_ComputerSystem()[0].Manufacturer])
for item in system_info[-6:]: print(f"{item[0]}: {item[1]}")

# Motherboard Information
print_section("Motherboard Information")
for board in computer.Win32_BaseBoard():
    system_info.append(["Motherboard Model", board.Product])
    system_info.append(["Manufacturer", board.Manufacturer])
    print(f"Model: {board.Product}")
    print(f"Manufacturer: {board.Manufacturer}")

# CPU Information
print_section("CPU Information")
for cpu in computer.Win32_Processor():
    system_info.append(["CPU Model", cpu.Name])
    system_info.append(["Number of Cores", cpu.NumberOfCores])
    system_info.append(["Clock Speed (MHz)", cpu.MaxClockSpeed])
    print(f"Model: {cpu.Name}")
    print(f"Cores: {cpu.NumberOfCores}")
    print(f"Clock Speed: {cpu.MaxClockSpeed} MHz")

# RAM Information
print_section("RAM Information")
for mem in computer.Win32_PhysicalMemory():
    system_info.append(["RAM Capacity (GB)", f"{int(mem.Capacity) / (1024**3):.2f}"])
    system_info.append(["RAM Type", mem.MemoryType])
    system_info.append(["RAM Speed (MHz)", mem.Speed])
    print(f"Capacity: {int(mem.Capacity) / (1024**3):.2f} GB")
    print(f"Type: {mem.MemoryType}")
    print(f"Speed: {mem.Speed} MHz")
total_ram = f"{psutil.virtual_memory().total / (1024**3):.2f} GB"
system_info.append(["Total RAM (GB)", total_ram])
print(f"Total RAM: {total_ram}")

# Storage Information
print_section("Storage Information")
for disk in computer.Win32_DiskDrive():
    disk_type = "SSD" if "SSD" in disk.Model.upper() else "HDD"
    system_info.append(["Storage Device", disk.Model])
    system_info.append(["Storage Interface", disk.InterfaceType])
    system_info.append(["Storage Size (GB)", f"{int(disk.Size) / (1024**3):.2f}"])
    system_info.append(["Storage Type", disk_type])
    print(f"Model: {disk.Model}")
    print(f"Interface: {disk.InterfaceType}")
    print(f"Size: {int(disk.Size) / (1024**3):.2f} GB")
    print(f"Type: {disk_type}")

# Partition Information
print_section("Partition Information")
for partition in computer.Win32_DiskPartition():
    system_info.append(["Partition Name", partition.Name])
    system_info.append(["Disk Index", partition.DiskIndex])
    system_info.append(["Partition Size (GB)", f"{int(partition.Size) / (1024**3):.2f}"])
    print(f"Partition Name: {partition.Name}")
    print(f"Disk Index: {partition.DiskIndex}")
    print(f"Size: {int(partition.Size) / (1024**3):.2f} GB")

# Logical Drive Information
print_section("Logical Drive Information")
for logical_disk in computer.Win32_LogicalDisk():
    system_info.append(["Drive Letter", logical_disk.DeviceID])
    system_info.append(["File System", logical_disk.FileSystem])
    system_info.append(["Total Space (GB)", f"{int(logical_disk.Size) / (1024**3):.2f}"])
    system_info.append(["Free Space (GB)", f"{int(logical_disk.FreeSpace) / (1024**3):.2f}"])
    print(f"Drive: {logical_disk.DeviceID}")
    print(f"File System: {logical_disk.FileSystem}")
    print(f"Total Space: {int(logical_disk.Size) / (1024**3):.2f} GB")
    print(f"Free Space: {int(logical_disk.FreeSpace) / (1024**3):.2f} GB")

# Network Information
print_section("Network Information")
for adapter in computer.Win32_NetworkAdapterConfiguration():
    if adapter.IPEnabled:
        system_info.append(["Adapter Name", adapter.Description])
        system_info.append(["MAC Address", adapter.MACAddress])
        system_info.append(["IP Address", ", ".join(adapter.IPAddress)])
        system_info.append(["Subnet Mask", ", ".join(adapter.IPSubnet)])
        system_info.append(["Default Gateway", ", ".join(adapter.DefaultIPGateway) if adapter.DefaultIPGateway else "N/A"])
        system_info.append(["DNS Servers", ", ".join(adapter.DNSServerSearchOrder) if isinstance(adapter.DNSServerSearchOrder, list) else "N/A"])
        print(f"Adapter Name: {adapter.Description}")
        print(f"MAC Address: {adapter.MACAddress}")
        print(f"IP Address: {', '.join(adapter.IPAddress)}")
        print(f"Subnet Mask: {', '.join(adapter.IPSubnet)}")
        print(f"Default Gateway: {', '.join(adapter.DefaultIPGateway) if isinstance(adapter.DefaultIPGateway, list) else 'N/A'}")
        print(f"DNS Servers: {', '.join(adapter.DNSServerSearchOrder) if isinstance(adapter.DNSServerSearchOrder, list) else 'N/A'}")

# GPU Information
print_section("GPU (Graphics Card) Information")
for gpu in computer.Win32_VideoController():
    system_info.append(["GPU Name", gpu.Name])
    system_info.append(["Driver Version", gpu.DriverVersion])
    system_info.append(["Driver Date", gpu.DriverDate])
    system_info.append(["Driver Provider", gpu.Description])
    system_info.append(["Adapter RAM (MB)", f"{int(gpu.AdapterRAM) / (1024**2):.2f}"])
    system_info.append(["Resolution", f"{gpu.CurrentHorizontalResolution}x{gpu.CurrentVerticalResolution}"])
    print(f"GPU Name: {gpu.Name}")
    print(f"Driver Version: {gpu.DriverVersion}")
    print(f"Driver Date: {gpu.DriverDate}")
    print(f"Driver Provider: {gpu.Description}")
    print(f"Adapter RAM: {int(gpu.AdapterRAM) / (1024**2):.2f} MB")
    print(f"Resolution: {gpu.CurrentHorizontalResolution}x{gpu.CurrentVerticalResolution}")

# Monitor Information
print_section("Monitor Information")
for monitor in computer.Win32_DesktopMonitor():
    system_info.append(["Monitor Name", monitor.Name])
    system_info.append(["Monitor Manufacturer", monitor.MonitorManufacturer])
    system_info.append(["Monitor Type", monitor.MonitorType])
    system_info.append(["Screen Width", monitor.ScreenWidth])
    system_info.append(["Screen Height", monitor.ScreenHeight])
    system_info.append(["Monitor Status", monitor.Status])
    print(f"Monitor Name: {monitor.Name}")
    print(f"Monitor Manufacturer: {monitor.MonitorManufacturer}")
    print(f"Monitor Type: {monitor.MonitorType}")
    print(f"Screen Width: {monitor.ScreenWidth}")
    print(f"Screen Height: {monitor.ScreenHeight}")
    print(f"Monitor Status: {monitor.Status}")

# Printer Information
print_section("Printer Information")
for printer in computer.Win32_Printer():
    system_info.append(["Printer Name", printer.Name])
    system_info.append(["Printer Status", printer.Status])
    system_info.append(["Default Printer", printer.Default])
    system_info.append(["Printer Port Name", printer.PortName])
    system_info.append(["Printer Driver Name", printer.DriverName])
    system_info.append(["Network Printer", printer.Network])
    print(f"Printer Name: {printer.Name}")
    print(f"Printer Status: {printer.Status}")
    print(f"Default Printer: {printer.Default}")
    print(f"Printer Port Name: {printer.PortName}")
    print(f"Printer Driver Name: {printer.DriverName}")
    print(f"Network Printer: {printer.Network}")

# Save to CSV
with open(csv_filename, "w", newline="") as file:
    writer = csv.writer(file)
    writer.writerow(["Property", "Value"])
    writer.writerows(system_info)

print(f"\nSystem information saved to {csv_filename}")
