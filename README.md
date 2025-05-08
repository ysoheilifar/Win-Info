# Windows System & Hardware Information Retriever
This repository contains Python and PowerShell scripts that retrieve detailed system, hardware, storage, network, GPU, monitor, and printer information on Windows machines. These scripts display information section-by-section and save results to a structured CSV file for reference

### Features
- System Details – OS version, architecture, and manufacturer
- Motherboard Information – Model and manufacturer
- CPU Specs – Model, cores, clock speed
- RAM Details – Capacity, type, and speed
- Storage & Partition Info – HDD/SSD detection, total space, partitions
- Logical Drive Info – File system, available/free space
- Network Details – Adapter name, MAC address, IP configuration
- GPU Information – Name, driver version, resolution
- Monitor Information – Manufacturer, model, screen resolution
- Printer Details – Name, status, driver, default printer

### Scripts
1. Python Script (win_info.py)
- Uses wmi and psutil to retrieve system data
- Displays information in a structured format
- Saves results to <Computer_Name>_system_info.csv

2. PowerShell Script (win_info.ps1)
- Uses Get-WmiObject to fetch system details
- Organizes output into sections for readability
- Saves system data to <Computer_Name>_System_Info.csv

### How to Use:
#### Python
1. Install dependencies:
```python
pip install wmi psutil
```
2. Run the script:
```python
python win_info.py
```
3. View system details in the terminal
4. Check the generated CSV file in the script director

#### PowerShell
1.  Open PowerShell as Administrator
2.  Run:
```powershell
.\win_info.ps1
```
3. View detailed system specs in the terminal
4. CSV file will be created in the current director
