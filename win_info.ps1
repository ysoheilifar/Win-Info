# Get Computer Name
$computerName = $env:COMPUTERNAME
$csvPath = "$computerName`_System_Info.csv"

# Function to Print Section Header
function Print-Section($title) {
    Write-Host "`n--------------------------------------------------"
    Write-Host "$title"
    Write-Host "--------------------------------------------------"
}

# Initialize system info array
$systemInfo = @()

# System Information
Print-Section "System Information"
$osInfo = Get-WmiObject Win32_OperatingSystem
$systemInfo += [PSCustomObject]@{Property="Computer Name"; Value=$computerName}
$systemInfo += [PSCustomObject]@{Property="OS Name"; Value=$osInfo.Caption}
$systemInfo += [PSCustomObject]@{Property="OS Version"; Value=$osInfo.Version}
$systemInfo += [PSCustomObject]@{Property="Architecture"; Value=$osInfo.OSArchitecture}
$systemInfo += [PSCustomObject]@{Property="System Manufacturer"; Value=(Get-WmiObject Win32_ComputerSystem).Manufacturer}
$systemInfo | ForEach-Object { Write-Host "$($_.Property): $($_.Value)" }

# Motherboard Information
Print-Section "Motherboard Information"
Get-WmiObject Win32_BaseBoard | ForEach-Object {
    Write-Host "Model: $($_.Product)"
    Write-Host "Manufacturer: $($_.Manufacturer)"
    $systemInfo += [PSCustomObject]@{Property="Motherboard Model"; Value=$_.Product}
    $systemInfo += [PSCustomObject]@{Property="Manufacturer"; Value=$_.Manufacturer}
}

# CPU Information
Print-Section "CPU Information"
Get-WmiObject Win32_Processor | ForEach-Object {
    Write-Host "Model: $($_.Name)"
    Write-Host "Cores: $($_.NumberOfCores)"
    Write-Host "Clock Speed: $($_.MaxClockSpeed) MHz"
    $systemInfo += [PSCustomObject]@{Property="CPU Model"; Value=$_.Name}
    $systemInfo += [PSCustomObject]@{Property="CPU Cores"; Value=$_.NumberOfCores}
    $systemInfo += [PSCustomObject]@{Property="Clock Speed (MHz)"; Value=$_.MaxClockSpeed}
}

# RAM Information
Print-Section "RAM Information"
Get-WmiObject Win32_PhysicalMemory | ForEach-Object {
    Write-Host "Capacity: $([math]::Round($_.Capacity / 1GB, 2)) GB"
    Write-Host "Type: $($_.MemoryType)"
    Write-Host "Speed: $($_.Speed) MHz"
    $systemInfo += [PSCustomObject]@{Property="RAM Capacity (GB)"; Value=[math]::Round($_.Capacity / 1GB, 2)}
    $systemInfo += [PSCustomObject]@{Property="RAM Type"; Value=$_.MemoryType}
    $systemInfo += [PSCustomObject]@{Property="RAM Speed (MHz)"; Value=$_.Speed}
}
$totalRam = [math]::Round((Get-WmiObject Win32_ComputerSystem).TotalPhysicalMemory / 1GB, 2)
Write-Host "Total RAM: $totalRam GB"
$systemInfo += [PSCustomObject]@{Property="Total RAM (GB)"; Value=$totalRam}

# Storage Information
Print-Section "Storage Information"
Get-WmiObject Win32_DiskDrive | ForEach-Object {
    $diskType = "HDD"
    if ($_.Model -match "SSD") { $diskType = "SSD" }

    Write-Host "Model: $($_.Model)"
    Write-Host "Interface: $($_.InterfaceType)"
    Write-Host "Size: $([math]::Round($_.Size / 1GB, 2)) GB"
    Write-Host "Type: $diskType"
    $systemInfo += [PSCustomObject]@{Property="Storage Device"; Value=$_.Model}
    $systemInfo += [PSCustomObject]@{Property="Interface"; Value=$_.InterfaceType}
    $systemInfo += [PSCustomObject]@{Property="Size (GB)"; Value=[math]::Round($_.Size / 1GB, 2)}
    $systemInfo += [PSCustomObject]@{Property="Type"; Value=$diskType}
}

# Partition Information
Print-Section "Partition Information"
Get-WmiObject Win32_DiskPartition | ForEach-Object {
    Write-Host "Partition Name: $($_.Name)"
    Write-Host "Disk Index: $($_.DiskIndex)"
    Write-Host "Size: $([math]::Round($_.Size / 1GB, 2)) GB"
    $systemInfo += [PSCustomObject]@{Property="Partition Name"; Value=$_.Name}
    $systemInfo += [PSCustomObject]@{Property="Disk Index"; Value=$_.DiskIndex}
    $systemInfo += [PSCustomObject]@{Property="Partition Size (GB)"; Value=[math]::Round($_.Size / 1GB, 2)}
}

# Logical Drive Information
Print-Section "Logical Drive Information"
Get-WmiObject Win32_LogicalDisk | ForEach-Object {
    Write-Host "Drive: $($_.DeviceID)"
    Write-Host "File System: $($_.FileSystem)"
    Write-Host "Total Space: $([math]::Round($_.Size / 1GB, 2)) GB"
    Write-Host "Free Space: $([math]::Round($_.FreeSpace / 1GB, 2)) GB"
    $systemInfo += [PSCustomObject]@{Property="Drive Letter"; Value=$_.DeviceID}
    $systemInfo += [PSCustomObject]@{Property="File System"; Value=$_.FileSystem}
    $systemInfo += [PSCustomObject]@{Property="Total Space (GB)"; Value=[math]::Round($_.Size / 1GB, 2)}
    $systemInfo += [PSCustomObject]@{Property="Free Space (GB)"; Value=[math]::Round($_.FreeSpace / 1GB, 2)}
}

# Network Information
Print-Section "Network Information"
Get-WmiObject Win32_NetworkAdapterConfiguration | Where-Object { $_.IPEnabled -eq $true } | ForEach-Object {
    Write-Host "Adapter Name: $($_.Description)"
    Write-Host "MAC Address: $($_.MACAddress)"
    Write-Host "IP Address: $($_.IPAddress -join ', ')"
    $systemInfo += [PSCustomObject]@{Property="Adapter Name"; Value=$_.Description}
    $systemInfo += [PSCustomObject]@{Property="MAC Address"; Value=$_.MACAddress}
    $systemInfo += [PSCustomObject]@{Property="IP Address"; Value=$_.IPAddress -join ", "}
}

# GPU (Graphics Card) Information
Print-Section "GPU (Graphics Card) Information"
Get-WmiObject Win32_VideoController | ForEach-Object {
    Write-Host "GPU Name: $($_.Name)"
    Write-Host "Driver Version: $($_.DriverVersion)"
    Write-Host "Driver Date: $($_.DriverDate)"
    Write-Host "Driver Provider: $($_.DriverProvider)"
    Write-Host "Adapter RAM (MB): $([math]::Round($_.AdapterRAM / 1MB, 2))"
    Write-Host "Resolution: $($_.CurrentHorizontalResolution)x$($_.CurrentVerticalResolution)"

    $systemInfo += [PSCustomObject]@{Property="GPU Name"; Value=$_.Name}
    $systemInfo += [PSCustomObject]@{Property="Driver Version"; Value=$_.DriverVersion}
    $systemInfo += [PSCustomObject]@{Property="Driver Date"; Value=$_.DriverDate}
    $systemInfo += [PSCustomObject]@{Property="Driver Provider"; Value=$_.DriverProvider}
    $systemInfo += [PSCustomObject]@{Property="Adapter RAM (MB)"; Value=[math]::Round($_.AdapterRAM / 1MB, 2)}
    $systemInfo += [PSCustomObject]@{Property="Resolution"; Value="$($_.CurrentHorizontalResolution)x$($_.CurrentVerticalResolution)"}
}

# Display (Monitor) Information
Print-Section "Display (Monitor) Information"
Get-WmiObject Win32_DesktopMonitor | ForEach-Object {
    Write-Host "Monitor Name: $($_.Name)"
    Write-Host "Monitor Manufacturer: $($_.MonitorManufacturer)"
    Write-Host "Monitor Type: $($_.MonitorType)"
    Write-Host "Screen Width: $($_.ScreenWidth)"
    Write-Host "Screen Height: $($_.ScreenHeight)"
    Write-Host "Monitor Status: $($_.Status)"
    $systemInfo += [PSCustomObject]@{Property="Monitor Name"; Value=$_.Name}
    $systemInfo += [PSCustomObject]@{Property="Monitor Manufacturer"; Value=$_.MonitorManufacturer}
    $systemInfo += [PSCustomObject]@{Property="Monitor Type"; Value=$_.MonitorType}
    $systemInfo += [PSCustomObject]@{Property="Screen Width"; Value=$_.ScreenWidth}
    $systemInfo += [PSCustomObject]@{Property="Screen Height"; Value=$_.ScreenHeight}
    $systemInfo += [PSCustomObject]@{Property="Monitor Status"; Value=$_.Status}
}

# Printer Information
Print-Section "Printer Information"
Get-WmiObject Win32_Printer | ForEach-Object {
    Write-Host "Printer Name: $($_.Name)"
    Write-Host "Printer Status: $($_.Status)"
    Write-Host "Printer Default: $($_.Default)"
    Write-Host "Printer Port Name: $($_.PortName)"
    Write-Host "Printer Driver Name: $($_.DriverName)"
    Write-Host "Printer Network: $($_.Network)"

    $systemInfo += [PSCustomObject]@{Property="Printer Name"; Value=$_.Name}
    $systemInfo += [PSCustomObject]@{Property="Printer Status"; Value=$_.Status}
    $systemInfo += [PSCustomObject]@{Property="Default Printer"; Value=$_.Default}
    $systemInfo += [PSCustomObject]@{Property="Printer Port Name"; Value=$_.PortName}
    $systemInfo += [PSCustomObject]@{Property="Printer Driver Name"; Value=$_.DriverName}
    $systemInfo += [PSCustomObject]@{Property="Network Printer"; Value=$_.Network}
}

# Save to CSV
$systemInfo | Export-Csv -Path $csvPath -NoTypeInformation

Write-Host "`nSystem information saved to $csvPath"
