# PowerShell Script to generate HTML report of installed software
# Created: $(Get-Date -Format "yyyy-MM-dd")

# Get hostname and generate timestamp for the report
$hostname = $env:COMPUTERNAME
$dateTimeFormat = Get-Date -Format "yyyyMMdd-HHmmss"
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$reportFile = "AkijSystemInfo_$hostname`_$dateTimeFormat.html"
$reportPath = Join-Path -Path $PWD -ChildPath $reportFile

# Get system information
try {
    Write-Host "Retrieving system information..."
    $systemInfo = Get-WmiObject -Class Win32_ComputerSystem | Select-Object Name, Manufacturer, Model
    $osInfo = Get-WmiObject -Class Win32_OperatingSystem | Select-Object Caption, Version, OSArchitecture, LastBootUpTime
    $processorInfo = Get-WmiObject -Class Win32_Processor | Select-Object Name, NumberOfCores, NumberOfLogicalProcessors
    $memoryInfo = Get-WmiObject -Class Win32_ComputerSystem | Select-Object TotalPhysicalMemory
    $diskInfo = Get-WmiObject -Class Win32_LogicalDisk -Filter "DriveType=3" | Select-Object DeviceID, VolumeName, Size, FreeSpace
}
catch {
    Write-Error "Error retrieving system information: $_"
}

# Format memory size
$totalMemoryGB = [math]::Round($memoryInfo.TotalPhysicalMemory / 1GB, 2)

# Retrieve installed software information
try {
    Write-Host "Retrieving installed software information..."
    $installedSoftware = Get-WmiObject -Class Win32_Product | Select-Object Name, Version, Vendor, InstallDate
    
    if (-not $installedSoftware) {
        Write-Warning "No software information found."
        exit
    }
}
catch {
    Write-Error "Error retrieving software information: $_"
    exit
}

# Function to format install date
function Format-InstallDate {
    param([string]$dateString)
    
    if ([string]::IsNullOrEmpty($dateString)) {
        return "N/A"
    }
    
    # WMI dates are in format YYYYMMDD
    if ($dateString -match '^\d{8}$') {
        try {
            $year = $dateString.Substring(0, 4)
            $month = $dateString.Substring(4, 2)
            $day = $dateString.Substring(6, 2)
            return "$year-$month-$day"
        }
        catch {
            return $dateString
        }
    }
    
    return $dateString
}

# Create HTML report content
$htmlContent = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>AKIJ Group System Information Report</title>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700&display=swap');
        body {
            font-family: 'Roboto', Arial, sans-serif;
            line-height: 1.6;
            margin: 20px;
            color: #333;
            background-color: #f9f9f9;
        }
        .container {
            max-width: 1200px;
            margin: 0 auto;
            background-color: white;
            padding: 20px;
            box-shadow: 0 0 10px rgba(0,0,0,0.1);
            border-radius: 5px;
        }
        .header {
            text-align: center;
            margin-bottom: 30px;
            padding: 20px;
            background-color: #2c3e50;
            color: white;
            border-radius: 5px;
        }
        .header h1 {
            margin: 0;
            padding: 0;
            color: white;
            border: none;
        }
        .header p {
            margin: 10px 0 0 0;
            font-size: 1.2em;
        }
        .section {
            margin-bottom: 30px;
        }
        .section-header {
            background-color: #3498db;
            color: white;
            padding: 10px 15px;
            font-size: 1.2em;
            border-radius: 5px 5px 0 0;
            margin-bottom: 0;
        }
        h1 {
            color: #2c3e50;
            border-bottom: 2px solid #3498db;
            padding-bottom: 10px;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin: 20px 0;
            box-shadow: 0 2px 3px rgba(0, 0, 0, 0.1);
        }
        th {
            background-color: #3498db;
            color: white;
            font-weight: bold;
            padding: 10px;
            text-align: left;
        }
        tr:nth-child(even) {
            background-color: #f2f2f2;
        }
        tr:hover {
            background-color: #e3f2fd;
        }
        td {
            padding: 8px 10px;
            border-bottom: 1px solid #ddd;
        }
        .system-info {
            background-color: #f8f9fa;
            padding: 15px;
            border-radius: 0 0 5px 5px;
            border: 1px solid #ddd;
            margin-top: 0;
        }
        .system-info-table {
            width: 100%;
            border-collapse: collapse;
        }
        .system-info-table td {
            padding: 8px;
            border-bottom: 1px solid #ddd;
        }
        .system-info-table tr:last-child td {
            border-bottom: none;
        }
        .system-info-label {
            font-weight: bold;
            width: 30%;
        }
        .footer {
            margin-top: 30px;
            padding: 20px;
            font-size: 0.9em;
            color: #7f8c8d;
            text-align: center;
            border-top: 1px solid #eee;
            background-color: #f8f9fa;
            border-radius: 0 0 5px 5px;
        }
        .copyright {
            margin-top: 10px;
            font-weight: bold;
        }
        .author-info {
            margin-top: 10px;
        }
        @media print {
            body {
                margin: 0;
                padding: 10px;
            }
            .no-print {
                display: none;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>AKIJ Group System Information Report</h1>
            <p>Generated on $timestamp for system: $hostname</p>
        </div>
        
        <div class="section">
            <h2 class="section-header">System Overview</h2>
            <div class="system-info">
                <table class="system-info-table">
                    <tr>
                        <td class="system-info-label">Computer Name</td>
                        <td>$($systemInfo.Name)</td>
                    </tr>
                    <tr>
                        <td class="system-info-label">Manufacturer</td>
                        <td>$($systemInfo.Manufacturer)</td>
                    </tr>
                    <tr>
                        <td class="system-info-label">Model</td>
                        <td>$($systemInfo.Model)</td>
                    </tr>
                    <tr>
                        <td class="system-info-label">Operating System</td>
                        <td>$($osInfo.Caption)</td>
                    </tr>
                    <tr>
                        <td class="system-info-label">OS Version</td>
                        <td>$($osInfo.Version)</td>
                    </tr>
                    <tr>
                        <td class="system-info-label">OS Architecture</td>
                        <td>$($osInfo.OSArchitecture)</td>
                    </tr>
                    <tr>
                        <td class="system-info-label">Processor</td>
                        <td>$($processorInfo.Name)</td>
                    </tr>
                    <tr>
                        <td class="system-info-label">CPU Cores</td>
                        <td>$($processorInfo.NumberOfCores) Physical, $($processorInfo.NumberOfLogicalProcessors) Logical</td>
                    </tr>
                    <tr>
                        <td class="system-info-label">Memory</td>
                        <td>$totalMemoryGB GB</td>
                    </tr>
                </table>
            </div>
        </div>
        
        <div class="section">
            <h2 class="section-header">Installed Software Report</h2>
            <table>
        <thead>
            <tr>
                <th>Name</th>
                <th>Version</th>
                <th>Publisher</th>
                <th>Install Date</th>
            </tr>
        </thead>
        <tbody>
"@

# Add rows for each software
foreach ($software in $installedSoftware) {
    $name = if ([string]::IsNullOrEmpty($software.Name)) { "N/A" } else { $software.Name }
    $version = if ([string]::IsNullOrEmpty($software.Version)) { "N/A" } else { $software.Version }
    $publisher = if ([string]::IsNullOrEmpty($software.Vendor)) { "N/A" } else { $software.Vendor }
    $installDate = Format-InstallDate -dateString $software.InstallDate
    
    $htmlContent += @"
            <tr>
                <td>$name</td>
                <td>$version</td>
                <td>$publisher</td>
                <td>$installDate</td>
            </tr>
"@
}

# Complete the HTML content
$htmlContent += @"
        </tbody>
    </table>
        </div>
        
        <div class="footer">
            <div class="copyright">&copy; 2025 AKIJ IT Team. All rights reserved.</div>
            <div class="author-info">
                Author: Boni Yeamin<br>
                Generated by AKIJ IT Team<br>
                Report Date: $timestamp
            </div>
        </div>
    </div>
</body>
</html>
"@

# Write HTML to file
try {
    Write-Host "Generating HTML report..."
    # Use .NET StreamWriter with UTF-8 encoding without BOM to prevent encoding issues with special characters
    $utf8WithoutBom = New-Object System.Text.UTF8Encoding $false
    [System.IO.File]::WriteAllText($reportPath, $htmlContent, $utf8WithoutBom)
    
    if (Test-Path $reportPath) {
        Write-Host "Report successfully generated at: $reportPath"
        # Open the report in the default browser
        # Uncomment the following line if you want the report to open automatically
        # Invoke-Item $reportPath
    }
    else {
        Write-Error "Failed to create report file."
    }
}
catch {
    Write-Error "Error generating report: $_"
}

