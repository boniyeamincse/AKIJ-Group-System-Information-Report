
# System Information and Installed Software Report Generator

This PowerShell script generates an HTML report containing detailed information about the system and the software installed on it. The report includes system overview, installed software details, and metadata about the report generation.

## Table of Contents
- [Prerequisites](#prerequisites)
- [Usage](#usage)
- [Output](#output)
- [License](#license)

## Prerequisites

- PowerShell 5.1 or later
- Administrative privileges to run the script

## Usage

1. Clone the repository or download the script file.
2. Open PowerShell with administrative privileges.
3. Navigate to the directory containing the script file.
4. Run the script using the following command:
   ```powershell
   .\install_software-report.ps1
   ```

## Output

The script generates an HTML report with the following sections:

### System Overview

- **Computer Name**: The name of the computer.
- **Manufacturer**: The manufacturer of the computer.
- **Model**: The model of the computer.
- **Operating System**: The operating system installed on the computer.
- **OS Version**: The version of the operating system.
- **OS Architecture**: The architecture of the operating system (e.g., 64-bit).
- **Processor**: Information about the processor, including the name, number of cores, and number of logical processors.
- **Memory**: Total physical memory installed on the computer.

### Installed Software Report

A table containing details of installed software, including:

- **Name**: The name of the software.
- **Version**: The version of the software.
- **Publisher**: The publisher or vendor of the software.
- **Install Date**: The date the software was installed.

### Metadata

- **Report Generation Date**: The date and time when the report was generated.
- **Hostname**: The name of the computer for which the report was generated.

## License
MIT License

Copyright (c) 2025 Boni Yeamin

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

&copy; 2025 AKIJ IT Team. All rights reserved.

Author: Boni Yeamin
