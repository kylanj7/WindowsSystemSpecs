# Windows System Data Collection

A comprehensive Python script for collecting detailed system information from Windows laptops and desktops. Generates detailed reports about hardware components, system specifications, and performance metrics.

## Features

- Comprehensive system information collection
- Detailed hardware specifications
- Component-specific analysis
- Automated report generation
- Battery health monitoring
- I/O port detection
- GPU monitoring

## Information Collected

### System Details
- OS version and release
- Manufacturer and model
- Serial number and SKU
- System specifications

### Hardware Components
- CPU details (model, speed, cores)
- RAM configuration
- GPU information
- Storage devices
- I/O ports and controllers
- Network adapters
- Battery status

## Prerequisites

- Windows operating system
- Python 3.7+
- Required packages:
```bash
pip install psutil wmi GPUtil
```

## Installation

1. Clone the repository or download the script
2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

Run the script:
```bash
python system_info.py
```

## Output

- Creates timestamped report in Documents folder
- Filename format: `laptop_info_YYYYMMDD_HHMMSS.txt`
- Displays summary in console
- Saves detailed report to file

## Key Functions

- `get_system_info()`: Basic system information
- `get_cpu_info()`: Detailed CPU specifications
- `get_memory_info()`: RAM configuration
- `get_gpu_info()`: Graphics card details
- `get_detailed_io_info()`: I/O ports and devices
- `get_battery_health()`: Battery status

## Features in Detail

### CPU Detection
- Model identification
- Generation detection
- Clock speed monitoring
- Core count

### Memory Analysis
- Total capacity
- DIMM slot information
- Memory type detection
- Speed information

### I/O Port Detection
- USB ports (including Type-C)
- Video outputs
- Network adapters
- Storage devices

### Battery Analysis
- Health percentage
- Capacity information
- Charging status

## Error Handling

- Permission verification
- Component detection fallbacks
- Missing information handling
- Access rights management

## Limitations

- Windows-specific functionality
- Some features require admin rights
- Hardware-dependent capabilities

## Contributing

Feel free to submit issues and enhancement requests!

## License

MIT License

Copyright (c) [2024] [Kylan Johnson]

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

## Acknowledgments

- WMI documentation
- Python psutil contributors
- GPUtil developers
