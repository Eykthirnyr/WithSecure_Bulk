# WithSecure Batch Management Tool

WithSecure Batch Management Tool is a **multi-organization, multi-device** management utility that automates bulk operations via the WithSecure API. It allows administrators to efficiently manage security endpoints, execute operations, and export inventory data across multiple organizations.

This project is partly based on [WithSecure API Export Tool](https://github.com/Eykthirnyr/WithSecure_API_Export_Tool/tree/main), expanding its capabilities with a broader feature set.

## Features

### 🔑 Authentication & Organization Management
- **OAuth2 Login** with **Client ID** & **Client Secret**
- Retrieve all organizations linked to the API credentials
- Select multiple organizations for bulk operations

### 🖥️ Device Management
- Fetch active devices from selected organizations
- Select individual or all devices for actions

### ⚡ Bulk Operations
Trigger predefined actions on multiple devices at once:
- **Malware Scan** (`scanForMalware`)
- **Send Message to User** (`showMessage`)
- **Network Isolation** (`isolateFromNetwork` / `releaseFromNetworkIsolation`)
- **Assign Security Profile** (`assignProfile`)
- **Enable Debug Logging** (`turnOnFeature`)
- **Check for Missing Software Updates** (`checkMissingUpdates`)

### 📋 Inventory Export
Export detailed hardware and system information of selected devices to a CSV file, including:
- Organization Name
- Device Name & ID
- BIOS Version
- Serial Number
- Total Memory (Bytes)
- OS Name & Version
- End-of-Life Status
- Last Logged-In User
- Online Status
- Computer Model
- System Drive Total & Free Space (GB)
- Disk Encryption Status

### 🗑️ Device State Management
- **Block, Inactivate, or Delete** devices in bulk.

### 📝 Logging & User Interface
- **Log output** for debugging and monitoring operations
- **Tkinter-based GUI** with checkboxes for selecting organizations and devices
- **Progress indicators** for longer tasks like inventory export or update checks

![GUI](https://github.com/user-attachments/assets/51cfc453-b50a-4f10-ae1d-3bd33863ea2c)

## 🔧 Requirements
- Python 3.x
- `requests` (automatically installed if missing)

###Contribution & Feedback

This project was designed for IT administrators using WithSecure solutions. Contributions, issues, and feature requests are welcome.

🔗 Made by Clément GHANEME (02/2025)
Visit: [clement.business](https://clement.business/)


