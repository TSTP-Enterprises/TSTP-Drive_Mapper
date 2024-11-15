# TSTP Drive Mapper 🌐

**TSTP Drive Mapper** is a powerful and intuitive tool designed to simplify the management of network drive mappings on Windows systems. Whether you're an IT professional or a user seeking efficient drive management, this tool offers a seamless experience with its user-friendly interface and robust features.

![Hero Image](https://www.tstp.xyz/images/drive_mapper_banner.png)

---

## 📋 Table of Contents

- [🚀 Key Features](#-key-features)
- [🛠 Installation Guide](#-installation-guide)
- [📖 User Guide](#-user-guide)
- [🎥 Demo Video](#-demo-video)
- [📦 Development](#-development)
- [🌟 Contributing](#-contributing)
- [💬 Support](#-support)
- [📜 License](#-license)
- [🔗 Links](#-links)
- [🙌 Acknowledgements](#-acknowledgements)

---

## 🚀 Key Features

- **Effortless Drive Management**  
  Easily add, edit, and remove network drives through a straightforward interface.

- **Batch Operations**  
  Simultaneously map or unmap multiple drives to enhance efficiency and productivity.

- **Seamless Startup Integration**  
  Set the application to launch automatically with Windows and re-add drives effortlessly.

- **Secure Credential Management**  
  Safely store credentials for drives that require authentication.

- **Advanced Export Options**  
  Export drive mappings as PowerShell scripts for automation and backup purposes.

- **Detailed Logging**  
  Keep comprehensive logs of all operations for auditing and troubleshooting.

- **Responsive Design**  
  Enjoy a smooth and intuitive experience across all devices with our responsive layout.

- **System Tray Integration**  
  Access quick controls and settings without occupying space on your taskbar.

---

## 🛠 Installation Guide

### **Step 1: Download the Latest Version**
Download the latest version of **TSTP Drive Mapper** from the [official website](https://www.tstp.xyz/downloads/TSTP-Drive_Mapper.zip).

### **Step 2: Extract the Files**
Unzip the downloaded archive and save the contents to your desired directory.

### **Step 3: Launch the Application**
Execute the `Drive_Mapper.exe` file to start using the application.

### **System Requirements**
- **Operating System:** Windows 10/11
- **Permissions:** Administrator privileges may be required for certain operations.

---

## 📖 User Guide

### **1. Adding a New Drive**
1. Click on the **"Add Drive"** button.
2. Choose an available drive letter from the dropdown menu.
3. Enter the UNC path (e.g., `\\server\share`).
4. If required, enter the necessary credentials.
5. Click **"Save"** to add the drive to your list.

### **2. Editing an Existing Drive**
1. Select the drive you wish to edit by checking the corresponding checkbox.
2. Click the **"Edit Drive"** button.
3. Modify the necessary details in the dialog that appears.
4. Click **"Save"** to apply the changes. If the drive is currently mapped, it will be unmapped and remapped with the new settings.

### **3. Removing a Drive**
1. Select the drive(s) you want to remove by checking the checkbox.
2. Click the **"Remove Drive"** button.
3. Confirm the removal in the prompt. If the drive is mapped, it will be unmapped before removal.

### **4. Mapping and Unmapping Drives**
- **Map Drives:** Click the **"Map Drives"** button to connect all selected drives. You can select specific drives or choose to map all drives in the list.
- **Unmap Drives:** Click the **"Unmap Drives"** button to disconnect all selected drives. Similar to mapping, you can select specific drives or unmap all.

### **5. Checking Drive Status**
- Click the **"Check Drives"** button to verify the current status of all drive mappings. The **"Mapped"** column will indicate whether each drive is currently connected.

### **6. Exporting Drive Mappings**
- Navigate to **File > Export Drives for PowerShell Script** to export your drive mappings as a PowerShell script. This script can be used for automation or shared with others.

### **7. Accessing Logs**
- The log console at the bottom of the main window displays real-time logs of all operations.
  - **Save Log:** Navigate to **File > Save Log** to export logs in various formats.
  - **Clear Log:** Navigate to **File > Clear Log** to clear the log history.
  - **Toggle Console:** Navigate to **File > Toggle Console** to show or hide the console.

### **8. Settings**
- Adjust your preferences under the **Settings** menu.
  - **Startup Settings:** Enable the application to start on Windows startup.
  - **Auto Re-Add Drives:** Automatically re-add drives upon startup.
  - **Theme Selection:** Switch between light and dark themes to suit your preference.

### **9. System Tray Integration**
- The application minimizes to the system tray, allowing you to access quick controls and settings without occupying space on your taskbar.
  - **Right-Click Tray Icon:** Open the main window, toggle startup settings, switch themes, or exit the application.

---

## 🎥 Demo Video

Watch the demo video to see **TSTP Drive Mapper** in action:  
[![Watch Demo](https://img.youtube.com/vi/DrImS-LSQ6c/0.jpg)](https://www.youtube.com/watch?v=DrImS-LSQ6c)

---

## 📦 Development

### **Clone the Repository**