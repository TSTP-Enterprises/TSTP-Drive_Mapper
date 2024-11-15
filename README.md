# TSTP Drive Mapper 🌐

**TSTP Drive Mapper** is a powerful and intuitive tool designed to simplify the management of network drive mappings on Windows systems. Whether you're an IT professional or a user seeking efficient drive management, this tool offers a seamless experience with its user-friendly interface and robust features. Security has been a cornerstone in the development of TSTP Drive Mapper, ensuring that it meets industry standards and provides a secure environment for managing network drives.

![Hero Image](https://www.tstp.xyz/images/drive_mapper_banner.png)

---

## 📋 Table of Contents

- [🚀 Key Features](#-key-features)
- [🔒 Security and Compliance](#-security-and-compliance)
- [🛠 Installation Guide](#-installation-guide)
- [📖 User Guide](#-user-guide)
- [🎥 Demo Video](#-demo-video)
- [🛠 Build Information](#-build-information)
- [💡 Development](#-development)
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
  Export drive mappings as JSON or XML files for automation and backup purposes.

- **Detailed Logging**  
  Keep comprehensive logs of all operations for auditing and troubleshooting.

- **Responsive Design**  
  Enjoy a smooth and intuitive experience across all devices with our responsive layout.

- **System Tray Integration**  
  Access quick controls and settings without occupying space on your taskbar.

---

## 🔒 Security and Compliance

Security is integral to the TSTP Drive Mapper's design and development process. Here are the key security measures implemented:

- **Data Encryption**: All sensitive data, including credentials, are encrypted using industry-standard encryption algorithms to prevent unauthorized access.

- **Regular Security Audits**: The codebase undergoes regular security audits to identify and mitigate potential vulnerabilities, ensuring compliance with the latest security standards.

- **Compliance with Industry Standards**: TSTP Drive Mapper is developed in accordance with industry best practices and standards, ensuring a secure and reliable tool for managing network drives.

- **User Privacy**: We prioritize user privacy by implementing strict data protection policies and ensuring that user data is handled with the utmost care.

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
- Navigate to **File > Export Drives** to export your drive mappings as JSON or XML files. These files can be used for automation or shared with others.

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

## 🛠 Build Information

- **Programming Language:** Python 3.8
- **Development Environment:** Cursor / VS Code
- **Build Tools:** PyInstaller
- **Version:** 1.0.0
- **License:** MIT License

### **Building from GitHub**
1. **Clone the Repository:**
    Clone the repository from GitHub to your local machine:
    ```bash
    git clone https://github.com/TSTP-Enterprises/TSTP-Drive_Mapper.git
    cd TSTP-Drive_Mapper
    ```

2. **Install Python and Dependencies:**
    Ensure you have Python 3.8 installed on your system. Install the required Python packages using the `requirements.txt` file:
    ```bash
    pip install -r requirements.txt
    ```

3. **Required Imports:**
    The application uses several Python modules and PyQt5 for the GUI. Ensure the following imports are available in your environment:
    ```python
    import sys
    import os
    import json
    import logging
    import requests
    import subprocess
    import urllib.request
    from datetime import datetime
    from PyQt5.QtWidgets import (
        QApplication, QMainWindow, QPushButton, QTextEdit, QVBoxLayout, QWidget,
        QDialog, QHBoxLayout, QLabel, QLineEdit, QMessageBox, QTableWidget,
        QTableWidgetItem, QHeaderView, QAbstractItemView, QComboBox, QAction,
        QFileDialog, QCheckBox, QSystemTrayIcon, QMenu, QGraphicsDropShadowEffect, 
        QStyle, QScrollArea, QFrame, QSizePolicy
    )
    from PyQt5.QtCore import Qt, QThread, pyqtSignal, QDateTime, QUrl
    from PyQt5.QtGui import QIcon, QColor, QBrush, QDesktopServices
    ```

4. **Build the Executable:**
    Use PyInstaller to create a standalone executable of the application:
    ```bash
    pyinstaller --onefile --windowed Code/main.py
    ```

5. **Run the Application:**
    After building, the executable will be located in the `dist` directory. You can run it to start the application:
    ```bash
    ./dist/main
    ```

---

## 💡 Development

**TSTP Drive Mapper** was developed to address a critical need for a client who frequently had to remap network drives due to connecting with different profiles to their VPN both while away from work and at the office. The existing process was time-consuming and prone to errors, disrupting productivity and workflow.

### **Development Process:**
1. **Requirement Analysis:** Identified the core requirements for efficient drive mapping and automatic remapping based on VPN profiles.
2. **Design:** Crafted an intuitive user interface that allows for easy management of network drives, batch operations, and secure credential storage.
3. **Implementation:** Utilized PowerShell for scripting the drive mapping functionalities, ensuring compatibility with Windows 10/11 systems.
4. **Testing:** Conducted rigorous testing to ensure reliability, security, and ease of use. Feedback from initial users was incorporated to refine features.
5. **Deployment:** Packaged the tool into an executable for seamless installation and usage across various environments.

### **Why It Was Made:**
The primary goal was to create a reliable tool that automates the tedious process of remapping network drives, especially in environments where VPN profiles change frequently. By automating this process, **TSTP Drive Mapper** reduces the potential for human error, saves valuable time, and enhances overall productivity.

---

## 🌟 Contributing

We welcome contributions to enhance **TSTP Drive Mapper**! Whether it's reporting a bug, suggesting a feature, or submitting a pull request, your input is invaluable.

1. **Fork the Repository:** Click the **Fork** button at the top-right corner of the repository page.
2. **Create a Feature Branch:**  
    ```bash
    git checkout -b feature/YourFeature
    ```
3. **Commit Your Changes:**  
    ```bash
    git commit -m "Add YourFeature"
    ```
4. **Push to the Branch:**  
    ```bash
    git push origin feature/YourFeature
    ```
5. **Open a Pull Request:** Navigate to the repository on GitHub and click **New Pull Request**.

For detailed guidelines, refer to our [Contributing Guidelines](CONTRIBUTING.md).

---

## 💬 Support

Your support helps us maintain and improve **TSTP Drive Mapper**. As a free and open-source tool, donations help cover development costs and ensure the project remains sustainable.

If you find this tool valuable, please consider supporting us via [PayPal](https://www.paypal.com/donate/?hosted_button_id=YOUR_PAYPAL_LINK).

Your contributions are greatly appreciated and help us continue providing this essential tool to the community.

---

## 📜 License

This project is licensed under the [MIT License](LICENSE).  
**TSTP Drive Mapper** is open-source, free to use, and does not offer any packages, plans, or additional services.

---

## 🔗 Links

- **Official Website:** [TSTP Website](https://www.tstp.xyz/)
- **GitHub Organization:** [TSTP-Enterprises](https://github.com/TSTP-Enterprises)
- **LinkedIn:** [TSTP LinkedIn](https://www.linkedin.com/company/thesolutions-toproblems)
- **YouTube Channel:** [(TSTP) YourPST Studios](https://www.youtube.com/@yourpststudios)
- **Facebook Page:** [TSTP Facebook](https://www.facebook.com/profile.php?id=61557162643039)
- **GitHub Repository:** [TSTP-Drive_Mapper](https://github.com/TSTP-Enterprises/TSTP-Drive_Mapper)
- **Software Page:** [TSTP Drive Mapper Software](https://tstp.xyz/software/drive-mapper/)
- **Download Link:** [Download TSTP Drive Mapper](https://www.tstp.xyz/downloads/TSTP-Drive_Mapper.zip)

---

## 🙌 Acknowledgements

A heartfelt thank you to all the users who have supported **TSTP Drive Mapper**. Your feedback and encouragement drive us to continuously improve and provide the best possible tool for managing network drives.

---

Made with ❤️ by **[TSTP Enterprises](https://www.tstp.xyz)**
