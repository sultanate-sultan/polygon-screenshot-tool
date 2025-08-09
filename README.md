# Polygon Screenshot Tool

A fast and lightweight Windows service for capturing screenshots of a polygon-defined area. This tool runs in the system tray and is activated by a global hotkey, allowing you to quickly capture irregular shapes on your screen and copy them directly to your clipboard.

## Features

- **Polygon Selection:** Accurately capture any irregular shape on your screen.
    
- **Global Hotkey:** Press `Ctrl+Alt+2` to instantly activate the capture tool.
    
- **System Tray Integration:** Runs silently in the background without a visible window.
    
- **High-DPI Support:** Designed to work correctly on high-resolution displays.
    
- **Clipboard Integration:** Captured images are automatically copied to your clipboard as a PNG.
    

## How to Use the Pre-built Executable

The easiest way to use this tool is by downloading the pre-built executable (`.exe`). This is a standalone application that does not require Python or any dependencies to be installed.

1. **Download:** Go to the [Releases page](https://www.google.com/search?q=https://github.com/your-username/your-repository-name/releases "null") and download the latest `polygon_screenshot_tool.exe` file.
    
2. **Run:** Simply double-click the executable file to start the service. The application will run minimized in your system tray.
    
## Demo



https://github.com/user-attachments/assets/704bdf3d-469a-42a4-a956-e2bf1a8a443f



## For Developers: How to Build from Source

If you want to contribute to the project, modify the code, or build the application yourself, follow these steps.

### Prerequisites

- Windows 10 or 11
    
- Python 3.x
    

### 1. Clone the repository

```
git clone https://github.com/your-username/your-repository-name.git
cd your-repository-name
```

### 2. Install dependencies

It is highly recommended to use a virtual environment.

```
python -m venv venv
# On Windows:
venv\Scripts\activate
# On macOS/Linux:
source venv/bin/activate
```

Install the required libraries:

```
pip install -r requirements.txt
```

### 3. Run the application

Run the script using the Pythonw interpreter to avoid a console window.

```
pythonw polygon_screenshot_service.pyw
```

The application will now be running in your system tray. Look for the Python icon.

### 4. Build the executable

If you wish to create a standalone executable, you can use PyInstaller.

1. **Install PyInstaller:**
    
    ```
    pip install pyinstaller
    ```
    
2. **Build the executable:**
    
    ```
    pyinstaller --onefile --noconsole --icon=icon.ico polygon_screenshot_service.pyw
    ```
    
    - `--onefile`: Packages the entire application into a single `.exe` file.
        
    - `--noconsole`: Prevents the console window from appearing.
        
    - `--icon=icon.ico`: (Optional) Specifies a custom icon file.
        

The final executable will be located in the `dist` folder.

## Contributing

We welcome contributions! If you have an idea for a new feature, a bug fix, or an improvement, please feel free to:

1. Fork the repository.
    
2. Create a new branch (`git checkout -b feature/your-feature-name`).
    
3. Make your changes.
    
4. Commit your changes (`git commit -am 'Add some feature'`).
    
5. Push to the branch (`git push origin feature/your-feature-name`).
    
6. Create a new Pull Request.
    

## License

This project is licensed under the [MIT License](https://www.google.com/search?q=LICENSE.txt "null").
