# onedrive-sharing-app/onedrive-sharing-app/README.md

# OneDrive Sharing App

This application allows users to generate sharing links for files and folders stored in OneDrive, utilizing the Microsoft 365 Developer subscription. The app features a simple graphical user interface (GUI) for easy interaction.

## Features

- Select files or folders from OneDrive
- Generate sharing links with a single click
- User-friendly GUI built with [Tkinter/PyQt]
- Secure authentication using tokens stored in a .env file

## Requirements

- Python 3.x
- Required libraries listed in `requirements.txt`

## Setup Instructions

1. Clone the repository:
   ```
   git clone <repository-url>
   cd onedrive-sharing-app
   ```

2. Create a `.env` file in the root directory and add your OneDrive API credentials:
   ```
   CLIENT_ID=<your-client-id>
   CLIENT_SECRET=<your-client-secret>
   TENANT_ID=<your-tenant-id>
   ```

3. Install the required dependencies:
   ```
   pip install -r requirements.txt
   ```

4. Run the application:
   ```
   python src/main.py
   ```

## Usage

1. Launch the application.
2. Use the GUI to select a file or folder from your OneDrive.
3. Click the "Generate Link" button to create a sharing link.
4. Copy the generated link for sharing.

## Contributing

Contributions are welcome! Please open an issue or submit a pull request for any enhancements or bug fixes.

## License

This project is licensed under the MIT License. See the LICENSE file for details.