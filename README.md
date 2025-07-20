# Backup Manager

Backup Manager is a simple automation tool designed to help users automate backup tasks, especially when working with report files such as `.xlsx` documents from various operators. The app is particularly useful for operational processes where frequent and automatic updates of reports are needed, without the need for manual file fetching or opening.

## Features

- **Automated File Backup:** Automatically backs up specified files to a secure location.
- **Supports .xlsx Files:** Handles Excel files from multiple sources/operators seamlessly.
- **Prevents Forced Processes:** Ensures files are not processed if they are currently open and being used by another user, preventing data corruption or overwrite issues.
- **Streamlined Workflow:** Eliminates the need to manually open files, fetch data, or run updates.

## Use Case

This tool was originally created to support a migration process that requires regularly updated reports from multiple operators. By automating the backup and update process, users can avoid manual intervention and reduce the risk of errors from files being locked by other users.

## Getting Started

1. **Clone the Repository**
   ```bash
   git clone https://github.com/abamakbar07/backup-manager.git
   ```

2. **Configure Your Backup Settings**
   - Specify the folder(s) and file types you want to include in the backup process.
   - Ensure that your `.xlsx` files are placed in the designated input directory.

3. **Run the Backup Manager**
   - Follow the instructions provided in the project to start the automation process.

## Requirements

- Python (or your app's main language, update if different)
- Any dependencies listed in `requirements.txt` or a similar file

## Usage

1. Place your operator `.xlsx` files in the input directory.
2. Run the backup tool as instructed.
3. The app will automatically check for file availability and avoid processing files that are currently open.
4. Backups will be created in your configured backup location.

## Why Use This Tool?

- Prevents accidental overwrites and data corruption.
- Saves time by automating repetitive backup tasks.
- Ideal for migration and reporting workflows requiring frequent updates.

## Contributing

Pull requests and suggestions are welcome! Please open an issue or contact the author for major changes.

## License

This project is licensed under the MIT License.

## Author

Created by [abamakbar07](https://github.com/abamakbar07)
