# Resume Repository

Welcome to the **Resume Repository**! This repository is designed to automatically manage and update my latest resume while maintaining a historical archive of older versions. The repository integrates with GitHub and ensures that my most recent resume is always available and linked to my portfolio website: [bharatgurbaxani.com](https://bharatgurbaxani.com).

---

## Features

- **Automated Resume Conversion**: A PowerShell script converts the latest Word document version of my resume into a PDF format.
- **Version Control**: 
  - The most recent resume is stored in the root directory of the repository.
  - Older versions are automatically moved to the `Older version` folder with timestamps for easy reference.
- **GitHub Integration**:
  - Automatically uploads the latest resume to this repository.
  - Archives older versions in a designated folder within the repository.
- **Portfolio Website Integration**: The latest resume is linked directly to my personal portfolio website, ensuring visitors always have access to the most up-to-date version.

---

## Repository Structure

```
Resume/
├── Bharat Gurbaxani resume.pdf         # Latest version of the resume
├── Older version/                      # Folder containing older versions of the resume
│   ├── Bharat Gurbaxani resume <timestamp>.pdf
```

---

## How It Works

1. **Resume Update Process**:
   - The PowerShell script automatically converts the Word document (`.docx`) to a PDF file.
   - The PDF file is uploaded to this GitHub repository as the latest version.

2. **Archiving Older Versions**:
   - If a previous version of the resume exists, it is moved to the `Older version` folder with a timestamped filename for easy identification.

3. **GitHub API Integration**:
   - The script uses GitHub's REST API to handle file uploads, updates, and deletions.
   - A commit message is generated dynamically based on the upload date and time.

---

## PowerShell Script Overview

The automation is powered by a robust PowerShell script that performs the following tasks:

1. **Converts Word Document to PDF**:
   - Leverages Microsoft Word's COM object model for seamless conversion.

2. **Manages File Versions**:
   - Checks if an existing resume file is present in the repository.
   - Moves previous versions to the `Older version` folder before uploading the new one.

3. **Interacts with GitHub API**:
   - Uses GitHub's REST API for uploading files, managing branches, and handling commits.

---

## Prerequisites

To use or modify this automation process, ensure you have:

1. **PowerShell Environment**: The script runs on Windows PowerShell with access to Microsoft Word (for COM object).
2. **GitHub Personal Access Token (PAT)**: Required for authenticating API requests. Replace the placeholder in the script with your token or retrieve it securely.
3. **File Paths Configuration**: Update file paths in the script to match your local environment.

---

## How to Use

1. Clone this repository:
   ```
   git clone https://github.com/bharat98/Resume.git
   ```
2. Update your local Word document (`Bharat Gurbaxani Resume.docx`) with any changes.
3. Run the PowerShell script:
   ```
   .\AutomationScript.ps1
   ```
4. The script will:
   - Convert your updated `.docx` file into a `.pdf`.
   - Upload it as the latest version to this repository.
   - Archive any previous versions in the `Older version` folder.

5. Visit [bharatgurbaxani.com](https://bharatgurbaxani.com) to verify that your portfolio reflects the updated resume.

---

## Example Commit Messages

The script dynamically generates commit messages based on timestamps, such as:

- `Resume Update: 27-Nov-2024-0600`
- `Move existing file to Older version folder`

This ensures clear tracking of updates over time.

---

## Future Enhancements

- Automate deployment of updated resumes directly to my portfolio website.
- Add email notifications for successful updates or errors during execution.
- Extend support for additional document formats (e.g., LaTeX).

---

## Contributing

Contributions are welcome! If you have suggestions for improving this automation process or adding new features, feel free to open an issue or submit a pull request.

---

## License

This project is licensed under the MIT License. Feel free to use, modify, and distribute it as needed.

---

Thank you for visiting this repository! If you have any questions or feedback, please don't hesitate to reach out via [bharatgurbaxani.com](https://bharatgurbaxani.com).
