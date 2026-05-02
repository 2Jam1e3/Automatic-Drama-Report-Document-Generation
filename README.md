# Drama Report Generator

A Node.js script designed for **Bombo Radyo Philippines** to automate the generation of weekly Drama Reports. The script parses radio automation logs (Raduga `.LOG` files) to track the airing of specific drama episodes, and automatically generates a beautifully formatted Microsoft Word (`.docx`) report.

## Features

- **Automated Log Parsing**: Reads and extracts episode data (Season & Chapter) directly from Raduga `.LOG` files.
- **Dynamic DOCX Generation**: Creates a heavily styled Microsoft Word document complete with floating header logos, correctly aligned text blocks, and structured tabular data.
- **Auto-Filing**: Automatically organizes and saves the generated reports into structured directories based on the date (`D:\Documents\Reports\Drama\YYYY\Month`).
- **Customizable**: Allows overriding the default network log directory and output paths via CLI arguments.

## Prerequisites

- **Node.js** (v14 or higher recommended)
- **Dependencies**: The script relies on the `docx` NPM package.

## Installation

1. Clone or download this repository.
2. Install the required dependencies:
   ```bash
   npm install docx
   ```
3. **Assets**: Ensure the required logo images are located in your `Pictures` folder at the exact paths expected by the script (e.g., `C:\Users\User\Pictures\Logo\...`).

## Usage

You can run the script directly using Node.js. By default, it will attempt to read logs from the network path `\\192.168.86.134\Raduga Log` and save the output to your local `D:\` drive.

```bash
node generate_drama_report.js
```

### Command Line Arguments

- `--folder <path>`: Override the default directory where Raduga `.LOG` files are stored.
- `--out <path>`: Override the default output filepath and name for the `.docx` file.

**Example:**
```bash
node generate_drama_report.js --folder "C:\Temp\Logs" --out "C:\Temp\Custom_Report.docx"
```

## How It Works

1. **Log Discovery**: The script scans the specified directory for all `.LOG` files.
2. **Parsing**: It parses the text to identify predefined drama titles, extracting the corresponding Season and Chapter numbers.
3. **Document Assembly**: Using the `docx` library, it constructs a multi-section document featuring:
   - A custom header with floating logos.
   - Dynamic date generation.
   - Pre-filled addressee and body text describing Facebook Live copyright management.
   - A summary table of the aired chapters.
   - A predefined signature block.
4. **Saving**: The document is packaged into a `.docx` buffer and saved to the auto-generated year/month directory on your `D:` drive.
