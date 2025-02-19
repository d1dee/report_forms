# Report Forms Generator

A Deno-based tool for generating report forms by merging data from an Excel file with a Word template. This script parses Excel files containing learners' data and uses a provided Word template to generate personalized report forms. The generated reports are saved in a specified output directory, organized by sheet names.

> **Disclaimer:** This project is not affiliated with, endorsed by, or supported by any educational institution or government body. It is provided for educational and informational purposes only.

## Table of Contents

- [Features](#features)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Usage](#usage)
- [Command-Line Arguments](#command-line-arguments)
- [Project Structure](#project-structure)
- [Contributing](#contributing)
- [License](#license)
- [Acknowledgements](#acknowledgements)

## Features

- **Excel Parsing:** Reads and parses Excel files (.xlsx) to extract learners' data.
- **Report Generation:** Merges extracted data with a Word template to generate personalized report forms.
- **Automatic Directory Creation:** Organizes output by creating subdirectories based on Excel sheet names.
- **Error Handling:** Validates file access permissions and provides error messages if required files or directories are missing.
- **Deno-Powered:** Leverages Deno for secure and efficient execution with modern TypeScript support.

## Prerequisites

- [Deno](https://deno.land/) (version 1.0 or higher)
- An Excel file (.xlsx) containing learners' data.
- A Word document (.docx) to be used as the template for generating reports.

## Installation

Clone the repository:

```bash
git clone https://github.com/d1dee/report_forms.git
cd report_forms
```

## Usage

Run the script with the required permissions. At a minimum, you'll need read and write permissions for file operations:

```bash
deno run --allow-read --allow-write main.ts --excel_file path/to/your/data.xlsx --word_template path/to/your/template.docx
```

By default, the output directory is set to `./out` in the current working directory. You can override this using the `--out_dir` flag.

## Command-Line Arguments

The script accepts the following command-line arguments:

- `--excel_file` (`-e`): **(Required)** Path to the Excel (.xlsx) file containing learners' data.
- `--word_template` (`-t`): **(Required)** Path to the Word (.docx) template file.
- `--out_dir` (`-o`): Output directory where the generated report forms will be saved. Defaults to `./out` if not provided.
- `--temp_dir` (`-T`): (Optional) Temporary directory to use during processing.
- `--skip_errors` (`-S`): (Optional) Skip errors encountered during Excel parsing.
- `--help` (`-h`): Show help information.

## Project Structure

```
report_forms/
├── .vscode/               # VSCode workspace settings (optional)
├── src/                   # Source code for Excel parsing and DOCX generation
│   ├── write_docx.ts      # Module for writing data to DOCX files
│   └── xlsx.ts            # Module for parsing Excel files
├── .gitignore             # Files and folders to ignore in Git
├── deno.json              # Deno configuration file
├── deno.lock              # Lock file for Deno dependencies
├── main.ts                # Main script for processing and generating reports
└── main_test.ts           # Tests for the main functionality
```

## Contributing

Contributions, bug reports, and feature requests are welcome! Feel free to open an issue or submit a pull request if you have suggestions or improvements.

## License

This project is provided under [MIT Lincese](https://mit-license.org/).

## Acknowledgements

- [Deno](https://deno.land/) for providing a modern and secure runtime.
- The developers of the [@std/cli](https://deno.land/std/cli) and [@std/path](https://deno.land/std/path) modules.
```
