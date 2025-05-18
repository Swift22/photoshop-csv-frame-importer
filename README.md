# Import JSX - Photoshop Automation Script

This script automates the process of importing and formatting images and text data from a CSV file into an Adobe Photoshop document. It's particularly useful for batch processing multiple entries with consistent formatting.

## Features

- CSV file import and parsing
- Automated image import and placement
- Text content placement in specific layers
- Automatic text size adjustment based on content length
- Support for multiple frames (Left, Middle, Right)
- Batch processing capability

## Requirements

- Adobe Photoshop (Compatible with CS6 and later versions)
- A properly formatted CSV file with the following columns:
  1. Name
  2. Profession
  3. Overdose
  4. Year of Death
  5. Age
  6. Image Path

## Usage

1. Open your target Photoshop document
2. Run the script in Photoshop (File > Scripts > Browse...)
3. Select your CSV file when prompted
4. The script will process the entries and create new documents as needed

## Layer Structure Requirements

Your Photoshop document should have the following layer structure:

- Groups named "1.", "2.", "3." (for each entry)
- Text layers named "Name", "Profession", "Type", "Year" in each group
- An "Age" subgroup containing an "Age" text layer
- Frame layers named "LeftFrame", "MiddleFrame", "RightFrame"

## Error Handling

The script includes error handling for:

- Missing files
- Invalid image paths
- Layer structure issues
- Text sizing adjustments

## Development

This is a JSX script that uses Adobe ExtendScript. To develop or modify this script, you can use:

- Visual Studio Code with ExtendScript Debugger extension
- Adobe ExtendScript Toolkit CC
- Any text editor with JavaScript syntax highlighting

## License

[MIT License](LICENSE)
