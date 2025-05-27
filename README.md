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

Your Photoshop document must have the following layer structure:

### Frame Layers

- `profile-frame-left` - Left frame for image placement
- `profile-frame-center` - Center frame for image placement
- `profile-frame-right` - Right frame for image placement

### Group Structure

For each profile (1-3), create a group with the following naming pattern:

- Main group name: `profile-1`, `profile-2`, `profile-3`
  - Text layers within each group:
    - `profile-name` - Person's name
    - `profile-occupation` - Person's profession
    - `profile-cause` - Cause information
    - `profile-year` - Year information
  - Subgroup named `age-details`:
    - Text layer `profile-age` - Age information

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

## Configuration

The script uses a configuration object (`LAYOUT_CONFIG`) that defines:

- Text settings (maximum width, initial font size)
- Frame names
- Layer names
- Group naming conventions
- Maximum number of profiles

## License

[MIT License](LICENSE)
