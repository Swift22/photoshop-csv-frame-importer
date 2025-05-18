/**
 * Adobe Photoshop CSV Data Import Script
 * This script imports data from a CSV file and places it into a properly structured Photoshop document.
 * Compatible with Adobe Photoshop CS6 and later versions.
 */

// Constants for configuration
var MAX_TEXT_WIDTH = 620; // Maximum width for text layers in pixels
var INITIAL_FONT_SIZE = 45; // Starting font size for text scaling
var MAX_ENTRIES = 3; // Maximum number of entries to process

/**
 * Opens a file dialog to select a CSV file
 * @returns {File} Selected CSV file or null if cancelled
 */
function selectCSVFile() {
  return File.openDialog("Select a CSV file", "*.csv");
}

/**
 * Reads and parses a CSV file into a 2D array
 * @param {File} file - The CSV file to read
 * @returns {Array<Array<string>>} Parsed CSV data as 2D array
 */
function readCSV(file) {
  file.open("r");
  var content = file.read();
  file.close();

  var lines = content.split("\n");
  var csvData = [];

  for (var i = 0; i < lines.length; i++) {
    var columns = lines[i].split(",");
    csvData.push(columns);
  }

  return csvData;
}

/**
 * Imports and processes an image into a specified frame
 * @param {string} imagePath - Path to the image file
 * @param {string} frameName - Name of the target frame layer
 */
function processImage(imagePath, frameName) {
  var originalDoc = app.activeDocument;
  var targetFrame = originalDoc.layers.getByName(frameName);

  var imageFile = new File(imagePath);
  if (!imageFile.exists) {
    alert("Image file not found: " + imagePath);
    return;
  }

  try {
    // Import and copy image
    var importedDoc = app.open(imageFile);
    importedDoc.selection.selectAll();
    importedDoc.selection.copy();
    importedDoc.close(SaveOptions.DONOTSAVECHANGES);

    // Paste and position image
    app.activeDocument = originalDoc;
    var imageLayer = originalDoc.paste();
    imageLayer.move(targetFrame, ElementPlacement.PLACEBEFORE);

    // Calculate frame and image dimensions
    var frameWidth = targetFrame.bounds[2] - targetFrame.bounds[0];
    var frameHeight = targetFrame.bounds[3] - targetFrame.bounds[1];
    var imageWidth = imageLayer.bounds[2] - imageLayer.bounds[0];
    var imageHeight = imageLayer.bounds[3] - imageLayer.bounds[1];

    // Calculate and apply scaling
    var widthRatio = frameWidth / imageWidth;
    var heightRatio = frameHeight / imageHeight;
    var scaleFactor = Math.max(widthRatio, heightRatio);

    imageLayer.resize(
      scaleFactor * 100,
      scaleFactor * 100,
      AnchorPosition.MIDDLECENTER
    );

    // Apply clipping mask
    imageLayer.grouped = true;

    // Center image in frame
    var frameCenterX = (targetFrame.bounds[0] + targetFrame.bounds[2]) / 2;
    var frameCenterY = (targetFrame.bounds[1] + targetFrame.bounds[3]) / 2;
    var imageCenterX = (imageLayer.bounds[0] + imageLayer.bounds[2]) / 2;
    var imageCenterY = (imageLayer.bounds[1] + imageLayer.bounds[3]) / 2;

    imageLayer.translate(
      frameCenterX - imageCenterX,
      frameCenterY - imageCenterY
    );
  } catch (error) {
    alert("Error processing image: " + imagePath + "\nError: " + error);
  }
}

/**
 * Places content in a specific text layer within a group
 * @param {string} content - Text content to place
 * @param {string} groupPrefix - Prefix of the group name
 * @param {string} textLayerName - Name of the target text layer
 * @param {string} [subGroupName] - Optional subgroup name
 * @param {boolean} [shouldAdjustSize] - Whether to adjust text size to fit
 * @returns {boolean} Success status
 */
function placeContentInTextLayer(
  content,
  groupPrefix,
  textLayerName,
  subGroupName,
  shouldAdjustSize
) {
  var doc = app.activeDocument;
  var groups = doc.layerSets;

  for (var i = 0; i < groups.length; i++) {
    if (groups[i].name.indexOf(groupPrefix) === 0) {
      var targetGroup = groups[i];

      // Navigate to subgroup if specified
      if (subGroupName) {
        var subGroups = targetGroup.layerSets;
        for (var k = 0; k < subGroups.length; k++) {
          if (subGroups[k].name === subGroupName) {
            targetGroup = subGroups[k];
            break;
          }
        }
      }

      // Find and update text layer
      var textLayers = targetGroup.artLayers;
      for (var j = 0; j < textLayers.length; j++) {
        if (
          textLayers[j].name === textLayerName &&
          textLayers[j].kind === LayerKind.TEXT
        ) {
          var textItem = textLayers[j].textItem;
          textItem.contents = content;

          if (!shouldAdjustSize) return true;

          // Adjust text size if needed
          function getTextWidth() {
            var bounds = textLayers[j].bounds;
            return bounds[2] - bounds[0];
          }

          textItem.size = INITIAL_FONT_SIZE;
          while (getTextWidth() > MAX_TEXT_WIDTH && textItem.size > 1) {
            textItem.size -= 0.5;
          }

          return true;
        }
      }
    }
  }

  alert(
    "Text layer '" +
      textLayerName +
      "' not found in group '" +
      groupPrefix +
      (subGroupName ? "' and subgroup '" + subGroupName : "") +
      "'"
  );
  return false;
}

/**
 * Creates a copy of the current document with a new name
 * @param {Document} sourceDoc - Source document
 * @param {string} newFileName - Name for the new file
 * @returns {Document} Newly created document
 */
function saveCopyAsNewFile(sourceDoc, newFileName) {
  var sourcePath = sourceDoc.path;
  var newFilePath = sourcePath + "/" + newFileName;
  var newFile = new File(newFilePath);

  sourceDoc.saveAs(newFile, new PhotoshopSaveOptions(), true);
  return app.open(newFile);
}

/**
 * Main execution function
 */
function main() {
  var sourceDoc = app.activeDocument;
  var frameNames = ["LeftFrame", "MiddleFrame", "RightFrame"];

  // Get CSV data
  var csvFile = selectCSVFile();
  if (csvFile === null) {
    alert("No CSV file selected. Script terminated.");
    return;
  }

  var csvData = readCSV(csvFile);
  var processedCount = 0;

  // Process entries
  for (var i = 1; i < csvData.length && processedCount < MAX_ENTRIES; i++) {
    // Create new document for first entry
    if (processedCount === 0) {
      var workingDoc = saveCopyAsNewFile(
        sourceDoc,
        processedCount + 1 + ".psd"
      );
    }

    app.activeDocument = workingDoc;

    var entryData = {
      name: csvData[i][0],
      profession: csvData[i][1],
      overdose: csvData[i][2],
      yearOfDeath: csvData[i][3],
      age: csvData[i][4],
      imagePath: csvData[i][5],
    };

    var groupPrefix = processedCount + 1 + ".";

    // Place text content
    placeContentInTextLayer(entryData.name, groupPrefix, "Name", null, true);
    placeContentInTextLayer(
      entryData.profession,
      groupPrefix,
      "Profession",
      null,
      true
    );
    placeContentInTextLayer(
      entryData.overdose,
      groupPrefix,
      "Type",
      null,
      true
    );
    placeContentInTextLayer(
      entryData.yearOfDeath,
      groupPrefix,
      "Year",
      null,
      false
    );
    placeContentInTextLayer(entryData.age, groupPrefix, "Age", "Age", false);

    // Process image if path provided
    if (entryData.imagePath) {
      try {
        processImage(entryData.imagePath, frameNames[processedCount]);
      } catch (error) {
        // Error already handled in processImage
      }
    }

    processedCount++;
  }

  alert("Script completed successfully!");
}

// Execute the script with history tracking
app.activeDocument.suspendHistory("Import CSV Data", "main()");
