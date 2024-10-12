// Function to select a CSV file
function selectCSVFile() {
  return File.openDialog("Select a CSV file", "*.csv");
}

// Function to read CSV file and store in array
function readCSV(file) {
  file.open("r");
  var content = file.read();
  file.close();
  var lines = content.split("\n");
  var csvArray = [];
  for (var i = 0; i < lines.length; i++) {
    var columns = lines[i].split(",");
    csvArray.push(columns);
  }

  return csvArray;
}

// Function to import image and process it
function processImage(imagePath, frame) {
  var originalDoc = app.activeDocument;
  var clippingFrame = originalDoc.layers.getByName(frame);

  // Import image
  var file = new File(imagePath);

  if (!file.exists) {
    alert("File does not exist: " + imagePath);
    return;
  }

  try {
    var importedDoc = app.open(file);
    importedDoc.selection.selectAll();
    importedDoc.selection.copy();
    importedDoc.close(SaveOptions.DONOTSAVECHANGES);

    app.activeDocument = originalDoc;
    var newLayer = originalDoc.paste();

    // Move layer above clippingFrame
    newLayer.move(clippingFrame, ElementPlacement.PLACEBEFORE);

    // Resize image to fit the frame
    var frameWidth = clippingFrame.bounds[2] - clippingFrame.bounds[0];
    var frameHeight = clippingFrame.bounds[3] - clippingFrame.bounds[1];
    var imageWidth = newLayer.bounds[2] - newLayer.bounds[0];
    var imageHeight = newLayer.bounds[3] - newLayer.bounds[1];

    var widthRatio = frameWidth / imageWidth;
    var heightRatio = frameHeight / imageHeight;
    var scaleFactor = Math.max(widthRatio, heightRatio);

    newLayer.resize(
      scaleFactor * 100,
      scaleFactor * 100,
      AnchorPosition.MIDDLECENTER
    );

    // Clip to MiddleFrame
    newLayer.grouped = true;

    // Center on MiddleFrame
    var frameBounds = clippingFrame.bounds;
    var frameCenterX = (frameBounds[0] + frameBounds[2]) / 2;
    var frameCenterY = (frameBounds[1] + frameBounds[3]) / 2;

    var imageBounds = newLayer.bounds;
    var imageCenterX = (imageBounds[0] + imageBounds[2]) / 2;
    var imageCenterY = (imageBounds[1] + imageBounds[3]) / 2;

    newLayer.translate(
      frameCenterX - imageCenterX,
      frameCenterY - imageCenterY
    );
  } catch (e) {
    alert("Error processing image: " + imagePath + "\nError: " + e);
  }
}

// Function to place content in a specific text layer within a group
function placeContentInTextLayer(
  content,
  groupPrefix,
  textLayerName,
  subGroupName,
  sizeCheck
) {
  var doc = app.activeDocument;
  var groups = doc.layerSets;

  for (var i = 0; i < groups.length; i++) {
    if (groups[i].name.indexOf(groupPrefix) === 0) {
      var targetGroup = groups[i];

      // If subGroupName is provided, look for the subgroup
      if (subGroupName) {
        var subGroups = targetGroup.layerSets;
        for (var k = 0; k < subGroups.length; k++) {
          if (subGroups[k].name === subGroupName) {
            targetGroup = subGroups[k];
            break;
          }
        }
      }

      var textLayers = targetGroup.artLayers;
      for (var j = 0; j < textLayers.length; j++) {
        if (
          textLayers[j].name === textLayerName &&
          textLayers[j].kind === LayerKind.TEXT
        ) {
          var textItem = textLayers[j].textItem;
          textItem.contents = content;

          if (!sizeCheck) return true;

          function getTextWidth() {
            var bounds = textLayers[j].bounds;
            return bounds[2] - bounds[0]; // right - left
          }

          // Check and adjust font size if width exceeds 630 pixels
          var maxWidth = 620;
          textItem.contents = content;
          textItem.size = 45;

          // alert(getTextWidth());
          while (getTextWidth() > maxWidth && textItem.size > 1) {
            textItem.size -= 0.5;
          }

          return true; // Content placed successfully
        }
      }
    }
  }

  alert(
    "Could not find text layer '" +
      textLayerName +
      "' in group starting with '" +
      groupPrefix +
      (subGroupName ? "' and subgroup '" + subGroupName : "") +
      "'"
  );
  return false; // Content not placed
}

// Function to save a copy of the current document
function saveCopyAsNewFile(originalDoc, newFileName) {
  var originalPath = originalDoc.path;
  var newFilePath = originalPath + "/" + newFileName;

  var newFile = new File(newFilePath);
  originalDoc.saveAs(newFile, new PhotoshopSaveOptions(), true);

  return app.open(newFile);
}

// Main script
function main() {
  var originalDoc = app.activeDocument;

  var frames = ["LeftFrame", "MiddleFrame", "RightFrame"];
  var csvFile = selectCSVFile();
  if (csvFile === null) {
    alert("No CSV file selected. Script terminated.");
    return;
  }

  var csvArray = readCSV(csvFile);

  var counter = 0;
  for (var i = 1; i < csvArray.length; i++) {
    if (counter == 3) break;
    if (counter == 0)
      var newDoc = saveCopyAsNewFile(originalDoc, counter + 1 + ".psd");
    app.activeDocument = newDoc;
    var name = csvArray[i][0];
    var profession = csvArray[i][1];
    var overdose = csvArray[i][2];
    var yearOfDeath = csvArray[i][3];
    var age = csvArray[i][4];
    var imagePath = csvArray[i][5];

    // Place name content
    placeContentInTextLayer(name, counter + 1 + ".", "Name", null, true);
    placeContentInTextLayer(
      profession,
      counter + 1 + ".",
      "Profession",
      null,
      true
    );
    placeContentInTextLayer(overdose, counter + 1 + ".", "Type", null, true);
    placeContentInTextLayer(
      yearOfDeath,
      counter + 1 + ".",
      "Year",
      null,
      false
    );
    placeContentInTextLayer(age, counter + 1 + ".", "Age", "Age", false);

    if (imagePath) {
      try {
        processImage(imagePath, frames[counter]);
      } catch (e) {}
    }
    counter++;
  }

  alert("Script completed successfully!");
}

app.activeDocument.suspendHistory("Import CSV Data", "main()");
