/**
 * Adobe Photoshop CSV Data Import Script
 * This script imports data from a CSV file and places it into a properly structured Photoshop document.
 * Compatible with Adobe Photoshop CS6 and later versions.
 * @version 1.0.0
 */

var Config = {
  TEXT: {
    MAX_WIDTH: 620,
    INITIAL_SIZE: 45,
    MIN_SIZE: 1,
  },
  FRAMES: {
    LEFT: "profile-frame-left",
    MIDDLE: "profile-frame-center",
    RIGHT: "profile-frame-right",
  },
  LAYERS: {
    PERSON_NAME: "profile-name",
    OCCUPATION: "profile-occupation",
    CAUSE: "profile-cause",
    DEATH_YEAR: "profile-year",
    AGE: "profile-age",
  },
  GROUP: {
    PREFIX: "profile-",
    AGE_SUBGROUP: "age-details",
  },
  MAX_PROFILES: 3,
  CSV: {
    EXPECTED_COLUMNS: 6,
    HEADERS: [
      "Name",
      "Profession",
      "Overdose",
      "Year of Death",
      "Age",
      "Image Path",
    ],
    REQUIRED_FIELDS: [0, 1, 2, 3, 4],
  },
  PATHS: {
    ALLOWED_IMAGE_EXTENSIONS: ["jpg", "jpeg", "png", "tif", "tiff", "psd"],
    MAX_PATH_LENGTH: 255,
  },
  DEBUG: {
    ENABLED: false,
    LOG_FILE: "photoshop_importer_log.txt",
    LOG_LEVELS: {
      INFO: "INFO",
      WARNING: "WARNING",
      ERROR: "ERROR",
      DEBUG: "DEBUG",
    },
  },
};

/**
 * Document handling namespace for managing Photoshop documents
 */
var DocumentManager = {
  /**
   * Creates a copy of the current document with a new name
   * @param {Document} sourceDoc - Source document
   * @param {string} newFileName - Name for the new file
   * @returns {Document} Newly created document
   */
  createCopy: function (sourceDoc, newFileName) {
    ErrorUtils.log(
      Config.DEBUG.LOG_LEVELS.INFO,
      "Creating document copy: " + newFileName
    );
    var sourcePath = sourceDoc.path;
    var newFilePath = sourcePath + "/" + newFileName;
    var newFile = new File(newFilePath);

    sourceDoc.saveAs(newFile, new PhotoshopSaveOptions(), true);
    return app.open(newFile);
  },

  /**
   * Gets a layer by name from a specific group
   * @param {LayerSet} group - Group to search in
   * @param {string} layerName - Name of the layer to find
   * @param {string} [subGroupName] - Optional subgroup name
   * @returns {ArtLayer|null} Found layer or null
   */
  getLayerFromGroup: function (group, layerName, subGroupName) {
    if (subGroupName) {
      var subGroups = group.layerSets;
      for (var i = 0; i < subGroups.length; i++) {
        if (subGroups[i].name === subGroupName) {
          group = subGroups[i];
          break;
        }
      }
    }

    var layers = group.artLayers;
    for (var j = 0; j < layers.length; j++) {
      if (layers[j].name === layerName) {
        return layers[j];
      }
    }
    return null;
  },
};

/**
 * Text handling namespace for managing text operations
 */
var TextManager = {
  /**
   * Places content in a specific text layer within a group
   * @param {string} content - Text content to place
   * @param {string} groupPrefix - Prefix of the group name
   * @param {string} textLayerName - Name of the target text layer
   * @param {string} [subGroupName] - Optional subgroup name
   * @param {boolean} [shouldAdjustSize] - Whether to adjust text size to fit
   * @returns {boolean} Success status
   */
  placeContent: function (
    content,
    groupPrefix,
    textLayerName,
    subGroupName,
    shouldAdjustSize
  ) {
    try {
      ErrorUtils.log(
        Config.DEBUG.LOG_LEVELS.DEBUG,
        "Placing content in layer: " +
          textLayerName +
          " (Group: " +
          groupPrefix +
          (subGroupName ? ", SubGroup: " + subGroupName : "") +
          ")"
      );

      var doc = app.activeDocument;
      var targetGroup = null;

      // Find target group
      var groups = doc.layerSets;
      for (var i = 0; i < groups.length; i++) {
        if (groups[i].name.indexOf(groupPrefix) === 0) {
          targetGroup = groups[i];
          break;
        }
      }

      if (!targetGroup) {
        throw new Error("Group not found: " + groupPrefix);
      }

      // Get target layer
      var textLayer = DocumentManager.getLayerFromGroup(
        targetGroup,
        textLayerName,
        subGroupName
      );
      if (!textLayer || textLayer.kind !== LayerKind.TEXT) {
        throw new Error("Text layer not found: " + textLayerName);
      }

      // Update content
      var textItem = textLayer.textItem;
      textItem.contents = content;

      // Adjust size if needed
      if (shouldAdjustSize) {
        this.adjustTextSize(textLayer);
      }

      return true;
    } catch (error) {
      ErrorUtils.handleError("Text Placement", error);
      return false;
    }
  },

  /**
   * Adjusts text size to fit within maximum width
   * @param {ArtLayer} textLayer - Text layer to adjust
   */
  adjustTextSize: function (textLayer) {
    var textItem = textLayer.textItem;
    textItem.size = Config.TEXT.INITIAL_SIZE;

    function getTextWidth() {
      var bounds = textLayer.bounds;
      return bounds[2] - bounds[0];
    }

    while (
      getTextWidth() > Config.TEXT.MAX_WIDTH &&
      textItem.size > Config.TEXT.MIN_SIZE
    ) {
      textItem.size -= 0.5;
    }
  },
};

/**
 * Image handling namespace for managing image operations
 */
var ImageManager = {
  /**
   * Imports and processes an image into a specified frame
   * @param {string} imagePath - Path to the image file
   * @param {string} frameName - Name of the target frame layer
   */
  processImage: function (imagePath, frameName) {
    try {
      ErrorUtils.log(
        Config.DEBUG.LOG_LEVELS.INFO,
        "Processing image for frame: " + frameName
      );

      var originalDoc = app.activeDocument;
      var targetFrame = originalDoc.layers.getByName(frameName);

      // Validate and normalize image path
      var validation = PathUtils.validatePath(
        imagePath,
        Config.PATHS.ALLOWED_IMAGE_EXTENSIONS
      );
      if (!validation.isValid) {
        throw new Error("Invalid image path: " + validation.error);
      }

      // Resolve relative paths against CSV location
      var resolvedPath = PathUtils.resolveRelativePath(
        app.activeDocument.path,
        imagePath
      );
      var imageFile = new File(resolvedPath);

      if (!imageFile.exists) {
        throw new Error("Image file not found: " + resolvedPath);
      }

      // Import and process image
      var importedDoc = app.open(imageFile);
      importedDoc.selection.selectAll();
      importedDoc.selection.copy();
      importedDoc.close(SaveOptions.DONOTSAVECHANGES);

      // Place and adjust image
      app.activeDocument = originalDoc;
      var imageLayer = originalDoc.paste();
      imageLayer.move(targetFrame, ElementPlacement.PLACEBEFORE);

      this.fitImageToFrame(imageLayer, targetFrame);

      ErrorUtils.log(
        Config.DEBUG.LOG_LEVELS.INFO,
        "Image processed successfully"
      );
    } catch (error) {
      ErrorUtils.handleError("Image Processing", error);
    }
  },

  /**
   * Fits an image layer to a frame
   * @param {ArtLayer} imageLayer - Image layer to adjust
   * @param {ArtLayer} frameLayer - Frame layer to fit to
   */
  fitImageToFrame: function (imageLayer, frameLayer) {
    // Calculate dimensions
    var frameWidth = frameLayer.bounds[2] - frameLayer.bounds[0];
    var frameHeight = frameLayer.bounds[3] - frameLayer.bounds[1];
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

    // Center in frame
    var frameCenterX = (frameLayer.bounds[0] + frameLayer.bounds[2]) / 2;
    var frameCenterY = (frameLayer.bounds[1] + frameLayer.bounds[3]) / 2;
    var imageCenterX = (imageLayer.bounds[0] + imageLayer.bounds[2]) / 2;
    var imageCenterY = (imageLayer.bounds[1] + imageLayer.bounds[3]) / 2;

    imageLayer.translate(
      frameCenterX - imageCenterX,
      frameCenterY - imageCenterY
    );
  },
};

/**
 * CSV handling namespace for managing CSV operations
 */
var CSVManager = {
  /**
   * Opens a file dialog to select a CSV file
   * @returns {File} Selected CSV file or null if cancelled
   */
  selectFile: function () {
    try {
      ErrorUtils.log(Config.DEBUG.LOG_LEVELS.INFO, "Opening CSV file dialog");

      var csvFile = File.openDialog("Select a CSV file", "*.csv");
      if (csvFile === null) {
        ErrorUtils.log(
          Config.DEBUG.LOG_LEVELS.INFO,
          "CSV file selection cancelled by user"
        );
        return null;
      }

      // Validate path
      var validation = PathUtils.validatePath(csvFile.fsName, ["csv"]);
      if (!validation.isValid) {
        throw new Error(validation.error);
      }

      ErrorUtils.log(
        Config.DEBUG.LOG_LEVELS.INFO,
        "Selected CSV file: " + csvFile.fsName
      );
      return csvFile;
    } catch (error) {
      ErrorUtils.handleError("CSV File Selection", error);
      return null;
    }
  },

  /**
   * Reads and parses a CSV file
   * @param {File} file - CSV file to read
   * @returns {Array<Array<string>>} Parsed CSV data
   */
  readAndParse: function (file) {
    try {
      ErrorUtils.log(
        Config.DEBUG.LOG_LEVELS.INFO,
        "Reading CSV file: " + file.fsName
      );

      var csvPath = PathUtils.normalize(file.fsName);
      file = new File(csvPath);

      file.open("r");
      var content = file.read();
      file.close();

      ErrorUtils.log(
        Config.DEBUG.LOG_LEVELS.DEBUG,
        "CSV content length: " + content.length
      );

      var lines = content.split("\n");
      var csvData = [];

      for (var i = 0; i < lines.length; i++) {
        var columns = lines[i].split(",");
        for (var j = 0; j < columns.length; j++) {
          columns[j] = columns[j].trim();
        }
        csvData.push(columns);
      }

      ErrorUtils.log(
        Config.DEBUG.LOG_LEVELS.INFO,
        "CSV parsing complete. Rows: " + csvData.length
      );
      return csvData;
    } catch (error) {
      ErrorUtils.handleError("CSV Reading", error);
      return null;
    }
  },

  /**
   * Validates CSV data structure and content
   * @param {Array<Array<string>>} csvData - CSV data to validate
   * @returns {Object} Validation result
   */
  validate: function (csvData) {
    try {
      if (!csvData || !csvData.length) {
        return {
          isValid: false,
          error: "CSV file is empty",
        };
      }

      // Check minimum rows
      if (csvData.length < 2) {
        return {
          isValid: false,
          error: "CSV file must contain headers and at least one data row",
        };
      }

      // Validate headers
      var headers = csvData[0];
      if (headers.length !== Config.CSV.EXPECTED_COLUMNS) {
        return {
          isValid: false,
          error:
            "CSV must have exactly " +
            Config.CSV.EXPECTED_COLUMNS +
            " columns: " +
            Config.CSV.HEADERS.join(", "),
        };
      }

      // Check header names
      for (var i = 0; i < Config.CSV.HEADERS.length; i++) {
        if (headers[i].trim() !== Config.CSV.HEADERS[i]) {
          return {
            isValid: false,
            error:
              "Invalid header: Expected '" +
              Config.CSV.HEADERS[i] +
              "' but found '" +
              headers[i].trim() +
              "'",
          };
        }
      }

      // Validate data rows
      for (var rowIndex = 1; rowIndex < csvData.length; rowIndex++) {
        var validationResult = this.validateRow(csvData[rowIndex], rowIndex);
        if (!validationResult.isValid) {
          return validationResult;
        }
      }

      return {
        isValid: true,
        error: null,
      };
    } catch (error) {
      ErrorUtils.handleError("CSV Validation", error);
      return {
        isValid: false,
        error: "Validation error: " + error.message,
      };
    }
  },

  /**
   * Validates a single CSV row
   * @param {Array<string>} row - Row to validate
   * @param {number} rowIndex - Index of the row (1-based)
   * @returns {Object} Validation result
   */
  validateRow: function (row, rowIndex) {
    // Skip empty rows
    if (row.length === 1 && !row[0].trim()) {
      return { isValid: true };
    }

    // Check column count
    if (row.length !== Config.CSV.EXPECTED_COLUMNS) {
      return {
        isValid: false,
        error:
          "Row " +
          (rowIndex + 1) +
          " has " +
          row.length +
          " columns instead of " +
          Config.CSV.EXPECTED_COLUMNS,
      };
    }

    // Validate required fields
    for (var i = 0; i < Config.CSV.REQUIRED_FIELDS.length; i++) {
      var colIndex = Config.CSV.REQUIRED_FIELDS[i];
      if (!row[colIndex] || !row[colIndex].trim()) {
        return {
          isValid: false,
          error:
            "Missing required value '" +
            Config.CSV.HEADERS[colIndex] +
            "' in row " +
            (rowIndex + 1),
        };
      }
    }

    // Validate numeric fields
    if (isNaN(parseInt(row[3]))) {
      // Year of Death
      return {
        isValid: false,
        error:
          "Invalid Year of Death in row " +
          (rowIndex + 1) +
          ": must be a number",
      };
    }

    if (isNaN(parseInt(row[4]))) {
      // Age
      return {
        isValid: false,
        error: "Invalid Age in row " + (rowIndex + 1) + ": must be a number",
      };
    }

    // Validate image path if provided
    if (row[5] && row[5].trim()) {
      var imageFile = new File(row[5].trim());
      if (!imageFile.exists) {
        return {
          isValid: false,
          error:
            "Invalid image path in row " + (rowIndex + 1) + ": file not found",
        };
      }
    }

    return { isValid: true };
  },
};

// Keep existing PathUtils and ErrorUtils as they are...

/**
 * Main application namespace
 */
var App = {
  /**
   * Processes a single profile entry
   * @param {Object} entryData - Profile data
   * @param {number} index - Profile index (0-based)
   * @param {Array<string>} frameNames - Available frame names
   * @returns {boolean} Success status
   */
  processProfile: function (entryData, index, frameNames) {
    try {
      ErrorUtils.log(
        Config.DEBUG.LOG_LEVELS.INFO,
        "Processing profile " + (index + 1)
      );

      var groupPrefix = Config.GROUP.PREFIX + (index + 1);

      // Place text content
      TextManager.placeContent(
        entryData.name,
        groupPrefix,
        Config.LAYERS.PERSON_NAME,
        null,
        true
      );

      TextManager.placeContent(
        entryData.profession,
        groupPrefix,
        Config.LAYERS.OCCUPATION,
        null,
        true
      );

      TextManager.placeContent(
        entryData.overdose,
        groupPrefix,
        Config.LAYERS.CAUSE,
        null,
        true
      );

      TextManager.placeContent(
        entryData.yearOfDeath,
        groupPrefix,
        Config.LAYERS.DEATH_YEAR,
        null,
        false
      );

      TextManager.placeContent(
        entryData.age,
        groupPrefix,
        Config.LAYERS.AGE,
        Config.GROUP.AGE_SUBGROUP,
        false
      );

      // Process image if provided
      if (entryData.imagePath) {
        ImageManager.processImage(entryData.imagePath, frameNames[index]);
      }

      ErrorUtils.log(
        Config.DEBUG.LOG_LEVELS.INFO,
        "Profile " + (index + 1) + " processed successfully"
      );
      return true;
    } catch (error) {
      ErrorUtils.handleError("Profile Processing", error);
      return false;
    }
  },

  /**
   * Main execution function
   */
  main: function () {
    try {
      // Initialize error handling
      ErrorUtils.init();
      ErrorUtils.log(Config.DEBUG.LOG_LEVELS.INFO, "Script execution started");

      var sourceDoc = app.activeDocument;
      ErrorUtils.log(
        Config.DEBUG.LOG_LEVELS.INFO,
        "Source document: " + sourceDoc.name
      );

      var frameNames = [
        Config.FRAMES.LEFT,
        Config.FRAMES.MIDDLE,
        Config.FRAMES.RIGHT,
      ];

      // Get and validate CSV data
      var csvFile = CSVManager.selectFile();
      if (!csvFile) return;

      var csvData = CSVManager.readAndParse(csvFile);
      if (!csvData || !csvData.length) {
        return ErrorUtils.handleError(
          "Main Execution",
          "Failed to read CSV data"
        );
      }

      ErrorUtils.log(Config.DEBUG.LOG_LEVELS.INFO, "Validating CSV data");
      var validationResult = CSVManager.validate(csvData);
      if (!validationResult.isValid) {
        return ErrorUtils.handleError("CSV Validation", validationResult.error);
      }

      var processedCount = 0;
      var workingDoc = null;

      // Process entries
      for (
        var i = 1;
        i < csvData.length && processedCount < Config.MAX_PROFILES;
        i++
      ) {
        // Skip empty rows
        if (csvData[i].length === 1 && !csvData[i][0].trim()) {
          continue;
        }

        // Create new document for first entry
        if (processedCount === 0) {
          ErrorUtils.log(Config.DEBUG.LOG_LEVELS.INFO, "Creating new document");
          workingDoc = DocumentManager.createCopy(
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

        if (this.processProfile(entryData, processedCount, frameNames)) {
          processedCount++;
        }
      }

      ErrorUtils.log(
        Config.DEBUG.LOG_LEVELS.INFO,
        "Script completed. Processed " + processedCount + " entries"
      );
      alert("Script completed successfully!");
    } catch (error) {
      ErrorUtils.handleError("Main Execution", error);
    } finally {
      ErrorUtils.cleanup();
    }
  },
};

// Execute the script with history tracking
app.activeDocument.suspendHistory("Import CSV Data", "App.main()");
