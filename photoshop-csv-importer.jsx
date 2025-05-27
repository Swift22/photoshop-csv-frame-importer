/**
 * Adobe Photoshop CSV Data Import Script
 * This script imports data from a CSV file and places it into a properly structured Photoshop document.
 * Compatible with Adobe Photoshop CS6 and later versions.
 */

// System Constants
var LAYOUT_CONFIG = {
  TEXT: {
    MAX_WIDTH: 620, // Maximum width for text layers in pixels
    INITIAL_SIZE: 45, // Starting font size for text scaling
    MIN_SIZE: 1, // Minimum allowed font size
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
    PREFIX: "profile-", // Will be followed by number (e.g., "profile-1")
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
    REQUIRED_FIELDS: [0, 1, 2, 3, 4], // Indices of required fields (all except image)
  },
  PATHS: {
    ALLOWED_IMAGE_EXTENSIONS: ["jpg", "jpeg", "png", "tif", "tiff", "psd"],
    MAX_PATH_LENGTH: 255,
  },
  DEBUG: {
    ENABLED: false, // Can be toggled via dialog
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
 * Error handling and logging utilities
 */
var ErrorUtils = {
  logFile: null,
  debugMode: false,

  /**
   * Initializes error handling and logging
   */
  init: function () {
    // Ask user if they want debug mode
    this.debugMode = confirm(
      "Enable debug mode?\nThis will create a detailed log file in your documents folder."
    );
    LAYOUT_CONFIG.DEBUG.ENABLED = this.debugMode;

    if (this.debugMode) {
      try {
        // Create log file in user's documents folder
        var logPath = Folder.myDocuments + "/" + LAYOUT_CONFIG.DEBUG.LOG_FILE;
        this.logFile = new File(logPath);

        // Append to existing log or create new
        this.logFile.open("a");
        this.log(
          LAYOUT_CONFIG.DEBUG.LOG_LEVELS.INFO,
          "=== New Session Started ==="
        );
        this.log(
          LAYOUT_CONFIG.DEBUG.LOG_LEVELS.INFO,
          "Photoshop Version: " + app.version
        );
        this.log(LAYOUT_CONFIG.DEBUG.LOG_LEVELS.INFO, "OS: " + $.os);
        this.log(LAYOUT_CONFIG.DEBUG.LOG_LEVELS.INFO, "Script Version: 1.0.0");
      } catch (e) {
        alert("Failed to initialize logging: " + e);
      }
    }
  },

  /**
   * Logs a message with timestamp and level
   * @param {string} level - Log level
   * @param {string} message - Message to log
   * @param {Error} [error] - Optional error object
   */
  log: function (level, message, error) {
    if (!this.debugMode) return;

    try {
      var timestamp = new Date().toISOString();
      var logMessage = timestamp + " [" + level + "] " + message;

      if (error) {
        logMessage += "\nError: " + error.message;
        if (error.stack) {
          logMessage += "\nStack: " + error.stack;
        }
      }

      logMessage += "\n";
      this.logFile.write(logMessage);
      this.logFile.flush(); // Ensure message is written immediately
    } catch (e) {
      alert("Logging failed: " + e);
    }
  },

  /**
   * Handles an error with optional logging
   * @param {string} context - Where the error occurred
   * @param {Error|string} error - Error object or message
   * @param {boolean} [showAlert=true] - Whether to show alert to user
   */
  handleError: function (context, error, showAlert) {
    var errorMsg = error.message || error;
    var fullMessage = context + ": " + errorMsg;

    // Log error
    if (this.debugMode) {
      this.log(LAYOUT_CONFIG.DEBUG.LOG_LEVELS.ERROR, context, error);
    }

    // Show alert if needed
    if (showAlert !== false) {
      alert(fullMessage);
    }

    return false; // For error handling chains
  },

  /**
   * Closes log file and performs cleanup
   */
  cleanup: function () {
    if (this.debugMode && this.logFile) {
      try {
        this.log(
          LAYOUT_CONFIG.DEBUG.LOG_LEVELS.INFO,
          "=== Session Ended ===\n"
        );
        this.logFile.close();
      } catch (e) {
        alert("Failed to close log file: " + e);
      }
    }
  },
};

/**
 * Utilities for handling file paths
 */
var PathUtils = {
  /**
   * Normalizes file path separators for current OS
   * @param {string} path - File path to normalize
   * @returns {string} Normalized path
   */
  normalize: function (path) {
    if (!path) return "";
    // Convert all slashes to system-specific separator
    return path.replace(/[\\/]+/g, Folder.fs === "Windows" ? "\\" : "/").trim();
  },

  /**
   * Resolves a relative path against the CSV file location
   * @param {string} basePath - Base directory path (CSV file location)
   * @param {string} relativePath - Relative path to resolve
   * @returns {string} Resolved absolute path
   */
  resolveRelativePath: function (basePath, relativePath) {
    if (!relativePath) return "";

    // If already absolute, return normalized
    if (this.isAbsolutePath(relativePath)) {
      return this.normalize(relativePath);
    }

    // Get CSV directory as base
    var baseDir = Folder(basePath).parent;
    var resolved =
      baseDir + Folder.fs === "Windows"
        ? "\\"
        : "/" + this.normalize(relativePath);
    return resolved;
  },

  /**
   * Checks if path is absolute
   * @param {string} path - Path to check
   * @returns {boolean} True if absolute path
   */
  isAbsolutePath: function (path) {
    if (!path) return false;
    if (Folder.fs === "Windows") {
      return /^[A-Za-z]:\\/.test(path) || /^\\\\/.test(path);
    }
    return path.charAt(0) === "/";
  },

  /**
   * Validates a file path
   * @param {string} path - Path to validate
   * @param {Array<string>} [allowedExtensions] - List of allowed file extensions
   * @returns {Object} Validation result with status and error message
   */
  validatePath: function (path, allowedExtensions) {
    var result = {
      isValid: false,
      error: null,
    };

    if (!path) {
      result.error = "Empty file path";
      return result;
    }

    // Check path length
    if (path.length > LAYOUT_CONFIG.PATHS.MAX_PATH_LENGTH) {
      result.error =
        "File path exceeds maximum length of " +
        LAYOUT_CONFIG.PATHS.MAX_PATH_LENGTH +
        " characters";
      return result;
    }

    // Check for invalid characters
    var invalidChars =
      Folder.fs === "Windows" ? /[<>:"|?*\x00-\x1F]/g : /[\x00\/]/g;
    if (invalidChars.test(path)) {
      result.error = "File path contains invalid characters";
      return result;
    }

    // Check extension if provided
    if (allowedExtensions && allowedExtensions.length > 0) {
      var ext = path.split(".").pop().toLowerCase();
      if (allowedExtensions.indexOf(ext) === -1) {
        result.error =
          "Invalid file extension. Allowed: " + allowedExtensions.join(", ");
        return result;
      }
    }

    result.isValid = true;
    return result;
  },
};

/**
 * Validates the structure and content of CSV data
 * @param {Array<Array<string>>} csvData - The parsed CSV data
 * @returns {Object} Validation result with status and error message
 */
function validateCSVData(csvData) {
  if (!csvData || !csvData.length) {
    return {
      isValid: false,
      error: "CSV file is empty",
    };
  }

  // Check if we have at least one data row
  if (csvData.length < 2) {
    return {
      isValid: false,
      error: "CSV file must contain headers and at least one data row",
    };
  }

  // Validate header row
  var headers = csvData[0];
  if (headers.length !== LAYOUT_CONFIG.CSV.EXPECTED_COLUMNS) {
    return {
      isValid: false,
      error:
        "CSV must have exactly " +
        LAYOUT_CONFIG.CSV.EXPECTED_COLUMNS +
        " columns: " +
        LAYOUT_CONFIG.CSV.HEADERS.join(", "),
    };
  }

  // Validate header names
  for (var i = 0; i < LAYOUT_CONFIG.CSV.HEADERS.length; i++) {
    if (headers[i].trim() !== LAYOUT_CONFIG.CSV.HEADERS[i]) {
      return {
        isValid: false,
        error:
          "Invalid header: Expected '" +
          LAYOUT_CONFIG.CSV.HEADERS[i] +
          "' but found '" +
          headers[i].trim() +
          "'",
      };
    }
  }

  // Validate data rows
  for (var rowIndex = 1; rowIndex < csvData.length; rowIndex++) {
    var row = csvData[rowIndex];

    // Skip empty rows
    if (row.length === 1 && !row[0].trim()) {
      continue;
    }

    // Check column count
    if (row.length !== LAYOUT_CONFIG.CSV.EXPECTED_COLUMNS) {
      return {
        isValid: false,
        error:
          "Row " +
          (rowIndex + 1) +
          " has " +
          row.length +
          " columns instead of " +
          LAYOUT_CONFIG.CSV.EXPECTED_COLUMNS,
      };
    }

    // Validate required fields
    for (
      var fieldIndex = 0;
      fieldIndex < LAYOUT_CONFIG.CSV.REQUIRED_FIELDS.length;
      fieldIndex++
    ) {
      var colIndex = LAYOUT_CONFIG.CSV.REQUIRED_FIELDS[fieldIndex];
      if (!row[colIndex] || !row[colIndex].trim()) {
        return {
          isValid: false,
          error:
            "Missing required value '" +
            LAYOUT_CONFIG.CSV.HEADERS[colIndex] +
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
  }

  return {
    isValid: true,
    error: null,
  };
}

/**
 * Opens a file dialog to select a CSV file
 * @returns {File} Selected CSV file or null if cancelled
 */
function selectCSVFile() {
  try {
    ErrorUtils.log(
      LAYOUT_CONFIG.DEBUG.LOG_LEVELS.INFO,
      "Opening CSV file dialog"
    );

    var csvFile = File.openDialog("Select a CSV file", "*.csv");
    if (csvFile === null) {
      ErrorUtils.log(
        LAYOUT_CONFIG.DEBUG.LOG_LEVELS.INFO,
        "CSV file selection cancelled by user"
      );
      return null;
    }

    // Validate path
    var validation = PathUtils.validatePath(csvFile.fsName, ["csv"]);
    if (!validation.isValid) {
      return ErrorUtils.handleError("CSV File Selection", validation.error);
    }

    ErrorUtils.log(
      LAYOUT_CONFIG.DEBUG.LOG_LEVELS.INFO,
      "Selected CSV file: " + csvFile.fsName
    );
    return csvFile;
  } catch (error) {
    return ErrorUtils.handleError("CSV File Selection", error);
  }
}

/**
 * Reads and parses a CSV file into a 2D array
 * @param {File} file - The CSV file to read
 * @returns {Array<Array<string>>} Parsed CSV data as 2D array
 */
function readCSV(file) {
  try {
    ErrorUtils.log(
      LAYOUT_CONFIG.DEBUG.LOG_LEVELS.INFO,
      "Reading CSV file: " + file.fsName
    );

    var csvPath = PathUtils.normalize(file.fsName);
    file = new File(csvPath);

    file.open("r");
    var content = file.read();
    file.close();

    ErrorUtils.log(
      LAYOUT_CONFIG.DEBUG.LOG_LEVELS.DEBUG,
      "CSV content length: " + content.length
    );

    var lines = content.split("\n");
    var csvData = [];

    for (var i = 0; i < lines.length; i++) {
      var columns = lines[i].split(",");
      // Trim whitespace from each column
      for (var j = 0; j < columns.length; j++) {
        columns[j] = columns[j].trim();
      }
      csvData.push(columns);
    }

    ErrorUtils.log(
      LAYOUT_CONFIG.DEBUG.LOG_LEVELS.INFO,
      "CSV parsing complete. Rows: " + csvData.length
    );
    return csvData;
  } catch (error) {
    return ErrorUtils.handleError("CSV Reading", error);
  }
}

/**
 * Imports and processes an image into a specified frame
 * @param {string} imagePath - Path to the image file
 * @param {string} frameName - Name of the target frame layer
 */
function processImage(imagePath, frameName) {
  var originalDoc = app.activeDocument;
  var targetFrame = originalDoc.layers.getByName(frameName);

  // Validate and normalize image path
  var validation = PathUtils.validatePath(
    imagePath,
    LAYOUT_CONFIG.PATHS.ALLOWED_IMAGE_EXTENSIONS
  );
  if (!validation.isValid) {
    alert("Invalid image path: " + validation.error);
    return;
  }

  // Resolve relative paths against CSV location
  var resolvedPath = PathUtils.resolveRelativePath(
    app.activeDocument.path,
    imagePath
  );
  var imageFile = new File(resolvedPath);

  if (!imageFile.exists) {
    alert("Image file not found: " + resolvedPath);
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
    alert("Error processing image: " + resolvedPath + "\nError: " + error);
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

          textItem.size = LAYOUT_CONFIG.TEXT.INITIAL_SIZE;
          while (
            getTextWidth() > LAYOUT_CONFIG.TEXT.MAX_WIDTH &&
            textItem.size > LAYOUT_CONFIG.TEXT.MIN_SIZE
          ) {
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
  try {
    // Initialize error handling
    ErrorUtils.init();
    ErrorUtils.log(
      LAYOUT_CONFIG.DEBUG.LOG_LEVELS.INFO,
      "Script execution started"
    );

    var sourceDoc = app.activeDocument;
    ErrorUtils.log(
      LAYOUT_CONFIG.DEBUG.LOG_LEVELS.INFO,
      "Source document: " + sourceDoc.name
    );

    var frameNames = [
      LAYOUT_CONFIG.FRAMES.LEFT,
      LAYOUT_CONFIG.FRAMES.MIDDLE,
      LAYOUT_CONFIG.FRAMES.RIGHT,
    ];

    // Get CSV data
    var csvFile = selectCSVFile();
    if (csvFile === null) {
      ErrorUtils.log(
        LAYOUT_CONFIG.DEBUG.LOG_LEVELS.INFO,
        "Script terminated - No CSV file selected"
      );
      return;
    }

    var csvData = readCSV(csvFile);
    if (!csvData || !csvData.length) {
      return ErrorUtils.handleError(
        "Main Execution",
        "Failed to read CSV data"
      );
    }

    // Validate CSV data
    ErrorUtils.log(LAYOUT_CONFIG.DEBUG.LOG_LEVELS.INFO, "Validating CSV data");
    var validationResult = validateCSVData(csvData);
    if (!validationResult.isValid) {
      return ErrorUtils.handleError("CSV Validation", validationResult.error);
    }

    var processedCount = 0;

    // Process entries
    for (
      var i = 1;
      i < csvData.length && processedCount < LAYOUT_CONFIG.MAX_PROFILES;
      i++
    ) {
      try {
        ErrorUtils.log(
          LAYOUT_CONFIG.DEBUG.LOG_LEVELS.INFO,
          "Processing entry " + i
        );

        // Skip empty rows
        if (csvData[i].length === 1 && !csvData[i][0].trim()) {
          ErrorUtils.log(
            LAYOUT_CONFIG.DEBUG.LOG_LEVELS.DEBUG,
            "Skipping empty row " + i
          );
          continue;
        }

        // Create new document for first entry
        if (processedCount === 0) {
          ErrorUtils.log(
            LAYOUT_CONFIG.DEBUG.LOG_LEVELS.INFO,
            "Creating new document"
          );
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

        var groupPrefix = LAYOUT_CONFIG.GROUP.PREFIX + (processedCount + 1);

        // Place text content
        placeContentInTextLayer(
          entryData.name,
          groupPrefix,
          LAYOUT_CONFIG.LAYERS.PERSON_NAME,
          null,
          true
        );
        placeContentInTextLayer(
          entryData.profession,
          groupPrefix,
          LAYOUT_CONFIG.LAYERS.OCCUPATION,
          null,
          true
        );
        placeContentInTextLayer(
          entryData.overdose,
          groupPrefix,
          LAYOUT_CONFIG.LAYERS.CAUSE,
          null,
          true
        );
        placeContentInTextLayer(
          entryData.yearOfDeath,
          groupPrefix,
          LAYOUT_CONFIG.LAYERS.DEATH_YEAR,
          null,
          false
        );
        placeContentInTextLayer(
          entryData.age,
          groupPrefix,
          LAYOUT_CONFIG.LAYERS.AGE,
          LAYOUT_CONFIG.GROUP.AGE_SUBGROUP,
          false
        );

        // Process image if path provided
        if (entryData.imagePath) {
          try {
            processImage(entryData.imagePath, frameNames[processedCount]);
          } catch (error) {
            // Error already handled in processImage
          }
        }

        processedCount++;
        ErrorUtils.log(
          LAYOUT_CONFIG.DEBUG.LOG_LEVELS.INFO,
          "Successfully processed entry " + i
        );
      } catch (error) {
        ErrorUtils.handleError(
          "Entry Processing",
          "Failed to process entry " + i + ": " + error.message,
          true
        );
        // Continue with next entry
      }
    }

    ErrorUtils.log(
      LAYOUT_CONFIG.DEBUG.LOG_LEVELS.INFO,
      "Script completed. Processed " + processedCount + " entries"
    );
    alert("Script completed successfully!");
  } catch (error) {
    ErrorUtils.handleError("Main Execution", error);
  } finally {
    // Cleanup
    ErrorUtils.cleanup();
  }
}

// Execute the script with history tracking
app.activeDocument.suspendHistory("Import CSV Data", "main()");
