/**
 * Config.gs - Core Configuration Settings
 */
var DEFAULT_TEMPLATE_ID = "1mjjPumCPBEvoChwuKyH-_i_S94yo6mnAGrZSiqY7gxA";

// PDF Destination Folders
var WORK_ORDER_PDF_FOLDER_ID = "1dj3qvrFBOIO0abiDF-DTQCJnQ2OjlWVD"; // Primary Active PDF Folder
var ARCHIVE_FOLDER_ID        = "1-MOTuaFbJfDHHxyY8bvXYTHuF3jybUBF"; // Previous Revisions Archive Folder

var CONFIG = {
  SHEET_NAME: "Serial Units",
  COLS: {
    UNIT_ID: 1,         // Col A: Serial Unit ID
    SERIAL_NUMBER: 2,   // Col B: Serial Number
    PRODUCT_NUMBER: 3,  // Col C: Product Number
    BUILD_WO_ID: 4,     // Col D: Build WO ID
    CONSUMED_WO_ID: 5,  // Col E: Consumed WO ID
    PARENT_SERIAL: 6,   // Col F: Parent Unit Serial
    STATUS: 7,          // Col G: Status
    CREATED_DATE: 8,    // Col H: Created Date
    CONSUMED_DATE: 9    // Col I: Consumed Date
  }
};

var PAGE2_TABLE_CONFIGS = [
  { targetKey: "WO Traveler Checks", sheetNames: ["WO Traveler Checks", "Traveler Checks", "Critical Assembly Checks"] },
  { targetKey: "WO Deviations", sheetNames: ["WO Deviations", "Deviations"] },
  { targetKey: "WO BOM Changes", sheetNames: ["WO BOM Changes", "BOM Changes"] }
];