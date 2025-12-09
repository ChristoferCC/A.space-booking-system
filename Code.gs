// REPLACE THIS WITH YOUR SPREADSHEET ID (from the URL)
const DB_SHEET_ID = '1SUUjkv2GIQDvUedoIuKdwTaWdF7SDDi64-2cZMaNxGc'; 
const DB_SHEET_NAME = 'Users'; 
const CATEGORIES_SHEET_NAME = 'Categories';
const RESOURCES_SHEET_NAME = 'Resources';
const RESERVATIONS_SHEET_NAME = 'Reservations'; 
const FLOOR_PLANS_SHEET_NAME = 'FloorPlans'; 

// NEW: Explicit Philippine Time Zone constant
const PHILIPPINES_TIMEZONE = 'Asia/Manila'; 

// Folder ID where Resource Photos (Vehicle and Seat) will be saved
const DRIVE_FOLDER_ID = '1rZvRS4nSrhfOEQE-SeD2ft6an3E1oJ7Z';

/**
 * This function serves the HTML file (index.html).
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle('A.Space Login Demo');
}

/**
 * Retrieves data from the Google Sheet and validates the login.
 * @param {Object} formData - The login form data (email and password).
 * @returns {Object} Login status and user details if successful.
 */
function checkLogin(formData) {
  try {
    const loginEmail = formData.email.trim().toLowerCase();

    // --- SECURITY PATCH: BLOCK INVALID DOMAINS ---
    // Even if they are in the database, if not @gmail.com, reject login.
    if (!loginEmail.endsWith('@gmail.com')) {
       return { success: false, message: 'Access Denied: This domain is not authorized.' };
    }
    // ---------------------------------------------

    const ss = SpreadsheetApp.openById(DB_SHEET_ID);
    const sheet = ss.getSheetByName(DB_SHEET_NAME);
    const data = sheet.getDataRange().getValues(); 

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const email = row[0] ? String(row[0]).trim() : '';
      const password = row[1] ? String(row[1]).trim() : '';
      
      // Compare Credentials
      if (loginEmail === email.toLowerCase() && formData.password === password) {
        
        const role = row[2] ? String(row[2]).trim() : 'user'; 
        const name = row[3] ? String(row[3]).trim() : '';
        const businessUnit = row[4] ? String(row[4]).trim() : '';

        return { 
          success: true, 
          email: email, 
          role: role,
          name: name,
          businessUnit: businessUnit
        };
      }
    }
    return { success: false, message: 'Invalid email or password. Please try again.' };

  } catch (e) {
    Logger.log("Error in checkLogin: " + e.toString());
    return { success: false, message: 'System Error: Could not access database.' };
  }
}

/**
 * Updates the user's password in the Google Sheet.
 * @param {string} email - The user's email.
 * @param {string} currentPassword - The current password.
 * @param {string} newPassword - The new password.
 * @returns {Object} Update status.
 */
function updatePassword(email, currentPassword, newPassword) {
  try {
    const ss = SpreadsheetApp.openById(DB_SHEET_ID);
    const sheet = ss.getSheetByName(DB_SHEET_NAME);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const userEmail = String(row[0]).trim();
      const storedPassword = String(row[1]).trim();
      const rowNumber = i + 1;

      if (email === userEmail) {
        if (currentPassword === storedPassword) {
          sheet.getRange(rowNumber, 2).setValue(newPassword);
          return { success: true, message: 'Password successfully updated in Google Sheet.' };
        } else {
          return { success: false, message: 'Invalid Current Password.' };
        }
      }
    }
    return { success: false, message: 'User not found in the database.' };

  } catch (e) {
    Logger.log("Error in updatePassword: " + e.toString());
    return { success: false, message: 'System Error: Failed to update password.' };
  }
}

/**
 * Retrieves all users for the Admin Dashboard.
 * @returns {Object} Success status and array of user objects.
 */
function getAllUsers() {
  try {
    const ss = SpreadsheetApp.openById(DB_SHEET_ID);
    const sheet = ss.getSheetByName(DB_SHEET_NAME);
    
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, users: [] };
    }
    
    // Get all data up to the last column (Col E = 5)
    const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5);
    const values = range.getValues();

    const users = values.map(row => ({
      email: String(row[0] || '').trim(),
      password: String(row[1] || '').trim(), 
      role: String(row[2] || 'user').trim(),
      name: String(row[3] || '').trim(), // Column D
      businessUnit: String(row[4] || '').trim() // Column E
    }));

    return { success: true, users: users };
  } catch (e) {
    Logger.log("Error in getAllUsers: " + e.toString());
    return { success: false, message: 'Error retrieving user data: ' + e.toString() };
  }
}

/**
 * Adds a new user to the Google Sheet.
 * UPDATED: With Domain Restriction (@gmail.com ONLY)
 * ADDED: Email Notification to New User
 * @param {Object} userData - The new user data.
 * @returns {Object} Success status and message.
 */
function addNewUser(userData) {
  try {
    const ss = SpreadsheetApp.openById(DB_SHEET_ID);
    const sheet = ss.getSheetByName(DB_SHEET_NAME);
    
    // 1. DOMAIN RESTRICTION CHECK (This was missing before)
    const emailCheck = userData.email.trim().toLowerCase();
    
    // Change this condition based on your preference.
    // For now, we are STRICT @gmail.com only:
    if (!emailCheck.endsWith('@gmail.com')) {
        return { success: false, message: 'Failed: Only @gmail.com addresses are allowed.' };
    }
    
    // Check if email already exists 
    if (sheet.getLastRow() > 1) {
        const emails = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat().map(String);
        if (emails.includes(userData.email)) {
          return { success: false, message: `Email ${userData.email} already exists.` };
        }
    }
    
    // Format the roles
    const rolesString = userData.roles.join(', ');
    
    // Combine Names (First + Middle + Last)
    const fullName = `${userData.firstName} ${userData.middleName ? userData.middleName + ' ' : ''}${userData.lastName}`.trim();
    
    // New row data: [Email, Password, Role, Name, Business Unit]
    const newRow = [
      userData.email,
      userData.password, 
      rolesString,
      fullName, 
      userData.businessUnit
    ];
    
    sheet.appendRow(newRow);
    
    // --- NEW: SEND WELCOME EMAIL ---
    sendWelcomeEmail(userData.email, fullName, userData.password);
    // -----------------------------
    
    return { success: true, message: 'New user successfully added to the database.' };

  } catch (e) {
    Logger.log("Error in addNewUser: " + e.toString());
    return { success: false, message: 'System Error: Failed to add new user.' };
  }
}

/**
 * Updates user details in the Google Sheet.
 * UPDATED: Handles First, Middle, and Last Name
 * @param {Object} userData - The user data to update.
 * @returns {Object} Success status and message.
 */
function updateUserDetails(userData) {
  try {
    const ss = SpreadsheetApp.openById(DB_SHEET_ID);
    const sheet = ss.getSheetByName(DB_SHEET_NAME);
    const data = sheet.getDataRange().getValues();

    let userRowIndex = -1;
    let originalPassword = '';

    // 1. Find the user row
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const userEmail = String(row[0]).trim();
      
      if (userData.originalEmail === userEmail) {
        userRowIndex = i; // Array index (0-based)
        originalPassword = String(row[1]).trim(); // Current password
        break;
      }
    }

    if (userRowIndex === -1) {
      return { success: false, message: 'User not found for update.' };
    }

    const rowNumber = userRowIndex + 1; // Sheet row number (1-based)

    // 2. Check for Email uniqueness (if new email is different)
    if (userData.newEmail !== userData.originalEmail) {
        for (let i = 1; i < data.length; i++) {
            if (i !== userRowIndex) { // Don't check against self
                const existingEmail = String(data[i][0]).trim();
                if (userData.newEmail === existingEmail) {
                    return { success: false, message: `Email ${userData.newEmail} already exists for another user.` };
                }
            }
        }
    }
    
    // Format the roles
    const rolesString = userData.roles.join(', ');
    
    // Combine Names
    const newFullName = `${userData.firstName} ${userData.middleName ? userData.middleName + ' ' : ''}${userData.lastName}`.trim();

    // Data to update
    const newEmail = userData.newEmail;
    const newPassword = userData.newPassword ? userData.newPassword : originalPassword; 
    const newRole = rolesString;
    const newName = newFullName;
    const newBusinessUnit = userData.businessUnit;

    // Update row
    sheet.getRange(rowNumber, 1, 1, 5).setValues([[
        newEmail,
        newPassword, 
        newRole,
        newName,
        newBusinessUnit
    ]]);

    return { success: true, message: `User ${newName} successfully updated.` };

  } catch (e) {
    Logger.log("Error in updateUserDetails: " + e.toString());
    return { success: false, message: 'System Error: Failed to update user details: ' + e.toString() };
  }
}

/**
 * Retrieves all Categories for Admin Categories Management.
 * @returns {Object} Success status and categorized object.
 */
function getAllCategories() {
  try {
    const ss = SpreadsheetApp.openById(DB_SHEET_ID);
    const sheet = ss.getSheetByName(CATEGORIES_SHEET_NAME);
    
    if (!sheet) {
      // If the Categories sheet doesn't exist, provide a default data structure
      return { 
        success: false, 
        message: 'Categories sheet not found. Please create a sheet named "Categories" in your spreadsheet.',
        categories: { room: [], desk: [], seat: [], vehicle: [] } 
      };
    }
    
    if (sheet.getLastRow() < 2) {
        return { 
            success: true, 
            categories: { room: [], desk: [], seat: [], vehicle: [] } 
        };
    }
    
    // Get all data (Only Columns A and B are needed)
    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    
    // Start data from Row 2 (i.e., i = 1)
    const categoryMap = { room: [], desk: [], seat: [], vehicle: [] };

    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const type = String(row[0] || '').toLowerCase().trim(); // Column A: Type
      const name = String(row[1] || '').trim(); // Column B: Category Name
      
      // Ensure Type is a valid key
      if (name && categoryMap[type]) {
        categoryMap[type].push(name);
      }
    }

    return { success: true, categories: categoryMap };
  } catch (e) {
    Logger.log("Error in getAllCategories: " + e.toString());
    return { success: false, message: 'Error retrieving category data: ' + e.toString(), categories: { room: [], desk: [], seat: [], vehicle: [] } };
  }
}

/**
 * Deletes a category from the Google Sheet.
 * @param {string} type - The resource type (room, desk, etc.).
 * @param {string} name - The category name.
 * @returns {Object} Success status and message.
 */
function removeCategory(type, name) {
  try {
    const ss = SpreadsheetApp.openById(DB_SHEET_ID);
    const sheet = ss.getSheetByName(CATEGORIES_SHEET_NAME);
    
    if (!sheet || sheet.getLastRow() < 2) {
        return { success: false, message: 'Categories sheet is empty or not found.' };
    }
    
    const data = sheet.getDataRange().getValues();

    let rowToDelete = -1;
    
    // Find the row index
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const categoryType = String(row[0] || '').toLowerCase().trim();
      const categoryName = String(row[1] || '').trim();
      const rowNumber = i + 1; // Row number in the sheet (1-based)

      if (categoryType === type && categoryName === name) {
        rowToDelete = rowNumber;
        break; 
      }
    }

    if (rowToDelete !== -1) {
      sheet.deleteRow(rowToDelete);
      return { success: true, message: `Category "${name}" removed successfully.` };
    } else {
      return { success: false, message: `Category "${name}" not found under type "${type}".` };
    }

  } catch (e) {
    Logger.log("Error in removeCategory: " + e.toString());
    return { success: false, message: 'System Error: Failed to remove category.' };
  }
}

/**
 * Adds a new category to the Google Sheet.
 * @param {string} type - The resource type (room, desk, etc.).
 * @param {string} name - The new category name.
 * @returns {Object} Success status and message.
 */
function addNewCategory(type, name) {
  try {
    const ss = SpreadsheetApp.openById(DB_SHEET_ID);
    const sheet = ss.getSheetByName(CATEGORIES_SHEET_NAME);
    
    if (!sheet) {
      return { success: false, message: 'Categories sheet not found. Please create a sheet named "Categories".' };
    }

    const typeLower = type.toLowerCase().trim();
    const nameTrimmed = name.trim();

    if (!nameTrimmed) {
      return { success: false, message: 'Category name cannot be empty.' };
    }
    
    // Check if category name already exists for that type
    if (sheet.getLastRow() > 1) {
        const data = sheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
          const row = data[i];
          const categoryType = String(row[0] || '').toLowerCase().trim();
          const categoryName = String(row[1] || '').trim();
          
          if (categoryType === typeLower && categoryName.toLowerCase() === nameTrimmed.toLowerCase()) {
            return { success: false, message: `Category "${nameTrimmed}" already exists under type "${type}".` };
          }
        }
    }
    
    // New row data: [Type, Category Name]
    const newRow = [
      typeLower, 
      nameTrimmed
    ];
    
    sheet.appendRow(newRow);
    
    return { success: true, message: `New category "${nameTrimmed}" successfully added to ${type} list.` };

  } catch (e) {
    Logger.log("Error in addNewCategory: " + e.toString());
    return { success: false, message: 'System Error: Failed to add new category. ' + e.toString() };
  }
}

/**
 * UPDATED: addNewResource now uses 16 columns (Col P for Seat Count).
 * @param {Object} form - The resource creation form data.
 * @returns {Object} Success status and message.
 */
function addNewResource(form) {
  try {
    const ss = SpreadsheetApp.openById(DB_SHEET_ID);
    let sheet = ss.getSheetByName(RESOURCES_SHEET_NAME);
    
    // Ensure the sheet has 16 columns
    if (!sheet || sheet.getLastRow() < 2) {
      sheet = ss.insertSheet(RESOURCES_SHEET_NAME);
      // New Headers (16 Columns)
      sheet.getRange(1, 1, 1, 16).setValues([['Type', 'Name/Description', 'Location/Garage', 'Capacity (Pax)', 'Category', 'Floor', 'Amenities', 'Accessible by BU', 'Requires Approval', 'Status', 'Photo URL', 'Vehicle Model', 'Plate Number', 'Layout Type', 'Layout Config', 'Seat Count']]);
    }
    
    // 1. Photo Upload Logic 
    let photoUrl = '';
    let fileBlob = null;
    
    if (form.vehiclePhoto && typeof form.vehiclePhoto.getName === 'function' && form.vehiclePhoto.getName() !== '') {
        fileBlob = form.vehiclePhoto;
    } else if (form.seatPhoto && typeof form.seatPhoto.getName === 'function' && form.seatPhoto.getName() !== '') { 
        fileBlob = form.seatPhoto;
    }

    if (fileBlob) {
      try {
        const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
        const file = folder.createFile(fileBlob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        photoUrl = "https://lh3.googleusercontent.com/d/" + file.getId(); 
      } catch (e) {
        Logger.log("Drive Upload Error: " + e.toString());
        photoUrl = 'Upload Error';
      }
    }
    
    // 2. Data Preparation
    const type = form.type; 
    const name = form.name.trim();
    const location = form.location;
    const category = form.category;
    const accessibleBU = Array.isArray(form.accessibleBU) ? form.accessibleBU.join(', ') : (form.accessibleBU || '');
    const requiresApproval = form.requiresApproval === 'on' ? 'Yes' : 'No'; 
    
    // --- Capacity (Col D) / Seat Count (Col P) Logic ---
    let capacity = form.capacity || ''; // Capacity (Col D)
    let seatCount = ''; // Seat Count (Col P)
    
    // --- NEW DEDICATED COLUMN DATA (Defaults) ---
    let vehicleModel = '';
    let plateNumber = '';
    let layoutType = '';
    let layoutConfig = '{}';
    
    // --- CLEANED OLD COLUMNS ---
    let floor = form.floor ? form.floor.trim() : '';
    let amenities = Array.isArray(form.amenities) ? form.amenities.join(', ') : (form.amenities || '');
    
    // --- VEHICLE LOGIC ---
    if (type === 'vehicle') {
        vehicleModel = form.model.trim();
        plateNumber = form.plateNumber.trim();
        capacity = form.capacity || ''; // Populate Col D
        seatCount = ''; // CLEAR P
        floor = ''; // CLEAR F
        amenities = ''; // CLEAR G
    }
    // --- SEAT LOGIC (UPDATED: Use the new seatCount variable) ---
    else if (type === 'seat' || type === 'car seat') {
        seatCount = form.seatCount || '0'; // Value in Col P
        capacity = ''; // CLEAR D for seat
        
        layoutType = form.vehicleLayout || 'Van';
        layoutConfig = form.seatLayoutData || '{}';
        
        floor = form.floor ? form.floor.trim() : ''; // Keep the Floor value (e.g., ATGI)
        amenities = ''; // CLEAR G
    }
    // --- ROOM/DESK Logic ---
    else {
        // Use Capacity (Col D)
        capacity = form.capacity || '';
        seatCount = ''; // CLEAR P
        vehicleModel = '';
        plateNumber = '';
        layoutType = '';
        layoutConfig = '';
        // Amenities and Floor are from form values
    }

    // 3. New Row Data (16 Columns)
    const newRow = [
        type,               // 1 (A)
        name,               // 2 (B)
        location,           // 3 (C)
        capacity,           // 4 (D) <-- Capacity (Room/Desk/Vehicle)
        category,           // 5 (E)
        floor,              // 6 (F)
        amenities,          // 7 (G)
        accessibleBU,       // 8 (H)
        requiresApproval,   // 9 (I)
        'Available',        // 10 (J)
        photoUrl,           // 11 (K)
        vehicleModel,       // 12 (L)
        plateNumber,        // 13 (M)
        layoutType,         // 14 (N)
        layoutConfig,       // 15 (O)
        seatCount           // 16 (P) <-- Seat Count
    ];
    
    sheet.appendRow(newRow);
    
    return { success: true, message: `New ${type} "${name}" added successfully.` };

  } catch (e) {
    return { success: false, message: 'System Error in addNewResource: ' + e.toString() };
  }
}
/**
 * UPDATED: getAllResources now reads 16 columns (Col P for Seat Count),
 * and configures capacity/amenities for the front-end based on the new structure.
 * @returns {Object} Success status and array of resource objects.
 */
function getAllResources() {
  try {
    const ss = SpreadsheetApp.openById(DB_SHEET_ID);
    const resourceSheet = ss.getSheetByName(RESOURCES_SHEET_NAME);
    const reservationSheet = ss.getSheetByName(RESERVATIONS_SHEET_NAME);
    
    const timezone = PHILIPPINES_TIMEZONE; 
    const now = new Date();
    const todayDateStr = Utilities.formatDate(now, timezone, 'yyyy-MM-dd');
    const nowMinutes = now.getHours() * 60 + now.getMinutes();

    if (!resourceSheet || resourceSheet.getLastRow() < 2) {
      return { success: true, resources: [] }; 
    }
    
    // 1. Get Resources Data (READ 16 COLUMNS: Col A up to Col P)
    const lastCol = resourceSheet.getLastColumn() < 16 ? 16 : resourceSheet.getLastColumn();
    const resourceValues = resourceSheet.getRange(2, 1, resourceSheet.getLastRow() - 1, lastCol).getValues(); 

    // 2. Get Reservations Data and Count Bookings per Resource
    const bookingCounts = {}; 
    const activeReservations = []; 

    if (reservationSheet && reservationSheet.getLastRow() > 1) {
        const reservationValues = reservationSheet.getRange(2, 1, reservationSheet.getLastRow() - 1, 10).getValues();
        
        reservationValues.forEach(row => {
            const status = String(row[0] || '').trim();
            const resourceType = String(row[1] || '').trim().toLowerCase();
            const resourceName = String(row[2] || '').trim();
            const approvalStatus = String(row[9] || '').trim();
            const resDate = row[4];
            
            // Filter: Only Active and Approved
            if (status !== 'Active' || approvalStatus !== 'Approved') return;
            
            // Check Date: Must be Today
            if (!resDate || !(resDate instanceof Date)) return;
            const resDateStr = Utilities.formatDate(resDate, timezone, 'yyyy-MM-dd');
            
            if (resDateStr === todayDateStr) {
                // A. Logic for Shuttle/Seat Count
                if (resourceType === 'seat' || resourceType === 'car seat') {
                    if (!bookingCounts[resourceName]) {
                        bookingCounts[resourceName] = 0;
                    }
                    bookingCounts[resourceName]++;
                }

                // B. Logic for Rooms/Desk (In Use Status based on Time)
                const startTime = row[5];
                const endTime = row[6];
                if (startTime instanceof Date && endTime instanceof Date) {
                    const startMinutes = startTime.getHours() * 60 + startTime.getMinutes();
                    const endMinutes = endTime.getHours() * 60 + endTime.getMinutes();
                    
                    if (nowMinutes >= startMinutes && nowMinutes < endMinutes) {
                        activeReservations.push({
                            resourceName: resourceName,
                            reservedByEmail: String(row[3] || '').trim(),
                            endTime: Utilities.formatDate(endTime, timezone, 'hh:mm a')
                        });
                    }
                }
            }
        });
    }

    // 3. Merge Counts with Resources
    const resourcesWithStatus = resourceValues.map(row => {
        const resourceName = String(row[1] || '').trim();
        let rawPhotoUrl = String(row[10] || '').trim(); // Col K (Index 10)
        
        // Photo URL fix (same as before)
        if (rawPhotoUrl.includes('drive.google.com')) {
            const idMatch = rawPhotoUrl.match(/id=([a-zA-Z0-9_-]+)/) || rawPhotoUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
            if (idMatch) rawPhotoUrl = "https://lh3.googleusercontent.com/d/" + idMatch[1];
        }
        
        const currentReservation = activeReservations.find(res => res.resourceName === resourceName);
        const bookedCount = bookingCounts[resourceName] || 0;

        // --- NEW DATA EXTRACTION (Use 16 columns for all configurations) ---
        const typeLower = String(row[0] || '').trim().toLowerCase(); // Col A (Index 0)
        
        let floorDisplay = String(row[5] || '').trim(); // Col F 
        let amenitiesConfig = String(row[6] || '').trim(); // Col G 
        let capacityDisplay = String(row[3] || '').trim(); // Col D (Capacity Pax)
        
        // Vehicle Type: Reconstruct the Floor/Amenities display
        if (typeLower === 'vehicle') {
            const model = String(row[11] || '').trim(); // Col L 
            const plate = String(row[12] || '').trim(); // Col M 
            floorDisplay = `Model: ${model} | Plate: ${plate}`;
            amenitiesConfig = ''; // Only Room amenities go in Col G
        }
        // Seat Type: Reconstruct the Amenities display
        else if (typeLower === 'seat' || typeLower === 'car seat') {
            const seatCount = String(row[15] || '0').trim(); // Col P (Index 15) <--- SEAT COUNT!
            const layoutType = String(row[13] || '').trim(); // Col N 
            const layoutConfig = String(row[14] || '{}').trim(); // Col O 
            
            // The capacity of a seat is the Seat Count itself (for front-end cards)
            capacityDisplay = seatCount; 
            
            // Re-combine the seat details into Amenities string (The seat selection logic relies on this)
            amenitiesConfig = `SeatCount: ${seatCount} | Layout: ${layoutType} | LayoutConfig: ${layoutConfig}`;
            floorDisplay = String(row[5] || '').trim(); // Keep the Floor/Area (e.g., ATGI)
        }
        // -----------------------------------------------------

        const resource = {
            type: String(row[0] || '').trim(),              // A
            name: resourceName,                             // B
            location: String(row[2] || '').trim(),          // C
            capacity: capacityDisplay,                      // D (Value: Pax or Seat Count)
            category: String(row[4] || '').trim(),          // E
            floor: floorDisplay,                            // F
            amenities: amenitiesConfig,                     // G 
            accessibleBU: String(row[7] || '').trim(),      // H
            requiresApproval: String(row[8] || '').trim(),  // I
            status: String(row[9] || 'Available').trim(),   // J
            photoUrl: rawPhotoUrl,                          // K
            bookedCount: bookedCount                        
        };
        
        if (currentReservation) {
            resource.status = 'In Use'; 
            resource.currentReservation = {
                reservedBy: currentReservation.reservedByEmail,
                endTime: currentReservation.endTime 
            };
        } else {
            resource.currentReservation = null; 
        }

        return resource;
    });

    return { success: true, resources: resourcesWithStatus };
  } catch (e) {
    Logger.log("Error in getAllResources: " + e.toString());
    return { success: false, message: 'Error: ' + e.toString(), resources: [] };
  }
}
/**
 * UPDATED: updateResource now uses 16 columns (Col P for Seat Count).
 * @param {Object} form - The resource update form data.
 * @returns {Object} Success status and message.
 */
function updateResource(form) {
  try {
    const ss = SpreadsheetApp.openById(DB_SHEET_ID);
    const sheet = ss.getSheetByName(RESOURCES_SHEET_NAME);
    
    if (!sheet || sheet.getLastRow() < 2) {
        return { success: false, message: 'Resources sheet is empty or not found.' };
    }
    
    // Read all 16 columns
    const data = sheet.getDataRange().getValues();
    
    const originalName = form.originalName;
    let rowToUpdate = -1;
    let originalPhotoUrl = '';
    let originalStatus = 'Available';
    
    // 1. Find the row
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const resourceName = String(row[1] || '').trim(); // Col B: Name
      const rowNumber = i + 1; 

      if (resourceName === originalName) {
        rowToUpdate = rowNumber;
        originalPhotoUrl = String(row[10] || '').trim(); // Col K: Photo URL
        originalStatus = String(row[9] || 'Available').trim(); // Col J: Status
        break; 
      }
    }
    
    if (rowToUpdate === -1) {
        return { success: false, message: `Resource "${originalName}" not found for update.` };
    }

    // 2. Photo Upload Logic 
    let photoUrl = originalPhotoUrl;
    let fileBlob = null;
    
    if (form.vehiclePhoto && form.vehiclePhoto.length > 0 && form.vehiclePhoto.getName() !== '') {
        fileBlob = form.vehiclePhoto;
    } else if (form.seatPhoto && form.seatPhoto.length > 0 && form.seatPhoto.getName() !== '') {
        fileBlob = form.seatPhoto;
    }

    if (fileBlob) {
      try {
        const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
        const file = folder.createFile(fileBlob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        photoUrl = "https://lh3.googleusercontent.com/d/" + file.getId(); 
      } catch (e) {
        Logger.log("Drive Upload Error: " + e.toString());
        if (photoUrl === originalPhotoUrl) {
             photoUrl = 'Upload Error';
        }
      }
    } else if (form.removePhoto === 'true') {
        photoUrl = ''; 
    }
    
    // 3. Data Preparation
    const type = form.type; 
    const name = form.name.trim();
    const location = form.location;
    const category = form.category;
    const accessibleBU = Array.isArray(form.accessibleBU) ? form.accessibleBU.join(', ') : (form.accessibleBU || '');
    const requiresApproval = form.requiresApproval === 'on' ? 'Yes' : 'No'; 
    const status = originalStatus;
    
    // --- Capacity (Col D) / Seat Count (Col P) Logic ---
    let capacity = form.capacity || ''; // Default to form capacity
    let seatCount = String(data[rowToUpdate - 1][15] || '').trim(); // Existing Seat Count (Col P, Index 15)
    
    // --- NEW DEDICATED COLUMN DATA (Defaults based on existing values) ---
    let vehicleModel = String(data[rowToUpdate - 1][11] || '').trim(); 
    let plateNumber = String(data[rowToUpdate - 1][12] || '').trim(); 
    let layoutType = String(data[rowToUpdate - 1][13] || '').trim(); 
    let layoutConfig = String(data[rowToUpdate - 1][14] || '{}').trim(); 
    
    // --- CLEANED OLD COLUMNS ---
    let floor = form.floor ? form.floor.trim() : ''; 
    let amenities = Array.isArray(form.amenities) ? form.amenities.join(', ') : (form.amenities || ''); 
    
    // Check for Name uniqueness (No change)
    if (name !== originalName) {
        for (let i = 1; i < data.length; i++) {
            if (i !== rowToUpdate - 1) { // -1 for array index
                const existingName = String(data[i][1]).trim();
                if (name === existingName) {
                    return { success: false, message: `Resource name "${name}" already exists.` };
                }
            }
        }
    }

    // --- VEHICLE LOGIC ---
    if (type === 'vehicle') {
        vehicleModel = form.model.trim();
        plateNumber = form.plateNumber.trim();
        capacity = form.capacity || ''; // Update Col D
        seatCount = ''; // CLEAR P
        floor = ''; // CLEAR F
        amenities = ''; // CLEAR G
    }
    // --- SEAT LOGIC (UPDATED: Use Col P) ---
    else if (type === 'seat' || type === 'car seat') {
        seatCount = form.seatCount || '0'; // New Seat Count from form (Update Col P)
        capacity = ''; // CLEAR D
        
        layoutType = form.vehicleLayout || 'Van';
        layoutConfig = form.seatLayoutData || '{}';
        
        floor = form.floor ? form.floor.trim() : ''; 
        amenities = ''; // CLEAR G
    }
    // --- ROOM/DESK Logic ---
    else {
        // Use Capacity (Col D)
        capacity = form.capacity || '';
        
        // Clear all dedicated columns (L, M, N, O, P)
        vehicleModel = '';
        plateNumber = '';
        layoutType = '';
        layoutConfig = '';
        seatCount = ''; // CLEAR P
    }

    // 4. Update the Row (16 Columns)
    const newRowValues = [
        type,               // 1 (A)
        name,               // 2 (B)
        location,           // 3 (C)
        capacity,           // 4 (D) <-- Capacity (Room/Desk/Vehicle)
        category,           // 5 (E)
        floor,              // 6 (F)
        amenities,          // 7 (G)
        accessibleBU,       // 8 (H)
        requiresApproval,   // 9 (I)
        status,             // 10 (J)
        photoUrl,           // 11 (K)
        vehicleModel,       // 12 (L)
        plateNumber,        // 13 (M)
        layoutType,         // 14 (N)
        layoutConfig,       // 15 (O)
        seatCount           // 16 (P) <-- Seat Count
    ];
    
    // Update range: Col 1 up to Col 16
    sheet.getRange(rowToUpdate, 1, 1, 16).setValues([newRowValues]); 
    
    return { success: true, message: `Resource "${name}" successfully updated.` };

  } catch (e) {
    Logger.log("Error in updateResource: " + e.toString());
    return { success: false, message: 'System Error: Failed to update resource: ' + e.toString() };
  }
}

/**
 * NEW: Deletes a resource from the Google Sheet.
 * @param {string} type - The resource type.
 * @param {string} name - The resource name.
 * @returns {Object} Success status and message.
 */
function removeResource(type, name) {
  try {
    const ss = SpreadsheetApp.openById(DB_SHEET_ID);
    const sheet = ss.getSheetByName(RESOURCES_SHEET_NAME);
    
    if (!sheet || sheet.getLastRow() < 2) {
        return { success: false, message: 'Resources sheet is empty or not found.' };
    }
    
    const data = sheet.getDataRange().getValues();
    let rowToDelete = -1;
    
    // Find the row index using Type (Col A) and Name (Col B)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const resourceType = String(row[0] || '').trim();
      const resourceName = String(row[1] || '').trim();
      const rowNumber = i + 1; 

      if (resourceType.toLowerCase() === type.toLowerCase() && resourceName === name) {
        rowToDelete = rowNumber;
        break; 
      }
    }

    if (rowToDelete !== -1) {
      sheet.deleteRow(rowToDelete);
      return { success: true, message: `Resource "${name}" removed successfully.` };
    } else {
      return { success: false, message: `Resource "${name}" not found under type "${type}".` };
    }

  } catch (e) {
    Logger.log("Error in removeResource: " + e.toString());
    return { success: false, message: 'System Error: Failed to remove resource.' };
  }
}

/**
 * HELPER: Convert Time String (HH:mm) or Date object to minutes from midnight.
 * @param {(string|Date)} timeInput - Time as "HH:mm" string or Date object.
 * @returns {number} Minutes from midnight or -1 if invalid.
 */
function convertTimeToMinutes(timeInput) {
    if (!timeInput) return -1;
    
    if (timeInput instanceof Date) {
        return timeInput.getHours() * 60 + timeInput.getMinutes();
    }
    
    // Assume string "HH:mm"
    if (typeof timeInput === 'string' && timeInput.includes(':')) {
        const parts = timeInput.split(':');
        return parseInt(parts[0]) * 60 + parseInt(parts[1]);
    }
    
    return -1;
}

/**
 * Adds a new reservation to the Reservations Sheet.
 * UPDATED:
 * 1. Past Dates/Time are not allowed.
 * 2. Future Dates beyond 1 MONTH are not allowed. (NEW)
 * 3. Uses LockService (Anti-Double Booking).
 * 4. Includes Conflict Checker (Time Overlap).
 * @param {Object} reservationData - The reservation details.
 * @returns {Object} Success status, resource name, approval status, and message.
 */
function addNewReservation(reservationData) {
  
  // Ensure PHILIPPINES_TIMEZONE is defined at the top of the file
  const timezone = (typeof PHILIPPINES_TIMEZONE !== 'undefined') ? PHILIPPINES_TIMEZONE : 'Asia/Manila';
  const now = new Date();

  // ----------------------------------------------------
  // 1. PAST DATE/TIME CHECK (Past times are not allowed)
  // ----------------------------------------------------
  // Create number format: YYYYMMDDHHmm (e.g., 202311261430) for easy comparison
  const currentTimestampStr = Utilities.formatDate(now, timezone, 'yyyyMMddHHmm');
  const currentTimestamp = parseFloat(currentTimestampStr);

  // Clean the requested date/time to also be a number
  // reservationData.date format: "yyyy-MM-dd" -> remove "-"
  // reservationData.startTime format: "HH:mm" -> remove ":"
  const reqDateClean = reservationData.date.replace(/-/g, ''); 
  const reqTimeClean = reservationData.startTime.replace(/:/g, '');
  const reqTimestamp = parseFloat(reqDateClean + reqTimeClean);

  // If Requested is less than Current, it means it's already passed.
  if (reqTimestamp < currentTimestamp) {
      return { success: false, message: 'Error: You cannot reserve a past date or time.' };
  }

  // ----------------------------------------------------
  // 2. (NEW) MAX 1 MONTH ADVANCE BOOKING CHECK
  // ----------------------------------------------------
  // Get today's date
  const todayDate = new Date();
  
  // Set a limit of 1 Month from now
  const maxAllowedDate = new Date(todayDate);
  maxAllowedDate.setMonth(maxAllowedDate.getMonth() + 1);
  
  // Convert requested date string to a Date object
  const requestedDateObj = new Date(reservationData.date);

  // Reset time to 00:00:00 for date-only comparison (avoids time conflicts)
  requestedDateObj.setHours(0,0,0,0);
  maxAllowedDate.setHours(23,59,59,999); // Set to end of the day of the max limit

  if (requestedDateObj > maxAllowedDate) {
      return { success: false, message: 'Error: Advance booking is limited to 1 month only.' };
  }
  // ----------------------------------------------------

  // ----------------------------------------------------
  // 3. ACQUIRE LOCK (To ensure sequential saving)
  // ----------------------------------------------------
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // Wait for a max of 10 seconds
  } catch (e) {
    return { success: false, message: 'System is busy. Please try again in a few seconds.' };
  }

  try {
    const ss = SpreadsheetApp.openById(DB_SHEET_ID);
    let sheet = ss.getSheetByName(RESERVATIONS_SHEET_NAME);
    
    if (!sheet) {
      sheet = ss.insertSheet(RESERVATIONS_SHEET_NAME);
      sheet.getRange(1, 1, 1, 10).setValues([[
        'Status', 'Resource Type', 'Resource Name', 'Requester Email', 
        'Date', 'Start Time', 'End Time', 'Participants', 'Notes', 'Approval Status' 
      ]]);
    }
    
    // ----------------------------------------------------
    // 4. CONFLICT CHECKING LOGIC (Avoid Double Booking)
    // ----------------------------------------------------
    const data = sheet.getDataRange().getValues();
    
    // Convert requested time to minutes for overlap check
    const newStartMin = convertTimeToMinutes(reservationData.startTime);
    const newEndMin = convertTimeToMinutes(reservationData.endTime);
    const newDateStr = reservationData.date; 
    
    const targetSeatNote = reservationData.seatNumber ? `Seat #${reservationData.seatNumber}` : null;
    const isSeatBooking = (reservationData.resourceType.toLowerCase() === 'seat' || reservationData.resourceType.toLowerCase() === 'car seat');

    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const status = String(row[0]).trim(); // Col A: Status
        const rName = String(row[2]).trim(); // Col C: Resource Name
        const approvalStatus = String(row[9]).trim(); // Col J: Approval Status

        // Filter: Only check Active and same Resource
        if (status !== 'Active') continue;
        if (approvalStatus === 'Rejected') continue; // Ignore rejected
        
        // IMPORTANT: Conflict check is only for actual resources (Room, Desk, Seat).
        // Do not check for conflict for generic "Vehicle Request"
        if (rName !== reservationData.resourceName && reservationData.resourceType.toLowerCase() !== 'vehicle') continue;
        
        // This will only be 'true' if the resourceName is found in the Resource Sheet.
        const isActualResource = (reservationData.resourceType.toLowerCase() !== 'vehicle' || rName !== "Vehicle Request");
        
        // If the booking is for an ACTUAL RESOURCE (not generic request), proceed with conflict check
        if (isActualResource && rName === reservationData.resourceName) {
            
            // Check Date (Must be the same date)
            const rowDate = new Date(row[4]);
            const rowDateStr = Utilities.formatDate(rowDate, timezone, 'yyyy-MM-dd');
            
            if (rowDateStr !== newDateStr) continue;

            // Check Time Overlap Formula: (StartA < EndB) and (EndA > StartB)
            const rowStartMin = convertTimeToMinutes(row[5]);
            const rowEndMin = convertTimeToMinutes(row[6]);

            if (newStartMin < rowEndMin && newEndMin > rowStartMin) {
                // IF THERE IS TIME OVERLAP:
                
                if (isSeatBooking) {
                    // SPECIAL LOGIC FOR SEATS: Conflict only if SAME SEAT NUMBER
                    const rowNotes = String(row[8] || '');
                    if (targetSeatNote && rowNotes.includes(targetSeatNote)) {
                        return { success: false, message: `Conflict: ${targetSeatNote} is already booked for this time.` };
                    }
                } else {
                    // ROOM/VEHICLE/DESK: Immediate conflict if there is an overlap
                    return { success: false, message: 'Conflict: This resource is already reserved for the selected time.' };
                }
            }
        }
    }

    // ----------------------------------------------------
    // 5. PROCEED TO SAVE (No Conflict)
    // ----------------------------------------------------
    
    // Check Approval Requirement from Resources Sheet
    const resourceSheet = ss.getSheetByName(RESOURCES_SHEET_NAME);
    const resourceData = resourceSheet.getDataRange().getValues();
    let requiresApproval = 'No';
    
    for (let i = 1; i < resourceData.length; i++) {
        if (String(resourceData[i][1]).trim() === reservationData.resourceName) {
            requiresApproval = String(resourceData[i][8] || 'No').trim();
            break;
        }
    }
    
    // ******************************************************************************
    // >>> CRITICAL ADDITION: FORCE APPROVAL FOR VEHICLE REQUEST <<<
    // ******************************************************************************
    if (reservationData.resourceType.toLowerCase() === 'vehicle') {
        // Since this is a Request Form and not a Resource Card, we need to force it to Pending.
        requiresApproval = 'Yes';
    }
    // ******************************************************************************
    
    const approvalStatus = (requiresApproval === 'Yes') ? 'Pending' : 'Approved';
    const reservationStatus = 'Active';

    // Process Notes (Add Seat # if applicable)
    let finalNotes = reservationData.notes || '';
    if (isSeatBooking && reservationData.seatNumber) {
         finalNotes = `Seat #${reservationData.seatNumber}${finalNotes ? ` | ${finalNotes}` : ''}`;
    }
    
    // Append Row
    const newRow = [
      reservationStatus, reservationData.resourceType, reservationData.resourceName, 
      reservationData.requesterEmail, reservationData.date, reservationData.startTime, 
      reservationData.endTime, reservationData.participants || '', finalNotes, approvalStatus
    ];
    
    sheet.appendRow(newRow);
    
    // IMPORTANT: Flush immediately so other users can read the update
    SpreadsheetApp.flush(); 
    
    // --- EMAIL & CALENDAR LOGIC ---
    const scriptUrl = ScriptApp.getService().getUrl();
    // Format the time for email display
    let emailTimeStr = `${reservationData.startTime} - ${reservationData.endTime}`;
    // Convert 24h to 12h AM/PM format for nicer email (Optional)
    try {
        const d1 = new Date(`2000-01-01T${reservationData.startTime}`);
        const d2 = new Date(`2000-01-01T${reservationData.endTime}`);
        const t1 = Utilities.formatDate(d1, timezone, 'hh:mm a');
        const t2 = Utilities.formatDate(d2, timezone, 'hh:mm a');
        emailTimeStr = `${t1} - ${t2}`;
    } catch(e) {} // Fallback to 24h if error

    const emailDetails = {
        "Resource": reservationData.resourceName,
        "Type": reservationData.resourceType.charAt(0).toUpperCase() + reservationData.resourceType.slice(1),
        "Date": reservationData.date,
        "Time": emailTimeStr,
        "Requester": reservationData.requesterEmail
    };
    if(reservationData.seatNumber) emailDetails["Seat Number"] = reservationData.seatNumber;
    
    // If generic Vehicle Request, add details to the email
    if (reservationData.resourceType.toLowerCase() === 'vehicle') {
        emailDetails["Passengers"] = reservationData.participants || '1 (Requester Only)';
        emailDetails["Details"] = finalNotes;
    }


    // A. Notify Admins & Approvers
    const adminEmails = getAdminAndApproverEmails();
    if (adminEmails && adminEmails.length > 0) {
        let adminTitle = "", adminMsg = "", adminColor = "";
        if (approvalStatus === 'Pending') {
            adminTitle = "ACTION REQUIRED: New Reservation Request";
            adminMsg = `A new reservation request has been submitted and is currently <strong>PENDING</strong> approval.`;
            adminColor = "#fbc02d"; 
        } else {
            adminTitle = "FYI: New Reservation (Auto-Approved)";
            adminMsg = `A new reservation has been successfully created and <strong>AUTO-APPROVED</strong>.`;
            adminColor = "#4CAF50"; 
        }
        const htmlBody = createEmailTemplate(adminTitle, adminMsg, emailDetails, adminColor, scriptUrl, "Go to Admin Dashboard");
        sendEmailNotification(adminEmails, `New Request: ${reservationData.resourceName}`, htmlBody);
    }

    // B. Notify Requester
    if (approvalStatus === 'Approved') {
        const userMsg = "Great news! Your reservation request has been automatically <strong>CONFIRMED</strong>.";
        const htmlBody = createEmailTemplate("Reservation Confirmed", userMsg, emailDetails, "#4CAF50", scriptUrl, "View My Reservations");
        sendEmailNotification(reservationData.requesterEmail, `Confirmed: ${reservationData.resourceName}`, htmlBody);

        let allGuests = reservationData.requesterEmail;
        // NOTE: We don't invite guests for Room/Desk/Car if we don't know their email
        // For Vehicle Request, participants is only a number, not an email
        if (reservationData.participants && reservationData.participants.trim() !== "" && reservationData.resourceType.toLowerCase() !== 'vehicle') {
            allGuests += "," + reservationData.participants;
        }
        
        // Set the Description
        let calendarDescription = `Resource: ${reservationData.resourceName}\nType: ${reservationData.resourceType}`;
        if(reservationData.resourceType.toLowerCase() === 'vehicle') {
            calendarDescription += `\nPassengers: ${reservationData.participants}\nDetails: ${finalNotes}`;
        } else {
            calendarDescription += `\nParticipants: ${reservationData.participants}\nNotes: ${finalNotes || 'None'}`;
        }


        addToGoogleCalendar(
            `A.Space Booking: ${reservationData.resourceName}`,
            reservationData.date, reservationData.startTime, reservationData.endTime,
            allGuests, 
            calendarDescription
        );
    }
    
    return { 
      success: true, 
      resourceName: reservationData.resourceName, 
      approvalStatus: approvalStatus, 
      message: `Reservation for "${reservationData.resourceName}" successfully created.` 
    };

  } catch (e) {
    Logger.log("Error in addNewReservation: " + e.toString());
    return { success: false, message: 'System Error: ' + e.toString() };
  } finally {
    // 6. Release Lock (ALWAYS do this whether there is an error or success)
    lock.releaseLock();
  }
}
// RESERVATION LIST & CANCEL
// ***************************************************************

/**
 * Retrieves all reservations for a single user.
 * @param {string} userEmail - The email of the requester.
 * @returns {Object} Success status and array of reservation objects.
 */
function getUsersReservations(userEmail) {
  try {
    const ss = SpreadsheetApp.openById(DB_SHEET_ID);
    const sheet = ss.getSheetByName(RESERVATIONS_SHEET_NAME);

    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, reservations: [] };
    }

    const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10);
    const values = range.getValues();
    
    const REQUESTER_EMAIL_COL_INDEX = 3; 
    const timezone = PHILIPPINES_TIMEZONE; 

    const userReservations = [];

    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const requester = String(row[REQUESTER_EMAIL_COL_INDEX] || '').trim();
      
      if (requester.toLowerCase() === userEmail.toLowerCase()) {
        
        // FIX: Format the Time values to a simple 'HH:mm' string
        const startTimeStr = row[5] ? Utilities.formatDate(row[5], timezone, 'HH:mm') : '';
        const endTimeStr = row[6] ? Utilities.formatDate(row[6], timezone, 'HH:mm') : '';
        // FIX: Format the Date object to 'YYYY-MM-DD' string.
        const dateStr = row[4] ? Utilities.formatDate(new Date(row[4]), timezone, 'yyyy-MM-dd') : '';


        userReservations.push({
          rowNumber: i + 2, 
          status: String(row[0] || '').trim(), 
          resourceType: String(row[1] || '').trim(), 
          resourceName: String(row[2] || '').trim(), 
          requesterEmail: requester,
          date: dateStr, 
          startTime: startTimeStr,     
          endTime: endTimeStr,         
          participants: String(row[7] || '').trim(), 
          notes: String(row[8] || '').trim(), 
          approvalStatus: String(row[9] || '').trim() 
        });
      }
    }

    return { success: true, reservations: userReservations };

  } catch (e) {
    Logger.log("Error in getUsersReservations: " + e.toString());
    return { success: false, message: 'Error retrieving reservations: ' + e.toString() };
  }
}

/**
 * Cancels a reservation in the Reservations Sheet.
 * @param {number} rowNumber - The 1-based sheet row number of the reservation.
 * @returns {Object} Success status and message.
 */
function cancelReservation(rowNumber) {
  try {
    const ss = SpreadsheetApp.openById(DB_SHEET_ID);
    const sheet = ss.getSheetByName(RESERVATIONS_SHEET_NAME);
    
    const rowNum = parseInt(rowNumber);

    if (isNaN(rowNum) || rowNum < 2 || rowNum > sheet.getLastRow()) {
        return { success: false, message: `Invalid row number: ${rowNumber}.` };
    }

    // Col A (Status) and Col J (Approval Status) will be updated.
    sheet.getRange(rowNum, 1).setValue('Cancelled'); // Column A: Status
    sheet.getRange(rowNum, 10).setValue('Cancelled'); // Column J: Approval Status
    
    const resourceName = sheet.getRange(rowNum, 3).getValue(); // Col C: Resource Name

    return { success: true, message: `Reservation for ${resourceName} successfully cancelled.` };

  } catch (e) {
    Logger.log("Error in cancelReservation: " + e.toString());
    return { success: false, message: 'System Error: Failed to cancel reservation. ' + e.toString() };
  }
}
// ***************************************************************
// APPROVALS FUNCTIONS
// ***************************************************************

/**
 * Retrieves all reservations with a "Pending" Approval Status.
 * @returns {Object} Success status and array of pending approval objects.
 */
function getPendingApprovals() {
  try {
    const ss = SpreadsheetApp.openById(DB_SHEET_ID);
    const sheet = ss.getSheetByName(RESERVATIONS_SHEET_NAME);
    const timezone = PHILIPPINES_TIMEZONE; 

    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, approvals: [] };
    }

    const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10);
    const values = range.getValues();
    
    const APPROVAL_STATUS_COL_INDEX = 9; 
    const pendingApprovals = [];

    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const approvalStatus = String(row[APPROVAL_STATUS_COL_INDEX] || '').trim();
      
      if (approvalStatus.toLowerCase() === 'pending') {
        
        // FIX: Format the Time values
        const startTimeStr = row[5] ? Utilities.formatDate(row[5], timezone, 'HH:mm') : '';
        const endTimeStr = row[6] ? Utilities.formatDate(row[6], timezone, 'HH:mm') : '';
        const dateStr = row[4] ? Utilities.formatDate(new Date(row[4]), timezone, 'yyyy-MM-dd') : '';


        pendingApprovals.push({
          rowNumber: i + 2, 
          resourceName: String(row[2] || '').trim(), 
          resourceType: String(row[1] || '').trim(), 
          requesterEmail: String(row[3] || '').trim(), 
          date: dateStr,
          startTime: startTimeStr, 
          endTime: endTimeStr,     
          participants: String(row[7] || '').trim(), 
          notes: String(row[8] || '').trim(), 
          approvalStatus: approvalStatus
        });
      }
    }
    
    // Get the Requester Name from the Users Sheet
    const userSheet = ss.getSheetByName(DB_SHEET_NAME);
    const userData = (userSheet && userSheet.getLastRow() > 1) ? userSheet.getDataRange().getValues() : [];
    
    pendingApprovals.forEach(approval => {
      // Find the row with the matching email (Col A)
      const userRow = userData.find(r => String(r[0]).trim() === approval.requesterEmail);
      if (userRow) {
          // Name (Col D: index 3) and Business Unit (Col E: index 4)
          approval.requesterName = String(userRow[3] || approval.requesterEmail).trim(); 
          approval.requesterBU = String(userRow[4] || '').trim();
      } else {
          approval.requesterName = approval.requesterEmail.split('@')[0];
          approval.requesterBU = 'N/A';
      }
    });

    return { success: true, approvals: pendingApprovals };

  } catch (e) {
    Logger.log("Error in getPendingApprovals: " + e.toString());
    return { success: false, message: 'Error retrieving pending approvals: ' + e.toString() };
  }
}

/**
 * Updates the Approval Status, sends an email, and adds to Google Calendar.
 * @param {number} rowNumber - The 1-based sheet row number of the reservation.
 * @param {string} newStatus - The new approval status ('Approved' or 'Rejected').
 * @returns {Object} Success status and message.
 */
function updateReservationApprovalStatus(rowNumber, newStatus) {
  try {
    const ss = SpreadsheetApp.openById(DB_SHEET_ID);
    const sheet = ss.getSheetByName(RESERVATIONS_SHEET_NAME);
    
    const rowNum = parseInt(rowNumber);
    const status = newStatus.trim(); 

    if (isNaN(rowNum) || rowNum < 2 || rowNum > sheet.getLastRow()) {
        return { success: false, message: `Invalid parameters.` };
    }

    // Update Status in Sheet
    sheet.getRange(rowNum, 10).setValue(status); 
    if (status === 'Rejected') {
        sheet.getRange(rowNum, 1).setValue('Cancelled'); 
    } else if (status === 'Approved') {
        sheet.getRange(rowNum, 1).setValue('Active'); 
    }

    // Fetch Data
    const resourceName = sheet.getRange(rowNum, 3).getValue();
    const requesterEmail = sheet.getRange(rowNum, 4).getValue();
    const resDateRaw = sheet.getRange(rowNum, 5).getValue();
    const startTime = sheet.getRange(rowNum, 6).getValue();
    const endTime = sheet.getRange(rowNum, 7).getValue();
    
    // Format Date & Time for Email Display
    const dateStr = Utilities.formatDate(new Date(resDateRaw), PHILIPPINES_TIMEZONE, 'MMMM dd, yyyy');
    // Note: If startTime/endTime are Date objects from the sheet, we format them for display
    let timeStr = "";
    if (startTime instanceof Date && endTime instanceof Date) {
        timeStr = `${Utilities.formatDate(startTime, PHILIPPINES_TIMEZONE, 'hh:mm a')} - ${Utilities.formatDate(endTime, PHILIPPINES_TIMEZONE, 'hh:mm a')}`;
    } else {
        timeStr = `${startTime} - ${endTime}`;
    }

    // --- EMAIL LOGIC ---
    
    const scriptUrl = ScriptApp.getService().getUrl();
    
    const emailDetails = {
        "Resource": resourceName,
        "Date": dateStr,
        "Time": timeStr,
        "Status": status.toUpperCase()
    };

    let emailSubject = "";
    let emailTitle = "";
    let emailMsg = "";
    let statusColor = "";

    if (status === 'Approved') {
        emailSubject = `Reservation Approved - ${resourceName}`;
        emailTitle = "Request Approved";
        emailMsg = "Your reservation request has been <strong>APPROVED</strong> by the administrator. You may now proceed with your schedule.";
        statusColor = "#4CAF50"; // Green

        // ---------------------------------------------------------
        // NEW: ADD TO GOOGLE CALENDAR (Only if Approved)
        // ---------------------------------------------------------
        addToGoogleCalendar(
            `A.Space Booking: ${resourceName}`,
            resDateRaw,   // Raw Date object from Sheet
            startTime,    // Raw Time object from Sheet
            endTime,      // Raw Time object from Sheet
            requesterEmail,
            `Resource: ${resourceName}\nStatus: Approved by Admin`
        );
        // ---------------------------------------------------------

    } else { // Rejected
        emailSubject = `Reservation Rejected - ${resourceName}`;
        emailTitle = "Request Rejected";
        emailMsg = "We regret to inform you that your reservation request has been <strong>REJECTED</strong> due to availability or policy reasons.";
        statusColor = "#e53935"; // Red
    }

    // Generate HTML
    const htmlBody = createEmailTemplate(
        emailTitle, 
        emailMsg, 
        emailDetails, 
        statusColor, 
        scriptUrl, 
        "Go to A.Space"
    );

    // Send Email
    if (requesterEmail) {
        sendEmailNotification(requesterEmail, emailSubject, htmlBody);
    }
    
    return { success: true, message: `Reservation marked as ${status}. Notification and Calendar Invite sent.` };

  } catch (e) {
    Logger.log("Error in updateReservationApprovalStatus: " + e.toString());
    return { success: false, message: 'System Error: ' + e.toString() };
  }
}
/**
 * Retrieves the list of booked seats AND the user who booked them.
 * @param {string} resourceName - The name of the shuttle/vehicle.
 * @param {string} dateString - The date of the booking (YYYY-MM-DD).
 * @returns {Object} Success status and object of booked seats ({seatNum: userName}).
 */
function getShuttleSeatStatus(resourceName, dateString) {
  try {
    const ss = SpreadsheetApp.openById(DB_SHEET_ID);
    const resSheet = ss.getSheetByName(RESERVATIONS_SHEET_NAME);
    const userSheet = ss.getSheetByName(DB_SHEET_NAME);
    
    if (!resSheet || resSheet.getLastRow() < 2) {
      return { success: true, bookedSeats: {} }; // Return empty object
    }
    
    // 1. Create a Map for User Names (Email -> Name)
    const userMap = {};
    if (userSheet && userSheet.getLastRow() > 1) {
        const userValues = userSheet.getDataRange().getValues();
        // Start loop at 1 to skip header
        for (let i = 1; i < userValues.length; i++) {
            const email = String(userValues[i][0]).trim();
            const name = String(userValues[i][3]).trim(); // Col D is Name
            if (email) userMap[email] = name || email.split('@')[0];
        }
    }
    
    const values = resSheet.getRange(2, 1, resSheet.getLastRow() - 1, 10).getValues();
    const bookedSeats = {}; // We made it an Object (Dictionary) instead of an Array
    
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const resName = String(row[2] || '').trim();
      const resDate = row[4];
      const status = String(row[0] || '').trim();
      const approvalStatus = String(row[9] || '').trim();
      const requesterEmail = String(row[3] || '').trim();
      const notes = String(row[8] || '').trim();
      
      if (resName !== resourceName || status !== 'Active' || approvalStatus !== 'Approved') {
          continue;
      }
      
      const reservationDateStr = Utilities.formatDate(resDate, PHILIPPINES_TIMEZONE, 'yyyy-MM-dd');
      if (reservationDateStr === dateString) {
          const match = notes.match(/Seat #(\d+)/);
          if (match && match[1]) {
              const seatNum = parseInt(match[1]);
              // Get the name from the map, or use the email if not found
              const bookedByName = userMap[requesterEmail] || requesterEmail.split('@')[0];
              
              // Save: Seat Number -> Name
              bookedSeats[seatNum] = bookedByName;
          }
      }
    }

    return { success: true, bookedSeats: bookedSeats };

  } catch (e) {
    Logger.log("Error in getShuttleSeatStatus: " + e.toString());
    return { success: false, message: 'Server Error retrieving seat status.' };
  }
}
/**
 * NEW: Retrieves all reservations for Admin Reservations Management.
 * @returns {Object} Success status and array of all reservation objects.
 */
function getAllReservationsForAdmin() {
  try {
    const ss = SpreadsheetApp.openById(DB_SHEET_ID);
    const resSheet = ss.getSheetByName(RESERVATIONS_SHEET_NAME);
    const userSheet = ss.getSheetByName(DB_SHEET_NAME);
    const timezone = PHILIPPINES_TIMEZONE; 

    if (!resSheet || resSheet.getLastRow() < 2) {
      return { success: true, reservations: [] };
    }

    // Get all data from the Reservations Sheet (10 Columns)
    const range = resSheet.getRange(2, 1, resSheet.getLastRow() - 1, 10);
    const values = range.getValues();
    
    // 1. Create User Map (Email -> Name)
    const userMap = {};
    if (userSheet && userSheet.getLastRow() > 1) {
        // Get all users (Col A: Email, Col D: Name)
        const userValues = userSheet.getDataRange().getValues().slice(1);
        userValues.forEach(row => { 
            // Ensure row[0] and row[3] are strings before trimming
            userMap[String(row[0] || '').trim()] = String(row[3] || '').trim(); 
        });
    }

    const allReservations = [];

    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const requester = String(row[3] || '').trim(); // Col D: Requester Email
      
      // Format the Time values for display (e.g., 9:00 AM)
      const startTimeStr = row[5] ? Utilities.formatDate(row[5], timezone, 'hh:mm a') : '';
      const endTimeStr = row[6] ? Utilities.formatDate(row[6], timezone, 'hh:mm a') : '';
      // Format the Date (e.g., 11/21/2025)
      const dateStr = row[4] ? Utilities.formatDate(new Date(row[4]), timezone, 'MM/dd/yyyy') : '';

      // Get the Requester Name from the map, use the email prefix if not found
      const userName = userMap[requester] || requester.split('@')[0];
      const type = String(row[1] || '').trim();
      
      allReservations.push({
        rowNumber: i + 2, // Sheet row number (this will be used for Cancel)
        resourceName: String(row[2] || '').trim(), // Col C
        resourceType: type, // Col B
        requesterName: userName,
        date: dateStr, 
        startTime: startTimeStr,     
        endTime: endTimeStr,         
        status: String(row[0] || '').trim(), // Col A (Active/Cancelled)
        approvalStatus: String(row[9] || '').trim() // Col J (Approved/Pending/Rejected)
      });
    }

    return { success: true, reservations: allReservations };

  } catch (e) {
    Logger.log("Error in getAllReservationsForAdmin: " + e.toString());
    return { success: false, message: 'Error retrieving all reservations: ' + e.toString() };
  }
}
/**
 * NEW: Retrieves all Floor Plan configurations.
 * @returns {Object} Success status and array of floor plan objects.
 */
function getAllFloorPlans() {
  try {
    const ss = SpreadsheetApp.openById(DB_SHEET_ID);
    const fpSheet = ss.getSheetByName(FLOOR_PLANS_SHEET_NAME);
    
    if (!fpSheet || fpSheet.getLastRow() < 2) {
      return { success: true, plans: [] }; 
    }
    
    // Get all data (Col A: Type, Col B: Category Name, Col C: Image URL, Col D: Layout Data)
    const values = fpSheet.getRange(2, 1, fpSheet.getLastRow() - 1, 4).getValues();

    const plans = values.map(row => ({
      type: String(row[0] || '').trim().toLowerCase(),
      categoryName: String(row[1] || '').trim(),
      imageUrl: String(row[2] || '').trim(),
      layoutData: String(row[3] || '{}').trim()
    }));

    return { success: true, plans: plans };
  } catch (e) {
    Logger.log("Error in getAllFloorPlans: " + e.toString());
    return { success: false, message: 'Error retrieving floor plans: ' + e.toString() };
  }
}

/**
 * NEW: Uploads a Floor Plan image and saves the URL to the FloorPlans sheet.
 * This will also be used to update an existing entry.
 * @param {Object} form - The floor plan upload form data.
 * @returns {Object} Success status and message.
 */
function uploadFloorPlanImage(form) {
    try {
        const ss = SpreadsheetApp.openById(DB_SHEET_ID);
        let fpSheet = ss.getSheetByName(FLOOR_PLANS_SHEET_NAME);
        
        // Ensure the sheet exists
        if (!fpSheet) {
            fpSheet = ss.insertSheet(FLOOR_PLANS_SHEET_NAME);
            fpSheet.getRange(1, 1, 1, 4).setValues([['Type', 'Category Name', 'Image URL', 'Layout Data']]);
        }
        
        const type = form.type.toLowerCase();
        const categoryName = form.categoryName.trim();
        const fileBlob = form.floorPlanPhoto;
        
        if (!fileBlob || typeof fileBlob.getDataAsString !== 'function' || fileBlob.getName() === '') {
            return { success: false, message: 'No file uploaded.' };
        }
        
        // 1. Upload Photo to Drive
        let photoUrl = '';
        try {
            const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
            const file = folder.createFile(fileBlob);
            file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
            photoUrl = "https://lh3.googleusercontent.com/d/" + file.getId(); 
        } catch (e) {
            Logger.log("Drive Upload Error: " + e.toString());
            return { success: false, message: 'Drive Upload Error: ' + e.toString() };
        }

        // 2. Find the existing row (Type + Category Name)
        const data = fpSheet.getDataRange().getValues();
        let rowNumber = -1;
        let layoutData = '{}';

        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            if (String(row[0]).trim().toLowerCase() === type && String(row[1]).trim() === categoryName) {
                rowNumber = i + 1;
                layoutData = String(row[3] || '{}').trim(); // Get the existing layout data
                break;
            }
        }
        
        // 3. Update or Append
        if (rowNumber !== -1) {
            // Update: Col C (Image URL)
            fpSheet.getRange(rowNumber, 3).setValue(photoUrl);
            return { success: true, message: `Floor plan for ${categoryName} updated successfully.` };
        } else {
            // Append: Create a new row
            const newRow = [type, categoryName, photoUrl, layoutData];
            fpSheet.appendRow(newRow);
            return { success: true, message: `New floor plan for ${categoryName} uploaded successfully.` };
        }

    } catch (e) {
        Logger.log("Error in uploadFloorPlanImage: " + e.toString());
        return { success: false, message: 'System Error: Failed to process floor plan upload. ' + e.toString() };
    }
}
/**
 * Saves the layout coordinates for a specific floor plan category.
 * @param {string} type - The resource type.
 * @param {string} categoryName - The category name associated with the floor plan.
 * @param {string} layoutDataString - The JSON string of the layout coordinates.
 * @returns {Object} Success status and message.
 */
function saveFloorPlanLayout(type, categoryName, layoutDataString) {
  try {
    const ss = SpreadsheetApp.openById(DB_SHEET_ID);
    const sheet = ss.getSheetByName(FLOOR_PLANS_SHEET_NAME); // 'FloorPlans'
    
    if (!sheet) return { success: false, message: 'FloorPlans sheet not found.' };
    
    const data = sheet.getDataRange().getValues();
    let rowToUpdate = -1;

    // Find the row with the correct Type and CategoryName
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (String(row[0]).trim().toLowerCase() === type.toLowerCase() && 
          String(row[1]).trim() === categoryName) {
        rowToUpdate = i + 1;
        break;
      }
    }

    if (rowToUpdate !== -1) {
      // Update Column D (Index 4) - Layout Data
      sheet.getRange(rowToUpdate, 4).setValue(layoutDataString);
      return { success: true, message: 'Layout saved successfully.' };
    } else {
      return { success: false, message: 'Floor plan category not found.' };
    }

  } catch (e) {
    Logger.log("Error in saveFloorPlanLayout: " + e.toString());
    return { success: false, message: 'System Error: ' + e.toString() };
  }
}
// ***************************************************************
// EMAIL NOTIFICATION HELPERS
// ***************************************************************

/**
 * Helper: Sends an email using the MailApp service.
 * @param {string} recipient - The email address(es) to send to.
 * @param {string} subject - The subject of the email.
 * @param {string} htmlBody - The HTML content of the email.
 */
function sendEmailNotification(recipient, subject, htmlBody) {
  try {
    if (!recipient || recipient.trim() === '') return;
    
    MailApp.sendEmail({
      to: recipient,
      subject: "A.Space Notification: " + subject,
      htmlBody: htmlBody,
      name: "A.Space Booking System"
    });
    Logger.log("Email sent to: " + recipient);
  } catch (e) {
    Logger.log("Failed to send email: " + e.toString());
  }
}

/**
 * Helper: Retrieves all ADMIN and APPROVER emails from the Users sheet.
 * @returns {string} Comma-separated string of admin/approver emails.
 */
function getAdminAndApproverEmails() {
  try {
    const ss = SpreadsheetApp.openById(DB_SHEET_ID);
    const sheet = ss.getSheetByName(DB_SHEET_NAME); // 'Users'
    
    if (!sheet) return "";
    
    const data = sheet.getDataRange().getValues();
    let emails = [];
    
    // Loop starting from row 1 (skip header)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const email = String(row[0] || '').trim(); // Col A: Email
      const role = String(row[2] || '').toLowerCase(); // Column C: Role
      
      // Check if Admin or Approver
      if (email && (role.includes('admin') || role.includes('approver'))) {
        emails.push(email);
      }
    }
    
    // Remove duplicates and join using a comma for the "To" field
    const result = [...new Set(emails)].join(',');
    
    Logger.log("Admin Emails found: " + result); 
    return result;

  } catch (e) {
    Logger.log("Error getting admin emails: " + e.toString());
    return ""; // Return empty string if error
  }
}

/**
 * A testing function to check admin email retrieval and sending.
 */
function testAdminEmail() {
  // 1. Try to retrieve the emails
  const emails = getAdminAndApproverEmails();
  Logger.log("Emails Found: " + emails);
  
  if (emails === "") {
    Logger.log("❌ ERROR: No admin email found. Check sheet name 'Users' or column positions.");
  } else {
    Logger.log("✅ SUCCESS: Emails found: " + emails);
    
    // 2. Try to send a test email
    try {
      MailApp.sendEmail({
        to: emails,
        subject: "Test Email from A.Space Script",
        htmlBody: "<h3>If you received this, email sending is WORKING!</h3>"
      });
      Logger.log("✅ Email sent command executed.");
    } catch (e) {
      Logger.log("❌ Email Failed: " + e.toString());
    }
  }
}
/**
 * GENERATES A PROFESSIONAL HTML EMAIL TEMPLATE
 * @param {string} title - The main title of the email.
 * @param {string} message - The main body message.
 * @param {Object} detailsObj - Key-value pair object for the details table.
 * @param {string} statusColor - The accent color for the title.
 * @param {string} actionLink - The URL for the call-to-action button.
 * @param {string} actionText - The text for the call-to-action button.
 * @returns {string} The complete HTML email body.
 */
function createEmailTemplate(title, message, detailsObj, statusColor, actionLink, actionText) {
  // A.Space Theme Colors
  const primaryRed = "#e53935";
  const bgGray = "#f4f6f8";
  const cardWhite = "#ffffff";
  const textDark = "#333333";
  const textGray = "#666666";

  // Build Details Table Rows
  let detailsHtml = "";
  for (const [key, value] of Object.entries(detailsObj)) {
    detailsHtml += `
      <tr>
        <td style="padding: 8px 0; color: ${textGray}; font-size: 14px; font-weight: bold; width: 35%;">${key}:</td>
        <td style="padding: 8px 0; color: ${textDark}; font-size: 14px;">${value}</td>
      </tr>
      <tr><td colspan="2" style="border-bottom: 1px solid #eeeeee;"></td></tr>
    `;
  }

  // Button HTML (Only if link provided)
  let buttonHtml = "";
  if (actionLink && actionText) {
    buttonHtml = `
      <div style="text-align: center; margin-top: 30px;">
        <a href="${actionLink}" style="background-color: ${primaryRed}; color: #ffffff; padding: 12px 24px; text-decoration: none; border-radius: 5px; font-weight: bold; font-size: 14px; display: inline-block;">${actionText}</a>
      </div>
    `;
  }

  // Full HTML Structure
  return `
    <div style="font-family: Arial, sans-serif; background-color: ${bgGray}; padding: 40px 0; margin: 0;">
      <div style="max-width: 600px; margin: 0 auto; background-color: ${cardWhite}; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 10px rgba(0,0,0,0.05);">
        
        <!-- Header with Logo -->
        <div style="background-color: ${primaryRed}; padding: 20px; text-align: center;">
          <h1 style="color: #ffffff; margin: 0; font-size: 24px; letter-spacing: 1px;">A.Space</h1>
        </div>

        <!-- Body Content -->
        <div style="padding: 40px;">
          <h2 style="color: ${statusColor}; margin-top: 0; font-size: 20px; text-align: center;">${title}</h2>
          <p style="color: ${textGray}; font-size: 15px; line-height: 1.6; text-align: center; margin-bottom: 30px;">
            ${message}
          </p>

          <!-- Details Box -->
          <div style="background-color: #fafafa; padding: 20px; border-radius: 6px; border: 1px solid #eeeeee;">
            <table style="width: 100%; border-collapse: collapse;">
              ${detailsHtml}
            </table>
          </div>

          ${buttonHtml}
        </div>

        <!-- Footer -->
        <div style="background-color: #333333; padding: 15px; text-align: center;">
          <p style="color: #999999; font-size: 12px; margin: 0;">
            © 2025 A.Space Booking System. All rights reserved.<br>
            This is an automated message. Please do not reply directly.
          </p>
        </div>

      </div>
    </div>
  `;
}
/**
 * HELPER: Adds event to System Calendar and invites the user.
 * @param {string} title - Event title.
 * @param {(string|Date)} dateStr - Event date.
 * @param {(string|Date)} startTimeStr - Event start time.
 * @param {(string|Date)} endTimeStr - Event end time.
 * @param {string} guestEmail - Comma-separated list of guest emails.
 * @param {string} description - Event description.
 * @returns {boolean} True if event was created.
 */
function addToGoogleCalendar(title, dateStr, startTimeStr, endTimeStr, guestEmail, description) {
  try {
    // 1. Parse Date String (YYYY-MM-DD)
    // Note: dateStr might be a Date object from sheet or String from form. Handle both.
    let dateBase = new Date(dateStr);
    
    // 2. Parse Time Strings (HH:mm or HH:mm a)
    // If from Form, it's a "14:00" string. If from Sheet, it might be a Date object.
    
    let startDate = new Date(dateBase);
    let endDate = new Date(dateBase);

    if (typeof startTimeStr === 'string' && startTimeStr.includes(':')) {
        // Handle "14:00" format
        const startParts = startTimeStr.split(':');
        startDate.setHours(parseInt(startParts[0]), parseInt(startParts[1]), 0);
    } else if (startTimeStr instanceof Date) {
        startDate.setHours(startTimeStr.getHours(), startTimeStr.getMinutes(), 0);
    }

    if (typeof endTimeStr === 'string' && endTimeStr.includes(':')) {
        // Handle "15:00" format
        const endParts = endTimeStr.split(':');
        endDate.setHours(parseInt(endParts[0]), parseInt(endParts[1]), 0);
    } else if (endTimeStr instanceof Date) {
        endDate.setHours(endTimeStr.getHours(), endTimeStr.getMinutes(), 0);
    }

    // 3. Create Event
    const event = CalendarApp.getDefaultCalendar().createEvent(
      title,
      startDate,
      endDate,
      {
        description: description,
        guests: guestEmail, // THIS IS WHERE IT WILL AUTO-ADD TO THEIR CALENDAR
        sendInvites: true
      }
    );
    
    Logger.log("Calendar Event Created ID: " + event.getId());
    return true;

  } catch (e) {
    Logger.log("Failed to create Calendar Event: " + e.toString());
    return false;
  }
}
/**
 * HELPER: Get all user emails for autocomplete suggestions
 * @returns {string[]} List of all user emails.
 */
function getAllUserEmailsList() {
  try {
    const ss = SpreadsheetApp.openById(DB_SHEET_ID);
    const sheet = ss.getSheetByName(DB_SHEET_NAME); // 'Users'
    // Get the Email (Col A) and Name (Col D)
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
    
    // Only return the emails (can use "Name <email>" format if preferred)
    const emailList = data.map(row => String(row[0]).trim()).filter(email => email !== "");
    
    return emailList;
  } catch (e) {
    return [];
  }
}
/**
 * 1. Generates OTP, saves it temporarily, and emails it.
 * @param {string} email - The user's email address.
 * @returns {Object} Success status and message.
 */
function sendForgotOTP(email) {
  try {
    const ss = SpreadsheetApp.openById(DB_SHEET_ID);
    const sheet = ss.getSheetByName(DB_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    
    let userExists = false;
    let userName = "";

    // Check if email exists
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === email) {
        userExists = true;
        userName = String(data[i][3]).trim();
        break;
      }
    }

    if (!userExists) {
      return { success: false, message: 'Email address not found.' };
    }

    // Generate 6-digit OTP
    const otp = Math.floor(100000 + Math.random() * 900000).toString();

    // Save OTP to Script Cache (Valid for 10 minutes)
    // Key: email, Value: otp
    const cache = CacheService.getScriptCache();
    cache.put(email, otp, 600); 

    // Send Email
    const subject = "A.Space: Password Reset Code";
    const htmlBody = `
      <div style="font-family: Arial, sans-serif; padding: 20px; background-color: #f4f6f8;">
        <div style="max-width: 500px; margin: 0 auto; background-color: white; padding: 30px; border-radius: 8px; border-top: 5px solid #e53935;">
          <h2 style="color: #333; margin-top: 0;">Verification Code</h2>
          <p>Hi ${userName || 'User'},</p>
          <p>You requested to reset your password. Use the code below to proceed:</p>
          <div style="background-color: #f0f0f0; padding: 15px; text-align: center; font-size: 24px; font-weight: bold; letter-spacing: 5px; color: #e53935;">
            ${otp}
          </div>
          <p>This code expires in 10 minutes.</p>
        </div>
      </div>
    `;

    MailApp.sendEmail({ to: email, subject: subject, htmlBody: htmlBody, name: "A.Space System" });

    return { success: true, message: 'Verification code sent to your email.' };

  } catch (e) {
    Logger.log("Error in sendForgotOTP: " + e.toString());
    return { success: false, message: 'System Error: ' + e.toString() };
  }
}

/**
 * Step 2: Checks if the OTP is correct (Without changing password yet).
 * @param {string} email - The user's email address.
 * @param {string} userOtp - The OTP provided by the user.
 * @returns {Object} Success status and message.
 */
function verifyOTPOnly(email, userOtp) {
  try {
    const cache = CacheService.getScriptCache();
    const cachedOtp = cache.get(email);

    if (cachedOtp && cachedOtp === userOtp) {
      return { success: true, message: 'Code verified.' };
    } else {
      return { success: false, message: 'Invalid or expired verification code.' };
    }
  } catch (e) {
    return { success: false, message: 'System Error: ' + e.toString() };
  }
}

/**
 * Step 3: Finalizes the password change.
 * Note: We check OTP again for security to prevent bypassing Step 2.
 * @param {string} email - The user's email address.
 * @param {string} userOtp - The OTP provided by the user.
 * @param {string} newPassword - The new password to set.
 * @returns {Object} Success status and message.
 */
function finalizePasswordChange(email, userOtp, newPassword) {
  try {
    const cache = CacheService.getScriptCache();
    const cachedOtp = cache.get(email);

    // Security Check
    if (!cachedOtp || cachedOtp !== userOtp) {
      return { success: false, message: 'Session expired or invalid code. Please try again.' };
    }

    const ss = SpreadsheetApp.openById(DB_SHEET_ID);
    const sheet = ss.getSheetByName(DB_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    let rowToUpdate = -1;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === email) {
        rowToUpdate = i + 1;
        break;
      }
    }

    if (rowToUpdate !== -1) {
      sheet.getRange(rowToUpdate, 2).setValue(newPassword); // Update Password
      cache.remove(email); // Clear OTP so it can't be used again
      return { success: true, message: 'Password successfully changed. Please login.' };
    } else {
      return { success: false, message: 'User not found.' };
    }

  } catch (e) {
    return { success: false, message: 'System Error: ' + e.toString() };
  }
}
/**
 * SSO Login: STRICT @gmail.com ONLY.
 * Blocks everything else immediately, even if they are in the database.
 * @returns {Object} Login status and user details if successful.
 */
function loginWithSSO() {
  try {
    // 1. Get the email
    const email = Session.getActiveUser().getEmail();
    
    if (!email) {
      return { success: false, message: 'Could not detect Google Account. Please login.' };
    }

    const emailLower = email.toLowerCase();

    // 2. SECURITY GATE: STRICT @GMAIL.COM ONLY
    // This will IMMEDIATELY block @agpglobal.com or any other domain.
    if (!emailLower.endsWith('@gmail.com')) {
      return { 
        success: false, 
        message: 'Access Denied: Only @gmail.com accounts are allowed via SSO.' 
      };
    }

    // ---------------------------------------------------------
    // If it passed the above check (meaning it's Gmail), 
    // THEN we check the Database.
    // ---------------------------------------------------------

    const ss = SpreadsheetApp.openById(DB_SHEET_ID);
    const sheet = ss.getSheetByName(DB_SHEET_NAME);
    const data = sheet.getDataRange().getValues();

    // 3. CHECK EXISTING (Login)
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).toLowerCase() === emailLower) {
        return { 
          success: true, 
          email: email, 
          role: data[i][2],
          name: data[i][3],
          businessUnit: data[i][4],
          method: 'SSO_EXISTING'
        };
      }
    }
    
    // 4. AUTO-ADD (For New Gmail Users ONLY)
    let nameParts = email.split('@')[0].replace(/[._]/g, ' ').split(' ');
    let generatedName = nameParts.map(n => n.charAt(0).toUpperCase() + n.slice(1)).join(' ');
    
    const newUserRow = [
      email,
      'SSO_Account',
      'user',
      generatedName,
      'LLI'
    ];
    
    sheet.appendRow(newUserRow);

    return { 
      success: true, 
      email: email, 
      role: 'user', 
      name: generatedName, 
      businessUnit: 'LLI', 
      message: 'Account created automatically.',
      method: 'SSO_NEW'
    };

  } catch (e) {
    Logger.log("Error in loginWithSSO: " + e.toString());
    return { success: false, message: 'System Error: ' + e.toString() };
  }
}
/**
 * NEW: Generates a Space Utilization Report (Total Reserved Hours per Resource/Day).
 * Reuses existing HELPER: convertTimeToMinutes.
 * @returns {Object} Success status and array of utilization report objects.
 */
function generateUtilizationReport() {
    try {
        const ss = SpreadsheetApp.openById(DB_SHEET_ID);
        const resSheet = ss.getSheetByName(RESERVATIONS_SHEET_NAME);
        
        if (!resSheet || resSheet.getLastRow() < 2) {
          return { success: true, report: [] };
        }

        const values = resSheet.getRange(2, 1, resSheet.getLastRow() - 1, 10).getValues();
        const timezone = PHILIPPINES_TIMEZONE; 
        
        // Data Structure: { "ResourceName": { "YYYY-MM-DD": TotalMinutes } }
        const utilizationMap = {}; 
        
        // 1. Iterate through Reservations
        values.forEach(row => {
            const status = String(row[0] || '').trim();
            const resourceName = String(row[2] || '').trim();
            const approvalStatus = String(row[9] || '').trim();
            const resDate = row[4]; // Date object
            const startTime = row[5]; // Date object or String
            const endTime = row[6]; // Date object or String

            // Filter: Only Active and Approved reservations are counted
            if (status !== 'Active' || approvalStatus !== 'Approved' || !resourceName || !resDate || !startTime || !endTime) {
                return;
            }

            // Convert Date object to standardized string (YYYY-MM-DD)
            const dateStr = Utilities.formatDate(new Date(resDate), timezone, 'yyyy-MM-dd');
            
            // Convert time to minutes (using existing helper)
            const startMinutes = convertTimeToMinutes(startTime);
            const endMinutes = convertTimeToMinutes(endTime);

            if (startMinutes === -1 || endMinutes === -1 || endMinutes <= startMinutes) {
                return; // Invalid time
            }

            const durationMinutes = endMinutes - startMinutes;
            
            // Store in map
            if (!utilizationMap[resourceName]) {
                utilizationMap[resourceName] = {};
            }
            
            if (!utilizationMap[resourceName][dateStr]) {
                utilizationMap[resourceName][dateStr] = 0;
            }
            
            utilizationMap[resourceName][dateStr] += durationMinutes;
        });

        // 2. Format the data for Report Display (Total Hours)
        const report = [];
        
        for (const resourceName in utilizationMap) {
            for (const dateStr in utilizationMap[resourceName]) {
                const totalMinutes = utilizationMap[resourceName][dateStr];
                const totalHours = (totalMinutes / 60).toFixed(2); // Two decimal places
                
                // Format the date: MMMM dd, yyyy
                const displayDate = Utilities.formatDate(new Date(dateStr), timezone, 'MMMM dd, yyyy');

                report.push({
                    resourceName: resourceName,
                    date: displayDate,
                    totalHours: totalHours
                });
            }
        }
        
        // Sort by Resource Name then Date
        report.sort((a, b) => {
            if (a.resourceName < b.resourceName) return -1;
            if (a.resourceName > b.resourceName) return 1;
            // Secondary sort by date (newest first for easier view)
            return new Date(b.date) - new Date(a.date);
        });

        return { success: true, report: report };

    } catch (e) {
        Logger.log("Error in generateUtilizationReport: " + e.toString());
        return { success: false, message: 'System Error: ' + e.toString(), report: [] };
    }
}
/**
 * ONE-TIME SCRIPT: UPDATED to migrate to 16 columns (Col P for Seat Count).
 * @returns {Object} Migration status and message.
 */
function migrateResourcesData() {
  try {
    const ss = SpreadsheetApp.openById(DB_SHEET_ID);
    const sheet = ss.getSheetByName(RESOURCES_SHEET_NAME);
    
    if (!sheet || sheet.getLastRow() < 2) {
      Logger.log("Resources sheet is empty or not found. Migration skipped.");
      return { success: true, message: "Resources sheet is empty. No migration needed." };
    }
    
    // Read all 11 columns (A-K) (even if there are 15/16 columns, only 11 might have content)
    const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, 11); 
    const values = dataRange.getValues();
    const numRows = values.length;
    
    // Prepare the new 16-column array for writing
    const newValues = [];
    
    // Ensure the new columns (L-P) have headers
    const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (header.length < 16) {
        sheet.getRange(1, 4).setValue('Capacity (Pax)'); // Update D header
        sheet.getRange(1, 12, 1, 5).setValues([['Vehicle Model', 'Plate Number', 'Layout Type', 'Layout Config', 'Seat Count']]); // L-P
        Logger.log("New headers added (L-P).");
    }

    values.forEach((row, index) => {
      const type = String(row[0] || '').trim().toLowerCase(); // Col A
      const capacity = String(row[3] || '').trim(); // Col D (Old Capacity/SeatCount)
      const floor = String(row[5] || '').trim(); // Col F
      const amenities = String(row[6] || '').trim(); // Col G
      
      let newRow = [...row]; // Copy original 11 columns
      
      // Extend to 16 columns (11 + 5 new empty columns)
      while (newRow.length < 16) {
          newRow.push('');
      }

      // --- VEHICLE MIGRATION ---
      if (type === 'vehicle' && floor.includes('Model:')) {
        const modelMatch = floor.match(/Model: (.*?) \|/);
        const plateMatch = floor.match(/Plate: (.*)/);
        
        if (modelMatch) newRow[11] = modelMatch[1].trim(); 
        if (plateMatch) newRow[12] = plateMatch[1].trim(); 
        
        newRow[5] = ''; // CLEAR OLD Floor (Col F)
        // Col D (Capacity) remains as is
      }

      // --- SEAT MIGRATION (UPDATED: Extract SeatCount and clear D/G) ---
      else if (type === 'seat' || type === 'car seat') {
        // 1. Extract Seat Count from old Amenities string (Priority 1) or Capacity (Priority 2)
        let extractedSeatCount = '';
        const countMatch = amenities.match(/SeatCount: (\d+)/);
        if (countMatch) {
             extractedSeatCount = countMatch[1].trim();
        } else {
             // Fallback: use the value from Col D (Capacity)
             extractedSeatCount = capacity;
        }

        // 2. Extract Layout details
        const layoutTypeMatch = amenities.match(/Layout: (.*?) \|/);
        const layoutConfigMatch = amenities.match(/LayoutConfig: ({.*})/);
        
        if (layoutTypeMatch) newRow[13] = layoutTypeMatch[1].trim(); 
        if (layoutConfigMatch) newRow[14] = layoutConfigMatch[1].trim(); 
        
        newRow[15] = extractedSeatCount; // Col P (Seat Count)
        
        newRow[3] = ''; // CLEAR OLD Capacity (Col D)
        newRow[6] = ''; // CLEAR OLD Amenities (Col G)
      }
      
      newValues.push(newRow);
    });

    // Write the migrated data back (A up to P)
    sheet.getRange(2, 1, numRows, 16).setValues(newValues);
    
    Logger.log("Migration successful. %s rows processed.", numRows);
    return { success: true, message: "Resource data migration to 16 columns successful." };
    
  } catch (e) {
    Logger.log("Error during migration: " + e.toString());
    return { success: false, message: 'Migration Error: ' + e.toString() };
  }
}
/**
 * NEW: Retrieves all Active and Approved/Pending reservations
 * for a specific resource type within a specific MONTH (based on dateString).
 * @param {string} resourceType - The type of resource to filter (e.g., 'room', 'seat').
 * @param {string} dateString - The start date of the target month (YYYY-MM-DD).
 * @returns {Object} Success status and array of calendar event objects.
 */
function getReservationsForCalendar(resourceType, dateString) {
  try {
    const ss = SpreadsheetApp.openById(DB_SHEET_ID);
    const sheet = ss.getSheetByName(RESERVATIONS_SHEET_NAME);
    const timezone = PHILIPPINES_TIMEZONE; 
    
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, events: [] };
    }

    // --- NEW: CALCULATE START & END DATE OF THE MONTH ---
    // dateString is the first day of the month (YYYY-MM-DD)
    const targetDate = new Date(dateString);
    const targetYear = targetDate.getFullYear();
    const targetMonth = targetDate.getMonth(); // 0-based
    
    const startOfMonth = new Date(targetYear, targetMonth, 1);
    const endOfMonth = new Date(targetYear, targetMonth + 1, 0); // Last day of the month
    // --- END OF NEW DATE LOGIC ---
    
    // Get all data (10 columns)
    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();
    const targetType = resourceType.toLowerCase();
    
    const events = [];

    values.forEach(row => {
        const status = String(row[0] || '').trim();
        const type = String(row[1] || '').trim().toLowerCase();
        const approval = String(row[9] || '').trim();
        const resDate = row[4];
        
        // Filter 1: Must be Active status
        if (status !== 'Active') return; 

        // Filter 2: Must be Approved or Pending
        if (approval !== 'Approved' && approval !== 'Pending') return;

        // Filter 3: Must be the correct resource type (room, desk, vehicle, or seat)
        if (type !== targetType) return;
        
        // --- NEW FILTER 4: Must be within the month ---
        if (!resDate || !(resDate instanceof Date)) return;
        
        // Reset resDate time for date-only comparison
        resDate.setHours(0, 0, 0, 0); 
        startOfMonth.setHours(0, 0, 0, 0);
        endOfMonth.setHours(23, 59, 59, 999); // Set the end-of-month to the end of the day
        
        if (resDate < startOfMonth || resDate > endOfMonth) return; 
        // --- END OF NEW FILTER 4 ---

        // Format the time for display
        const startTimeStr = row[5] ? Utilities.formatDate(row[5], timezone, 'hh:mm a') : '';
        const endTimeStr = row[6] ? Utilities.formatDate(row[6], timezone, 'hh:mm a') : '';
        
        // Find the Seat # (if present)
        const notes = String(row[8] || '').trim();
        const seatMatch = notes.match(/Seat #(\d+)/);
        const seatDetail = seatMatch ? ` (Seat #${seatMatch[1]})` : '';

        // Format the date for Client-Side filtering
        const resDateStr = Utilities.formatDate(resDate, timezone, 'yyyy-MM-dd');
        
        events.push({
            date: resDateStr, // <--- CRITICAL: Added date for client-side sorting
            resourceName: String(row[2] || '').trim(),
            time: `${startTimeStr} - ${endTimeStr}`,
            requester: String(row[3] || '').trim().split('@')[0],
            approvalStatus: approval,
            seatDetail: seatDetail
        });
    });

    return { success: true, events: events };
  } catch (e) {
    Logger.log("Error in getReservationsForCalendar: " + e.toString());
    return { success: false, message: 'System Error: ' + e.toString(), events: [] };
  }
}
/**
 * NEW: Imports users from an uploaded CSV file.
 * Required CSV Columns (in order): Email, Password, Role, Name, Business Unit
 * ADDED: Email Notification to Successfully Imported Users
 * @param {Object} form - The form submission object containing the CSV file blob.
 * @returns {Object} Import status, message, and report summary.
 */
function importUsersFromCSV(form) {
  const ss = SpreadsheetApp.openById(DB_SHEET_ID);
  const sheet = ss.getSheetByName(DB_SHEET_NAME);
  
  if (!sheet) {
    return { success: false, message: 'DB_SHEET_NAME sheet not found.' };
  }

  // FILE UPLOAD FIX: form.csvFile IS ALREADY THE BLOB.
  const fileBlob = form.csvFile; 
  
  if (!fileBlob || typeof fileBlob.getDataAsString !== 'function') {
     return { success: false, message: 'File processing failed. Please ensure you selected a valid CSV file.' };
  }

  try {
    const csvData = fileBlob.getDataAsString();
    const records = Utilities.parseCsv(csvData);

    // 1. Get existing emails for duplicate check
    const existingEmails = sheet.getLastRow() > 1 ? 
      sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat().map(String).map(e => e.trim().toLowerCase()) : 
      [];

    const newUsersData = [];
    const usersToEmail = []; // NEW ARRAY: This will store the details for the email
    const report = {
      totalRecords: records.length - 1, 
      imported: 0,
      errors: [],
      skipped: 0
    };

    // Iterate starting from Row 2 (index 1) to skip the header
    for (let i = 1; i < records.length; i++) {
      const row = records[i];
      const rowNumber = i + 1;
      
      if (row.length < 5) {
        report.errors.push(`Row ${rowNumber}: Skipped (Missing one or more required columns. Expected 5, got ${row.length}).`);
        report.skipped++;
        continue;
      }
      
      const email = String(row[0] || '').trim().toLowerCase();
      const password = String(row[1] || '').trim();
      const role = String(row[2] || 'user').trim().toLowerCase().split(',').map(r => r.trim()).join(', '); 
      const name = String(row[3] || '').trim();
      const businessUnit = String(row[4] || 'N/A').trim();
      
      // Validation checks (Email format, Domain, Password, Duplicate)
      if (!email || !email.includes('@')) {
        report.errors.push(`Row ${rowNumber} (${row[0]}): Skipped (Invalid email format).`);
        report.skipped++;
        continue;
      }
      if (!email.endsWith('@gmail.com')) {
         report.errors.push(`Row ${rowNumber} (${email}): Skipped (Domain is not authorized, only @gmail.com is allowed).`);
         report.skipped++;
         continue;
      }
      if (!password) {
        report.errors.push(`Row ${rowNumber} (${email}): Skipped (Password cannot be empty).`);
        report.skipped++;
        continue;
      }
      if (existingEmails.includes(email)) {
        report.errors.push(`Row ${rowNumber} (${email}): Skipped (Email already exists in the database).`);
        report.skipped++;
        continue;
      }

      // No error, store in batch array
      newUsersData.push([email, password, role, name, businessUnit]);
      existingEmails.push(email); 
      report.imported++;
      
      // --- NEW: Store details for email ---
      usersToEmail.push({ email: email, name: name, password: password }); 
      // ---------------------------------------------
    }

    // 2. Bulk Append (If any)
    if (newUsersData.length > 0) {
      try {
        sheet.getRange(sheet.getLastRow() + 1, 1, newUsersData.length, 5).setValues(newUsersData);
        SpreadsheetApp.flush(); 
      } catch (e) {
        report.imported -= newUsersData.length; 
        report.errors.push('Database Write Error: Could not save imported data. ' + e.toString());
      }
    }
    
    // 3. --- NEW: Send Welcome Emails to Imported Users ---
    if (report.imported > 0) {
        usersToEmail.forEach(user => {
          sendWelcomeEmail(user.email, user.name, user.password);
        });
    }
    // ---------------------------------------------------
    
    // 4. Final Report
    let finalMessage = `Import finished: ${report.imported} user(s) imported. ${report.skipped} skipped.`;
    if (report.errors.length > 0) {
      finalMessage += ` Total errors: ${report.errors.length}.`;
    }
    
    return { success: report.imported > 0, message: finalMessage, report: report };

  } catch (e) {
    Logger.log("Error in importUsersFromCSV: " + e.toString());
    return { success: false, message: 'System Error during file processing: ' + e.toString(), report: null };
  }
}
/**
 * NEW HELPER: Sends a Welcome Email including the Credentials.
 * @param {string} email - The recipient's email.
 * @param {string} name - The recipient's full name.
 * @param {string} password - The initial password.
 */
function sendWelcomeEmail(email, name, password) {
  try {
    const subject = "Welcome to A.Space Booking System!";
    const scriptUrl = ScriptApp.getService().getUrl();
    const displayName = name || email.split('@')[0];

    const htmlBody = `
      <div style="font-family: Arial, sans-serif; background-color: #f7f9fb; padding: 40px 0; margin: 0;">
        <div style="max-width: 600px; margin: 0 auto; background-color: white; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 10px rgba(0,0,0,0.05);">
          <!-- Header (Reuse primary red style) -->
          <div style="background-color: #e53935; padding: 20px; text-align: center;">
            <h1 style="color: #ffffff; margin: 0; font-size: 24px; letter-spacing: 1px;">A.Space Registration</h1>
          </div>
          
          <!-- Body Content -->
          <div style="padding: 30px;">
            <h2 style="color: #4CAF50; margin-top: 0; font-size: 20px; text-align: center;">Account Created Successfully!</h2>
            <p style="color: #666666; font-size: 15px; line-height: 1.6;">
              Hi ${displayName},
            </p>
            <p style="color: #666666; font-size: 15px; line-height: 1.6; margin-bottom: 25px;">
              Your account has been successfully registered with the A.Space Resource Booking System. You can now sign in using the credentials below.
            </p>

            <!-- Credentials Box -->
            <div style="background-color: #f0f4f8; padding: 20px; border-radius: 6px; border: 1px dashed #ccc; text-align: left;">
              <p style="font-size: 14px; color: #333; margin: 5px 0;"><strong>Email:</strong> ${email}</p>
              <p style="font-size: 14px; color: #e53935; margin: 5px 0;"><strong>Initial Password:</strong> <span style="font-weight: bold;">${password}</span></p>
            </div>

            <div style="text-align: center; margin-top: 30px;">
              <a href="${scriptUrl}" style="background-color: #2196F3; color: #ffffff; padding: 12px 24px; text-decoration: none; border-radius: 5px; font-weight: bold; font-size: 14px; display: inline-block;">Go to A.Space Login</a>
            </div>
          </div>

          <!-- Footer -->
          <div style="background-color: #333333; padding: 15px; text-align: center;">
            <p style="color: #999999; font-size: 12px; margin: 0;">
              Please change your password after your first login.
            </p>
          </div>

        </div>
      </div>
    `;

    // Call existing helper to send the email
    sendEmailNotification(email, subject, htmlBody);
    
  } catch (e) {
    Logger.log("Failed to send Welcome Email to " + email + ": " + e.toString());
  }
}
