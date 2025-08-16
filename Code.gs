// --- Global Configuration: IDs for your three separate Google Spreadsheets ---
// IMPORTANT: These IDs MUST match the Google Sheets you provided.
// Make sure each spreadsheet has a default sheet named 'Sheet1' or the specific tab name you prefer.
const USERS_SPREADSHEET_ID = 'add a google sheet 1 url';
const ANNOUNCEMENTS_SPREADSHEET_ID = 'add a google sheet 2 url';
const PRODUCTS_SPREADSHEET_ID = 'add a googlesheet 3 url';

// New: Dedicated Spreadsheet ID for Orders
const ORDERS_SPREADSHEET_ID = 'add a google sheet 4 url';


// Define the exact names of the sheets (tabs) within each of these spreadsheets
// For simplicity, we assume 'Sheet1' if not specified, or you can rename them.
const USERS_SHEET_NAME = 'Sheet1'; // Or 'Users' if you rename the tab in the Users spreadsheet
const PRODUCTS_SHEET_NAME = 'Sheet1'; // Or 'Products' if you rename the tab in the Products spreadsheet
const ANNOUNCEMENTS_SHEET_NAME = 'Sheet1'; // Or 'Announcements' if you rename the tab in the Announcements spreadsheet
const ORDERS_SHEET_NAME = 'Sheet1'; // The default sheet name within the NEW Orders spreadsheet


// Configuration for the Pending Recharges Sheet
// This sheet will be created within the USERS_SPREADSHEET_ID
const PENDING_RECHARGES_SHEET_NAME = 'PendingRecharges';

// Configuration for the Purchase Requests Sheet
// This sheet will be created within the USERS_SPREADSHEET_ID
const PURCHASE_REQUESTS_SHEET_NAME = 'PurchaseRequests';

// Configuration for the Chat Messages Sheet
// This sheet will be created within the USERS_SPREADSHEET_ID
const CHAT_MESSAGES_SHEET_NAME = 'ChatMessages';


// Define headers for each sheet - CRITICAL FOR DATA CONSISTENCY
// These headers correspond to the columns in your Google Sheets.
// The order here should match the desired order in your sheets.
const USER_HEADERS = ['id', 'name', 'email', 'password', 'createdAt', 'lastLogin', 'balance', 'isActive', 'pending', 'notes'];
const PRODUCT_HEADERS = ['id', 'name', 'description', 'price', 'discountPrice', 'imageUrl', 'rating', 'tags', 'category', 'stock', 'createdAt', 'updatedAt'];
const ANNOUNCEMENT_HEADERS = ['id', 'title', 'content', 'createdAt', 'updatedAt'];
const PENDING_RECHARGE_HEADERS = ['id', 'userEmail', 'amount', 'transactionId', 'screenshotUrl', 'requestDate', 'status', 'adminNotes'];
const PURCHASE_REQUEST_HEADERS = ['id', 'userEmail', 'userName', 'productId', 'productName', 'productPrice', 'requestDate', 'status', 'adminNotes'];
const CHAT_MESSAGES_HEADERS = ['id', 'senderEmail', 'recipientEmail', 'messageContent', 'timestamp'];
const ORDERS_HEADERS = ['id', 'userEmail', 'orderDate', 'productName', 'quantity', 'price', 'totalAmount', 'status']; // Adjust these to match your actual order data structure


// --- Helper Functions for Google Sheet Interaction ---

/**
 * Gets a specific sheet by ID and name.
 * If the sheet does not exist in the specified spreadsheet, it attempts to create it.
 * @param {string} spreadsheetId The ID of the Google Spreadsheet.
 * @param {string} sheetName The name of the sheet (tab) within that spreadsheet.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The Sheet object.
 * @throws {Error} If the spreadsheet itself is not found or cannot be accessed.
 */
function getSheet(spreadsheetId, sheetName) {
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  let sheet = spreadsheet.getSheetByName(sheetName);
  // If sheet doesn't exist, create it (useful for dynamically adding sheets like PendingRecharges)
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    Logger.log(`Created new sheet: ${sheetName} in spreadsheet ID: ${spreadsheetId}`);
  }
  return sheet;
}

/**
 * Reads all data from a specified sheet and converts it to an array of objects.
 * Assumes the first row contains headers. If the sheet is empty, it initializes it with headers.
 * @param {string} spreadsheetId The ID of the Google Spreadsheet.
 * @param {string} sheetName The name of the sheet (tab) within that spreadsheet.
 * @param {Array<string>} headers Expected headers for the sheet.
 * @returns {Array<Object>} An array of JavaScript objects, each representing a row.
 * @throws {Error} If sheet access fails.
 */
function readAllDataFromSheet(spreadsheetId, sheetName, headers) {
  try {
    const sheet = getSheet(spreadsheetId, sheetName);
    const range = sheet.getDataRange();
    const values = range.getValues();

    if (values.length === 0 || values[0].every(cell => !cell)) { // Check if sheet is truly empty
      // If sheet is empty, write headers and return empty array
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      return [];
    }

    const actualHeaders = values[0].map(h => String(h).trim());

    // Basic check for header matching, log warning if mismatch
    if (JSON.stringify(actualHeaders) !== JSON.stringify(headers)) {
      Logger.log(`Warning: Headers for '${sheetName}' in spreadsheet '${spreadsheetId}' do not match expected. Expected: ${headers}, Actual: ${actualHeaders}.`);
      // For robustness, you might add more complex mapping or throw a fatal error here.
      // For now, we proceed assuming `headers` array defines the target object keys.
    }

    const data = [];
    // Start from the second row to skip headers
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const obj = {};
      for (let j = 0; j < headers.length; j++) {
        // Ensure index `j` is within bounds of the `row`
        obj[headers[j]] = (j < row.length) ? row[j] : null; // Use null if column data is missing
      }
      data.push(obj);
    }
    return data;
  } catch (e) {
    Logger.log(`Error reading data from sheet '${sheetName}' (ID: ${spreadsheetId}): ${e.message}`);
    throw new Error(`Failed to read data from '${sheetName}' sheet: ${e.message}`);
  }
}

/**
 * Writes an array of objects to a specified sheet, clearing previous data rows (below headers).
 * Assumes the first row is headers.
 * @param {string} spreadsheetId The ID of the Google Spreadsheet.
 * @param {string} sheetName The name of the sheet (tab) within that spreadsheet.
 * @param {Array<Object>} data The array of JavaScript objects to write.
 * @param {Array<string>} headers Expected headers for the sheet.
 * @throws {Error} If sheet access or write operation fails.
 */
function writeAllDataToSheet(spreadsheetId, sheetName, data, headers) {
  try {
    const sheet = getSheet(spreadsheetId, sheetName);

    // Clear existing content below the header row
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
    }

    // Prepare data for writing to sheet based on headers
    const rows = data.map(obj => headers.map(header => {
      // Handle potential undefined/null values gracefully
      const value = obj[header];
      // Convert boolean to string for Sheets, as Sheets doesn't have native boolean type display
      if (typeof value === 'boolean') {
        return value ? 'TRUE' : 'FALSE';
      }
      return value === undefined || value === null ? '' : value;
    }));

    // Ensure headers are present
    if (sheet.getLastRow() === 0 || sheet.getRange(1, 1).getDisplayValue() === '') {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }

    if (rows.length > 0) {
      // Calculate start row and number of columns to write to
      const numRows = rows.length;
      const numCols = headers.length;
      sheet.getRange(2, 1, numRows, numCols).setValues(rows);
    }
    SpreadsheetApp.flush(); // Ensure changes are written immediately
  } catch (e) {
    Logger.log(`Error writing data to sheet '${sheetName}' (ID: ${spreadsheetId}): ${e.message}`);
    throw new Error(`Failed to write data to '${sheetName}' sheet: ${e.message}`);
  }
}


// --- Web App Service Function (doGet) ---

/**
 * Handles GET requests to the web app.
 * Serves the appropriate HTML file based on the 'admin' or 'page' parameter.
 * @param {GoogleAppsScript.Events.AppsScriptHttpRequestEvent} e The event object.
 * @returns {GoogleAppsScript.HTML.HtmlOutput} The HTML output to serve.
 */
function doGet(e) {
  Logger.log("doGet called with parameters: " + JSON.stringify(e.parameter));

  if (e.parameter.admin === "true") {
    // Serves the Admin.html page for administration.
    return HtmlService.createHtmlOutputFromFile('Admin')
      .setTitle('Admin Panel')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else if (e.parameter.page === "shop") {
    // Serve the Shop.html page.
    return HtmlService.createHtmlOutputFromFile('Shop')
      .setTitle('Our Products')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  // Default: Serves the Index.html for user authentication (login/signup).
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('User Authentication')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


// --- Dashboard Stats Function ---

/**
 * Aggregates various statistics for the admin dashboard.
 * @returns {{success: boolean, stats?: Object, message?: string}} Dashboard statistics.
 */
function getDashboardStats() {
  try {
    // User Stats
    const users = readAllDataFromSheet(USERS_SPREADSHEET_ID, USERS_SHEET_NAME, USER_HEADERS);
    let totalUsers = users.length;
    let activeUsers = 0;
    let pendingUsers = 0;
    users.forEach(user => {
      // Ensure boolean conversion as Sheets might store as string "TRUE"/"FALSE"
      if (String(user.isActive).toLowerCase() === 'true') activeUsers++;
      if (String(user.pending).toLowerCase() === 'true') pendingUsers++;
    });

    // Recharge Request Stats
    const rechargeRequests = readAllDataFromSheet(USERS_SPREADSHEET_ID, PENDING_RECHARGES_SHEET_NAME, PENDING_RECHARGE_HEADERS);
    let totalRechargeRequests = rechargeRequests.length;
    let pendingRechargeRequests = 0;
    let approvedRechargeRequests = 0;
    let rejectedRechargeRequests = 0;
    let totalApprovedRechargeAmount = 0;

    rechargeRequests.forEach(req => {
      if (req.status === 'Pending') {
        pendingRechargeRequests++;
      } else if (req.status === 'Approved') {
        approvedRechargeRequests++;
        totalApprovedRechargeAmount += parseFloat(req.amount || 0);
      } else if (req.status === 'Rejected') {
        rejectedRechargeRequests++;
      }
    });

    // Purchase Request Stats
    const purchaseRequests = readAllDataFromSheet(USERS_SPREADSHEET_ID, PURCHASE_REQUESTS_SHEET_NAME, PURCHASE_REQUEST_HEADERS);
    let totalPurchaseRequests = purchaseRequests.length;
    let pendingPurchaseRequests = 0;
    let approvedPurchaseRequests = 0;
    let rejectedPurchaseRequests = 0;
    let totalApprovedPurchaseValue = 0;

    purchaseRequests.forEach(req => {
      if (req.status === 'Pending') {
        pendingPurchaseRequests++;
      } else if (req.status === 'Approved') {
        approvedPurchaseRequests++;
        totalApprovedPurchaseValue += parseFloat(req.productPrice || 0);
      } else if (req.status === 'Rejected') {
        rejectedPurchaseRequests++;
      }
    });

    return {
      success: true,
      stats: {
        users: {
          total: totalUsers,
          active: activeUsers,
          pending: pendingUsers
        },
        rechargeRequests: {
          total: totalRechargeRequests,
          pending: pendingRechargeRequests,
          approved: approvedRechargeRequests,
          rejected: rejectedRechargeRequests,
          totalApprovedAmount: totalApprovedRechargeAmount,
          revenueGenerated: totalApprovedRechargeAmount // This is the core revenue metric
        },
        purchaseRequests: {
          total: totalPurchaseRequests,
          pending: pendingPurchaseRequests,
          approved: approvedPurchaseRequests,
          rejected: rejectedPurchaseRequests,
          totalApprovedValue: totalApprovedPurchaseValue
        }
      }
    };
  } catch (e) {
    Logger.log(`Error getting dashboard stats: ${e.message}`);
    return { success: false, message: 'Failed to retrieve dashboard stats: ' + e.message };
  }
}


// --- User Management Functions ---

/**
 * Handles user login.
 * @param {string} email User's email.
 * @param {string} password User's password.
 * @returns {{success: boolean, message: string, user?: Object}} Login result.
 */
function loginUser(email, password) {
  try {
    const users = readAllDataFromSheet(USERS_SPREADSHEET_ID, USERS_SHEET_NAME, USER_HEADERS);
    const userIndex = users.findIndex(u => u.email === email);

    if (userIndex === -1) {
      Logger.log(`Login failed: User ${email} not found.`);
      return { success: false, message: 'User not found.' };
    }

    const userData = users[userIndex];
    if (userData.password === password) { // Direct password comparison (for this example)
      if (!userData.isActive || String(userData.isActive).toLowerCase() !== 'true') { // Ensure isActive is treated as boolean true
        Logger.log(`Login failed: User ${email} is inactive.`);
        return { success: false, message: 'Your account is inactive. Please contact support.' };
      }
      userData.lastLogin = new Date().toISOString(); // Update last login time
      writeAllDataToSheet(USERS_SPREADSHEET_ID, USERS_SHEET_NAME, users, USER_HEADERS); // Save updated user data
      Logger.log(`User ${email} logged in successfully.`);
      return { success: true, message: 'Login successful!', user: { name: userData.name, email: userData.email, balance: parseFloat(userData.balance) || 0 } };
    } else {
      Logger.log(`Login failed: Invalid password for ${email}.`);
      return { success: false, message: 'Invalid password.' };
    }
  } catch (e) {
    Logger.log(`Error during login for ${email}: ${e.message}`);
    return { success: false, message: 'An error occurred during login.' };
  }
}

/**
 * Handles user signup.
 * @param {string} name User's name.
 * @param {string} email User's email.
 * @param {string} password User's password.
 * @returns {{success: boolean, message: string}} Signup result.
 */
function signupUser(name, email, password) {
  try {
    const users = readAllDataFromSheet(USERS_SPREADSHEET_ID, USERS_SHEET_NAME, USER_HEADERS);
    if (users.some(u => u.email === email)) {
      Logger.log(`Signup failed: User ${email} already exists.`);
      return { success: false, message: 'User with this email already exists.' };
    }

    // Determine the next available ID
    const maxId = users.reduce((max, u) => Math.max(max, u.id || 0), 0);
    const newUserId = maxId + 1;

    const newUserData = {
      id: newUserId,
      name: name,
      email: email,
      password: password, // In a real app, hash this password!
      createdAt: new Date().toISOString(),
      lastLogin: null,
      balance: 0,
      isActive: true, // Default to active upon signup
      pending: false, // Default to not pending
      notes: '',
      // transactionHistory is not stored in sheet directly, would need separate sheet or stringify
    };

    users.push(newUserData);
    writeAllDataToSheet(USERS_SPREADSHEET_ID, USERS_SHEET_NAME, users, USER_HEADERS);
    Logger.log(`User ${email} signed up successfully.`);
    return { success: true, message: 'Signup successful! You can now log in.' };
  } catch (e) {
    Logger.log(`Error during signup for ${email}: ${e.message}`);
    return { success: false, message: 'An error occurred during signup.' };
  }
}

/**
 * Checks if a user is currently logged in based on session data (e.g., email in LocalStorage).
 * This function is called by client-side JS from shop.html.
 * @param {string} userEmail The email stored in client-side LocalStorage.
 * @returns {{isLoggedIn: boolean, user?: Object}} Login status and basic user data if logged in.
 */
function checkUserSession(userEmail) {
  if (!userEmail) {
    Logger.log("checkUserSession: No email provided from client.");
    return { isLoggedIn: false };
  }

  try {
    const users = readAllDataFromSheet(USERS_SPREADSHEET_ID, USERS_SHEET_NAME, USER_HEADERS);
    const userData = users.find(u => u.email === userEmail);

    if (userData && (userData.isActive === true || String(userData.isActive).toLowerCase() === 'true')) {
      Logger.log(`checkUserSession: User ${userEmail} is active and session valid.`);
      return {
        isLoggedIn: true,
        user: {
          name: userData.name,
          email: userData.email,
          balance: parseFloat(userData.balance) || 0,
        }
      };
    }
  }
  // No catch here, as we want to return isLoggedIn: false on any failure to read data or user not found
  catch (e) {
    Logger.log(`checkUserSession: Error checking session for ${userEmail}: ${e.message}`);
  }
  Logger.log(`checkUserSession: User ${userEmail} not found or inactive.`);
  return { isLoggedIn: false };
}


/**
 * Fetches details for a single user by email.
 * @param {string} email User's email.
 * @returns {Object|null} User data if found, otherwise null.
 */
function getUserDetails(email) {
  try {
    const users = readAllDataFromSheet(USERS_SPREADSHEET_ID, USERS_SHEET_NAME, USER_HEADERS);
    const userData = users.find(u => u.email === email);
    if (userData) {
      // Remove password for security before sending to client
      const userCopy = { ...userData };
      delete userCopy.password;
      return userCopy;
    }
  } catch (e) {
    Logger.log(`Error fetching user details for ${email}: ${e.message}`);
  }
  Logger.log(`User details not found for ${email}.`);
  return null;
}

/**
 * Fetches all users from the user sheet.
 * @returns {Array<Object>} An array of user objects.
 */
function getAllUsers() {
  try {
    const users = readAllDataFromSheet(USERS_SPREADSHEET_ID, USERS_SHEET_NAME, USER_HEADERS);
    // Remove password for security before sending to client
    return users.map(u => {
      const userCopy = { ...u };
      delete userCopy.password;
      return userCopy;
    });
  } catch (e) {
    Logger.log(`Error fetching all users: ${e.message}`);
    return [];
  }
}

/**
 * Updates an existing user's details (admin-side).
 * @param {Object} updatedUserData User data to update.
 * @returns {{success: boolean, message: string}} Result of the update.
 */
function updateUser(updatedUserData) {
  try {
    const users = readAllDataFromSheet(USERS_SPREADSHEET_ID, USERS_SHEET_NAME, USER_HEADERS);
    const userIndex = users.findIndex(u => u.email === updatedUserData.email);

    if (userIndex === -1) {
      Logger.log(`Update failed: User ${updatedUserData.email} not found.`);
      return { success: false, message: 'User not found.' };
    }

    // Update only allowed fields, ensuring types are correct for Sheets
    users[userIndex].name = updatedUserData.name;
    users[userIndex].balance = parseFloat(updatedUserData.balance);
    users[userIndex].isActive = updatedUserData.isActive; // Boolean true/false
    users[userIndex].pending = updatedUserData.pending; // Boolean true/false
    users[userIndex].notes = updatedUserData.notes;

    writeAllDataToSheet(USERS_SPREADSHEET_ID, USERS_SHEET_NAME, users, USER_HEADERS);
    Logger.log(`User ${updatedUserData.email} updated successfully.`);
    return { success: true, message: 'User updated successfully.' };
  } catch (e) {
    Logger.log(`Error updating user ${updatedUserData.email}: ${e.message}`);
    return { success: false, message: 'An error occurred during user update.' };
  }
}

/**
 * Toggles a user's active status.
 * @param {string} email User's email.
 * @param {boolean} newStatus The new active status.
 * @returns {{success: boolean, message: string}} Result of the operation.
 */
function toggleUserStatus(email, newStatus) {
  try {
    const users = readAllDataFromSheet(USERS_SPREADSHEET_ID, USERS_SHEET_NAME, USER_HEADERS);
    const userIndex = users.findIndex(u => u.email === email);

    if (userIndex === -1) {
      Logger.log(`Status toggle failed: User ${email} not found.`);
      return { success: false, message: 'User not found.' };
    }

    users[userIndex].isActive = newStatus;
    writeAllDataToSheet(USERS_SPREADSHEET_ID, USERS_SHEET_NAME, users, USER_HEADERS);
    Logger.log(`User ${email} active status set to ${newStatus}.`);
    return { success: true, message: `User status set to ${newStatus ? 'active' : 'inactive'}.` };
  } catch (e) {
    Logger.log(`Error toggling user status for ${email}: ${e.message}`);
    return { success: false, message: 'An error occurred while updating user status.' };
  }
}

/**
 * Marks a user's pending status as false.
 * @param {string} email User's email.
 * @returns {{success: boolean, message: string}} Result of the operation.
 */
function markAsDone(email) {
  try {
    const users = readAllDataFromSheet(USERS_SPREADSHEET_ID, USERS_SHEET_NAME, USER_HEADERS);
    const userIndex = users.findIndex(u => u.email === email);

    if (userIndex === -1) {
      Logger.log(`Mark done failed: User ${email} not found.`);
      return { success: false, message: 'User not found.' };
    }

    users[userIndex].pending = false;
    writeAllDataToSheet(USERS_SPREADSHEET_ID, USERS_SHEET_NAME, users, USER_HEADERS);
    Logger.log(`User ${email} marked as done.`);
    return { success: true, message: 'User marked as done.' };
  } catch (e) {
    Logger.log(`Error marking user done for ${email}: ${e.message}`);
    return { success: false, message: 'An error occurred while marking user done.' };
  }
}

/**
 * Admin function to delete a user.
 * @param {string} email User's email to delete.
 * @returns {{success: boolean, message: string}} Deletion result.
 */
function adminDeleteUser(email) {
  try {
    let users = readAllDataFromSheet(USERS_SPREADSHEET_ID, USERS_SHEET_NAME, USER_HEADERS);
    const initialLength = users.length;
    users = users.filter(u => u.email !== email);

    if (users.length === initialLength) {
      Logger.log(`Deletion failed: User ${email} not found.`);
      return { success: false, message: 'User not found.' };
    }

    writeAllDataToSheet(USERS_SPREADSHEET_ID, USERS_SHEET_NAME, users, USER_HEADERS);
    Logger.log(`User ${email} deleted successfully by Admin.`);
    return { success: true, message: 'User deleted successfully.' };
  } catch (e) {
    Logger.log(`Error in adminDeleteUser for ${email}: ${e.message}`);
    return { success: false, message: 'An error occurred while deleting the user.' };
  }
}

/**
 * Admin function to add a new user.
 * @param {Object} userData New user details.
 * @returns {{success: boolean, message: string}} Result.
 */
function adminAddUser(userData) {
  try {
    const users = readAllDataFromSheet(USERS_SPREADSHEET_ID, USERS_SHEET_NAME, USER_HEADERS);
    const email = userData.email;

    if (users.some(u => u.email === email)) {
      Logger.log(`Admin add user failed: User ${email} already exists.`);
      return { success: false, message: 'User with this email already exists.' };
    }

    // Determine the next available ID
    const maxId = users.reduce((max, u) => Math.max(max, u.id || 0), 0);
    const newUserId = maxId + 1;

    const newUserData = {
      id: newUserId,
      name: userData.name,
      email: email,
      password: userData.password, // Admin sets initial password
      createdAt: new Date().toISOString(),
      lastLogin: null,
      balance: parseFloat(userData.balance) || 0,
      isActive: true, // Admin-added users are active by default
      pending: false, // Admin-added users are not pending by default
      notes: 'Added by Admin',
    };

    users.push(newUserData);
    writeAllDataToSheet(USERS_SPREADSHEET_ID, USERS_SHEET_NAME, users, USER_HEADERS);
    Logger.log(`Admin added new user: ${email}.`);
    return { success: true, message: 'User added successfully by Admin.' };
  } catch (e) {
    Logger.log(`Error in adminAddUser for ${email}: ${e.message}`);
    return { success: false, message: 'An error occurred while adding the user.' };
  }
}

// --- Product Management Functions ---

/**
 * Adds a new product.
 * @param {Object} productData Product details.
 * @returns {{success: boolean, message: string}} Result.
 */
function addProduct(productData) {
  try {
    const products = readAllDataFromSheet(PRODUCTS_SPREADSHEET_ID, PRODUCTS_SHEET_NAME, PRODUCT_HEADERS);

    // Determine the next available ID
    const maxId = products.reduce((max, p) => Math.max(max, p.id || 0), 0);
    const newProductId = maxId + 1;

    const newProduct = {
      id: newProductId,
      name: productData.name,
      description: productData.description,
      price: parseFloat(productData.price),
      discountPrice: productData.discountPrice ? parseFloat(productData.discountPrice) : null,
      imageUrl: productData.imageUrl,
      rating: parseFloat(productData.rating) || 0,
      tags: Array.isArray(productData.tags) ? productData.tags.join(',') : (productData.tags || ''), // Store as comma-separated string
      category: productData.category || 'Uncategorized',
      stock: productData.stock === '-1' ? -1 : parseInt(productData.stock) || 0, // -1 for unlimited
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    };

    products.push(newProduct);
    writeAllDataToSheet(PRODUCTS_SPREADSHEET_ID, PRODUCTS_SHEET_NAME, products, PRODUCT_HEADERS);
    Logger.log('Product added successfully: ' + newProduct.name);
    return { success: true, message: 'Product added successfully!' };
  } catch (e) {
    Logger.log('Error adding product: ' + e.message);
    return { success: false, message: 'Failed to add product.' };
  }
}

/**
 * Fetches all products.
 * @returns {{success: boolean, products: Array<Object>, message?: string}} List of products.
 */
function getAllProducts() {
  try {
    const products = readAllDataFromSheet(PRODUCTS_SPREADSHEET_ID, PRODUCTS_SHEET_NAME, PRODUCT_HEADERS);
    // Convert tags string back to array if needed for client-side processing
    return {
      success: true,
      products: products.map(p => ({
        ...p,
        tags: typeof p.tags === 'string' && p.tags ? p.tags.split(',') : []
      }))
    };
  } catch (e) {
    Logger.log('Error fetching products: ' + e.message);
    return { success: false, message: 'Failed to fetch products.' };
  }
}

/**
 * Updates an existing product.
 * @param {Object} updatedProductData Updated product details.
 * @returns {{success: boolean, message: string}} Result.
 */
function updateProduct(updatedProductData) {
  try {
    let products = readAllDataFromSheet(PRODUCTS_SPREADSHEET_ID, PRODUCTS_SHEET_NAME, PRODUCT_HEADERS);
    const index = products.findIndex(p => p.id === updatedProductData.id);

    if (index === -1) {
      return { success: false, message: 'Product not found.' };
    }

    products[index] = {
      ...products[index], // Keep existing properties not explicitly updated
      name: updatedProductData.name,
      description: updatedProductData.description,
      price: parseFloat(updatedProductData.price),
      discountPrice: updatedProductData.discountPrice !== null ? parseFloat(updatedProductData.discountPrice) : null,
      imageUrl: updatedProductData.imageUrl,
      rating: parseFloat(updatedProductData.rating) || 0,
      tags: Array.isArray(updatedProductData.tags) ? updatedProductData.tags.join(',') : (updatedProductData.tags || ''), // Convert array back to comma-separated string
      category: updatedProductData.category || 'Uncategorized',
      stock: updatedProductData.stock === -1 ? -1 : parseInt(updatedProductData.stock) || 0,
      updatedAt: new Date().toISOString()
    };

    writeAllDataToSheet(PRODUCTS_SPREADSHEET_ID, PRODUCTS_SHEET_NAME, products, PRODUCT_HEADERS);
    Logger.log('Product updated successfully: ' + updatedProductData.id);
    return { success: true, message: 'Product updated successfully!' };
  } catch (e) {
    Logger.log('Error updating product: ' + e.message);
    return { success: false, message: 'Failed to update product.' };
  }
}

/**
 * Deletes a product by ID.
 * @param {number} productId The ID of the product to delete.
 * @returns {{success: boolean, message: string}} Result.
 */
function deleteProduct(productId) {
  try {
    let products = readAllDataFromSheet(PRODUCTS_SPREADSHEET_ID, PRODUCTS_SHEET_NAME, PRODUCT_HEADERS);
    const initialLength = products.length;
    products = products.filter(p => p.id !== productId);

    if (products.length === initialLength) {
      return { success: false, message: 'Product not found.' };
    }

    writeAllDataToSheet(PRODUCTS_SPREADSHEET_ID, PRODUCTS_SHEET_NAME, products, PRODUCT_HEADERS);
    Logger.log('Product deleted successfully: ' + productId);
    return { success: true, message: 'Product deleted successfully!' };
  } catch (e) {
    Logger.log('Error deleting product: ' + e.message);
    return { success: false, message: 'Failed to delete product.' };
  }
}


// --- Announcement Management Functions ---

/**
 * Adds a new announcement.
 * @param {Object} announcementData Announcement details.
 * @returns {{success: boolean, message: string}} Result.
 */
function addAnnouncement(announcementData) {
  try {
    const announcements = readAllDataFromSheet(ANNOUNCEMENTS_SPREADSHEET_ID, ANNOUNCEMENTS_SHEET_NAME, ANNOUNCEMENT_HEADERS);

    const maxId = announcements.reduce((max, a) => Math.max(max, a.id || 0), 0);
    const newAnnouncementId = maxId + 1;

    const newAnnouncement = {
      id: newAnnouncementId,
      title: announcementData.title,
      content: announcementData.content,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    };

    announcements.push(newAnnouncement);
    writeAllDataToSheet(ANNOUNCEMENTS_SPREADSHEET_ID, ANNOUNCEMENTS_SHEET_NAME, announcements, ANNOUNCEMENT_HEADERS);
    Logger.log('Announcement added successfully: ' + newAnnouncement.title);
    return { success: true, message: 'Announcement added successfully!' };
  } catch (e) {
    Logger.log('Error adding announcement: ' + e.message);
    return { success: false, message: 'Failed to add announcement.' };
  }
}

/**
 * Fetches all announcements.
 * @returns {{success: boolean, announcements: Array<Object>, message?: string}} List of announcements.
 */
function getAllAnnouncements() {
  try {
    const announcements = readAllDataFromSheet(ANNOUNCEMENTS_SPREADSHEET_ID, ANNOUNCEMENTS_SHEET_NAME, ANNOUNCEMENT_HEADERS);
    return { success: true, announcements: announcements };
  } catch (e) {
    Logger.log('Error fetching announcements: ' + e.message);
    return { success: false, message: 'Failed to fetch announcements.' };
  }
}

/**
 * Updates an existing announcement.
 * @param {Object} updatedAnnouncementData Updated announcement details.
 * @returns {{success: boolean, message: string}} Result.
 */
function updateAnnouncement(updatedAnnouncementData) {
  try {
    let announcements = readAllDataFromSheet(ANNOUNCEMENTS_SPREADSHEET_ID, ANNOUNCEMENTS_SHEET_NAME, ANNOUNCEMENT_HEADERS);
    const index = announcements.findIndex(a => a.id === updatedAnnouncementData.id);

    if (index === -1) {
      return { success: false, message: 'Announcement not found.' };
    }

    announcements[index] = {
      ...announcements[index], // Keep existing properties not explicitly updated
      title: updatedAnnouncementData.title,
      content: updatedAnnouncementData.content,
      updatedAt: new Date().toISOString()
    };
    writeAllDataToSheet(ANNOUNCEMENTS_SPREADSHEET_ID, ANNOUNCEMENTS_SHEET_NAME, announcements, ANNOUNCEMENT_HEADERS);
    Logger.log('Announcement updated successfully: ' + updatedAnnouncementData.id);
    return { success: true, message: 'Announcement updated successfully!' };
  } catch (e) {
    Logger.log('Error updating announcement: ' + e.message);
    return { success: false, message: 'Failed to update announcement.' };
  }
}

/**
 * Deletes an announcement by ID.
 * @param {number} announcementId The ID of the announcement to delete.
 * @returns {{success: boolean, message: string}} Result.
 */
function deleteAnnouncement(announcementId) {
  try {
    let announcements = readAllDataFromSheet(ANNOUNCEMENTS_SPREADSHEET_ID, ANNOUNCEMENTS_SHEET_NAME, ANNOUNCEMENT_HEADERS);
    const initialLength = announcements.length;
    announcements = announcements.filter(a => a.id !== announcementId);

    if (announcements.length === initialLength) {
      return { success: false, message: 'Announcement not found.' };
    }

    writeAllDataToSheet(ANNOUNCEMENTS_SPREADSHEET_ID, ANNOUNCEMENTS_SHEET_NAME, announcements, ANNOUNCEMENT_HEADERS);
    Logger.log('Announcement deleted successfully: ' + announcementId);
    return { success: true, message: 'Announcement deleted successfully!' };
  } catch (e) {
    Logger.log('Error deleting announcement: ' + e.message);
    return { success: false, message: 'Failed to delete announcement.' };
  }
}


// --- Recharge Request Functions (User & Admin) ---

/**
 * Submits a new recharge request from a user.
 * @param {string} userEmail The email of the user requesting recharge.
 * @param {number} amount The amount to recharge.
 * @param {string} transactionId The transaction ID from the payment.
 * @param {string} screenshotUrl URL of the payment screenshot (e.g., Google Drive link).
 * @returns {{success: boolean, message: string}} Result of the submission.
 */
function submitRechargeRequest(userEmail, amount, transactionId, screenshotUrl) {
  try {
    const requests = readAllDataFromSheet(USERS_SPREADSHEET_ID, PENDING_RECHARGES_SHEET_NAME, PENDING_RECHARGE_HEADERS);

    // Generate new ID for the request
    const maxId = requests.reduce((max, r) => Math.max(max, r.id || 0), 0);
    const newRequestId = maxId + 1;

    const newRequest = {
      id: newRequestId,
      userEmail: userEmail,
      amount: parseFloat(amount),
      transactionId: transactionId,
      screenshotUrl: screenshotUrl,
      requestDate: new Date().toISOString(),
      status: 'Pending', // Initial status
      adminNotes: ''
    };

    requests.push(newRequest);
    writeAllDataToSheet(USERS_SPREADSHEET_ID, PENDING_RECHARGES_SHEET_NAME, requests, PENDING_RECHARGE_HEADERS);
    Logger.log(`Recharge request submitted by ${userEmail} for ${amount}. ID: ${newRequestId}`);
    return { success: true, message: 'Recharge request submitted successfully! It will be reviewed by admin.' };
  } catch (e) {
    Logger.log(`Error submitting recharge request for ${userEmail}: ${e.message}`);
    return { success: false, message: 'Failed to submit recharge request: ' + e.message };
  }
}

/**
 * Fetches all pending recharge requests for the admin panel.
 * @returns {{success: boolean, requests: Array<Object>, message?: string}} List of pending requests.
 */
function getPendingRechargeRequests() {
  try {
    const requests = readAllDataFromSheet(USERS_SPREADSHEET_ID, PENDING_RECHARGES_SHEET_NAME, PENDING_RECHARGE_HEADERS);
    // Return all requests to show them in admin panel (pending, approved, rejected)
    // Admin can then filter further if needed client-side.
    return { success: true, requests: requests };
  } catch (e) {
    Logger.log(`Error fetching pending recharge requests: ${e.message}`);
    return { success: false, message: 'Failed to fetch recharge requests.' };
  }
}

/**
 * Approves a recharge request and updates user balance.
 * @param {number} requestId The ID of the recharge request.
 * @param {string} adminEmail The email of the admin approving the request.
 * @param {string} notes Optional notes from the admin.
 * @returns {{success: boolean, message: string}} Result of the approval.
 */
function approveRechargeRequest(requestId, adminEmail, notes = '') {
  try {
    const rechargeRequests = readAllDataFromSheet(USERS_SPREADSHEET_ID, PENDING_RECHARGES_SHEET_NAME, PENDING_RECHARGE_HEADERS);
    const requestIndex = rechargeRequests.findIndex(r => r.id === requestId);

    if (requestIndex === -1) {
      return { success: false, message: 'Recharge request not found.' };
    }

    const request = rechargeRequests[requestIndex];
    if (request.status === 'Approved') {
      return { success: false, message: 'This recharge request has already been approved.' };
    }
    if (request.status === 'Rejected') {
      return { success: false, message: 'This recharge request has been rejected. Cannot approve a rejected request.' };
    }

    // Update request status and admin notes
    request.status = 'Approved';
    request.adminNotes = `Approved by ${adminEmail} on ${new Date().toLocaleString()}. Notes: ${notes}`;
    writeAllDataToSheet(USERS_SPREADSHEET_ID, PENDING_RECHARGES_SHEET_NAME, rechargeRequests, PENDING_RECHARGE_HEADERS);

    // Update user's balance
    const users = readAllDataFromSheet(USERS_SPREADSHEET_ID, USERS_SHEET_NAME, USER_HEADERS);
    const userIndex = users.findIndex(u => u.email === request.userEmail);

    if (userIndex === -1) {
      Logger.log(`User ${request.userEmail} not found for recharge approval. Request ID: ${requestId}`);
      return { success: false, message: 'User for recharge request not found. Request approved but balance not updated.' };
    }

    // Ensure balance is treated as a number
    users[userIndex].balance = parseFloat(users[userIndex].balance || 0) + parseFloat(request.amount);
    writeAllDataToSheet(USERS_SPREADSHEET_ID, USERS_SHEET_NAME, users, USER_HEADERS);

    Logger.log(`Recharge request ${requestId} approved. User ${request.userEmail} balance updated.`);
    return { success: true, message: 'Recharge request approved and user balance updated.' };
  } catch (e) {
    Logger.log(`Error approving recharge request ${requestId}: ${e.message}`);
    return { success: false, message: 'Failed to approve recharge request: ' + e.message };
  }
}

/**
 * Rejects a recharge request.
 * @param {number} requestId The ID of the recharge request.
 * @param {string} adminEmail The email of the admin rejecting the request.
 * @param {string} notes Optional notes from the admin explaining rejection.
 * @returns {{success: boolean, message: string}} Result of the rejection.
 */
function rejectRechargeRequest(requestId, adminEmail, notes = '') {
  try {
    const rechargeRequests = readAllDataFromSheet(USERS_SPREADSHEET_ID, PENDING_RECHARGES_SHEET_NAME, PENDING_RECHARGE_HEADERS);
    const requestIndex = rechargeRequests.findIndex(r => r.id === requestId);

    if (requestIndex === -1) {
      return { success: false, message: 'Recharge request not found.' };
    }

    const request = rechargeRequests[requestIndex];
    if (request.status === 'Rejected') {
      return { success: false, message: 'This recharge request has already been rejected.' };
    }
    if (request.status === 'Approved') {
      return { success: false, message: 'This recharge request has already been approved. Cannot reject an approved request.' };
    }

    request.status = 'Rejected';
    request.adminNotes = `Rejected by ${adminEmail} on ${new Date().toLocaleString()}. Reason: ${notes}`;
    writeAllDataToSheet(USERS_SPREADSHEET_ID, PENDING_RECHARGES_SHEET_NAME, rechargeRequests, PENDING_RECHARGE_HEADERS);

    Logger.log(`Recharge request ${requestId} rejected.`);
    return { success: true, message: 'Recharge request rejected.' };
  } catch (e) {
    Logger.log(`Error rejecting recharge request ${requestId}: ${e.message}`);
    return { success: false, message: 'Failed to reject recharge request: ' + e.message };
  }
}


// --- Purchase Request Functions (User & Admin) ---

/**
 * Submits a new purchase request from a user.
 * @param {string} userEmail The email of the user.
 * @param {string} userName The name of the user.
 * @param {number} productId The ID of the product being requested.
 * @param {string} productName The name of the product.
 * @param {number} productPrice The price of the product at the time of request.
 * @returns {{success: boolean, message: string}} Result of the submission.
 */
function submitPurchaseRequest(userEmail, userName, productId, productName, productPrice) {
  try {
    const requests = readAllDataFromSheet(USERS_SPREADSHEET_ID, PURCHASE_REQUESTS_SHEET_NAME, PURCHASE_REQUEST_HEADERS);

    // Generate new ID for the request
    const maxId = requests.reduce((max, r) => Math.max(max, r.id || 0), 0);
    const newRequestId = maxId + 1;

    const newRequest = {
      id: newRequestId,
      userEmail: userEmail,
      userName: userName,
      productId: productId,
      productName: productName,
      productPrice: parseFloat(productPrice),
      requestDate: new Date().toISOString(),
      status: 'Pending', // Initial status
      adminNotes: ''
    };

    requests.push(newRequest);
    writeAllDataToSheet(USERS_SPREADSHEET_ID, PURCHASE_REQUESTS_SHEET_NAME, requests, PURCHASE_REQUEST_HEADERS);
    Logger.log(`Purchase request submitted by ${userEmail} for product ${productName}. ID: ${newRequestId}`);
    return { success: true, message: 'Purchase request submitted successfully! Admin will review it.' };
  } catch (e) {
    Logger.log(`Error submitting purchase request for ${userEmail}: ${e.message}`);
    return { success: false, message: 'Failed to submit purchase request: ' + e.message };
  }
}

/**
 * Fetches all purchase requests for the admin panel.
 * @returns {{success: boolean, requests: Array<Object>, message?: string}} List of purchase requests.
 */
function getPurchaseRequests() {
  try {
    const requests = readAllDataFromSheet(USERS_SPREADSHEET_ID, PURCHASE_REQUESTS_SHEET_NAME, PURCHASE_REQUEST_HEADERS);
    // Return all requests (pending, approved, rejected) to show them in admin panel.
    return { success: true, requests: requests };
  } catch (e) {
    Logger.log(`Error fetching purchase requests: ${e.message}`);
    return { success: false, message: 'Failed to fetch purchase requests.' };
  }
}

/**
 * Approves a purchase request, updates user balance and product stock.
 * @param {number} requestId The ID of the purchase request.
 * @param {string} adminEmail The email of the admin approving the request.
 * @param {string} notes Optional notes from the admin.
 * @returns {{success: boolean, message: string}} Result of the approval.
 */
function approvePurchaseRequest(requestId, adminEmail, notes = '') {
  try {
    const purchaseRequests = readAllDataFromSheet(USERS_SPREADSHEET_ID, PURCHASE_REQUESTS_SHEET_NAME, PURCHASE_REQUEST_HEADERS);
    const requestIndex = purchaseRequests.findIndex(r => r.id === requestId);

    if (requestIndex === -1) {
      return { success: false, message: 'Purchase request not found.' };
    }

    const request = purchaseRequests[requestIndex];
    if (request.status === 'Approved') {
      return { success: false, message: 'This purchase request has already been approved.' };
    }
    if (request.status === 'Rejected') {
      return { success: false, message: 'This purchase request has been rejected. Cannot approve a rejected request.' };
    }

    // 2. Get user and product data
    const users = readAllDataFromSheet(USERS_SPREADSHEET_ID, USERS_SHEET_NAME, USER_HEADERS);
    const userIndex = users.findIndex(u => u.email === request.userEmail);
    if (userIndex === -1) {
      Logger.log(`User ${request.userEmail} not found for purchase approval. Request ID: ${requestId}`);
      return { success: false, message: `User not found for request ${requestId}. Purchase not approved.` };
    }
    const user = users[userIndex];

    const products = readAllDataFromSheet(PRODUCTS_SPREADSHEET_ID, PRODUCTS_SHEET_NAME, PRODUCT_HEADERS);
    const productIndex = products.findIndex(p => p.id === request.productId);
    if (productIndex === -1) {
      Logger.log(`Product ${request.productName} (ID: ${request.productId}) not found for purchase approval. Request ID: ${requestId}`);
      return { success: false, message: `Product not found for request ${requestId}. Purchase not approved.` };
    }
    const product = products[productIndex];

    // 3. Check balance and stock
    const purchasePrice = parseFloat(request.productPrice);
    if (user.balance < purchasePrice) {
      return { success: false, message: `User ${user.email} has insufficient balance (₹${user.balance.toFixed(2)}) for purchase (₹${purchasePrice.toFixed(2)}).` };
    }

    if (product.stock !== -1 && product.stock <= 0) { // Check if stock is limited and <= 0
      return { success: false, message: `Product ${product.name} is out of stock.` };
    }

    // 4. Perform transactions
    user.balance -= purchasePrice; // Deduct from user balance
    if (product.stock !== -1) {
      product.stock -= 1; // Decrement stock if not unlimited
    }

    request.status = 'Approved';
    request.adminNotes = `Approved by ${adminEmail} on ${new Date().toLocaleString()}. Notes: ${notes}`;

    // 5. Write back updated data
    writeAllDataToSheet(USERS_SPREADSHEET_ID, PURCHASE_REQUESTS_SHEET_NAME, purchaseRequests, PURCHASE_REQUEST_HEADERS);
    writeAllDataToSheet(USERS_SPREADSHEET_ID, USERS_SHEET_NAME, users, USER_HEADERS); // Update user balance
    writeAllDataToSheet(PRODUCTS_SPREADSHEET_ID, PRODUCTS_SHEET_NAME, products, PRODUCT_HEADERS); // Update product stock

    Logger.log(`Purchase request ${requestId} approved. User ${request.userEmail} purchased ${request.productName}.`);
    return { success: true, message: 'Purchase request approved, user balance updated, and stock adjusted.' };
  } catch (e) {
    Logger.log(`Error approving purchase request ${requestId}: ${e.message}`);
    return { success: false, message: 'Failed to approve purchase request: ' + e.message };
  }
}

/**
 * Rejects a purchase request.
 * @param {number} requestId The ID of the purchase request.
 * @param {string} adminEmail The email of the admin rejecting the request.
 * @param {string} notes Optional notes from the admin explaining rejection.
 * @returns {{success: boolean, message: string}} Result of the rejection.
 */
function rejectPurchaseRequest(requestId, adminEmail, notes = '') {
  try {
    const purchaseRequests = readAllDataFromSheet(USERS_SPREADSHEET_ID, PURCHASE_REQUESTS_SHEET_NAME, PURCHASE_REQUEST_HEADERS);
    const requestIndex = purchaseRequests.findIndex(r => r.id === requestId);

    if (requestIndex === -1) {
      return { success: false, message: 'Purchase request not found.' };
    }

    const request = purchaseRequests[requestIndex];
    if (request.status === 'Rejected') {
      return { success: false, message: 'This purchase request has already been rejected.' };
    }
    if (request.status === 'Approved') {
      return { success: false, message: 'This purchase request has already been approved. Cannot reject an approved request.' };
    }

    request.status = 'Rejected';
    request.adminNotes = `Rejected by ${adminEmail} on ${new Date().toLocaleString()}. Reason: ${notes}`;
    writeAllDataToSheet(USERS_SPREADSHEET_ID, PURCHASE_REQUESTS_SHEET_NAME, purchaseRequests, PURCHASE_REQUEST_HEADERS);

    Logger.log(`Purchase request ${requestId} rejected. User ${request.userEmail}.`);
    return { success: true, message: 'Purchase request rejected.' };
  } catch (e) {
    Logger.log(`Error rejecting purchase request ${requestId}: ${e.message}`);
    return { success: false, message: 'Failed to reject purchase request: ' + e.message };
  }
}


// --- Chat Functions ---
const ADMIN_EMAIL = "admin@example.com"; // Set your admin email here

/**
 * Sends a chat message.
 * @param {string} senderEmail The email of the sender (user or admin).
 * @param {string} recipientEmail The email of the recipient (user or admin).
 * @param {string} messageContent The content of the message.
 * @returns {{success: boolean, message?: string}} Result of sending the message.
 */
function sendChatMessage(senderEmail, recipientEmail, messageContent) {
  try {
    const messages = readAllDataFromSheet(USERS_SPREADSHEET_ID, CHAT_MESSAGES_SHEET_NAME, CHAT_MESSAGES_HEADERS);

    const maxId = messages.reduce((max, msg) => Math.max(max, msg.id || 0), 0);
    const newId = maxId + 1;

    const newMessage = {
      id: newId,
      senderEmail: senderEmail,
      recipientEmail: recipientEmail,
      messageContent: messageContent,
      timestamp: new Date().toISOString()
    };

    messages.push(newMessage);
    writeAllDataToSheet(USERS_SPREADSHEET_ID, CHAT_MESSAGES_SHEET_NAME, messages, CHAT_MESSAGES_HEADERS);

    Logger.log(`Message sent from ${senderEmail} to ${recipientEmail}`);
    return { success: true };
  } catch (e) {
    Logger.log(`Error sending chat message: ${e.message}`);
    return { success: false, message: 'Failed to send message.' };
  }
}

/**
 * Gets chat history between a specific user and the admin.
 * @param {string} userEmail The email of the user.
 * @param {string} adminEmail The email of the admin.
 * @returns {{success: boolean, chatHistory: Array<Object>, message?: string}} Chat history.
 */
function getChatHistory(userEmail, adminEmail) {
  try {
    const messages = readAllDataFromSheet(USERS_SPREADSHEET_ID, CHAT_MESSAGES_SHEET_NAME, CHAT_MESSAGES_HEADERS);

    // Filter messages that are either from user to admin or from admin to user
    const filteredHistory = messages.filter(msg =>
      (msg.senderEmail === userEmail && msg.recipientEmail === adminEmail) ||
      (msg.senderEmail === adminEmail && msg.recipientEmail === userEmail)
    ).sort((a, b) => new Date(a.timestamp).getTime() - new Date(b.timestamp).getTime()); // Sort by timestamp

    return { success: true, chatHistory: filteredHistory };
  } catch (e) {
    Logger.log(`Error fetching chat history for ${userEmail}: ${e.message}`);
    return { success: false, chatHistory: [], message: 'Failed to fetch chat history: ' + e.message };
  }
}

/**
 * Gets a list of unique user emails who have sent or received chat messages.
 * This is useful for the admin to see who to chat with.
 * @param {string} adminEmail The email of the admin.
 * @returns {{success: boolean, users: Array<string>, message?: string}} List of user emails.
 */
function getUsersWithChatHistory(adminEmail) {
  try {
    const messages = readAllDataFromSheet(USERS_SPREADSHEET_ID, CHAT_MESSAGES_SHEET_NAME, CHAT_MESSAGES_HEADERS);
    const userEmails = new Set();
    messages.forEach(msg => {
      if (msg.senderEmail !== adminEmail) {
        userEmails.add(msg.senderEmail);
      }
      if (msg.recipientEmail !== adminEmail) {
        userEmails.add(msg.recipientEmail);
      }
    });
    return { success: true, users: Array.from(userEmails) };
  } catch (e) {
    Logger.log(`Error getting users with chat history: ${e.message}`);
    return { success: false, users: [], message: 'Failed to get users with chat history: ' + e.message };
  }
}

/**
 * Fetches order history for a given user email.
 * @param {string} userEmail The email of the user.
 * @returns {{success: boolean, orders: Array<Object>, message?: string}} An object containing success status, order data, and an optional message.
 */
function getOrderHistory(userEmail) {
  try {
    const orders = readAllDataFromSheet(ORDERS_SPREADSHEET_ID, ORDERS_SHEET_NAME, ORDERS_HEADERS); // Use the dedicated ORDERS_SPREADSHEET_ID

    // Filter orders by the user's email
    const userOrders = orders.filter(order => order.userEmail === userEmail);

    return { success: true, orders: userOrders };

  } catch (e) {
    Logger.log(`Error fetching order history for ${userEmail}: ${e.message}`);
    return { success: false, orders: [], message: 'Failed to fetch order history: ' + e.message };
  }

}
