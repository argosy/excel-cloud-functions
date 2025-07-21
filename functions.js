// functions.js - Excel Custom Functions Implementation

// IMPORTANT: Your actual Cloud Run service URL
const CLOUD_RUN_URL = 'https://excel-add-one-function-449328337363.us-central1.run.app/add-one';

// Wait for CustomFunctions to be available
function waitForCustomFunctions() {
  if (typeof CustomFunctions !== 'undefined') {
    console.log('CustomFunctions is available, registering functions...');
    registerFunctions();
  } else {
    console.log('CustomFunctions not yet available, waiting...');
    setTimeout(waitForCustomFunctions, 100);
  }
}

/**
 * Adds 1 to each element in a 2D range using Google Cloud Run
 * @customfunction ADDONE
 * @param {number[][]} range The 2D range of numbers to process
 * @returns {Promise<number[][]>} 2D array with 1 added to each element
 */
async function addOne(range) {
  try {
    // Input validation
    if (!range || !Array.isArray(range)) {
      return [["Error: Please provide a valid range"]];
    }

    // Handle single cell input (convert to 2D array)
    if (!Array.isArray(range[0])) {
      range = [range];
    }

    // Log for debugging (visible in browser console)
    console.log('ADDONE called with range:', range);
    console.log('Calling Cloud Run URL:', CLOUD_RUN_URL);

    // Call Cloud Run service
    const response = await fetch(CLOUD_RUN_URL, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Accept': 'application/json',
      },
      body: JSON.stringify({
        data: range
      })
    });

    // Handle HTTP errors
    if (!response.ok) {
      const errorText = await response.text();
      console.error('HTTP Error:', response.status, errorText);
      return [[\`HTTP Error ${response.status}: ${errorText}\`]];
    }

    const result = await response.json();
    console.log('Cloud Run response:', result);
    
    // Handle service errors
    if (result.status !== 'success') {
      console.error('Service Error:', result.error);
      return [['Service Error: ' + (result.error || 'Unknown error')]];
    }

    // Return the processed 2D array
    return result.result;
    
  } catch (error) {
    // Handle network or parsing errors
    console.error('Custom function error:', error);
    return [['Error: ' + error.message]];
  }
}

/**
 * Adds a custom value to each element in a 2D range (local processing)
 * @customfunction ADDVALUE
 * @param {number[][]} range The 2D range of numbers to process
 * @param {number} value The value to add to each element (default: 1)
 * @returns {number[][]} 2D array with the value added to each element
 */
function addValue(range, value = 1) {
  try {
    if (!range || !Array.isArray(range)) {
      return [["Error: Please provide a valid range"]];
    }

    // Convert single cell to 2D array
    if (!Array.isArray(range[0])) {
      range = [range];
    }

    // Process locally for better performance with simple operations
    const result = range.map(row => 
      row.map(cell => {
        if (typeof cell === 'number') {
          return cell + value;
        } else if (typeof cell === 'string' && !isNaN(parseFloat(cell))) {
          return parseFloat(cell) + value;
        } else if (cell === '' || cell === null) {
          return ''; // Keep empty cells empty
        } else {
          return cell; // Return unchanged for non-numeric values
        }
      })
    );

    return result;
    
  } catch (error) {
    console.error('Add value function error:', error);
    return [['Error: ' + error.message]];
  }
}

/**
 * Test function to verify the add-in is working
 * @customfunction TEST
 * @param {string} message A test message (default: "Hello")
 * @returns {string} Confirmation message
 */
function test(message = "Hello") {
  const timestamp = new Date().toLocaleTimeString();
  return \`${message} from Cloud Functions at ${timestamp}\`;
}

// Register functions with better error handling
function registerFunctions() {
  try {
    if (typeof CustomFunctions !== 'undefined' && CustomFunctions.associate) {
      CustomFunctions.associate("ADDONE", addOne);
      CustomFunctions.associate("ADDVALUE", addValue);
      CustomFunctions.associate("TEST", test);
      console.log('✅ All custom functions registered successfully');
    } else {
      console.error('❌ CustomFunctions.associate is not available');
    }
  } catch (error) {
    console.error('❌ Error registering custom functions:', error);
  }
}

// Start the registration process when the script loads
console.log('Functions.js loaded, waiting for CustomFunctions...');
waitForCustomFunctions();