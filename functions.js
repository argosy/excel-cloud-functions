/**
 * Adds 1 to each element using Cloud Run
 * @customfunction
 * @param {number[][]} range The range to process
 * @returns {Promise<number[][]>} Processed range
 */
async function ADDONE(range) {
  try {
    console.log('ADDONE called with:', range);
    
    const response = await fetch('https://excel-add-one-function-449328337363.us-central1.run.app/add-one', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ data: range })
    });
    
    if (!response.ok) {
      throw new Error(`HTTP ${response.status}`);
    }
    
    const result = await response.json();
    console.log('Cloud response:', result);
    
    if (result.status === 'success') {
      return result.result;
    } else {
      return [["Error: " + (result.error || 'Unknown error')]];
    }
  } catch (error) {
    console.error('Function error:', error);
    return [["Error: " + error.message]];
  }
}

/**
 * Simple test function
 * @customfunction
 * @param {string} message Test message
 * @returns {string} Response
 */
function TEST(message = "Hello") {
  return "SUCCESS: " + message + " at " + new Date().toLocaleTimeString();
}

/**
 * Add custom value to each element
 * @customfunction
 * @param {number[][]} range The range to process
 * @param {number} value Value to add (default 1)
 * @returns {number[][]} Processed range
 */
function ADDVALUE(range, value = 1) {
  try {
    if (!range || !Array.isArray(range)) {
      return [["Error: Invalid range"]];
    }
    
    // Handle single cell
    if (!Array.isArray(range[0])) {
      range = [range];
    }
    
    return range.map(row => 
      row.map(cell => 
        typeof cell === 'number' ? cell + value : cell
      )
    );
  } catch (error) {
    return [["Error: " + error.message]];
  }
}