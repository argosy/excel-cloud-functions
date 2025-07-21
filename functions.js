// Minimal functions.js for testing
console.log('📦 Functions.js loaded');

/**
 * Simple test function
 * @customfunction
 * @param {string} message Test message
 * @returns {string} Response message
 */
function TEST(message = "Hello") {
  console.log('🚀 TEST function called with:', message);
  return "SUCCESS: " + message + " at " + new Date().toLocaleTimeString();
}

/**
 * Add one to a number
 * @customfunction
 * @param {number} value Input number
 * @returns {number} Number plus one
 */
function ADDONE(value) {
  console.log('🔢 ADDONE function called with:', value);
  if (typeof value === 'number') {
    return value + 1;
  }
  return "Error: Not a number";
}

// Wait for CustomFunctions to be available
function registerFunctions() {
  console.log('🔍 Checking for CustomFunctions...');
  
  if (typeof CustomFunctions !== 'undefined' && CustomFunctions.associate) {
    console.log('✅ CustomFunctions found! Registering...');
    
    try {
      CustomFunctions.associate("TEST", TEST);
      CustomFunctions.associate("ADDONE", ADDONE);
      console.log('🎉 Functions registered successfully!');
    } catch (error) {
      console.error('❌ Registration error:', error);
    }
  } else {
    console.log('⏳ CustomFunctions not ready, trying again...');
    setTimeout(registerFunctions, 500);
  }
}

// Start registration when script loads
registerFunctions();