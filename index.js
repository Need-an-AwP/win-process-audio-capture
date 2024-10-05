const path = require('path');
const fs = require('fs');

let binding;
const addonPath = path.resolve(__dirname, 'build/Release/test_addon.node');

try {
  if (fs.existsSync(addonPath)) {
    binding = require(addonPath);
  } else {
    throw new Error('Prebuilt addon not found');
  }
} catch (err) {
  throw new Error(`Failed to load the addon: ${err.message}`);
}

module.exports = binding;
