// Script to extract ALL HyperFormula functions with their metadata
import { HyperFormula } from 'hyperformula';
import * as fs from 'fs';

// Initialize HyperFormula
const hf = HyperFormula.buildEmpty({
  licenseKey: 'gpl-v3',
  useArrayArithmetic: true
});

// Access the function registry through the instance
// The registry is stored in _functionRegistry
const functionRegistry = hf._functionRegistry;

// Get all registered function IDs
const allFunctionIds = functionRegistry.getRegisteredFunctionIds();

console.log(`Found ${allFunctionIds.length} functions`);

// Get function translations using getRegisteredFunctionNames
let functionTranslations = {};
try {
  functionTranslations = hf.getRegisteredFunctionNames() || {};
} catch (e) {
  console.warn('Could not get function translations:', e.message);
}

console.log(`Found ${allFunctionIds.length} functions`);

// Extract function metadata
const functions = [];

for (const functionId of allFunctionIds) {
  const translation = functionTranslations[functionId];
  
  // Translation structure: { name, description, parameters }
  let description = '';
  let signature = '';
  
  if (translation) {
    description = translation.description || '';
    // Parameters might be in translation.parameters or we build from name
    if (translation.parameters) {
      signature = translation.parameters;
    } else if (translation.name && translation.name !== functionId) {
      // Sometimes the name contains the signature
      signature = translation.name;
    }
  }
  
  // If no signature from translation, try to build from metadata
  if (!signature) {
    // Try to get plugin metadata
    try {
      const plugin = hf.getFunctionPlugin(functionId);
      if (plugin && plugin.implementedFunctions && plugin.implementedFunctions[functionId]) {
        const metadata = plugin.implementedFunctions[functionId];
        if (metadata.parameters && metadata.parameters.length > 0) {
          const params = metadata.parameters.map((param) => {
            const paramName = param.name || 
              (param.argumentType === 'NUMBER' ? 'number' :
               param.argumentType === 'STRING' ? 'text' :
               param.argumentType === 'BOOLEAN' ? 'logical' :
               param.argumentType === 'RANGE' ? 'range' :
               param.argumentType === 'ANY' ? 'value' :
               param.argumentType === 'SCALAR' ? 'value' :
               param.argumentType === 'INTEGER' ? 'number' : 'value');
            if (param.optional) {
              return `[${paramName}]`;
            }
            return paramName;
          }).join(', ');
          
          // Handle repeatLastArgs
          if (metadata.repeatLastArgs) {
            const lastParam = params.split(', ').pop();
            signature = `${functionId}(${params}, ${lastParam}...)`;
          } else {
            signature = `${functionId}(${params})`;
          }
        }
      }
    } catch (e) {
      // If we can't get metadata, use default
    }
  }
  
  // Fallback to simple signature
  if (!signature) {
    signature = `${functionId}()`;
  }
  
  functions.push({
    name: functionId,
    signature: signature,
    description: description
  });
}

// Sort alphabetically
functions.sort((a, b) => a.name.localeCompare(b.name));

// Output as JSON
const output = {
  totalFunctions: functions.length,
  extractedAt: new Date().toISOString(),
  functions: functions
};

// Write to file
fs.writeFileSync('hyperformula-functions.json', JSON.stringify(output, null, 2));
console.log(`\nExtracted ${functions.length} functions to hyperformula-functions.json`);

// Also create a simpler format for Monaco Editor
const monacoFunctions = functions.map(func => ({
  name: func.name,
  signature: func.signature,
  description: func.description
}));

fs.writeFileSync('hyperformula-functions-monaco.js', 
  `// Auto-generated list of ALL HyperFormula functions\n` +
  `// Total: ${monacoFunctions.length} functions\n` +
  `export const hyperFormulaFunctions = ${JSON.stringify(monacoFunctions, null, 2)};`
);

console.log(`Created hyperformula-functions-monaco.js for Monaco Editor integration`);

