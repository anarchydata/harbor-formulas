// Script to scrape ALL HyperFormula functions from their documentation
import * as https from 'https';
import * as fs from 'fs';

// The documentation page with all functions
const docsUrl = 'https://hyperformula.handsontable.com/guide/built-in-functions.html';

console.log('Fetching HyperFormula documentation...');

https.get(docsUrl, (res) => {
  let data = '';
  
  res.on('data', (chunk) => {
    data += chunk;
  });
  
  res.on('end', () => {
    // Extract function names and descriptions from HTML
    // Look for function definitions in the HTML
    const functionMatches = data.match(/<h[23][^>]*>([A-Z_]+)<\/h[23]>/g) || [];
    const functions = [];
    
    functionMatches.forEach((match, index) => {
      const functionName = match.match(/>([A-Z_]+)</)?.[1];
      if (functionName) {
        // Try to extract description that follows
        const nextMatch = functionMatches[index + 1];
        const startIdx = data.indexOf(match);
        const endIdx = nextMatch ? data.indexOf(nextMatch) : data.length;
        const section = data.substring(startIdx, endIdx);
        
        // Extract description from paragraph after heading
        const descMatch = section.match(/<p[^>]*>(.*?)<\/p>/);
        const description = descMatch ? descMatch[1].replace(/<[^>]+>/g, '').trim() : '';
        
        functions.push({
          name: functionName,
          description: description
        });
      }
    });
    
    console.log(`Found ${functions.length} functions from documentation`);
    
    // Merge with existing function list
    const existingFunctions = JSON.parse(fs.readFileSync('hyperformula-functions.json', 'utf8'));
    
    // Create a map of descriptions
    const descMap = {};
    functions.forEach(f => {
      if (f.description) {
        descMap[f.name] = f.description;
      }
    });
    
    // Update existing functions with descriptions
    existingFunctions.functions.forEach(func => {
      if (descMap[func.name]) {
        func.description = descMap[func.name];
      }
    });
    
    // Save updated file
    fs.writeFileSync('hyperformula-functions.json', JSON.stringify(existingFunctions, null, 2));
    console.log('Updated hyperformula-functions.json with descriptions');
    
    // Update Monaco file
    const monacoFunctions = existingFunctions.functions.map(func => ({
      name: func.name,
      signature: func.signature,
      description: func.description
    }));
    
    fs.writeFileSync('hyperformula-functions-monaco.js', 
      `// Auto-generated list of ALL HyperFormula functions\n` +
      `// Total: ${monacoFunctions.length} functions\n` +
      `export const hyperFormulaFunctions = ${JSON.stringify(monacoFunctions, null, 2)};`
    );
    
    console.log('Updated hyperformula-functions-monaco.js');
  });
}).on('error', (err) => {
  console.error('Error fetching documentation:', err);
});




