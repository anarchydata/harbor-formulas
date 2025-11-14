/**
 * Helper function to create snippet template from function signature
 * This must be defined before it's used in completion providers and Tab handlers
 */
export function createFunctionSnippet(func) {
  const signature = func.signature;
  // Extract function name and parameters from signature like "IF(logical, value_if_true, value_if_false)"
  const match = signature.match(/^(\w+)\s*\((.*)\)$/);
  if (!match) {
    // If signature doesn't match, return simple template
    return func.name + '($1)';
  }
  
  const params = match[2];
  if (!params || params.trim() === '') {
    // No parameters
    return func.name + '()';
  }
  
  // Parse parameters (handle ... for variable args)
  const paramList = params.split(',').map(p => p.trim());
  let placeholderIndex = 1;
  const snippetParams = paramList.map((param, index) => {
    if (param.endsWith('...')) {
      // Variable arguments - use $1 for first, rest can be added
      return '${' + placeholderIndex++ + ':' + param.replace('...', '') + '}';
    } else {
      return '${' + placeholderIndex++ + ':' + param + '}';
    }
  });
  
  return func.name + '(' + snippetParams.join(', ') + ')';
}

