global.window = {};
const diagnostics = require('./utils/diagnostics.js');
const sample = `IFERROR(
CONCATENATE(
"TopName=",
INDEX($A$2:$A$6,
MATCH(
MAX(
$C$2:$C$6*$D$2:$D$6,
$C$2:$C$6*$D$2:$D$6,
0
)
),
" | TotalRevenue=",
TEXT(SUMPRODUCT($C$2:$C$6,$D$2:$D$6),"0.00"),
" | AvgPriceQty>=10=",
TEXT(AVERAGEIF($C$2:$C$6,">=10",$D$2:$D$6),"0.00"),
" | OrdersAfter2025-01-01=",
TEXT(COUNTIF($E$2:$E$6,">="&DATE(2025,1,1)),"0"),
" | MaxQty=",
TEXT(MAX($C$2:$C$6),"0"),
)
"Error"
)

IF(), value, value)
// HELLo I am a comment
`;
const monaco = { MarkerSeverity: { Error: 'error', Warning: 'warning' } };
const provider = diagnostics.createDiagnosticsProvider(monaco);
const model = { getValue: () => sample };
const result = provider.provideDiagnostics(model);
console.log(JSON.stringify(result, null, 2));
