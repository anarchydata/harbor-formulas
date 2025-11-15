import { HyperFormula } from 'hyperformula';

const hf = HyperFormula.buildEmpty({ licenseKey: 'gpl-v3', useArrayArithmetic: true });
const sheetName = 'Sheet1';
hf.addSheet(sheetName);
const sheetId = hf.getSheetId(sheetName);
const data = [
  ['A','B','C','D','E'],
  ['Dept1','','10',100,new Date('2025-02-01')],
  ['Dept2','','20',200,new Date('2024-12-31')],
  ['Dept3','','30',300,new Date('2025-03-15')],
  ['Dept4','','40',400,new Date('2025-04-10')],
  ['Dept5','','50',500,new Date('2025-05-20')],
];
hf.setSheetContent(sheetId, data);
const formula = `=IFERROR(
  CONCATENATE(
    "TopName=",
    INDEX($A$2:$A$6,
      MATCH(
        MAX($C$2:$C$6*$D$2:$D$6),
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
    TEXT(MAX($C$2:$C$6),"0")
  ),
  "Error"
)`;
const addr = { sheet: sheetId, col: 0, row: 8 };
hf.setCellContents(addr, [[formula]]);
console.log('Value:', hf.getCellValue(addr));
