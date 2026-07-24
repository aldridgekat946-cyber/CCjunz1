const fs = require('fs');
let code = fs.readFileSync('utils/packingProcessor.ts', 'utf8');

code = code.replace(
  /let headerRowIndex = 0;\s+let xxColIdx = -1;/,
  `let headerRowIndex = 0;
  let xxColIdx = -1;
  let qtyColIdx = -1;`
);

code = code.replace(
  /xxColIdx = findColIndex\(rowValues, \['XX CODE', 'XX编码', 'XX 编码'\]\);/,
  `if (xxColIdx === -1) xxColIdx = findColIndex(rowValues, ['XX CODE', 'XX编码', 'XX 编码']);
    if (qtyColIdx === -1) qtyColIdx = findColIndex(rowValues, ['数量', 'QTY', 'Quantity']);`
);

code = code.replace(
  /const xxCodeRaw = String\(getCellValue\(row\.getCell\(xxColIdx\)\) \|\| ""\)\.trim\(\);\n\s+if \(!xxCodeRaw\) return;\n\n\s+const baseCodeStr = xxCodeRaw;/,
  `const xxCodeRaw = String(getCellValue(row.getCell(xxColIdx)) || "").trim();
    if (!xxCodeRaw) return;

    let initialQuantity = '';
    if (qtyColIdx !== -1) {
        initialQuantity = String(getCellValue(row.getCell(qtyColIdx)) || "").trim();
    }

    const baseCodeStr = xxCodeRaw;`
);

fs.writeFileSync('utils/packingProcessor.ts', code);
