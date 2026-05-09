const fs = require('fs');

let code = fs.readFileSync('app.js', 'utf8');

// 1. Move fbKey to the top
code = code.replace(/    function fbKey\(inv\) \{\r?\n        return String\(inv\)\.replace\(\/\[\.\#\$\\\[\\\]\\\/\]\/g, '_'\);\r?\n    \}/g, '');
code = code.replace('let coStatusMap = {};', `let coStatusMap = {};\n    function fbKey(inv) {\n        return String(inv || '').replace(/[.#$\\[\\]\\/]/g, '_');\n    }`);

// 2. Replace coStatusMap[r.invoice || ''] with coStatusMap[fbKey(r.invoice)]
code = code.replace(/coStatusMap\[r\.invoice \|\| ''\]/g, 'coStatusMap[fbKey(r.invoice)]');

// 3. Fix saveCoStatus to just use fbKey directly without the dual sync hack
code = code.replace(/\/\/ Keep coStatusMap in sync under BOTH the raw key and sanitized key[\s\S]*?coStatusMap\[safeKey\] = status;/g, '');
code = code.replace(/const status = coStatusMap\[inv\]/g, 'const status = coStatusMap[fbKey(inv)]');

// 4. Replace coStatusMap[inv] with coStatusMap[fbKey(inv)] everywhere else
code = code.replace(/coStatusMap\[inv\]/g, 'coStatusMap[fbKey(inv)]');

// 5. In child_changed, fix the updateCoSingleRow call
code = code.replace(/if \(typeof updateCoSingleRow === 'function'\) updateCoSingleRow\(inv\);/g, 
`const rowData = (coCurrentRows || []).find(r => fbKey(r.invoice) === inv);\n                        if (rowData && typeof updateCoSingleRow === 'function') updateCoSingleRow(rowData.invoice);`);

// 6. Fix the row ID in buildCoRowHTML and child_changed to use fbKey(inv)
// buildCoRowHTML has id="co-row-${inv}"
code = code.replace(/id="co-row-\$\{inv\}"/g, 'id="co-row-${fbKey(inv)}"');
// In JS concatenations like 'co-row-' + inv
code = code.replace(/'co-row-'\s*\+\s*inv/g, `'co-row-' + fbKey(inv)`);

fs.writeFileSync('app.js', code);
console.log('Fixed app.js');
