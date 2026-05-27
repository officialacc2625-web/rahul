const fs = require('fs');
let text = fs.readFileSync('c:/Users/rahul_myg/Downloads/New folder/app.js', 'utf8');

// Define fmtShortHtml
const fmtShortHtmlDef = `
    function fmtShortHtml(n) {
        return \`<span title="₹\${Number(n).toLocaleString('en-IN')}" style="cursor:help; border-bottom:1px dotted #94a3b8;">\${fmtShort(n)}</span>\`;
    }
`;

if (!text.includes('function fmtShortHtml')) {
    text = text.replace('function formatNumber(n)', fmtShortHtmlDef + '\n    function formatNumber(n)');
}

text = text.replace(/\$\{fmtShort\(/g, '${fmtShortHtml(');

fs.writeFileSync('c:/Users/rahul_myg/Downloads/New folder/app.js', text, 'utf8');
console.log('Done');
