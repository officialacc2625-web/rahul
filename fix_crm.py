import re

with open('app.js', 'r', encoding='utf-8') as f:
    code = f.read()

# 1. Move fbKey to the top
code = re.sub(r'    function fbKey\(inv\) \{\n        return String\(inv\)\.replace\(\/\[\.\#\$\\\\[\\\\]\\\/\]\/g, \'_' + r"'\);\n    \}\n?", '', code)
code = code.replace('let coStatusMap = {};', "let coStatusMap = {};\n    function fbKey(inv) {\n        return String(inv || '').replace(/[.#$\\[\\]\\/]/g, '_');\n    }")

# 2. Replace coStatusMap[r.invoice || ''] with coStatusMap[fbKey(r.invoice)]
code = re.sub(r"coStatusMap\[r\.invoice \|\| ''\]", "coStatusMap[fbKey(r.invoice)]", code)

# 3. Fix saveCoStatus to just use fbKey directly without the dual sync hack
code = re.sub(r"// Keep coStatusMap in sync under BOTH the raw key and sanitized key[\s\S]*?coStatusMap\[safeKey\] = status;", "", code)
code = re.sub(r"const status = coStatusMap\[inv\]", "const status = coStatusMap[fbKey(inv)]", code)

# 4. Replace coStatusMap[inv] with coStatusMap[fbKey(inv)] everywhere else
code = re.sub(r"coStatusMap\[inv\]", "coStatusMap[fbKey(inv)]", code)

# 5. In child_changed, fix the updateCoSingleRow call
# Because safeKey is returned by Firebase, we need to find the raw invoice in coCurrentRows to generate the HTML properly
code = re.sub(
    r"if \(typeof updateCoSingleRow === 'function'\) updateCoSingleRow\(inv\);",
    "const rowData = (coCurrentRows || []).find(r => fbKey(r.invoice) === inv);\n                        if (rowData && typeof updateCoSingleRow === 'function') updateCoSingleRow(rowData.invoice);",
    code
)

# 6. Fix the row ID in buildCoRowHTML and child_changed to use fbKey(inv)
code = re.sub(r'id="co-row-\$\{inv\}"', 'id="co-row-${fbKey(inv)}"', code)
code = re.sub(r"'co-row-'\s*\+\s*inv", "'co-row-' + fbKey(inv)", code)

with open('app.js', 'w', encoding='utf-8') as f:
    f.write(code)

print("Fixed app.js using Python")
