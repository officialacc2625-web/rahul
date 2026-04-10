import re
with open('app.js', 'r', encoding='utf-8') as f:
    text = f.read()

# I need to update renderCustomersOSGPage filters
# Currently it has:
# const selRBM = $('coRBM').value;
# const selBDM = $('coBDM').value;
# const selProduct = $('coProduct').value;
# const selBranch = $('coBranch').value;

old_filter_lines = """            const selRBM = $('coRBM').value;
            const selBDM = $('coBDM').value;
            const selProduct = $('coProduct').value;
            const selBranch = $('coBranch').value;"""

new_filter_lines = """            const selBrand = document.getElementById('coBrand') ? document.getElementById('coBrand').value : '';
            const selRBM = $('coRBM').value;
            const selBDM = $('coBDM').value;
            const selProduct = $('coProduct').value;
            const selBranch = $('coBranch').value;
            const selDate = document.getElementById('coDateFilter') ? document.getElementById('coDateFilter').value : '';
            const selStatusFilter = document.getElementById('coStatusFilter') ? document.getElementById('coStatusFilter').value : '';"""

text = text.replace(old_filter_lines, new_filter_lines)

# Populate dropdown options for Brand
# It currently has:
# const rbmSet = [...new Set(productData.map(r => r.rbm).filter(Boolean))].sort();

old_dropdown_lines = """            const rbmSet = [...new Set(productData.map(r => r.rbm).filter(Boolean))].sort();
            const bdmSet = [...new Set(productData.map(r => r.bdm).filter(Boolean))].sort();
            const prodSet = [...new Set(productData.map(r => r.product).filter(Boolean))].sort();
            const branchSet = [...new Set(productData.map(r => r.branch).filter(Boolean))].sort();

            $('coRBM').innerHTML = '<option value="">All RBMs</option>' + rbmSet.map(r => `<option value="${r}" ${r === selRBM ? 'selected' : ''}>${r}</option>`).join('');
            $('coBDM').innerHTML = '<option value="">All BDMs</option>' + bdmSet.map(b => `<option value="${b}" ${b === selBDM ? 'selected' : ''}>${b}</option>`).join('');
            $('coProduct').innerHTML = '<option value="">All Products</option>' + prodSet.map(p => `<option value="${p}" ${p === selProduct ? 'selected' : ''}>${p}</option>`).join('');
            $('coBranch').innerHTML = '<option value="">All Branches</option>' + branchSet.map(b => `<option value="${b}" ${b === selBranch ? 'selected' : ''}>${b}</option>`).join('');"""

new_dropdown_lines = """            const brandSet = [...new Set(productData.map(r => r.brand).filter(Boolean))].sort();
            const rbmSet = [...new Set(productData.map(r => r.rbm).filter(Boolean))].sort();
            const bdmSet = [...new Set(productData.map(r => r.bdm).filter(Boolean))].sort();
            const prodSet = [...new Set(productData.map(r => r.product).filter(Boolean))].sort();
            const branchSet = [...new Set(productData.map(r => r.branch).filter(Boolean))].sort();

            if($('coBrand')) $('coBrand').innerHTML = '<option value="">All Brands</option>' + brandSet.map(b => `<option value="${b}" ${b === selBrand ? 'selected' : ''}>${b}</option>`).join('');
            $('coRBM').innerHTML = '<option value="">All RBMs</option>' + rbmSet.map(r => `<option value="${r}" ${r === selRBM ? 'selected' : ''}>${r}</option>`).join('');
            $('coBDM').innerHTML = '<option value="">All BDMs</option>' + bdmSet.map(b => `<option value="${b}" ${b === selBDM ? 'selected' : ''}>${b}</option>`).join('');
            $('coProduct').innerHTML = '<option value="">All Products</option>' + prodSet.map(p => `<option value="${p}" ${p === selProduct ? 'selected' : ''}>${p}</option>`).join('');
            $('coBranch').innerHTML = '<option value="">All Branches</option>' + branchSet.map(b => `<option value="${b}" ${b === selBranch ? 'selected' : ''}>${b}</option>`).join('');"""

text = text.replace(old_dropdown_lines, new_dropdown_lines)

# Apply Date and Brand filters
old_filt_logic = """            let filtP = productData;
            if (selRBM) filtP = filtP.filter(r => r.rbm === selRBM);
            if (selBDM) filtP = filtP.filter(r => r.bdm === selBDM);
            if (selProduct) filtP = filtP.filter(r => r.product === selProduct);
            if (selBranch) filtP = filtP.filter(r => r.branch === selBranch);"""

new_filt_logic = """            let filtP = productData;
            if (selBrand) filtP = filtP.filter(r => r.brand === selBrand);
            if (selRBM) filtP = filtP.filter(r => r.rbm === selRBM);
            if (selBDM) filtP = filtP.filter(r => r.bdm === selBDM);
            if (selProduct) filtP = filtP.filter(r => r.product === selProduct);
            if (selBranch) filtP = filtP.filter(r => r.branch === selBranch);
            
            if (selDate) {
                filtP = filtP.filter(r => {
                    let dStr = '';
                    if (r.invoiceDate) { dStr = new Date(r.invoiceDate).toISOString().split('T')[0]; }
                    else if (r.time) { dStr = new Date(r.time).toISOString().split('T')[0]; }
                    return dStr === selDate;
                });
            }"""

text = text.replace(old_filt_logic, new_filt_logic)

with open('app.js', 'w', encoding='utf-8') as f:
    f.write(text)
print('Phase 2 filter logic applied.')
