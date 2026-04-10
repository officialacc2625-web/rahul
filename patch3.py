import re

with open('app.js', 'r', encoding='utf-8') as f:
    text = f.read()

# 1. Add follow-up logic
old_toggle_fns = """        window.toggleCoInterest = function(inv, status) {
            if (!coStatusMap[inv]) coStatusMap[inv] = { callStatus: null, interest: null, calledBy: null, remarks: '' };
            coStatusMap[inv].interest = status === "" ? null : status;
            saveCoStatus(inv);
            updateCoSingleRow(inv);
            updateCoStatsInPlace();
        };"""

new_toggle_fns = """        window.toggleCoInterest = function(inv, status) {
            if (!currentCaller) {
                alert('Please select your name first.');
                updateCoSingleRow(inv);
                return;
            }
            if (!coStatusMap[inv]) coStatusMap[inv] = { callStatus: null, interest: null, calledBy: null, remarks: '', followUpDate: '' };
            coStatusMap[inv].interest = status === "" ? null : status;
            coStatusMap[inv].calledBy = currentCaller; // assign row since they interacted
            saveCoStatus(inv);
            updateCoSingleRow(inv);
            updateCoStatsInPlace();
        };
        
        window.setCoFollowUpDate = function(inv, dateStr) {
            if (!coStatusMap[inv]) coStatusMap[inv] = { callStatus: null, interest: null, calledBy: null, remarks: '', followUpDate: '' };
            coStatusMap[inv].followUpDate = dateStr;
            saveCoStatus(inv);
            updateCoSingleRow(inv);
        };"""

text = text.replace(old_toggle_fns, new_toggle_fns)


# 2. Update toggleCoCall missing assign
old_toggle_call = """        window.toggleCoCall = function(inv, status) {
            if (!currentCaller && status !== "") {
                alert('Please select your name (Harmiya / Aswathi / Shikha) at the top of the page before logging a call.');
                // Revert UI to match state since we rejected it
                updateCoSingleRow(inv);
                return;
            }
            if (!coStatusMap[inv]) coStatusMap[inv] = { callStatus: null, interest: null, calledBy: null, remarks: '' };
            coStatusMap[inv].callStatus = status === "" ? null : status;
            coStatusMap[inv].calledBy   = coStatusMap[inv].callStatus ? currentCaller : null;
            saveCoStatus(inv);
            updateCoSingleRow(inv);
            updateCoStatsInPlace();
        };"""

new_toggle_call = """        window.toggleCoCall = function(inv, status) {
            if (!currentCaller && status !== "") {
                alert('Please select your name (Harmiya / Aswathi / Shikha) at the top of the page before logging a call.');
                updateCoSingleRow(inv);
                return;
            }
            if (!coStatusMap[inv]) coStatusMap[inv] = { callStatus: null, interest: null, calledBy: null, remarks: '' };
            coStatusMap[inv].callStatus = status === "" ? null : status;
            if (status !== "") coStatusMap[inv].calledBy = currentCaller;
            saveCoStatus(inv);
            updateCoSingleRow(inv);
            updateCoStatsInPlace();
        };"""
        
text = text.replace(old_toggle_call, new_toggle_call)

with open('app.js', 'w', encoding='utf-8') as f:
    f.write(text)
