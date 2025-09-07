// FORGE Audit Report Generator - Updated JavaScript File

let findingCounter = 0;
let currentFieldId = '';
let currentFieldName = '';

function openAIModal(fieldId, fieldName) {
    currentFieldId = fieldId;
    currentFieldName = fieldName;
    
    // Get current content of the field
    const field = document.querySelector(`[name="${fieldId}"]`) || document.getElementById(fieldId);
    const currentContent = field ? field.value : '';
    
    // Display current content in left panel
    const currentContentDiv = document.getElementById('currentContent');
    if (currentContent.trim()) {
        currentContentDiv.textContent = currentContent;
        currentContentDiv.classList.remove('empty');
    } else {
        currentContentDiv.textContent = `No content entered for ${fieldName}`;
        currentContentDiv.classList.add('empty');
    }
    
    // Set default prompt based on field type
    let defaultPrompt = '';
    if (currentContent) {
        defaultPrompt = `Improve and enhance the following ${fieldName.toLowerCase()}: "${currentContent}"`;
    } else {
        defaultPrompt = `Help me write a ${fieldName.toLowerCase()} for an audit report`;
    }
    
    document.getElementById('aiPromptText').value = defaultPrompt;
    
    // Reset enhancement section
    document.getElementById('enhancedContent').classList.remove('show');
    document.getElementById('enhancementPlaceholder').style.display = 'block';
    document.querySelector('.ai-apply-btn').style.display = 'none';
    
    document.getElementById('aiModal').style.display = 'block';
}

function closeAIModal() {
    document.getElementById('aiModal').style.display = 'none';
    currentFieldId = '';
    currentFieldName = '';
}

function enhanceSection() {
    const prompt = document.getElementById('aiPromptText').value;
    
    if (!prompt.trim()) {
        alert('Please enter an enhancement prompt');
        return;
    }
    
    // Call the backend Amazon Bedrock Claude API for real AI enhancement
    const field = document.querySelector(`[name="${currentFieldId}"]`) || document.getElementById(currentFieldId);
    const currentContent = field ? field.value : '';
    
    // Show processing state
    const enhanceBtn = document.querySelector('.ai-enhance-btn');
    const originalText = enhanceBtn.textContent;
    enhanceBtn.textContent = 'Enhancing...';
    enhanceBtn.disabled = true;
    
    // Hide placeholder and show enhanced content area
    document.getElementById('enhancementPlaceholder').style.display = 'none';
    
    // Map field names to section names for the backend
    const sectionMap = {
        'auditName': 'Audit Name',
        'executiveSummary': 'Executive Summary',
        'background': 'Background',
        'scopeIncluded': 'Scope',
        'scopeExcluded': 'Scope',
        'appendix': 'Appendix'
    };
    
    const section = sectionMap[currentFieldId] || currentFieldName;
    
    // Call the backend API
    fetch('/api/report/enhance', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
        },
        body: JSON.stringify({
            section: section,
            prompt: prompt,
            currentContent: currentContent || ''
        })
    })
    .then(response => {
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        return response.json();
    })
    .then(result => {
        if (result.success) {
            // Display enhanced content
            const enhancedContentDiv = document.getElementById('enhancedContent');
            enhancedContentDiv.textContent = result.enhancedContent;
            enhancedContentDiv.classList.add('show');
            
            // Show apply button
            document.querySelector('.ai-apply-btn').style.display = 'inline-block';
            
            // Show success message
            console.log(`${currentFieldName} has been enhanced with Claude AI!`);
        } else {
            throw new Error(result.error || 'Enhancement failed');
        }
    })
    .catch(error => {
        console.error('Error enhancing section:', error);
        alert(`Failed to enhance section: ${error.message}`);
        
        // Fallback to simulation for demo purposes
        let simulatedEnhancement;
        if (currentContent.trim()) {
            simulatedEnhancement = `Enhanced Version: ${currentContent}\n\n[AI has improved the clarity, structure, and professional tone of your content. The enhanced version includes better formatting, more precise language, and additional relevant details that strengthen the overall message while maintaining the original intent.]`;
        } else {
            simulatedEnhancement = `[AI Generated Content]\n\nThis is a professionally crafted ${currentFieldName.toLowerCase()} generated based on your prompt: "${prompt}"\n\nThe content follows industry best practices for audit documentation and includes appropriate structure, terminology, and level of detail expected in professional audit reports.`;
        }
        
        // Display simulated enhanced content
        const enhancedContentDiv = document.getElementById('enhancedContent');
        enhancedContentDiv.textContent = simulatedEnhancement;
        enhancedContentDiv.classList.add('show');
        
        // Show apply button
        document.querySelector('.ai-apply-btn').style.display = 'inline-block';
    })
    .finally(() => {
        enhanceBtn.textContent = originalText;
        enhanceBtn.disabled = false;
    });
}

function applyEnhancement() {
    const enhancedContent = document.getElementById('enhancedContent').textContent;
    const field = document.querySelector(`[name="${currentFieldId}"]`) || document.getElementById(currentFieldId);
    
    if (field && enhancedContent) {
        field.value = enhancedContent;
        closeAIModal();
        
        // Show success message
        alert(`${currentFieldName} has been enhanced and applied!`);
    }
}

function toggleSection(sectionId) {
    // Only toggle if not clicking on Edit/Save buttons
    if (event.target.classList.contains('edit-btn') || event.target.classList.contains('save-btn')) {
        return;
    }
    
    const content = document.getElementById(sectionId + '-content');
    const icon = content.previousElementSibling.querySelector('.toggle-icon');
    
    if (content.classList.contains('collapsed')) {
        content.classList.remove('collapsed');
        icon.textContent = '▼';
    } else {
        content.classList.add('collapsed');
        icon.textContent = '▶';
    }
}

function toggleEdit(sectionId, event) {
    event.stopPropagation();
    
    const content = document.getElementById(sectionId + '-content');
    const editBtn = event.target;
    const saveBtn = editBtn.nextElementSibling;
    
    // Enable all form elements in this section
    const formElements = content.querySelectorAll('input, textarea, select, button');
    formElements.forEach(element => {
        if (!element.classList.contains('edit-btn') && !element.classList.contains('save-btn')) {
            element.disabled = false;
        }
    });
    
    // Update button visibility
    editBtn.style.display = 'none';
    saveBtn.style.display = 'block';
    
    // Add edit mode styling
    content.classList.add('edit-mode');
}

function saveSection(sectionId, event) {
    event.stopPropagation();
    
    const content = document.getElementById(sectionId + '-content');
    const saveBtn = event.target;
    const editBtn = saveBtn.previousElementSibling;
    
    // Disable all form elements in this section
    const formElements = content.querySelectorAll('input, textarea, select, button');
    formElements.forEach(element => {
        if (!element.classList.contains('edit-btn') && !element.classList.contains('save-btn')) {
            element.disabled = true;
        }
    });
    
    // Update button visibility
    saveBtn.style.display = 'none';
    editBtn.style.display = 'block';
    
    // Remove edit mode styling
    content.classList.remove('edit-mode');
    
    // Show save confirmation
    const sectionName = content.previousElementSibling.querySelector('h2').textContent;
    alert(`${sectionName} saved successfully!`);
}

function addFinding() {
    findingCounter++;
    const findingsContainer = document.getElementById('findingsContainer');
    
    const findingDiv = document.createElement('div');
    findingDiv.className = 'finding-container';
    findingDiv.innerHTML = `
        <div class="finding-header">
            <span class="finding-number">Finding ${findingCounter}</span>
            <button type="button" class="remove-finding" onclick="removeFinding(this)">Remove</button>
        </div>
        <div class="form-group">
            <div class="field-with-ai">
                <label>Finding Short Name</label>
                <button type="button" class="enhance-ai-button" onclick="openAIModal('findingName${findingCounter}', 'Finding Short Name')">
                    ✨ Enhance with AI
                </button>
            </div>
            <input type="text" name="findingName${findingCounter}" placeholder="Enter finding short name..." disabled>
        </div>
        <div class="form-group">
            <label>Risk Rating</label>
            <select name="findingRisk${findingCounter}" disabled>
                <option value="">Select Risk Rating</option>
                <option value="Critical">Critical</option>
                <option value="High">High</option>
                <option value="Medium">Medium</option>
                <option value="Low">Low</option>
            </select>
        </div>
        <div class="form-group">
            <div class="field-with-ai">
                <label>Finding Description</label>
                <button type="button" class="enhance-ai-button" onclick="openAIModal('findingDescription${findingCounter}', 'Finding Description')">
                    ✨ Enhance with AI
                </button>
            </div>
            <textarea name="findingDescription${findingCounter}" placeholder="Enter detailed finding description..." disabled></textarea>
        </div>
        <div class="form-group">
            <div class="field-with-ai">
                <label>Recommendation(s)</label>
                <button type="button" class="enhance-ai-button" onclick="openAIModal('findingRecommendation${findingCounter}', 'Recommendations')">
                    ✨ Enhance with AI
                </button>
            </div>
            <textarea name="findingRecommendation${findingCounter}" placeholder="Enter recommendations..." disabled></textarea>
        </div>

        <div class="form-group">
            <div class="action-items-header" onclick="toggleActionItems(${findingCounter})">
                <h4>Action Items</h4>
                <div class="action-items-toggle" id="actionToggle${findingCounter}">▼</div>
            </div>
            <div class="action-items-content" id="actionItemsContent${findingCounter}">
                <div class="action-items-container" id="actionItems${findingCounter}">
                    <!-- Action items will be added here -->
                </div>
            </div>
            <button type="button" class="add-action-item" onclick="addActionItem(${findingCounter})" disabled>+ Add Action Item</button>
        </div>
    `;
    
    findingsContainer.appendChild(findingDiv);
    addActionItem(findingCounter); // Add initial action item
}

function addActionItem(findingNumber) {
    const actionItemsContainer = document.getElementById(`actionItems${findingNumber}`);
    const actionItemCount = actionItemsContainer.children.length + 1;
    
    const actionItemDiv = document.createElement('div');
    actionItemDiv.className = 'action-item-container';
    actionItemDiv.innerHTML = `
        <div class="action-item-header">
            <span class="action-item-number">Action Item ${actionItemCount}</span>
            <button type="button" class="remove-action-item" onclick="removeActionItem(this, ${findingNumber})">Remove</button>
        </div>
        <div class="form-group">
            <div class="field-with-ai">
                <label>Action Item Short Name</label>
                <button type="button" class="enhance-ai-button" onclick="openAIModal('actionItemName${findingNumber}_${actionItemCount}', 'Action Item Short Name')">
                    ✨ Enhance with AI
                </button>
            </div>
            <input type="text" name="actionItemName${findingNumber}_${actionItemCount}" placeholder="Enter action item short name..." disabled>
        </div>
        <div class="form-group">
            <div class="field-with-ai">
                <label>Action Item Description</label>
                <button type="button" class="enhance-ai-button" onclick="openAIModal('actionItem${findingNumber}_${actionItemCount}', 'Action Item Description')">
                    ✨ Enhance with AI
                </button>
            </div>
            <textarea name="actionItem${findingNumber}_${actionItemCount}" placeholder="Enter action item description..." disabled></textarea>
        </div>
        <div class="form-group">
            <label>Due Date</label>
            <input type="date" name="actionItemDueDate${findingNumber}_${actionItemCount}" disabled>
        </div>
        <div class="form-group">
            <label>Owner (L8+) Team</label>
            <input type="text" name="actionItemOwner${findingNumber}_${actionItemCount}" placeholder="Enter owner and team..." disabled>
        </div>
    `;
    
    actionItemsContainer.appendChild(actionItemDiv);
}

function toggleActionItems(findingNumber) {
    const content = document.getElementById(`actionItemsContent${findingNumber}`);
    const toggle = document.getElementById(`actionToggle${findingNumber}`);
    
    if (content.classList.contains('collapsed')) {
        content.classList.remove('collapsed');
        toggle.classList.remove('collapsed');
        toggle.textContent = '▼';
    } else {
        content.classList.add('collapsed');
        toggle.classList.add('collapsed');
        toggle.textContent = '▶';
    }
}

function removeActionItem(button, findingNumber) {
    button.closest('.action-item-container').remove();
    // Renumber remaining action items
    const actionItemsContainer = document.getElementById(`actionItems${findingNumber}`);
    const actionItems = actionItemsContainer.children;
    for (let i = 0; i < actionItems.length; i++) {
        const numberSpan = actionItems[i].querySelector('.action-item-number');
        numberSpan.textContent = `Action Item ${i + 1}`;
    }
}

function removeFinding(button) {
    button.closest('.finding-container').remove();
}

function clearForm() {
    if (confirm('Are you sure you want to clear all form data?')) {
        document.getElementById('auditForm').reset();
        document.getElementById('findingsContainer').innerHTML = '';
        findingCounter = 0;
    }
}

function previewReport() {
    // Get form data
    const auditNameInput = document.getElementById('auditName') || document.querySelector('input[name="auditName"]');
    const auditName = auditNameInput ? auditNameInput.value || '[AUDIT NAME]' : '[AUDIT NAME]';
    
    const executiveSummaryInput = document.getElementById('executiveSummary') || document.querySelector('textarea[name="executiveSummary"]');
    const backgroundInput = document.getElementById('background') || document.querySelector('textarea[name="background"]');
    const scopeIncludedInput = document.getElementById('scopeIncluded') || document.querySelector('textarea[name="scopeIncluded"]');
    const scopeExcludedInput = document.getElementById('scopeExcluded') || document.querySelector('textarea[name="scopeExcluded"]');
    
    const executiveSummary = executiveSummaryInput ? executiveSummaryInput.value || 'No executive summary provided.' : 'No executive summary provided.';
    const background = backgroundInput ? backgroundInput.value || 'No background information provided.' : 'No background information provided.';
    const scopeIncluded = scopeIncludedInput ? scopeIncludedInput.value || 'No scope inclusions specified.' : 'No scope inclusions specified.';
    const scopeExcluded = scopeExcludedInput ? scopeExcludedInput.value || 'No scope exclusions specified.' : 'No scope exclusions specified.';
    
    // Calculate line numbers dynamically
    let currentLine = 1;
    let contentHTML = '';
    
    // Helper function to add a line with inline line number
    function addLine(content, isHeader = false, isComment = false) {
        let lineClass = '';
        if (isHeader) lineClass = 'header-line';
        if (isComment) lineClass = 'comment-line';
        
        contentHTML += `<div class="content-line ${lineClass}"><span class="line-num">${currentLine}</span><span class="line-content">${content}</span></div>`;
        currentLine++;
    }
    
    let previewHTML = `
        <style>
            .preview-report {
                font-family: 'Calibri', Arial, sans-serif;
                line-height: 1.2;
                color: #000;
                max-width: 100%;
                margin: 0;
                padding: 20px;
                background: white;
            }
            .content-line {
                display: flex;
                margin: 0;
                padding: 1px 0;
            }
            .line-num {
                width: 25px;
                font-size: 10px;
                color: #666;
                text-align: right;
                margin-right: 10px;
                flex-shrink: 0;
            }
            .line-content {
                flex: 1;
                font-size: 11px;
                margin: 0;
            }
            .header-line .line-content {
                font-weight: bold;
                text-transform: uppercase;
                color: #90EE90;
                font-size: 14px;
            }
            .header-line.audit-name .line-content {
                color: #FF8C00;
                font-size: 18px;
            }
            .comment-line .line-content {
                font-style: italic;
                color: #666;
                font-size: 10px;
            }
        </style>
        <div class="preview-report">`;
    
    // Line 1: Audit Name
    addLine(auditName, true, false);
    contentHTML = contentHTML.replace('header-line', 'header-line audit-name');
    
    // Line 2: Empty space
    addLine('&nbsp;');
    
    // Line 3: Executive Summary
    addLine('EXECUTIVE SUMMARY', true);
    
    // Line 4: Comment
    addLine('// Include reason for audit, high level scope, high level findings //', false, true);
    
    // Lines 5+: Executive Summary Content (dynamic)
    const execSummaryLines = executiveSummary.split('\n');
    execSummaryLines.forEach(line => {
        addLine(line || '&nbsp;');
    });
    
    // One empty space after executive summary
    addLine('&nbsp;');
    
    // Background section
    addLine('BACKGROUND', true);
    
    // Background comment
    addLine('// Include context and background of the teams, process, systems //', false, true);
    
    // Background content (dynamic)
    const backgroundLines = background.split('\n');
    backgroundLines.forEach(line => {
        addLine(line || '&nbsp;');
    });
    
    // One empty space after background
    addLine('&nbsp;');
    
    // Scope section
    addLine('SCOPE', true);
    
    // Scope introduction
    addLine('CSI gained an understanding of the Critical Vendor processes, systems, and controls applicable to the 12-month period from Jan to Dec 2022. CSI\'s scope included:');
    
    // Scope included content (dynamic)
    const scopeIncludedLines = scopeIncluded.split('\n');
    scopeIncludedLines.forEach((line, index) => {
        if (line.trim()) {
            addLine(`${index + 1}. ${line}`);
        }
    });
    
    // Empty space before out of scope
    addLine('&nbsp;');
    
    // Out of scope introduction
    addLine('For this inspection, the following areas were out of scope:');
    
    // Scope excluded content (dynamic)
    const scopeExcludedLines = scopeExcluded.split('\n');
    scopeExcludedLines.forEach((line, index) => {
        if (line.trim()) {
            addLine(`${index + 1}. ${line}`);
        }
    });
    
    // Add final spacing
    addLine('&nbsp;');
    
    // Findings and Action Items section
    addLine('FINDINGS AND ACTION ITEMS', true);
    addLine('Findings and action items resulting from the inspection are detailed in the table below. All metrics reported in the findings are based on information related to the XX-month period ended MM/DD/YY, unless otherwise noted. Action items will be tracked in GRC.');
    
    // Build findings table content
    previewHTML += contentHTML;
    
    // Add findings table
    previewHTML += `
            <table style="width: 100%; border-collapse: collapse; margin: 10px 0; font-size: 10px; border: 1px solid #000;">
                <thead>
                    <tr style="background-color: #f0f0f0;">
                        <th style="border: 1px solid #000; padding: 6px; text-align: center; font-weight: bold;">Finding, Name, Rating, Ref</th>
                        <th style="border: 1px solid #000; padding: 6px; text-align: center; font-weight: bold;">Finding, Recommendation(s) & Business Owner Response</th>
                        <th style="border: 1px solid #000; padding: 6px; text-align: center; font-weight: bold;">Action Item(s)</th>
                        <th style="border: 1px solid #000; padding: 6px; text-align: center; font-weight: bold;">Due Date</th>
                        <th style="border: 1px solid #000; padding: 6px; text-align: center; font-weight: bold;">Owner (L8+) Team</th>
                    </tr>
                </thead>
                <tbody>`;
    
    // Get all findings from the form
    const findings = document.querySelectorAll('.finding-container');
    
    findings.forEach((finding, findingIndex) => {
        const findingNumber = findingIndex + 1;
        const findingNumberFormatted = findingNumber < 10 ? `0${findingNumber}` : `${findingNumber}`;
        
        // Get finding data
        const nameInput = finding.querySelector('input[name*="findingName"]');
        const riskSelect = finding.querySelector('select[name*="findingRisk"]');
        const descriptionTextarea = finding.querySelector('textarea[name*="findingDescription"]');
        const recommendationTextarea = finding.querySelector('textarea[name*="findingRecommendation"]');
        
        const name = nameInput ? nameInput.value || `Finding ${findingNumber}` : `Finding ${findingNumber}`;
        const risk = riskSelect ? riskSelect.value || 'Not Set' : 'Not Set';
        const description = descriptionTextarea ? descriptionTextarea.value || 'No description provided' : 'No description provided';
        const recommendation = recommendationTextarea ? recommendationTextarea.value || 'No recommendations provided' : 'No recommendations provided';
        
        // Get action items for this finding
        const actionItems = finding.querySelectorAll('.action-item-container');
        
        // Build action items content
        let actionItemsContent = '';
        let actionItemsOwners = '';
        let actionItemsDueDates = '';
        
        if (actionItems.length === 0) {
            actionItemsContent = 'No action items specified';
            actionItemsOwners = 'Not assigned';
            actionItemsDueDates = '-';
        } else {
            const actionItemsList = [];
            const ownersList = [];
            const dueDatesList = [];
            
            actionItems.forEach((actionItem, actionIndex) => {
                const actionNameInput = actionItem.querySelector('input[name*="actionName"]');
                const actionDescInput = actionItem.querySelector('textarea[name*="actionDescription"]');
                const actionDueDateInput = actionItem.querySelector('input[name*="actionDueDate"]');
                const actionOwnerInput = actionItem.querySelector('input[name*="actionOwner"]');
                
                const actionName = actionNameInput ? actionNameInput.value || `Action Item ${actionIndex + 1}` : `Action Item ${actionIndex + 1}`;
                const actionDescription = actionDescInput ? actionDescInput.value || 'No description' : 'No description';
                const actionDueDate = actionDueDateInput ? actionDueDateInput.value || 'MM/DD/YY' : 'MM/DD/YY';
                const actionOwner = actionOwnerInput ? actionOwnerInput.value || 'Not assigned' : 'Not assigned';
                
                actionItemsList.push(`${findingNumber}.${actionIndex + 1}. ${actionName}<br>${actionDescription}`);
                ownersList.push(actionOwner);
                dueDatesList.push(actionDueDate);
            });
            
            actionItemsContent = actionItemsList.join('<br><br>');
            actionItemsOwners = ownersList.join('<br>');
            actionItemsDueDates = dueDatesList.join('<br>');
        }
        
        // Create single row for this finding
        previewHTML += `
            <tr>
                <td style="border: 1px solid #000; padding: 6px; text-align: center; font-weight: bold; width: 15%; vertical-align: top;">
                    ${findingNumberFormatted}<br>
                    ${name}<br><br>
                    <span style="color: #FF0000; font-weight: bold;">${risk}</span>
                </td>
                <td style="border: 1px solid #000; padding: 6px; width: 50%; vertical-align: top;">
                    <strong>${description}</strong><br><br>
                    <strong>Recommendations</strong><br>
                    ${recommendation}<br><br>
                    <strong>Business Owner Response</strong><br>
                    We agree with the finding.
                </td>
                <td style="border: 1px solid #000; padding: 6px; width: 20%; vertical-align: top;">
                    ${actionItemsContent}
                </td>
                <td style="border: 1px solid #000; padding: 6px; width: 10%; text-align: center; vertical-align: top;">${actionItemsDueDates}</td>
                <td style="border: 1px solid #000; padding: 6px; width: 15%; text-align: center; vertical-align: top;">${actionItemsOwners}</td>
            </tr>`;
    });
    
    previewHTML += `
                </tbody>
            </table>`;
    
    // Add Appendix if provided
    const appendixInput = document.getElementById('appendix') || document.querySelector('textarea[name="appendix"]');
    const appendix = appendixInput ? appendixInput.value : '';
    
    if (appendix && appendix.trim()) {
        previewHTML += `
            <div style="margin-top: 30px;">
                <h2 style="color: #FF8C00; font-size: 14px; font-weight: bold; text-transform: uppercase; margin: 10px 0;">APPENDIX A -- RATING</h2>
                <p style="font-size: 11px; margin: 0;">${appendix}</p>
            </div>`;
    }
    
    previewHTML += `</div>`;
    
    document.getElementById('previewContent').innerHTML = previewHTML;
    document.getElementById('previewModal').style.display = 'block';
}

function closePreview() {
    document.getElementById('previewModal').style.display = 'none';
}

function exportReport() {
    const formData = new FormData(document.getElementById('auditForm'));
    const reportData = {};
    
    // Collect all form data
    for (let [key, value] of formData.entries()) {
        reportData[key] = value;
    }
    
    // Create downloadable JSON file
    const dataStr = JSON.stringify(reportData, null, 2);
    const dataBlob = new Blob([dataStr], {type: 'application/json'});
    const url = URL.createObjectURL(dataBlob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `forge_audit_report_${new Date().toISOString().split('T')[0]}.json`;
    link.click();
    
    alert('Report exported successfully!');
}

// Initialize with one finding
document.addEventListener('DOMContentLoaded', function() {
    addFinding();
});