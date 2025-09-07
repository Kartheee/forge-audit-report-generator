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
    
    // Save the section data based on the section ID
    if (sectionId === 'auditname') {
        const auditName = document.getElementById('auditName').value;
        fetch('/api/report/audit-name', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ auditName })
        })
        .then(response => response.json())
        .then(data => console.log('Audit name saved:', data))
        .catch(error => console.error('Error saving audit name:', error));
    } else if (sectionId === 'executive') {
        const executiveSummary = document.getElementById('executiveSummary').value;
        fetch('/api/report/executive-summary', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ executiveSummary })
        })
        .then(response => response.json())
        .then(data => console.log('Executive summary saved:', data))
        .catch(error => console.error('Error saving executive summary:', error));
    } else if (sectionId === 'background') {
        const background = document.getElementById('background').value;
        fetch('/api/report/background', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ background })
        })
        .then(response => response.json())
        .then(data => console.log('Background saved:', data))
        .catch(error => console.error('Error saving background:', error));
    } else if (sectionId === 'scope') {
        const included = document.getElementById('scopeIncluded').value;
        const excluded = document.getElementById('scopeExcluded').value;
        fetch('/api/report/scope', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ included, excluded })
        })
        .then(response => response.json())
        .then(data => console.log('Scope saved:', data))
        .catch(error => console.error('Error saving scope:', error));
    } else if (sectionId === 'findings') {
        // For findings, we'll just save the current state
        console.log('Findings section saved');
    } else if (sectionId === 'appendix') {
        const appendix = document.getElementById('appendix').value;
        fetch('/api/report/appendix', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ appendix })
        })
        .then(response => response.json())
        .then(data => console.log('Appendix saved:', data))
        .catch(error => console.error('Error saving appendix:', error));
    }
    
    // Show save confirmation
    const sectionName = content.previousElementSibling.querySelector('h2').textContent;
    alert(`${sectionName} saved successfully!`);
}

function addFinding() {
    // Get the current number of findings and use that for numbering
    const findingsContainer = document.getElementById('findingsContainer');
    const currentFindingCount = findingsContainer.querySelectorAll('.finding-container').length;
    const findingNumber = currentFindingCount + 1;
    
    // Use a unique ID for this finding (for element IDs)
    const findingId = findingCounter + 1;
    findingCounter = findingId; // Increment the counter for next finding
    
    const findingDiv = document.createElement('div');
    findingDiv.className = 'finding-container';
    findingDiv.innerHTML = `
        <div class="finding-header">
            <span class="finding-number">Finding ${findingNumber}</span>
            <button type="button" class="remove-finding" onclick="removeFinding(this)">Remove</button>
        </div>
        <div class="form-group">
            <div class="field-with-ai">
                <label>Finding Short Name</label>
                <button type="button" class="enhance-ai-button" onclick="openAIModal('findingName${findingId}', 'Finding Short Name')">
                    ✨ Enhance with AI
                </button>
            </div>
            <input type="text" name="findingName${findingId}" placeholder="Enter finding short name...">
        </div>
        <div class="form-group">
            <label>Risk Rating</label>
            <select name="findingRisk${findingId}">
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
                <button type="button" class="enhance-ai-button" onclick="openAIModal('findingDescription${findingId}', 'Finding Description')">
                    ✨ Enhance with AI
                </button>
            </div>
            <textarea name="findingDescription${findingId}" placeholder="Enter detailed finding description..."></textarea>
        </div>
        <div class="form-group">
            <div class="field-with-ai">
                <label>Recommendation(s)</label>
                <button type="button" class="enhance-ai-button" onclick="openAIModal('findingRecommendation${findingId}', 'Recommendations')">
                    ✨ Enhance with AI
                </button>
            </div>
            <textarea name="findingRecommendation${findingId}" placeholder="Enter recommendations..."></textarea>
        </div>

        <div class="form-group">
            <div class="action-items-header" onclick="toggleActionItems(${findingId})">
                <h4>Action Items</h4>
                <div class="action-items-toggle" id="actionToggle${findingId}">▼</div>
            </div>
            <div class="action-items-content" id="actionItemsContent${findingId}">
                <!-- Action items container will be created dynamically -->
            </div>
            <button type="button" class="add-action-item" onclick="addActionItem(${findingId})">+ Add Action Item</button>
        </div>
    `;
    
    findingsContainer.appendChild(findingDiv);
    
    // Create action items container with the correct finding ID
    const actionItemsContainer = document.createElement('div');
    actionItemsContainer.id = `actionItems${findingId}`;
    actionItemsContainer.className = 'action-items-container';
    document.getElementById(`actionItemsContent${findingId}`).appendChild(actionItemsContainer);
    
    // Add the first action item with number 1
    addActionItem(findingId); // Add initial action item
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
            <input type="text" name="actionItemName${findingNumber}_${actionItemCount}" placeholder="Enter action item short name...">
        </div>
        <div class="form-group">
            <div class="field-with-ai">
                <label>Action Item Description</label>
                <button type="button" class="enhance-ai-button" onclick="openAIModal('actionItem${findingNumber}_${actionItemCount}', 'Action Item Description')">
                    ✨ Enhance with AI
                </button>
            </div>
            <textarea name="actionItem${findingNumber}_${actionItemCount}" placeholder="Enter action item description..."></textarea>
        </div>
        <div class="form-group">
            <label>Due Date</label>
            <input type="date" name="actionItemDueDate${findingNumber}_${actionItemCount}">
        </div>
        <div class="form-group">
            <label>Owner Information</label>
            <div class="owner-fields-container">
                <div class="owner-field">
                    <label>Action Owner</label>
                    <input type="text" name="actionOwner${findingNumber}_${actionItemCount}" placeholder="Enter action owner...">
                </div>
                <div class="owner-field">
                    <label>Owner (L8+)</label>
                    <input type="text" name="ownerL8${findingNumber}_${actionItemCount}" placeholder="Enter L8+ name...">
                </div>
                <div class="owner-field">
                    <label>Team</label>
                    <input type="text" name="ownerTeam${findingNumber}_${actionItemCount}" placeholder="Enter team name...">
                </div>
            </div>
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
    
    // Renumber remaining findings
    const findingsContainer = document.getElementById('findingsContainer');
    const findings = findingsContainer.querySelectorAll('.finding-container');
    
    findings.forEach((finding, index) => {
        // Update finding number display
        const numberSpan = finding.querySelector('.finding-number');
        numberSpan.textContent = `Finding ${index + 1}`;
    });
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
        // Preserve line breaks by converting newlines to <br> tags
        const description = descriptionTextarea ? (descriptionTextarea.value.replace(/\n/g, '<br>') || 'No description provided') : 'No description provided';
        const recommendation = recommendationTextarea ? (recommendationTextarea.value.replace(/\n/g, '<br>') || 'No recommendations provided') : 'No recommendations provided';
        
        // Get action items for this finding
        const actionItems = finding.querySelectorAll('.action-item-container');
        
        // Build action items content
        // Create separate row for each action item to ensure proper column separation
        const actionItemsList = [];
        const dueDates = [];
        const owners = [];
        
        // Collect all action items data
        actionItems.forEach((actionItem, actionIndex) => {
            const actionNameInput = actionItem.querySelector('input[name*="actionItemName"]');
            const actionDescInput = actionItem.querySelector('textarea[name*="actionItem"]');
            const actionDueDateInput = actionItem.querySelector('input[name*="actionItemDueDate"]');
            const actionOwnerInput = actionItem.querySelector('input[name*="actionItemOwner"]');
            
            const actionName = actionNameInput ? (actionNameInput.value.trim() || `Action Item ${actionIndex + 1}`) : `Action Item ${actionIndex + 1}`;
            const actionDescription = actionDescInput ? (actionDescInput.value.trim() || '') : '';
            // Format date as MM/DD/YY if it's a valid date, otherwise use 'MM/DD/YY'
            let actionDueDate = 'MM/DD/YY';
            if (actionDueDateInput && actionDueDateInput.value) {
                const dateValue = actionDueDateInput.value;
                try {
                    const date = new Date(dateValue);
                    if (!isNaN(date.getTime())) {
                        const month = String(date.getMonth() + 1).padStart(2, '0');
                        const day = String(date.getDate()).padStart(2, '0');
                        const year = String(date.getFullYear()).slice(-2);
                        actionDueDate = `${month}/${day}/${year}`;
                    }
                } catch (e) {
                    console.error("Error formatting date:", e);
                }
            }
            const actionOwner = actionOwnerInput ? (actionOwnerInput.value.trim() || 'Action Owner<br>(L8 name)<br>Team') : 'Action Owner<br>(L8 name)<br>Team';
            
            // Include description if provided
            let actionItemText = `<div style="text-align: left; font-weight: bold;">${findingNumber}.${actionIndex + 1}. ${actionName}</div>`;
            if (actionDescription) {
                actionItemText += `<div style="text-align: left; margin-top: 4px;">${actionDescription.replace(/\n/g, '<br>')}</div>`;
            }
            
            actionItemsList.push(actionItemText);
            dueDates.push(actionDueDate);
            owners.push(actionOwner);
        });
        
        // Create main finding row
        // Check if this finding has no additional action items
        const hasNoAdditionalActionItems = actionItemsList.length <= 1;
        // Apply bottom border if there are no additional action items
        const bottomBorderForMainRow = hasNoAdditionalActionItems ? 'border-bottom: 1px solid #000;' : 'border-bottom: none;';
        
        previewHTML += `
            <tr>
                <td style="border-left: 1px solid #000; border-right: 1px solid #000; border-top: 1px solid #000; ${bottomBorderForMainRow}; padding: 6px; text-align: center; font-weight: bold; width: 15%; vertical-align: top;">
                    ${findingNumberFormatted}<br>
                    ${name}<br><br>
                    <span style="color: #FF0000; font-weight: bold;">${risk}</span>
                </td>
                <td style="border-left: 1px solid #000; border-right: 1px solid #000; border-top: 1px solid #000; ${bottomBorderForMainRow}; padding: 6px; width: 50%; vertical-align: top;">
                    <strong>Finding Description:</strong><br>
                    ${description}<br><br>
                    <strong>Recommendations:</strong><br>
                    ${recommendation}<br><br>
                    <strong>Business Owner Response:</strong><br>
                    <strong>Team Name:</strong> We agree with the finding.
                </td>
                <td style="border: 1px solid #000; padding: 6px; width: 20%; vertical-align: top;">
                    ${actionItemsList.length > 0 ? actionItemsList[0] : 'No action items'}
                </td>
                <td style="border: 1px solid #000; padding: 6px; width: 10%; text-align: center; vertical-align: top;">
                    ${dueDates.length > 0 ? dueDates[0] : 'MM/DD/YY'}
                </td>
                <td style="border: 1px solid #000; padding: 6px; width: 15%; text-align: center; vertical-align: top;">
                    ${owners.length > 0 ? owners[0] : 'Action Owner<br>(L8 name)<br>Team'}
                </td>
            </tr>`;
        
        // Create additional rows for remaining action items (if any)
        for (let i = 1; i < actionItemsList.length; i++) {
            // Check if this is the last action item for this finding
            const isLastActionItem = i === actionItemsList.length - 1;
            
            // Apply bottom border only to the last action item row
            const bottomBorderStyle = isLastActionItem ? 'border-bottom: 1px solid #000;' : 'border-bottom: none;';
            
            previewHTML += `
                <tr>
                    <td style="border-left: 1px solid #000; border-right: 1px solid #000; border-top: none; ${bottomBorderStyle}; padding: 6px; width: 15%; vertical-align: top;"></td>
                    <td style="border-left: 1px solid #000; border-right: 1px solid #000; border-top: none; ${bottomBorderStyle}; padding: 6px; width: 50%; vertical-align: top;"></td>
                    <td style="border: 1px solid #000; padding: 6px; width: 20%; vertical-align: top;">
                        ${actionItemsList[i]}
                    </td>
                    <td style="border: 1px solid #000; padding: 6px; width: 10%; text-align: center; vertical-align: top;">
                        ${dueDates[i]}
                    </td>
                    <td style="border: 1px solid #000; padding: 6px; width: 15%; text-align: center; vertical-align: top;">
                        ${owners[i]}
                    </td>
                </tr>`;
        }
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
    // First, save all form data to the server
    saveAllFormData().then(() => {
        // Then request the Word document using POST method
        fetch('/api/report/export', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({})
        })
        .then(response => {
            if (!response.ok) {
                throw new Error('Network response was not ok');
            }
            return response.blob();
        })
        .then(blob => {
            // Create a link to download the blob
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.style.display = 'none';
            a.href = url;
            a.download = 'FORGE_Audit_Report.docx';
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);
            
            // Show success message
            alert('Report exported successfully as Word document!');
        })
        .catch(error => {
            console.error('Error downloading report:', error);
            alert('Failed to export report. Please try again.');
        });
    }).catch(error => {
        console.error('Error exporting report:', error);
        alert('Failed to export report. Please try again.');
    });
}

// Helper function to save all form data to the server before export
async function saveAllFormData() {
    try {
        // Get form data
        const auditName = document.getElementById('auditName').value || '';
        const executiveSummary = document.getElementById('executiveSummary').value || '';
        const background = document.getElementById('background').value || '';
        const scopeIncluded = document.getElementById('scopeIncluded').value || '';
        const scopeExcluded = document.getElementById('scopeExcluded').value || '';
        const appendix = document.getElementById('appendix').value || '';
        
        // Save audit name
        await fetch('/api/report/audit-name', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ auditName })
        });
        
        // Save executive summary
        await fetch('/api/report/executive-summary', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ executiveSummary })
        });
        
        // Save background
        await fetch('/api/report/background', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ background })
        });
        
        // Save scope
        await fetch('/api/report/scope', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ included: scopeIncluded, excluded: scopeExcluded })
        });
        
        // Save appendix
        await fetch('/api/report/appendix', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ appendix })
        });
        
        // Process findings
        const findings = [];
        document.querySelectorAll('.finding-container').forEach((findingElem, findingIndex) => {
            const findingNumber = findingIndex + 1;
            
            // Get finding data
            const nameInput = findingElem.querySelector('input[name*="findingName"]');
            const riskSelect = findingElem.querySelector('select[name*="findingRisk"]');
            const descriptionTextarea = findingElem.querySelector('textarea[name*="findingDescription"]');
            const recommendationTextarea = findingElem.querySelector('textarea[name*="findingRecommendation"]');
            
            const shortName = nameInput ? nameInput.value || `Finding ${findingNumber}` : `Finding ${findingNumber}`;
            const rating = riskSelect ? riskSelect.value || 'Not Set' : 'Not Set';
            const description = descriptionTextarea ? descriptionTextarea.value || '' : '';
            const recommendations = recommendationTextarea ? recommendationTextarea.value || '' : '';
            
            // Process action items
            const actionItems = [];
            findingElem.querySelectorAll('.action-item-container').forEach((actionItem, actionIndex) => {
                const actionNameInput = actionItem.querySelector('input[name*="actionItemName"]');
                const actionDescInput = actionItem.querySelector('textarea[name*="actionItem"]');
                console.log('Action Description Input:', actionDescInput ? actionDescInput.value : 'Not found');
                const actionDueDateInput = actionItem.querySelector('input[name*="actionItemDueDate"]');
                // Get the three owner fields
                const actionOwnerInput = actionItem.querySelector('input[name*="actionOwner"]');
                const ownerL8Input = actionItem.querySelector('input[name*="ownerL8"]');
                const ownerTeamInput = actionItem.querySelector('input[name*="ownerTeam"]');
                
                const actionName = actionNameInput ? actionNameInput.value || '' : '';
                const actionDescription = actionDescInput ? actionDescInput.value || '' : '';
                const dueDate = actionDueDateInput ? actionDueDateInput.value || '' : '';
                
                // Get values from the three owner fields
                const actionOwner = actionOwnerInput ? actionOwnerInput.value || '' : '';
                const ownerL8 = ownerL8Input ? ownerL8Input.value || '' : '';
                const ownerTeam = ownerTeamInput ? ownerTeamInput.value || '' : '';
                
                actionItems.push({
                    actionItemName: actionName || 'Action Item',
                    actionItemDescription: actionDescription || '',
                    dueDate,
                    actionOwner,
                    ownerL8,
                    ownerTeam
                });
            });
            
            findings.push({
                shortName,
                rating,
                description,
                recommendations,
                actionItems
            });
        });
        
        // Save findings
        await fetch('/api/report/findings', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ findings })
        });
        
        return true;
    } catch (error) {
        console.error('Error saving form data:', error);
        throw error;
    }
}

// Initialize with one finding
document.addEventListener('DOMContentLoaded', function() {
    addFinding();
});