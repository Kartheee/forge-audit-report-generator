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
    // Get audit name directly from the input element
    const auditNameInput = document.getElementById('auditName') || document.querySelector('input[name="auditName"]');
    const auditName = auditNameInput ? auditNameInput.value || '[Audit Name]' : '[Audit Name]';
    
    // Get all other form data directly from DOM elements
    const executiveSummaryInput = document.getElementById('executiveSummary') || document.querySelector('textarea[name="executiveSummary"]');
    const backgroundInput = document.getElementById('background') || document.querySelector('textarea[name="background"]');
    const scopeIncludedInput = document.getElementById('scopeIncluded') || document.querySelector('textarea[name="scopeIncluded"]');
    const scopeExcludedInput = document.getElementById('scopeExcluded') || document.querySelector('textarea[name="scopeExcluded"]');
    const appendixInput = document.getElementById('appendix') || document.querySelector('textarea[name="appendix"]');
    
    const executiveSummary = executiveSummaryInput ? executiveSummaryInput.value || 'No executive summary provided.' : 'No executive summary provided.';
    const background = backgroundInput ? backgroundInput.value || 'No background information provided.' : 'No background information provided.';
    const scopeIncluded = scopeIncludedInput ? scopeIncludedInput.value || 'No scope inclusions specified.' : 'No scope inclusions specified.';
    const scopeExcluded = scopeExcludedInput ? scopeExcludedInput.value || 'No scope exclusions specified.' : 'No scope exclusions specified.';
    const appendix = appendixInput ? appendixInput.value : '';
    
    let previewHTML = `
        <style>
            .preview-report {
                font-family: 'Calibri', Arial, sans-serif;
                line-height: 1.15;
                color: #000;
                max-width: 100%;
                margin: 0;
                padding: 0;
                padding-left: 50px;
                position: relative;
            }
            .line-number {
                position: absolute;
                left: 10px;
                color: #666;
                font-size: 10px;
                width: 30px;
                text-align: right;
            }
            .preview-report h1 {
                font-size: 18px;
                font-weight: bold;
                text-align: left;
                margin-bottom: 20px;
                color: #FF8C00;
                position: relative;
            }
            .preview-report h2 {
                font-size: 14px;
                font-weight: bold;
                margin-top: 20px;
                margin-bottom: 10px;
                text-transform: uppercase;
                position: relative;
            }
            .preview-report h2.audit-name {
                color: #FF8C00;
            }
            .preview-report h2.green-section {
                color: #90EE90;
            }
            .preview-report h3 {
                font-size: 12px;
                font-weight: bold;
                margin-top: 15px;
                margin-bottom: 8px;
                color: #000;
            }
            .preview-report p {
                margin-bottom: 10px;
                text-align: justify;
                font-size: 11px;
            }
            .preview-report .subheading {
                font-style: italic;
                color: #666;
                font-size: 10px;
                margin-bottom: 8px;
            }
            .findings-table {
                width: 100%;
                border-collapse: collapse;
                margin: 15px 0;
                font-size: 10px;
            }
            .findings-table th, .findings-table td {
                border: 1px solid #000;
                padding: 6px;
                vertical-align: top;
                text-align: left;
            }
            .findings-table th {
                background-color: #f0f0f0;
                font-weight: bold;
                text-align: center;
                font-size: 10px;
            }
            .finding-name {
                font-weight: bold;
                text-align: center;
                width: 15%;
            }
            .finding-content {
                width: 50%;
            }
            .action-items {
                width: 20%;
            }
            .due-date {
                width: 10%;
                text-align: center;
            }
            .owner {
                width: 15%;
                text-align: center;
            }
            .rating-high {
                color: #FF0000;
                font-weight: bold;
            }
            .rating-medium {
                color: #FF0000;
                font-weight: bold;
            }
            .rating-low {
                color: #FF0000;
                font-weight: bold;
            }
            .rating-critical {
                color: #FF0000;
                font-weight: bold;
            }
            .scope-list {
                margin-left: 15px;
            }
            .scope-list ol {
                margin: 8px 0;
            }
            .scope-list li {
                margin-bottom: 4px;
            }
            .finding-number {
                font-weight: normal;
                color: #000;
                background: transparent;
            }
            .scope-content {
                white-space: pre-line;
            }
        </style>
        <div class="preview-report">
            <div class="line-number">1</div>
            <h1 class="audit-name">${auditName}</h1>
            <div class="line-number">2</div>
            <p>&nbsp;</p>
            <div class="line-number">3</div>
            <h2 class="green-section">EXECUTIVE SUMMARY</h2>
            <div class="line-number">4</div>
            <div class="subheading"><em>// Include reason for audit, high level scope, high level findings //</em></div>
            <div class="line-number">5</div>
            <p>${executiveSummary}</p>
            <div class="line-number">6</div>
            <p>&nbsp;</p>
            <div class="line-number">7</div>
            <p>&nbsp;</p>
            <div class="line-number">8</div>
            <p>&nbsp;</p>
            <div class="line-number">9</div>
            <p>&nbsp;</p>
            <div class="line-number">10</div>
            <p>&nbsp;</p>
            <div class="line-number">11</div>
            <h2 class="green-section">BACKGROUND</h2>
            <div class="line-number">12</div>
            <div class="subheading"><em>// Include context and background of the teams, process, systems //</em></div>
            <div class="line-number">13</div>
            <p>${background}</p>
            <div class="line-number">14</div>
            <p>&nbsp;</p>
            <div class="line-number">15</div>
            <p>&nbsp;</p>
            <div class="line-number">16</div>
            <p>&nbsp;</p>
            <div class="line-number">17</div>
            <p>&nbsp;</p>
            <div class="line-number">18</div>
            <p>&nbsp;</p>
            <div class="line-number">19</div>
            <p>&nbsp;</p>
            <div class="line-number">20</div>
            <p>&nbsp;</p>
            <div class="line-number">21</div>
            <p>&nbsp;</p>
            <div class="line-number">22</div>
            <p>&nbsp;</p>
            <div class="line-number">23</div>
            <p>&nbsp;</p>
            <div class="line-number">24</div>
            <p>&nbsp;</p>
            <div class="line-number">25</div>
            <h2 class="green-section">SCOPE</h2>
            <div class="line-number">26</div>
            <p>CSI gained an understanding of the Critical Vendor processes, systems, and controls applicable to the 12-month period from Jan to Dec 2022. CSI's scope included:</p>
            <div class="line-number">27</div>
            <p>1. ${scopeIncluded.split('\n')[0] || 'XX'}</p>
            <div class="line-number">28</div>
            <p>2. ${scopeIncluded.split('\n')[1] || 'XX'}</p>
            <div class="line-number">29</div>
            <p>3. ${scopeIncluded.split('\n')[2] || 'xx'}</p>
            <div class="line-number">30</div>
            <p>&nbsp;</p>
            <div class="line-number">31</div>
            <p>&nbsp;</p>
            <div class="line-number">32</div>
            <p>For this inspection, the following areas were out of scope:</p>
            <div class="line-number">33</div>
            <p>1. ${scopeExcluded.split('\n')[0] || 'XX'}</p>
            <div class="line-number">34</div>
            <p>2. ${scopeExcluded.split('\n')[1] || 'XX'}</p>
            <div class="line-number">35</div>
            <p>3. ${scopeExcluded.split('\n')[2] || 'xx'}</p>
            <div class="line-number">36</div>
            <p>&nbsp;</p>
            
            <table class="findings-table">
                <thead>
                    <tr>
                        <th>Finding, Name, Rating, Ref</th>
                        <th>Finding, Recommendation(s) & Business Owner Response</th>
                        <th>Action Item(s)</th>
                        <th>Due Date</th>
                        <th>Owner (L8+) Team</th>
                    </tr>
                </thead>
                <tbody>
    `;
    
    const findings = document.querySelectorAll('.finding-container');
    let findingNumber = 1;
    
    findings.forEach((finding) => {
        // Get finding data directly from DOM elements with better selectors
        const nameInput = finding.querySelector('input[name*="findingName"]');
        const riskSelect = finding.querySelector('select[name*="findingRisk"]');
        const descriptionTextarea = finding.querySelector('textarea[name*="findingDescription"]');
        const recommendationTextarea = finding.querySelector('textarea[name*="findingRecommendation"]');
        
        const name = nameInput ? nameInput.value || `Finding ${findingNumber}` : `Finding ${findingNumber}`;
        const risk = riskSelect ? riskSelect.value || 'Not Set' : 'Not Set';
        const description = descriptionTextarea ? descriptionTextarea.value || 'No description provided' : 'No description provided';
        const recommendation = recommendationTextarea ? recommendationTextarea.value || 'No recommendations provided' : 'No recommendations provided';
        
        let ratingClass = 'rating-medium';
        switch(risk.toLowerCase()) {
            case 'critical': ratingClass = 'rating-critical'; break;
            case 'high': ratingClass = 'rating-high'; break;
            case 'medium': ratingClass = 'rating-medium'; break;
            case 'low': ratingClass = 'rating-low'; break;
        }
        
        const actionItems = finding.querySelectorAll('.action-item-container');
        let isFirstActionItem = true;
        
        if (actionItems.length === 0) {
            // No action items - create a single row
            previewHTML += `
                <tr>
                    <td class="finding-name">
                        <span class="finding-number">${String(findingNumber).padStart(2, '0')}</span><br><br>
                        ${name}<br><br>
                        <span class="${ratingClass}"><strong>${risk}</strong></span>
                    </td>
                    <td class="finding-content">
                        <strong>${description}</strong><br><br>
                        <strong>Recommendations</strong><br>
                        ${recommendation}<br><br>
                        <strong>Business Owner Response</strong><br>
                        Response pending.
                    </td>
                    <td class="action-items">No action items specified</td>
                    <td class="due-date">-</td>
                    <td class="owner">Not assigned</td>
                </tr>
            `;
        } else {
            actionItems.forEach((actionItem, actionIndex) => {
                // Get action item data directly from DOM elements
                const actionNameInput = actionItem.querySelector('input[name*="actionItemName"]');
                const actionDescriptionTextarea = actionItem.querySelector('textarea[name*="actionItem"]:not([name*="Name"])');
                const actionDueDateInput = actionItem.querySelector('input[type="date"]');
                const actionOwnerInput = actionItem.querySelector('input[name*="actionItemOwner"]');
                
                const actionName = actionNameInput ? actionNameInput.value || `Action Item ${actionIndex + 1}` : `Action Item ${actionIndex + 1}`;
                const actionDescription = actionDescriptionTextarea ? actionDescriptionTextarea.value || 'No description' : 'No description';
                const actionDueDate = actionDueDateInput ? actionDueDateInput.value || 'Not set' : 'Not set';
                const actionOwner = actionOwnerInput ? actionOwnerInput.value || 'Not assigned' : 'Not assigned';
                
                if (isFirstActionItem) {
                    // First action item row includes finding details
                    previewHTML += `
                        <tr>
                            <td class="finding-name">
                                <span class="finding-number">${String(findingNumber).padStart(2, '0')}</span><br><br>
                                ${name}<br><br>
                                <span class="${ratingClass}"><strong>${risk}</strong></span>
                            </td>
                            <td class="finding-content">
                                <strong>${description}</strong><br><br>
                                <strong>Recommendations</strong><br>
                                ${recommendation}<br><br>
                                <strong>Business Owner Response</strong><br>
                                We agree with the finding.
                            </td>
                            <td class="action-items">
                                ${findingNumber}.${actionIndex + 1}. ${actionName}<br>
                                ${actionDescription}
                            </td>
                            <td class="due-date">${actionDueDate !== 'Not set' ? actionDueDate : 'MM/DD/YY'}</td>
                            <td class="owner">${actionOwner}</td>
                        </tr>
                    `;
                    isFirstActionItem = false;
                } else {
                    // Subsequent action item rows
                    previewHTML += `
                        <tr>
                            <td class="finding-name"></td>
                            <td class="finding-content"></td>
                            <td class="action-items">
                                ${findingNumber}.${actionIndex + 1}. ${actionName}<br>
                                ${actionDescription}
                            </td>
                            <td class="due-date">${actionDueDate !== 'Not set' ? actionDueDate : 'MM/DD/YY'}</td>
                            <td class="owner">${actionOwner}</td>
                        </tr>
                    `;
                }
            });
        }
        findingNumber++;
    });
    
    previewHTML += `
                </tbody>
            </table>
            
            <h2>Appendix A -- Rating</h2>
            <p>Findings are rated based on the table below.</p>
            
            <table class="findings-table">
                <thead>
                    <tr>
                        <th style="width: 20%;">Rating</th>
                        <th style="width: 80%;">Description</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td><strong>High (H)</strong></td>
                        <td><strong>Significant</strong> impact and/or <strong>high</strong> likelihood of significant impact to occur OR <strong>High</strong> impact and/or <strong>significant</strong> likelihood of high impact to occur.</td>
                    </tr>
                    <tr>
                        <td><strong>Medium (M)</strong></td>
                        <td><strong>High</strong> impact potential but <strong>moderate</strong> likelihood to occur OR <strong>Medium</strong> impact and/or <strong>low</strong> likelihood of medium impact to occur.</td>
                    </tr>
                    <tr>
                        <td><strong>Low (L)</strong></td>
                        <td>Identified <strong>improvement opportunity</strong>.</td>
                    </tr>
                </tbody>
            </table>
            
            <p><em>*The importance rating is similar to Internal Audit's approach and based on the facts and circumstances of the individual item/event under inspection, impact (quantitative/qualitative) and likelihood that the event will occur (e.g. control failure).</em></p>
            
            ${appendix ? `
                <h2>Appendix B -- Additional Information</h2>
                <p>${appendix}</p>
            ` : ''}
        </div>
    `;
    
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