const express = require('express');
const cors = require('cors');
const multer = require('multer');
const { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, AlignmentType, Header, Footer, BorderStyle, PageNumber, TableLayoutType } = require('docx');

// Custom function to convert inches to twips (1 inch = 1440 twips)
function convertInchesToTwip(inches) {
  return Math.round(inches * 1440);
}
const { BedrockRuntimeClient, InvokeModelCommand } = require('@aws-sdk/client-bedrock-runtime');
const path = require('path');
require('dotenv').config({ path: path.join(__dirname, '.env') });

// Load environment variables manually since dotenv isn't working properly
const fs = require('fs');
const envPath = path.join(__dirname, '.env');
if (fs.existsSync(envPath)) {
  const envContent = fs.readFileSync(envPath, 'utf8');
  console.log('Raw .env content length:', envContent.length);
  console.log('Raw .env content (first 100 chars):', envContent.substring(0, 100));
  
  // Remove all hidden characters and normalize
  const cleanContent = envContent
    .replace(/^\uFEFF/, '')  // Remove BOM
    .replace(/^\uFFFE/, '')  // Remove BOM
    .replace(/[\u200B-\u200D\uFEFF]/g, '') // Remove zero-width spaces
    .replace(/\r\n/g, '\n')  // Normalize line endings
    .replace(/\r/g, '\n');   // Normalize line endings
  
  const lines = cleanContent.split('\n');
  console.log('Number of lines:', lines.length);
  
  lines.forEach((line, index) => {
    const trimmedLine = line.trim();
    if (trimmedLine && !trimmedLine.startsWith('#')) {
      const [key, value] = trimmedLine.split('=');
      if (key && value) {
        const cleanKey = key.trim();
        const cleanValue = value.trim();
        process.env[cleanKey] = cleanValue;
        console.log(`Set env var: ${cleanKey} = ${cleanValue}`);
      }
    }
  });
}

// Debug: Check what was loaded
console.log('After manual loading:');
console.log('AWS_REGION:', process.env.AWS_REGION);
console.log('AWS_ACCESS_KEY_ID:', process.env.AWS_ACCESS_KEY_ID ? 'SET' : 'NOT SET');
console.log('AWS_SECRET_ACCESS_KEY:', process.env.AWS_SECRET_ACCESS_KEY ? 'SET' : 'NOT SET');

const app = express();
const PORT = process.env.PORT || 3000;

// Middleware
app.use(cors());
app.use(express.json());
app.use(express.static('public'));

// Add a simple health check route
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Amazon Bedrock configuration
console.log('AWS Environment Variables Status:');
console.log('AWS_REGION:', process.env.AWS_REGION);
console.log('AWS_ACCESS_KEY_ID:', process.env.AWS_ACCESS_KEY_ID ? 'SET' : 'NOT SET');
console.log('AWS_SECRET_ACCESS_KEY:', process.env.AWS_SECRET_ACCESS_KEY ? 'SET' : 'NOT SET');

// Create Bedrock client with environment variables (SECURE)
let bedrockClient = null;
if (process.env.AWS_ACCESS_KEY_ID && process.env.AWS_SECRET_ACCESS_KEY) {
  bedrockClient = new BedrockRuntimeClient({
    region: process.env.AWS_REGION || 'us-east-1',
    credentials: {
      accessKeyId: process.env.AWS_ACCESS_KEY_ID,
      secretAccessKey: process.env.AWS_SECRET_ACCESS_KEY,
    },
  });
  console.log('AWS Bedrock client initialized with environment variables');
} else {
  console.log('AWS credentials not found, using fallback AI enhancement');
}

// In-memory storage for report data
let currentReport = {
  auditName: '',
  executiveSummary: '',
  background: '',
  scope: {
    included: '',
    excluded: ''
  },
  findings: [],
  appendix: ''
};

// Routes
app.get('/api/report', (req, res) => {
  res.json(currentReport);
});

app.post('/api/report/audit-name', (req, res) => {
  const { auditName } = req.body;
  currentReport.auditName = auditName;
  res.json({ success: true, message: 'Audit Name updated' });
});

app.post('/api/report/executive-summary', (req, res) => {
  const { executiveSummary } = req.body;
  currentReport.executiveSummary = executiveSummary;
  res.json({ success: true, message: 'Executive Summary updated' });
});

app.post('/api/report/background', (req, res) => {
  const { background } = req.body;
  currentReport.background = background;
  res.json({ success: true, message: 'Background updated' });
});

app.post('/api/report/scope', (req, res) => {
  const { included, excluded } = req.body;
  currentReport.scope = { included, excluded };
  res.json({ success: true, message: 'Scope updated' });
});

app.post('/api/report/findings', (req, res) => {
  const { findings } = req.body;
  currentReport.findings = findings;
  res.json({ success: true, message: 'Findings updated' });
});

app.post('/api/report/appendix', (req, res) => {
  const { appendix } = req.body;
  currentReport.appendix = appendix;
  res.json({ success: true, message: 'Appendix updated' });
});

app.post('/api/report/enhance', async (req, res) => {
  try {
    const { section, prompt, currentContent } = req.body;
    
    if (!section || !prompt) {
      return res.status(400).json({ success: false, error: 'Section and prompt are required' });
    }

    // Try AWS Bedrock first, fallback to simple enhancement
    let enhancedContent;
    
    if (bedrockClient) {
      try {
        const systemPrompt = `You are an expert audit analyst. Enhance the following ${section} section based on the user's prompt. 
        Make it more professional, comprehensive, and suitable for an executive audience. 
        Maintain the original structure and key information while improving clarity and impact.
        
        Return only the enhanced content, no explanations or additional text.`;
        
        const userPrompt = `Section: ${section}\nCurrent content: ${currentContent || 'No content provided'}\nUser enhancement request: ${prompt}\n\nPlease enhance this content based on the user's request.`;
        
        const modelRequest = {
          anthropic_version: "bedrock-2023-05-31",
          max_tokens: 1500,
          system: systemPrompt,
          messages: [
            {
              role: "user",
              content: userPrompt
            }
          ]
        };
        
        const command = new InvokeModelCommand({
          modelId: "us.anthropic.claude-3-7-sonnet-20250219-v1:0",
          body: JSON.stringify(modelRequest),
          contentType: "application/json",
          accept: "application/json",
        });
        
        const response = await bedrockClient.send(command);
        const responseData = JSON.parse(new TextDecoder().decode(response.body));
        enhancedContent = responseData.content[0].text;
        
      } catch (bedrockError) {
        console.log('AWS Bedrock failed, using fallback enhancement:', bedrockError.message);
        enhancedContent = simpleAIEnhancement(section, prompt, currentContent);
      }
    } else {
      enhancedContent = simpleAIEnhancement(section, prompt, currentContent);
    }
    
    // Update the current report with enhanced content
    if (section === 'Executive Summary') {
      currentReport.executiveSummary = enhancedContent;
    } else if (section === 'Background') {
      currentReport.background = enhancedContent;
    } else if (section === 'Scope') {
      if (currentContent && currentContent.toLowerCase().includes('excluded')) {
        currentReport.scope.excluded = enhancedContent;
      } else {
        currentReport.scope.included = enhancedContent;
      }
    } else if (section === 'Appendix') {
      currentReport.appendix = enhancedContent;
    } else if (section === 'Audit Name') {
      currentReport.auditName = enhancedContent;
    }
    
    res.json({ success: true, enhancedContent });
  } catch (error) {
    console.error('Error enhancing section:', error);
    res.status(500).json({ success: false, error: 'Failed to enhance section: ' + error.message });
  }
});

// Simple AI Enhancement Function (Fallback)
function simpleAIEnhancement(section, prompt, currentContent) {
  const templates = {
    'Executive Summary': currentContent ? 
      `${currentContent}\n\nThis comprehensive executive summary provides strategic insights into the audit findings and recommendations for organizational improvement.` :
      'This comprehensive audit review identifies key operational areas requiring attention and provides strategic recommendations for enhanced efficiency and risk mitigation.',
    
    'Background': currentContent ?
      `${currentContent}\n\nAdditional context: This background section establishes the foundation for understanding the audit scope, methodology, and strategic importance of the reviewed processes.` :
      'This audit was conducted to evaluate critical business processes, internal controls, and compliance measures to ensure operational excellence and risk management.',
    
    'Audit Name': currentContent ?
      `${currentContent} - Comprehensive Process Review` :
      'Comprehensive Business Process Audit Review',
    
    'Finding Short Name': currentContent ?
      `${currentContent} - Critical Assessment` :
      'Process Control Assessment',
    
    'Finding Description': currentContent ?
      `Detailed Analysis: ${currentContent}\n\nThis finding represents a significant operational area requiring immediate attention to ensure regulatory compliance and process effectiveness.` :
      'This finding requires immediate attention to ensure operational compliance and process integrity.',
    
    'Action Item Short Name': currentContent ?
      `${currentContent} - Implementation Strategy` :
      'Process Improvement Implementation',
    
    'Action Item Description': currentContent ?
      `Implementation Plan: ${currentContent}\n\nThis action item should be prioritized with clear timelines, responsible parties, and success metrics to address identified risks and improve operational controls.` :
      'This action item addresses critical process improvements and risk mitigation measures with defined implementation timelines.',
    
    'Recommendations': currentContent ?
      `Enhanced Recommendations: ${currentContent}\n\nAdditional strategic considerations include process automation, regular monitoring protocols, comprehensive staff training, and continuous improvement initiatives.` :
      'Implement comprehensive process improvements including enhanced monitoring, staff training, and regular compliance assessments.',
    
    'Scope': currentContent ?
      `${currentContent}\n\nThis comprehensive scope encompasses detailed analysis of business processes, internal control evaluation, compliance assessment, and risk management review within the defined audit parameters.` :
      'Comprehensive audit scope including process analysis, control evaluation, compliance assessment, and strategic risk management review.',
    
    'Appendix': currentContent ?
      `${currentContent}\n\nThis appendix provides detailed supporting documentation, audit methodology references, and additional technical specifications for comprehensive review.` :
      'Detailed supporting documentation including audit methodology, compliance frameworks, and technical specifications.'
  };

  let enhancedContent = templates[section] || 
    (currentContent ? 
      `${currentContent}\n\nThis ${section.toLowerCase()} has been enhanced for professional presentation and executive review.` : 
      `Professional ${section.toLowerCase()} content developed for strategic decision-making and executive presentation.`);

  // Apply prompt-specific enhancements
  if (prompt.toLowerCase().includes('professional') || prompt.toLowerCase().includes('formal')) {
    enhancedContent = enhancedContent.replace(/\w+/g, word => 
      word.charAt(0).toUpperCase() + word.slice(1).toLowerCase()
    );
  }
  
  if (prompt.toLowerCase().includes('detailed') || prompt.toLowerCase().includes('comprehensive')) {
    enhancedContent += '\n\nAdditional detailed analysis and supporting evidence have been incorporated to provide comprehensive insights and actionable recommendations.';
  }
  
  if (prompt.toLowerCase().includes('executive') || prompt.toLowerCase().includes('summary')) {
    enhancedContent = `Executive Overview: ${enhancedContent}`;
  }
  
  if (prompt.toLowerCase().includes('improve') || prompt.toLowerCase().includes('enhance')) {
    enhancedContent += '\n\nThis enhancement focuses on clarity, professional presentation, and strategic value for stakeholder review.';
  }

  return enhancedContent;
}

app.post('/api/report/export', async (req, res) => {
  try {
    const doc = new Document({
      styles: {
        default: {
          document: {
            run: {
              font: "Calibri",
              size: 22, // 11pt in half-points
            },
          },
        },
      },

      sections: [{
        properties: {
          page: {
            margin: {
              top: 720, // 1.27 cm (720 twips = 0.5 inches = 1.27 cm)
              right: 907, // 1.6 cm (907 twips = 0.63 inches = 1.6 cm)
              bottom: 963, // 1.7 cm (963 twips = 0.67 inches = 1.7 cm)
              left: 775, // 1.37 cm (775 twips = 0.54 inches = 1.37 cm)
            },
            size: {
              width: 12240, // 8.5 inches
              height: 15840, // 11 inches
            },
            gutter: 0,
            gutterPosition: "left",
            orientation: "portrait",
          },
          headerDistance: 547, // 0.38 inches * 1440 = 547 twips
          footerDistance: 259, // 0.18 inches * 1440 = 259 twips
          lineNumbers: {
            countBy: 1,
            start: 0, // Set to 0 to ensure line numbers start at 1 in Word
            restart: "continuous",
            distance: 288 // 0.2 inches * 1440 = 288 twips
          }
        },
        footers: {
          default: new Footer({
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                  new TextRun({ 
                    text: "Amazon.com Confidential",
                    size: 20
                  })
                ]
              }),
              new Paragraph({
                alignment: AlignmentType.RIGHT,
                children: [
                  // Removed "Page " prefix to show just the number
                  new TextRun({
                    children: [
                      PageNumber.CURRENT
                    ],
                    size: 20
                  })
                ]
              })
            ]
          })
        },
        children: [
          // AUDIT NAME
          new Paragraph({
            children: [
              new TextRun({ 
                text: currentReport.auditName || "AUTOMATED CASH APPLICATION REVIEW", 
                bold: false, 
                color: "D2691E", // Orange-brown color matching template
                size: 26 // 13pt
              })
            ]
          }),
          
          // Single blank line
          new Paragraph({ text: "" }),
          
          // EXECUTIVE SUMMARY
          new Paragraph({
            children: [
              new TextRun({ 
                text: "EXECUTIVE SUMMARY", 
                bold: false, 
                color: "556B2F", // Olive green/dark khaki color matching template
                size: 24 // 12pt
              })
            ]
          }),
          
          // Executive Summary content - preserve line breaks
          ...(currentReport.executiveSummary ? 
            currentReport.executiveSummary.split('\n').map(line => 
              new Paragraph({
                text: line.trim(),
                alignment: AlignmentType.JUSTIFIED
              })
            ) : 
            [new Paragraph({
              text: "VP, FinOps Services requested an audit of Critical Vendor processes in 2025. This document summarizes FORGE's findings, business owner's responses and actions items on the global audit performed on Critical Vendor Processes. FORGE reviewed critical vendor processes, controls, and payments for the 12-month period ending 31-Dec-2025. Overall, FORGE identified XX findings rated according to the methodology included in Appendix A (XX High, XX Medium) and XX improvement opportunities.",
              alignment: AlignmentType.JUSTIFIED
            })]
          ),
          
          // Single blank line
          new Paragraph({ text: "" }),
          
          // BACKGROUND
          new Paragraph({
            children: [
              new TextRun({ 
                text: "BACKGROUND", 
                bold: false, 
                color: "556B2F", // Olive green/dark khaki color matching template
                size: 24 // 12pt
              })
            ]
          }),
          
          // Background content - preserve line breaks
          ...(currentReport.background ? 
            currentReport.background.split('\n').map(line => 
              new Paragraph({
                text: line.trim(),
                alignment: AlignmentType.JUSTIFIED
              })
            ) : 
            [new Paragraph({
              text: "A critical vendor (CV) is defined as a vendor who, as a result of non/late payment, could impact Amazon by causing building and office space to be unsafe for use, increasing the risk of reputational damage to Amazon or non-negotiable penalties/fees.",
              alignment: AlignmentType.JUSTIFIED
            })]
          ),
          
          // Single blank line
          new Paragraph({ text: "" }),
          
          // SCOPE
          new Paragraph({
            children: [
              new TextRun({ 
                text: "SCOPE", 
                bold: false, 
                color: "556B2F", // Olive green/dark khaki color matching template
                size: 24 // 12pt
              })
            ]
          }),
          
          // Scope included
          new Paragraph({
            children: [
              new TextRun({ 
                text: "FORGE's scope included:", 
                bold: true
              })
            ]
          }),
          
          // Scope included content - preserve line breaks
          ...(currentReport.scope.included ? 
            currentReport.scope.included.split('\n').map(line => 
              new Paragraph({
                text: line.trim(),
                alignment: AlignmentType.JUSTIFIED
              })
            ) : 
            [new Paragraph({
              text: "No scope inclusions specified.",
              alignment: AlignmentType.JUSTIFIED
            })]
          ),
          
          // Single blank line before out of scope
          new Paragraph({ text: "" }),
          
          // Out of scope
          new Paragraph({
            children: [
              new TextRun({ text: "For this inspection, the following areas were " }),
              new TextRun({ text: "out of scope", bold: true }),
              new TextRun({ text: ":" })
            ]
          }),
          
          // Scope excluded content - preserve line breaks
          ...(currentReport.scope.excluded ? 
            currentReport.scope.excluded.split('\n').map(line => 
              new Paragraph({
                text: line.trim(),
                alignment: AlignmentType.JUSTIFIED
              })
            ) : 
            [new Paragraph({
              text: "No scope exclusions specified.",
              alignment: AlignmentType.JUSTIFIED
            })]
          ),
          
          // Single blank line
          new Paragraph({ text: "" }),
          
          // FINDINGS AND ACTION ITEMS
          new Paragraph({
            children: [
              new TextRun({ 
                text: "FINDINGS AND ACTION ITEMS", 
                bold: false, 
                color: "556B2F", // Olive green/dark khaki color matching template
                size: 24 // 12pt
              })
            ]
          }),
          
          // Findings intro text
          new Paragraph({
            text: "Findings and action items resulting from the inspection are detailed in the table below. All metrics reported in the findings are based on information related to the XX-month period ended MM/DD/YY, unless otherwise noted. Action items will be tracked in AIM.",
            alignment: AlignmentType.JUSTIFIED
          }),
          
          // Create findings table with header
          createFindingsTable(currentReport.findings || []),
           
          new Paragraph({ text: "" }), // Single space
          
          // Line 18: APPENDIX
          new Paragraph({
            children: [
              new TextRun({ text: "18", bold: true, size: 20, color: "666666" }),
              new TextRun({ text: "    " }),
              new TextRun({ 
                text: "APPENDIX A -- RATING", 
                bold: true, 
                color: "FF8C00",
                size: 24 // 12pt
              })
            ]
          }),
          
          new Paragraph({
            children: [
              new TextRun({ text: "19", bold: true, size: 20, color: "666666" }),
              new TextRun({ text: "    " }),
              new TextRun({
                text: currentReport.appendix || "Not specified"
              })
            ]
          })
        ]
      }]
    });
    
    const buffer = await Packer.toBuffer(doc);
    
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', 'attachment; filename="FORGE_Audit_Report.docx"');
    res.send(buffer);
    
  } catch (error) {
    console.error('Error generating document:', error);
    res.status(500).json({ success: false, error: 'Failed to generate document' });
  }
});

// Helper function to create the findings table
function createFindingsTable(findings) {
  // Create header row with exact specifications
  const headerRow = new TableRow({
    tableHeader: true,
    children: [
      createHeaderCell("Finding, Name,\n   Rating, Ref"),
      createHeaderCell("Finding, Recommendation(s) &\nBusiness Owner Response"),
      createHeaderCell("Action Item(s)"),
      createHeaderCell("Due Date¹"),
      createHeaderCell("Owner (L8+)\n   Team")
    ]
  });
  
  // Set the header row height explicitly through direct properties
  if (!headerRow.properties) {
    headerRow.properties = {};
  }
  headerRow.properties = {
    height: {
      value: 68, // 0.12 cm = 0.047 inches * 1440 = 68 twips
      rule: "atLeast" // Changed to "atLeast" to accommodate multi-line text
    },
    cantSplit: true // Prevent the row from splitting across pages
  };
  
  // Create an array to hold all rows
  const rows = [headerRow];
  
  // Process each finding
  findings.forEach((finding, findingIndex) => {
    const findingNumber = findingIndex + 1;
    const findingNumberFormatted = findingNumber < 10 ? `0${findingNumber}` : `${findingNumber}`;
    
    // Convert risk rating to single letter
    let riskLetter = 'N/A';
    if (finding.rating === 'High') riskLetter = 'H';
    else if (finding.rating === 'Medium') riskLetter = 'M';
    else if (finding.rating === 'Low') riskLetter = 'L';
    
    // If there are no action items, create a single row
    if (!finding.actionItems || finding.actionItems.length === 0) {
      rows.push(
        new TableRow({
          children: [
            createFindingCell(findingNumberFormatted, finding.shortName, riskLetter),
            createDescriptionCell(finding.description),
            createEmptyCell("No action items"),
            createEmptyCell("MM/DD/YY"),
            createEmptyCell("N/A")
          ]
        })
      );
      
      // Make sure the bottom borders are visible for this row
      const lastRowIndex = rows.length - 1;
      const lastRow = rows[lastRowIndex];
      lastRow.children.forEach(cell => {
        if (cell && cell.borders) {
          cell.borders.bottom = { style: BorderStyle.SINGLE, size: 6, color: "000000" };
        }
      });
      
      return; // Skip to next finding
    }
    
    // Process action items
    finding.actionItems.forEach((actionItem, actionIndex) => {
      const isFirstActionItem = actionIndex === 0;
      const isLastActionItem = actionIndex === finding.actionItems.length - 1;
      
      // For first action item, include finding details
      if (isFirstActionItem) {
        rows.push(
          new TableRow({
            children: [
              createFindingCell(findingNumberFormatted, finding.shortName, riskLetter, finding.actionItems.length),
              createDescriptionCell(finding.description, finding.actionItems.length),
              createActionItemCell(
                `${findingNumber}.${actionIndex + 1}. ${actionItem.actionItemName || 'Action Item'}`,
                actionItem.actionItemDescription,
                !isLastActionItem
              ),
              createSimpleCell(formatDateToMMDDYY(actionItem.dueDate), !isLastActionItem),
              createOwnerCell(actionItem, !isLastActionItem)
            ]
          })
        );
      } else {
        // For subsequent action items, only include action item details
        rows.push(
          new TableRow({
            children: [
              // These cells are covered by rowSpan from the first row
              createActionItemCell(
                `${findingNumber}.${actionIndex + 1}. ${actionItem.actionItemName || 'Action Item'}`,
                actionItem.actionItemDescription,
                !isLastActionItem
              ),
              createSimpleCell(formatDateToMMDDYY(actionItem.dueDate), !isLastActionItem),
              createOwnerCell(actionItem, !isLastActionItem)
            ]
          })
        );
      }
    });
    
    // Instead of adding a separator row, add bottom borders to the last action item's cells
    if (finding.actionItems && finding.actionItems.length > 0) {
      const lastActionItemIndex = finding.actionItems.length - 1;
      const lastRowIndex = rows.length - 1;
      
      // Get the last row we just added
      const lastRow = rows[lastRowIndex];
      
      // Add bottom borders to all cells in the last row
      if (lastRow && lastRow.children && Array.isArray(lastRow.children)) {
        lastRow.children.forEach(cell => {
          if (cell && cell.borders) {
            cell.borders.bottom = { style: BorderStyle.SINGLE, size: 6, color: "000000" };
          }
        });
      }
    }
  });
  
  // Create and return the table
  return new Table({
    width: {
      size: 10714, // 18.89 cm = 7.44 inches * 1440 = 10714 twips
      type: WidthType.DXA,
    },
    indent: {
      size: -6, // -0.01 cm = -0.004 inches * 1440 = -6 twips
      type: WidthType.DXA,
    },
    tableHeader: true, // Enable repeat header rows
    layout: TableLayoutType.FIXED,
    borders: {
      top: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
      bottom: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
      left: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
      right: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
      insideVertical: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
      outside: { style: BorderStyle.SINGLE, size: 6, color: "000000" }
    },
    columnWidths: [
      Math.round(2.5 * 0.3937 * 1440),  // Finding, Name, Rating, Ref - 2.5 cm
      Math.round(8.75 * 0.3937 * 1440), // Finding, Recommendation(s) & Business Owner Response - 8.75 cm
      Math.round(3.5 * 0.3937 * 1440),  // Action Item(s) - 3.5 cm
      Math.round(2.06 * 0.3937 * 1440), // Due Date - 2.06 cm
      Math.round(2.28 * 0.3937 * 1440)  // Owner (L8+) Team - 2.28 cm
    ],
    margins: {
      top: 0,
      bottom: 0,
      left: Math.round(0.19 * 0.3937 * 1440), // 0.19 cm = 108 twips
      right: Math.round(0.19 * 0.3937 * 1440), // 0.19 cm = 108 twips
    },
    rows: rows
  });
}

// Helper function to create a header cell
function createHeaderCell(text) {
  return new TableCell({
    verticalAlign: "top", // Align top
    shading: { fill: "DCE6F1" }, // Light blue background matching template
    borders: {
      top: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
      left: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
      right: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
      bottom: { style: BorderStyle.SINGLE, size: 6, color: "000000" }
    },
    children: text.split('\n').map(line => new Paragraph({ 
      spacing: {
        before: 0,
        after: 0
      },
      children: [new TextRun({
        text: line,
        bold: true,
        color: "000000", // Black text matching template
        size: 18, // 9pt (18 half-points)
        font: "Calibri" // Explicitly set Calibri font
      })],
      alignment: AlignmentType.CENTER
    }))
  });
}

// Helper function to create a finding cell (first column)
function createFindingCell(findingNumber, shortName, riskLetter, rowSpan = 1) {
  return new TableCell({
    verticalAlign: "top",
    borders: {
      top: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
      left: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
      right: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
      bottom: { style: BorderStyle.NONE }
    },
    rowSpan: rowSpan,
    children: [
      new Paragraph({ 
        children: [new TextRun({
          text: findingNumber,
          size: 18 // 9pt
        })],
        alignment: AlignmentType.CENTER
      }),
      new Paragraph({ 
        children: [new TextRun({
          text: shortName || "Finding Name",
          size: 18 // 9pt
        })],
        alignment: AlignmentType.CENTER
      }),
      new Paragraph({ 
        children: [new TextRun({ 
          text: riskLetter, 
          color: "FF0000",
          bold: true,
          size: 18 // 9pt
        })],
        alignment: AlignmentType.CENTER
      })
    ]
  });
}

// Helper function to create a description cell (second column)
function createDescriptionCell(description, rowSpan = 1) {
  return new TableCell({
    verticalAlign: "top",
    borders: {
      top: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
      left: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
      right: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
      bottom: { style: BorderStyle.NONE }
    },
    rowSpan: rowSpan,
    children: [
      new Paragraph({ 
        children: [new TextRun({
          text: `Finding Description: ${description || "The finding goes here. Use complete sentences. Include risk to Amazon, & FinOps, sampling/data analysis details, what we audited against (e.g., policy, regulation), and any additional details that would convince the owners to action on the findings."}`,
          size: 18 // 9pt
        })]
      }),
      new Paragraph({ 
        children: [new TextRun({
          text: "",
          size: 18 // 9pt
        })]
      }),
      new Paragraph({ 
        children: [new TextRun({ 
          text: "Recommendations:",
          bold: true,
          size: 18 // 9pt
        })],
        size: 18 // 9pt
      }),
      new Paragraph({ 
        children: [new TextRun({
          text: "1. XX",
          size: 18 // 9pt
        })]
      }),
      new Paragraph({ 
        children: [new TextRun({
          text: "2. XX.",
          size: 18 // 9pt
        })]
      }),
      new Paragraph({ 
        children: [new TextRun({
          text: "3. XX.",
          size: 18 // 9pt
        })]
      }),
      new Paragraph({ 
        children: [new TextRun({
          text: "",
          size: 18 // 9pt
        })]
      }),
      new Paragraph({ 
        children: [new TextRun({ 
          text: "Business Owner Response: Team (Name): ",
          bold: true,
          size: 18 // 9pt
        })],
        size: 18 // 9pt
      }),
      new Paragraph({ 
        children: [new TextRun({ 
          text: "We agree with the finding.",
          size: 18 // 9pt
        })]
      })
    ]
  });
}

// Helper function to create an action item cell
function createActionItemCell(title, content, showBottomBorder) {
  // Create array of children paragraphs
  const children = [
    new Paragraph({ 
      children: [new TextRun({ 
        text: title,
        bold: true,
        size: 18 // 9pt
      })],
      alignment: AlignmentType.LEFT
    })
  ];
  
  // Debug the content
  console.log('Action Item Content:', content);
  
  // Always add a paragraph for content, even if empty
  children.push(
    new Paragraph({ 
      children: [new TextRun({
        text: content || '',
        bold: false,
        size: 18 // 9pt
      })],
      alignment: AlignmentType.LEFT
    })
  );
  
  return new TableCell({
    verticalAlign: "top",
    borders: {
      top: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
      left: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
      right: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
      bottom: showBottomBorder ? { style: BorderStyle.SINGLE, size: 6, color: "000000" } : { style: BorderStyle.NONE }
    },
    children: children
  });
}

// Helper function to format date from YYYY-MM-DD to MM/DD/YY
function formatDateToMMDDYY(dateString) {
  if (!dateString || dateString === "MM/DD/YY" || dateString === "N/A") {
    return "MM/DD/YY";
  }
  
  // Check if it's already in MM/DD/YY format
  if (dateString.includes('/')) {
    return dateString;
  }
  
  try {
    const date = new Date(dateString);
    if (isNaN(date.getTime())) {
      return "MM/DD/YY";
    }
    
    const month = (date.getMonth() + 1).toString().padStart(2, '0');
    const day = date.getDate().toString().padStart(2, '0');
    const year = date.getFullYear().toString().slice(-2); // Get last 2 digits
    
    return `${month}/${day}/${year}`;
  } catch (error) {
    return "MM/DD/YY";
  }
}

// Helper function to create an owner cell with the three fields
function createOwnerCell(actionItem, showBottomBorder) {
  return new TableCell({
    verticalAlign: "top",
    borders: {
      top: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
      left: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
      right: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
      bottom: showBottomBorder ? { style: BorderStyle.SINGLE, size: 6, color: "000000" } : { style: BorderStyle.NONE }
    },
    children: [
      new Paragraph({ 
        children: [new TextRun({
          text: actionItem.actionOwner || "Action Owner",
          size: 18 // 9pt
        })],
        alignment: AlignmentType.CENTER
      }),
      // Add empty paragraph for line break
      new Paragraph({
        children: [new TextRun({
          text: "",
          size: 18
        })],
        alignment: AlignmentType.CENTER
      }),
      new Paragraph({ 
        children: [new TextRun({
          text: actionItem.ownerL8 || "L8 name",
          size: 18 // 9pt
        })],
        alignment: AlignmentType.CENTER
      }),
      new Paragraph({ 
        children: [new TextRun({
          text: actionItem.ownerTeam ? `(${actionItem.ownerTeam})` : "(Team)",
          size: 18 // 9pt
        })],
        alignment: AlignmentType.CENTER
      })
    ]
  });
}

// Helper function to create a simple cell (due date or other)
function createSimpleCell(text, showBottomBorder) {
  
  // Regular cell for due dates or other content
  return new TableCell({
    verticalAlign: "top",
    borders: {
      top: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
      left: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
      right: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
      bottom: showBottomBorder ? { style: BorderStyle.SINGLE, size: 6, color: "000000" } : { style: BorderStyle.NONE }
    },
    children: [new Paragraph({ 
      children: [new TextRun({
        text: text,
        size: 18 // 9pt
      })],
      alignment: AlignmentType.CENTER
    })]
  });
}

// Helper function to create an empty cell
function createEmptyCell(text) {
  // Special handling for empty owner cell
  if (text === "N/A") {
    return new TableCell({
      verticalAlign: "top",
      borders: {
        top: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
        left: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
        right: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
        bottom: { style: BorderStyle.SINGLE, size: 6, color: "000000" }
      },
      children: [
        new Paragraph({ 
          children: [new TextRun({
            text: "Action Owner",
            size: 18 // 9pt
          })],
          alignment: AlignmentType.CENTER
        }),
        // Add empty paragraph for line break
        new Paragraph({
          children: [new TextRun({
            text: "",
            size: 18
          })],
          alignment: AlignmentType.CENTER
        }),
        new Paragraph({ 
          children: [new TextRun({
            text: "L8 name",
            size: 18 // 9pt
          })],
          alignment: AlignmentType.CENTER
        }),
        new Paragraph({ 
          children: [new TextRun({
            text: "(Team)",
            size: 18 // 9pt
          })],
          alignment: AlignmentType.CENTER
        })
      ]
    });
  }
  
  return new TableCell({
    verticalAlign: "top",
    borders: {
      top: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
      left: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
      right: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
      bottom: { style: BorderStyle.SINGLE, size: 6, color: "000000" }
    },
    children: [new Paragraph({ 
      children: [new TextRun({
        text: text,
        size: 18 // 9pt
      })],
      alignment: AlignmentType.CENTER
    })]
  });
}

app.listen(PORT, () => {
  console.log(`FORGE Audit Report Generator running on port ${PORT}`);
});
