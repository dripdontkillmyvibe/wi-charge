// This code runs when someone opens your spreadsheet
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🔋 Wi-Charge Evaluator')
    .addItem('Start New Evaluation', 'showSidebar')
    .addItem('View Instructions', 'showInstructions')
    .addSeparator()
    .addItem('View Technical Reference', 'showTechReference')
    .addItem('Test AI Connection', 'testAIConnection')
    .addItem('Reset Results Sheet', 'resetResultsSheet')
    .addItem('About AI Mode', 'showAIInfo')
    .addToUi();
}

// Add this new function
function showTechReference() {
  const refData = getReferenceData();
  const message = `
Wi-Charge Technical Reference

TRANSMITTER SPECIFICATIONS:
- R1: 100mW per device, 1-10m range, 80° FOV
- R1-HP: 300mW per device, 1-5m range, 80° FOV

KEY APPLICATIONS:
- Smart Locks: 10-50mW (R1 suitable)
- Digital Displays: 100-300mW (R1-HP recommended)
- IoT Sensors: 1-20mW (R1 suitable)
- Security Cameras: 200-300mW (R1-HP needed)

PROVEN USE CASES:
- Retail Digital Signage: 10% sales increase
- Hotel Smart Locks: $100/lock/year savings
- Meeting Room Displays: 50% install cost reduction

For detailed specifications, see the Technical Specs sheet.
  `;

  SpreadsheetApp.getUi().alert('📚 Technical Reference', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

// This shows the evaluation form
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Device Evaluation')
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

// This shows instructions
function showInstructions() {
  const message = `
Wi-Charge Evaluation Instructions:

1. Click "Wi-Charge Evaluator" menu → "Start New Evaluation"
2. Enter device information (minimum: device name)
3. AI will research missing specifications if needed
4. Results appear in the Results sheet

Scoring: 3=Strong, 2=Workable, 1=Weak
Pass threshold: Total score ≥ 27

💡 TIP: You can enter just a device name and the AI will research typical specifications!
  `;

  SpreadsheetApp.getUi().alert('How to Use', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

// This is the UPDATED main evaluation function that uses AI
function processEvaluation(data) {
  try {
    // Use ChatGPT for evaluation
    const aiEvaluation = getAIEvaluation(data);

    // Check if it passed the gate
    const gatePassed = aiEvaluation.gateResult.powerBudget.pass && aiEvaluation.gateResult.lineOfSight.pass;

    if (!gatePassed) {
      const gateReason = !aiEvaluation.gateResult.powerBudget.pass ?
        aiEvaluation.gateResult.powerBudget.details :
        aiEvaluation.gateResult.lineOfSight.details;
      writeFailedResult(data, gateReason);
      return { success: true };
    }

    // Write the AI-powered results
    writeAIResult(aiEvaluation);

    return { success: true };

  } catch (error) {
    // If AI fails, fall back to simple evaluation
    console.error('AI evaluation failed, using fallback:', error);

    // Use the original simple evaluation
    const gateCheck = checkShowStoppers(data);
    if (!gateCheck.pass) {
      writeFailedResult(data, gateCheck.reason);
      return { success: true };
    }

    const scores = calculateScores(data);
    writeSuccessResult(data, scores);

    return { success: true };
  }
}

// Check if device passes basic requirements
function checkShowStoppers(data) {
  const avgPower = parseFloat(data.avgPower);
  const peakPower = parseFloat(data.peakPower);
  const los = parseFloat(data.los);

  // Check power requirements
  if (avgPower > 300) {
    return { pass: false, reason: 'Average power exceeds 300mW limit' };
  }

  if (peakPower > 5) {
    return { pass: false, reason: 'Peak power exceeds 5W limit' };
  }

  // Check line of sight
  if (los < 60) {
    return { pass: false, reason: 'Line of sight less than 60% requirement' };
  }

  return { pass: true };
}

// Calculate scores for each factor
function calculateScores(data) {
  const scores = {};

  // 1. Pain Intensity (battery/OPEX)
  const swapsYear = parseFloat(data.swapsYear) || 0;
  const swapCost = parseFloat(data.swapCost) || 0;
  const annualCost = swapsYear * swapCost;

  if (swapsYear < 1 || annualCost < 10) {
    scores.pain = { score: 1, reason: 'Low battery replacement frequency/cost' };
  } else if (swapsYear >= 4 || annualCost >= 100) {
    scores.pain = { score: 3, reason: 'High battery pain point' };
  } else {
    scores.pain = { score: 2, reason: 'Moderate battery maintenance needs' };
  }

  // 2. Coverage Geometry (simplified)
  scores.coverage = { score: 2, reason: 'Standard coverage geometry assumed' };

  // 3. Install Leverage
  scores.install = { score: 2, reason: 'Standard installation assumed' };

  // 4. Regulatory
  scores.regulatory = { score: 2, reason: 'Standard regulatory path assumed' };

  // 5. ROI Math
  if (annualCost >= 100) {
    scores.roi = { score: 3, reason: 'Strong ROI potential' };
  } else if (annualCost >= 50) {
    scores.roi = { score: 2, reason: 'Moderate ROI' };
  } else {
    scores.roi = { score: 1, reason: 'Weak ROI' };
  }

  // 6. Fleet Size
  scores.fleet = { score: 2, reason: 'Fleet size needs verification' };

  // 7. Strategic Halo
  scores.halo = { score: 2, reason: 'Standard strategic value' };

  // 8. Competitive Threat
  scores.threat = { score: 2, reason: 'Moderate competitive landscape' };

  // Calculate weighted total
  const weights = {
    pain: 2, coverage: 1, install: 2, regulatory: 1,
    roi: 2, fleet: 1, halo: 1, threat: 1
  };

  let totalScore = 0;
  for (let factor in scores) {
    totalScore += scores[factor].score * weights[factor];
  }

  scores.total = totalScore;
  scores.verdict = totalScore >= 27 ? 'ADVANCE' :
                   totalScore >= 21 ? 'NEEDS WORK' : 'PARK';

  return scores;
}

// Write results for devices that failed the gate check
function writeFailedResult(data, reason) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Results');

  // Add headers if this is the first result
  if (sheet.getLastRow() === 0) {
    addResultHeaders(sheet);
  }

  // Add the failed result
  const row = [
    new Date(),
    data.name,
    'FAILED',
    reason,
    '-', '-', '-', '-', '-', '-', '-', '-', '-', '-',
    'PARK - Failed Gate Check'
  ];

  sheet.appendRow(row);
}

// Write results for devices that passed
function writeSuccessResult(data, scores) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Results');

  // Add headers if this is the first result
  if (sheet.getLastRow() === 0) {
    addResultHeaders(sheet);
  }

  // Add the successful result
  const row = [
    new Date(),
    data.name,
    'PASSED',
    'Passed all gate checks',
    scores.pain.score,
    scores.coverage.score,
    scores.install.score,
    scores.regulatory.score,
    scores.roi.score,
    scores.fleet.score,
    scores.halo.score,
    scores.threat.score,
    scores.total,
    scores.verdict,
    'Next: Verify actual coverage geometry and fleet expansion plans'
  ];

  sheet.appendRow(row);

  // Format the new row
  const lastRow = sheet.getLastRow();
  const range = sheet.getRange(lastRow, 1, 1, 15);

  // Color code based on verdict
  if (scores.verdict === 'ADVANCE') {
    range.setBackground('#d4f8d4'); // Light green
  } else if (scores.verdict === 'NEEDS WORK') {
    range.setBackground('#fff3cd'); // Light yellow
  } else {
    range.setBackground('#f8d7da'); // Light red
  }
}

// Add headers to results sheet
function addResultHeaders(sheet) {
  const headers = [
    'Date',
    'Device Name',
    'Gate Result',
    'Gate Details',
    'Pain Score',
    'Coverage Score',
    'Install Score',
    'Regulatory Score',
    'ROI Score',
    'Fleet Score',
    'Halo Score',
    'Threat Score',
    'Total Score',
    'Verdict',
    'Next Steps'
  ];

  sheet.appendRow(headers);

  // Format headers
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#e7e7e7');
}

// COMPLETELY REVISED writeAIResult function
function writeAIResult(result) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Results');

  // Determine the next available row without clearing previous results

  // Write evaluation header
  const lastRow = sheet.getLastRow();
  const newRow = lastRow + 1;

  // Title
  sheet.getRange(newRow, 1).setValue('🔋 Wi-Charge Evaluation Report');
  sheet.getRange(newRow, 1).setFontSize(16).setFontWeight('bold');

  // Device name and date
  sheet.getRange(newRow + 1, 1).setValue('Device:');
  sheet.getRange(newRow + 1, 2).setValue(result.deviceName);
  sheet.getRange(newRow + 1, 2).setFontWeight('bold');
  sheet.getRange(newRow + 1, 4).setValue('Date:');
  sheet.getRange(newRow + 1, 5).setValue(new Date());

  // Gate Results
  sheet.getRange(newRow + 3, 1).setValue('⛔ Gate Check Results');
  sheet.getRange(newRow + 3, 1).setFontWeight('bold').setFontSize(12);

  sheet.getRange(newRow + 4, 1).setValue('Power Budget:');
  sheet.getRange(newRow + 4, 2).setValue(result.gateResult.powerBudget.pass ? '✅ PASS' : '❌ FAIL');
  sheet.getRange(newRow + 4, 3, 1, 3).merge();
  sheet.getRange(newRow + 4, 3).setValue(result.gateResult.powerBudget.details);

  sheet.getRange(newRow + 5, 1).setValue('Line of Sight:');
  sheet.getRange(newRow + 5, 2).setValue(result.gateResult.lineOfSight.pass ? '✅ PASS' : '❌ FAIL');
  sheet.getRange(newRow + 5, 3, 1, 3).merge();
  sheet.getRange(newRow + 5, 3).setValue(result.gateResult.lineOfSight.details);

  // Scores Table
  sheet.getRange(newRow + 7, 1).setValue('📊 Scoring Analysis');
  sheet.getRange(newRow + 7, 1).setFontWeight('bold').setFontSize(12);

  // Score headers
  const scoreHeaders = ['Factor', 'Score', 'Weight', 'Weighted', 'Rationale'];
  sheet.getRange(newRow + 8, 1, 1, 5).setValues([scoreHeaders]);
  sheet.getRange(newRow + 8, 1, 1, 5).setFontWeight('bold').setBackground('#e7e7e7');

  // Score data
  const scoreRows = [
    ['Pain Intensity', result.scores.painIntensity.score, result.scores.painIntensity.weight, result.scores.painIntensity.weighted, result.scores.painIntensity.reason],
    ['Coverage Geometry', result.scores.coverageGeometry.score, result.scores.coverageGeometry.weight, result.scores.coverageGeometry.weighted, result.scores.coverageGeometry.reason],
    ['Install Leverage', result.scores.installLeverage.score, result.scores.installLeverage.weight, result.scores.installLeverage.weighted, result.scores.installLeverage.reason],
    ['Regulatory & Shipping', result.scores.regulatoryShipping.score, result.scores.regulatoryShipping.weight, result.scores.regulatoryShipping.weighted, result.scores.regulatoryShipping.reason],
    ['ROI Math', result.scores.roiMath.score, result.scores.roiMath.weight, result.scores.roiMath.weighted, result.scores.roiMath.reason],
    ['Fleet Size', result.scores.fleetSize.score, result.scores.fleetSize.weight, result.scores.fleetSize.weighted, result.scores.fleetSize.reason],
    ['Strategic Halo', result.scores.strategicHalo.score, result.scores.strategicHalo.weight, result.scores.strategicHalo.weighted, result.scores.strategicHalo.reason],
    ['Competitive Threat', result.scores.competitiveThreat.score, result.scores.competitiveThreat.weight, result.scores.competitiveThreat.weighted, result.scores.competitiveThreat.reason]
  ];

  sheet.getRange(newRow + 9, 1, 8, 5).setValues(scoreRows);
  sheet.getRange(newRow + 9, 5, 8, 1).setWrap(true);

  // Total Score
  sheet.getRange(newRow + 18, 1).setValue('TOTAL SCORE:');
  sheet.getRange(newRow + 18, 1).setFontWeight('bold');
  sheet.getRange(newRow + 18, 2).setValue(result.totalScore);
  sheet.getRange(newRow + 18, 2).setFontWeight('bold').setFontSize(14);

  // Verdict
  sheet.getRange(newRow + 19, 1).setValue('VERDICT:');
  sheet.getRange(newRow + 19, 1).setFontWeight('bold');
  sheet.getRange(newRow + 19, 2).setValue(result.verdict);
  sheet.getRange(newRow + 19, 2).setFontWeight('bold').setFontSize(14);

  // Color code verdict
  const verdictCell = sheet.getRange(newRow + 19, 2);
  if (result.verdict.includes('Advance')) {
    verdictCell.setBackground('#d4f8d4'); // Light green
  } else if (result.verdict.includes('creativity')) {
    verdictCell.setBackground('#fff3cd'); // Light yellow
  } else {
    verdictCell.setBackground('#f8d7da'); // Light red
  }

  // Top Risks
  sheet.getRange(newRow + 21, 1).setValue('⚠️ Top Risks:');
  sheet.getRange(newRow + 21, 1).setFontWeight('bold');
  sheet.getRange(newRow + 21, 2, 1, 4).merge();
  sheet.getRange(newRow + 21, 2).setValue(result.topRisks.join('\n'));
  sheet.getRange(newRow + 21, 2).setWrap(true);

  // Recommendations
  if (result.recommendations && result.recommendations.length > 0) {
    sheet.getRange(newRow + 23, 1).setValue('💡 Recommendations:');
    sheet.getRange(newRow + 23, 1).setFontWeight('bold');
    sheet.getRange(newRow + 23, 2, 1, 4).merge();
    sheet.getRange(newRow + 23, 2).setValue(result.recommendations.join('\n'));
    sheet.getRange(newRow + 23, 2).setWrap(true);
  }

  // Estimated Values (if present)
  if (result.estimatedValues) {
    const estimatedRow = newRow + 25;
    sheet.getRange(estimatedRow, 1).setValue('📊 Estimated Values:');
    sheet.getRange(estimatedRow, 1).setFontWeight('bold');
    sheet.getRange(estimatedRow, 1).setBackground('#FFF3CD'); // Light yellow

    let estimatedText = '';
    if (result.estimatedValues.avgPower) {
      estimatedText += `Average Power: ${result.estimatedValues.avgPower}\n`;
    }
    if (result.estimatedValues.peakPower) {
      estimatedText += `Peak Power: ${result.estimatedValues.peakPower}\n`;
    }
    if (result.estimatedValues.los) {
      estimatedText += `Line of Sight: ${result.estimatedValues.los}\n`;
    }
    if (result.estimatedValues.assumptions && result.estimatedValues.assumptions.length > 0) {
      estimatedText += '\nKey Assumptions:\n';
      result.estimatedValues.assumptions.forEach(assumption => {
        estimatedText += `• ${assumption}\n`;
      });
    }

    sheet.getRange(estimatedRow + 1, 1, 1, 6).merge();
    sheet.getRange(estimatedRow + 1, 1).setValue(estimatedText);
    sheet.getRange(estimatedRow + 1, 1).setWrap(true);
    sheet.getRange(estimatedRow + 1, 1).setBackground('#FFF3CD'); // Light yellow background
  }

  // Auto-resize columns
  sheet.autoResizeColumns(1, 6);

  // Set column widths for better readability
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(5, 300);
}

/**
 * Retrieves text from up to five PDFs stored in a Drive folder named "ReferencePDFs".
 * Requires the Drive advanced service to be enabled for the script.
 * @return {Object[]} Array of objects containing file name and extracted text.
 */
function getReferenceData() {
  const results = [];
  const folderIt = DriveApp.getFoldersByName('ReferencePDFs');
  if (!folderIt.hasNext()) {
    Logger.log('ReferencePDFs folder not found. No reference data will be included.');
    return results;
  }
  const folder = folderIt.next();
  const files = folder.getFilesByType(MimeType.PDF);
  let count = 0;
  while (files.hasNext() && count < 5) {
    const file = files.next();
    try {
      // Convert PDF to Google Doc to read its text
      const temp = Drive.Files.copy({}, file.getId(), {convert: true});
      const doc = DocumentApp.openById(temp.id);
      const text = doc.getBody().getText();
      Drive.Files.remove(temp.id);
      results.push({name: file.getName(), text: text});
    } catch (e) {
      results.push({name: file.getName(), text: 'Error extracting text: ' + e.message});
    }
    count++;
  }
  return results;
}

/**
 * Displays information about the AI evaluation mode.
 */
function showAIInfo() {
  const msg = 'This add-on uses OpenAI to evaluate your device. When key specs are missing, ' +
              'the AI researches typical values online and combines them with text from PDFs ' +
              'placed in the "ReferencePDFs" folder on Drive.';
  SpreadsheetApp.getUi().alert('About AI Mode', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Simple connectivity test to the OpenAI API.
 */
function testAIConnection() {
  const ui = SpreadsheetApp.getUi();
  try {
    const props = PropertiesService.getScriptProperties();
    const apiKey = props.getProperty('OPENAI_API_KEY');
    const model = props.getProperty('OPENAI_MODEL') || 'gpt-4o';
    const tokenParam = props.getProperty('TOKEN_PARAM') || 'max_tokens';
    if (!apiKey) {
      ui.alert('OpenAI API key not found. Set OPENAI_API_KEY in Script Properties.');
      return;
    }

    const url = 'https://api.openai.com/v1/chat/completions';
    const payload = {
      model: model,
      messages: [{role: 'user', content: 'Hello'}]
    };
    payload[tokenParam] = 1;
    const options = {
      method: 'post',
      headers: { 'Authorization': 'Bearer ' + apiKey, 'Content-Type': 'application/json' },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const resp = UrlFetchApp.fetch(url, options);
    if (resp.getResponseCode() === 200) {
      ui.alert('✅ AI connection successful');
    } else {
      ui.alert('❌ Connection failed: ' + resp.getContentText());
    }
  } catch (e) {
    ui.alert('Error testing AI connection: ' + e.message);
  }
}

// Reset the Results sheet by clearing all content and formatting
function resetResultsSheet() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('Reset Results Sheet?', 'This will remove all existing results. Continue?', ui.ButtonSet.OK_CANCEL);
  if (response !== ui.Button.OK) {
    return;
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Results');
  if (sheet) {
    sheet.clear();
  }
}
