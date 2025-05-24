function getAIEvaluation(deviceData) {
  try {
    const props = PropertiesService.getScriptProperties();
    const apiKey = props.getProperty('OPENAI_API_KEY');
    const model = props.getProperty('OPENAI_MODEL') || 'gpt-4o';
    const tokenParam = props.getProperty('TOKEN_PARAM') || 'max_tokens';
    if (!apiKey) {
      throw new Error('OpenAI API key not found. Please set it in Script Properties.');
    }

    // Get reference data and user settings
    const referenceData = getReferenceData();
    const settings = getSettings();

    // Check if we need research mode
    const needsResearch = !deviceData.avgPower || !deviceData.peakPower || !deviceData.los;

    // Build the system prompt with research capabilities
    const systemPrompt = `You are the Wi-Charge Device Evaluation Expert. ${needsResearch ? 'RESEARCH MODE ACTIVATED: You need to research and estimate missing device specifications based on the device name and description provided.' : ''}

${needsResearch ? `RESEARCH INSTRUCTIONS:
When device specifications are missing, you should:
1. Identify the device type and typical use cases
2. Research or estimate typical power consumption for similar devices
3. Estimate line-of-sight availability based on typical installation scenarios
4. Make reasonable assumptions about battery replacement frequency and costs
5. Clearly state when you are making estimates vs using provided data
6. Use industry standards and typical values for the device category

For example:
- Smart locks typically consume 150-300mW average, 1-3W peak
- IoT sensors usually need 50-200mW average, 0.5-2W peak
- Security cameras need 2-5W continuous power
- Line-of-sight depends on mounting location (ceiling=90%, wall=60-80%, desk=40-60%)
- Battery replacement costs typically $15-50 including labor

` : ''}

EVALUATION FRAMEWORK:

# Show-Stopper Gate (Must Pass Both)
1. Power-Budget Fit: R1 delivers 1W, R1-HP delivers 3W within certified range
2. Line-of-Sight: Needs >40% unobstructed time

# Scoring Factors (1=weak/blocker, 2=workable, 3=strong)
1. Pain Intensity (${Math.round(settings.weights.painIntensity*100)}%): How painful/costly is current solution?
2. Coverage Geometry (${Math.round(settings.weights.coverageGeometry*100)}%): How well does omnidirectional coverage help?
3. Install Leverage (${Math.round(settings.weights.installLeverage*100)}%): How easy to deploy transmitters?
4. Regulatory & Shipping (${Math.round(settings.weights.regulatoryShipping*100)}%): Complexity of compliance?
5. ROI Math (${Math.round(settings.weights.roiMath*100)}%): Do savings justify costs?
6. Fleet Size (${Math.round(settings.weights.fleetSize*100)}%): How many devices benefit?
7. Strategic Halo (${Math.round(settings.weights.strategicHalo*100)}%): Deeper benefits/differentiation?
8. Competitive Threat (${Math.round(settings.weights.competitiveThreat*100)}%): Risk if competitors get wireless power?

# Available Reference Data:
${JSON.stringify(referenceData, null, 2)}

# Output Format:
Provide a valid JSON response with this structure. Include an "overview" field summarizing market context and ROI in short bullets.
{
  "deviceName": "string",
  "gateResult": {
    "powerBudget": {"pass": boolean, "details": "string"},
    "lineOfSight": {"pass": boolean, "details": "string"}
  },
  "scores": {
    "painIntensity": {"score": 1-3, "weight": ${settings.weights.painIntensity}, "weighted": number, "reason": "string"},
    "coverageGeometry": {"score": 1-3, "weight": ${settings.weights.coverageGeometry}, "weighted": number, "reason": "string"},
    "installLeverage": {"score": 1-3, "weight": ${settings.weights.installLeverage}, "weighted": number, "reason": "string"},
    "regulatoryShipping": {"score": 1-3, "weight": ${settings.weights.regulatoryShipping}, "weighted": number, "reason": "string"},
    "roiMath": {"score": 1-3, "weight": ${settings.weights.roiMath}, "weighted": number, "reason": "string"},
    "fleetSize": {"score": 1-3, "weight": ${settings.weights.fleetSize}, "weighted": number, "reason": "string"},
    "strategicHalo": {"score": 1-3, "weight": ${settings.weights.strategicHalo}, "weighted": number, "reason": "string"},
    "competitiveThreat": {"score": 1-3, "weight": ${settings.weights.competitiveThreat}, "weighted": number, "reason": "string"}
  },
  "totalScore": number,
  "verdict": "string (Advance/Needs creativity/Park)",
  "topRisks": ["risk1", "risk2", "risk3"],
  "recommendations": ["specific transmitter model", "mounting suggestion", "use case reference"],
  "overview": "string with bullet-style business case summary",
  "estimatedValues": ${needsResearch ? '{' : 'null'}
    ${needsResearch ? '"avgPower": "estimated value with unit",' : ''}
    ${needsResearch ? '"peakPower": "estimated value with unit",' : ''}
    ${needsResearch ? '"los": "estimated percentage",' : ''}
    ${needsResearch ? '"assumptions": ["list of key assumptions made"]' : ''}
  ${needsResearch ? '}' : ''}
}`;

    // Build user prompt with available data
    let userPrompt = `Evaluate this device for Wi-Charge wireless power:\n\n`;
    userPrompt += `Device Name: ${deviceData.name}\n`;

    if (deviceData.avgPower) {
      userPrompt += `Average Power: ${deviceData.avgPower} mW\n`;
    } else {
      userPrompt += `Average Power: NOT PROVIDED - Please research and estimate\n`;
    }

    if (deviceData.peakPower) {
      userPrompt += `Peak Power: ${deviceData.peakPower} W\n`;
    } else {
      userPrompt += `Peak Power: NOT PROVIDED - Please research and estimate\n`;
    }

    if (deviceData.los) {
      userPrompt += `Line of Sight: ${deviceData.los}%\n`;
    } else {
      userPrompt += `Line of Sight: NOT PROVIDED - Please estimate based on typical installation\n`;
    }

    if (deviceData.swapsYear) {
      userPrompt += `Battery Swaps/Year: ${deviceData.swapsYear}\n`;
    }
    if (deviceData.swapCost) {
      userPrompt += `Cost per Swap: $${deviceData.swapCost}\n`;
    }
    if (deviceData.description) {
      userPrompt += `Description: ${deviceData.description}\n`;
    }
    if (deviceData.deviceType) {
      userPrompt += `Device Type: ${deviceData.deviceType}\n`;
    }

    if (needsResearch) {
      userPrompt += `\nNOTE: This is a RESEARCH MODE evaluation. Please research the device specifications and provide your best estimates for missing values. Be transparent about what is estimated vs provided.`;
    }

    // Make API call
    const url = 'https://api.openai.com/v1/chat/completions';
    const payload = {
      'model': model,
      'messages': [
          {
            'role': 'system',
            'content': systemPrompt
          },
          {
            'role': 'user',
            'content': userPrompt
          }
        ],
      'temperature': 0.7
    };
    payload[tokenParam] = 2000;

    const options = {
      'method': 'post',
      'headers': {
        'Authorization': 'Bearer ' + apiKey,
        'Content-Type': 'application/json'
      },
      'payload': JSON.stringify(payload)
    };

    options.muteHttpExceptions = true;
    const response = UrlFetchApp.fetch(url, options);

    if (response.getResponseCode() !== 200) {
      throw new Error('OpenAI API error: ' + response.getContentText());
    }

    const result = JSON.parse(response.getContentText());

    if (result.choices && result.choices[0] && result.choices[0].message) {
      const content = result.choices[0].message.content;

      // Try to parse the JSON response
      try {
        // Remove any markdown code blocks if present
        const jsonMatch = content.match(/```json\n?([\s\S]*?)\n?```/) || content.match(/({[\s\S]*})/);
        if (jsonMatch) {
          return JSON.parse(jsonMatch[1]);
        }
        return JSON.parse(content);
      } catch (parseError) {
        console.error('Failed to parse AI response:', content);
        throw new Error('AI response was not in the expected format');
      }
    }

    throw new Error('Unexpected API response structure');

  } catch (error) {
    console.error('AI evaluation error:', error);
    throw error;
  }
}
