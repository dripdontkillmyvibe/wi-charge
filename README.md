# Wi-Charge Device Evaluator

This Google Sheets add‑on evaluates device ideas for use with Wi‑Charge wireless power transmitters using OpenAI. Missing specifications can be researched automatically and reference PDFs are included in the analysis.

## Setup

1. Open the spreadsheet and choose **Extensions → Apps Script**.
2. In **Project Settings**, add a script property named `OPENAI_API_KEY` with your OpenAI key.
3. (Optional) add `OPENAI_MODEL` to override the default model (`gpt-4o`).
4. (Optional) add `TOKEN_PARAM` to control which token limit parameter is sent (`max_tokens` or `max_completion_tokens`, defaults to `max_tokens`).
5. (Optional) Enable the **Drive API** in **Services** to allow PDF extraction.
6. Create a Drive folder named **ReferencePDFs** and place up to five PDF files you want included as reference material. If this folder is missing, the script logs `ReferencePDFs folder not found` and skips loading reference PDFs.
7. Save the project and reload the spreadsheet.

## Usage

1. A new menu **Wi‑Charge Evaluator** will appear.
2. Select **Start New Evaluation** and fill in device information. Only the device name is required.
3. Choose **View Technical Reference** to see transmitter specs or **About AI Mode** for details about the evaluation process.
4. Run **Test AI Connection** to verify the API key works.
5. After running an evaluation, a new sheet containing the full report is created
   and a summary row is added to the **evaluations_history** sheet.


## Settings

Open the settings panel from the spreadsheet menu via **Wi‑Charge Evaluator → Settings**.  
Each field controls how much that factor influences the final score:

* **Pain Intensity Weight (%)** – importance of the battery replacement cost or operational pain.
* **Coverage Geometry Weight (%)** – how well the device fits the transmitter’s coverage pattern.
* **Install Leverage Weight (%)** – impact of installation complexity or synergy with existing infrastructure.
* **Regulatory & Shipping Weight (%)** – hurdles related to compliance and shipping restrictions.
* **ROI Math Weight (%)** – contribution of the calculated return on investment.
* **Fleet Size Weight (%)** – relative importance of the number of units that will be deployed.
* **Strategic Halo Weight (%)** – credit for broader strategic or branding value.
* **Competitive Threat Weight (%)** – consideration of the competitive landscape.

The checkbox **Continue even if gate checks fail** lets the analysis run to completion even when the power-budget or line-of-sight gate fails.  
All values are saved as script properties so they persist across sessions.

=======

The report now also includes a brief market overview section summarizing battery challenges, the Wi-Charge angle and ROI in bullet form.
The AI will extract text from the PDFs in the `ReferencePDFs` folder and combine
them with online research when building the evaluation. To keep the system
prompt within token limits, each PDF's text is truncated to a few thousand
characters before inclusion.

## License

This project is licensed under the [MIT License](LICENSE).
