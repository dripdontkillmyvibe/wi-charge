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
5. After running an evaluation, results are written to the **Results** sheet.

The AI will extract text from the PDFs in the `ReferencePDFs` folder and combine
them with online research when building the evaluation. To keep the system
prompt within token limits, each PDF's text is truncated to a few thousand
characters before inclusion.

## License

This project is licensed under the [MIT License](LICENSE).
