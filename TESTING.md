# Verification — 29 October 2026

The final `test.ps1` run passed with exit code 0. See `test-log.txt` for the complete execution log. The executable version is 1.0.0.0.

Environment: Windows, .NET Framework compiler, Microsoft Word and PowerPoint 16.0.20527.20088, and the locally available Poppler PDF renderer. LibreOffice is not installed here.

| Route | Result | Checks |
| --- | --- | --- |
| DOCX → PDF | Passed | PDF signature, page rendering, expected paragraphs visually verified |
| DOCX → PPTX | Passed | Valid package, expected slide text, opened in PowerPoint |
| PPTX → PDF | Passed | PDF signature, page rendering, expected text and picture visually verified |
| PPTX → DOCX | Passed | Valid package, expected text and embedded picture, opened in Word |
| PDF → DOCX | Passed | Valid package, expected text, opened in Word |
| PDF → PPTX | Passed | Valid package, one embedded page image, expected slide count, opened in PowerPoint |

Fixtures contain two Word paragraphs, a PowerPoint slide with text and a picture, and a one-page PDF exported by Word. PDF text extraction was not available in this Poppler installation; PDF exports were rendered and visually inspected instead. These are representative smoke tests, not a claim of universal layout fidelity.

Additional passing checks:

- Source SHA-256 hashes unchanged after all six routes.
- Existing outputs retained; new outputs automatically renamed.
- Duplicate rejection when automatic renaming is disabled.
- Corrupted DOCX, unsupported extension, missing input, invalid output folder, missing configured engine, and missing configured renderer fail clearly.
- Actual WinForms batch queue: valid DOCX, corrupted DOCX, valid PPTX. Results are Completed / Failed / Completed, with final progress 3/3.
- A UI timer fired 123 times during that final batch, confirming that the message loop continued processing while conversions ran.
- Deliberately stalled test engine: timeout, no published output, worker and engine processes cleaned up.
- Converter and Settings tabs rendered and visually inspected; previews were inspected during verification. Generated test files and previews were subsequently removed during cleanup; rerun the tests to regenerate them.

The initial console UI harness used a manually pumped message loop, which lost the WinForms synchronization context. The final harness uses `Application.Run`, matching the production application. The final complete run passes.

Not verified here: LibreOffice integration, machines without Microsoft Office, other Office languages, encrypted files, scan OCR (not implemented), complex or unusually large documents, mixed-size/multiple-page PDFs, and intentionally injected permission failures. Handling is present for ordinary permission exceptions, but that failure mode was not explicitly forced during verification.
