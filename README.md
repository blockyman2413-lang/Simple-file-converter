# RianFile 1.0

A compact Windows Forms converter for DOCX, PPTX and PDF. Documents stay on your computer. No account, API key, upload, or background service is used.

## Run and build

Run `..\Ready\RianFile.exe` (or `bin\RianFile.exe` after building from source) on Windows with .NET Framework 4.8. Windows 10/11 normally already include a compatible framework. Build from source with `powershell -ExecutionPolicy Bypass -File .\build.ps1`; this uses the Windows framework compiler and needs no SDK or NuGet downloads. The SDK project is also provided for Visual Studio.

## Local conversion engines

| Conversion | Engine | Output behavior |
| --- | --- | --- |
| DOCX → PDF | LibreOffice when configured; otherwise Microsoft Word | Document PDF export |
| PPTX → PDF | LibreOffice when configured; otherwise Microsoft PowerPoint | Presentation PDF export |
| DOCX → PPTX | Word and PowerPoint | Text is divided into simple slides; original layout, images and styling are not retained |
| PPTX → DOCX | PowerPoint and Word | Slide headings, text, tables as text, and picture shapes; charts and other embedded objects are omitted |
| PDF → DOCX | Microsoft Word PDF import | Editable reconstruction where possible; complex layouts may change and scanned pages are not OCR processed |
| PDF → PPTX | Poppler `pdftoppm.exe` and PowerPoint | One rendered image per page; mixed page sizes are fitted without cropping |

RianFile itself is free. Microsoft Office is a separately licensed, locally installed dependency for the Office-based routes. LibreOffice is an optional free engine for the two PDF export routes. It is not bundled. Set the full path to `soffice.exe` in Settings to use it.

For PDF → PPTX, install Poppler for Windows and set the full path to `pdftoppm.exe` in Settings. It can also be discovered on PATH or in `bin\tools`. Keep its accompanying DLLs in the distribution supplied by its provider. Poppler is not bundled or redistributed here.

Close PowerPoint before starting an Office-based presentation conversion. PowerPoint can reuse an existing interactive process, so RianFile refuses to attach to it. Close existing Word PDF import dialogs before PDF → DOCX as well. The converter acknowledges Word's English-language PDF import information dialog only in its tracked conversion processes; it does not change Word's registry preferences. Other Office languages may require acknowledging this information dialog manually before the job times out. Missing engines fail with a visible per-file error; no substitute or renamed file is produced. Same-format jobs are marked Unsupported and skipped.

## Use

1. Add files or drag them onto the list. Multiple selection is supported.
2. Choose **Convert To**, then select an output folder. Alternatively enable **Save beside original files** in Settings.
3. Click **Convert**. Files run sequentially in background worker processes. Failed jobs do not stop later jobs. The progress bar counts finished jobs, including failures.
4. Double-click a row for its full error or output path. Use **Open Folder** to open the last successful output folder, or the selected folder before a run.

Sources are copied into a private temporary working directory and opened read-only. Completed outputs are validated and moved into place without replacing existing files. Duplicate filenames receive `(1)`, `(2)`, etc. Turning off automatic renaming makes a collision fail instead of overwrite. Temporary work is removed after success or failure. Each job has a timeout; tracked worker/engine processes are cleaned up without terminating pre-existing Office sessions.

Settings are stored in `%LOCALAPPDATA%\RianFile\settings.xml`. The default folder in Settings and the output-folder field are synchronized. Timeout is configurable from 10 to 3600 seconds. Closing the app during a running queue is blocked so work can finish or time out cleanly.

## Verification

Run `powershell -ExecutionPolicy Bypass -File .\test.ps1` with Word, PowerPoint and Poppler available and PowerPoint closed. It creates real fixtures and checks all six routes, output signatures/packages, reopening Office documents, rendered PDF output, expected text/images, source hashes, duplicate names, and failure cases. It also exercises the actual WinForms queue with a damaged document between valid jobs, pumps the UI message loop, and verifies final progress. A deliberately stalled test engine checks timeout and process cleanup; it produces no converted files. Tests generate fixtures and outputs in `test-results`; the execution log is `test-log.txt`. Generated fixtures were removed after verification to keep this project clean.

The engine choices follow [LibreOffice's PDF export documentation](https://help.libreoffice.org/latest/en-US/text/shared/guide/pdf_params.html) and [Microsoft's description of PDF import and its layout limitations](https://support.microsoft.com/en-gb/word/opening-pdfs-in-word?ver=15).

This is a deliberately simple converter, not a fidelity guarantee for every document. Password-protected files, damaged files, unusual fonts, complex graphics and scan-only PDFs can fail or lose fidelity. Test important documents before relying on the output. See `TESTING.md` for the actual verification results on this machine.
