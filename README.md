# main

## WordTemplateToPdf

A .NET 8 console app that:

1. Loads a `.docx` Word template.
2. Replaces known placeholders with dummy data using `MiniWord`.
3. Converts the updated document to PDF using LibreOffice Headless (`soffice` / `libreoffice`).

### Run

```bash
cd /tmp/workspace/Ota85/main/WordTemplateToPdf
dotnet run
```

When prompted, enter the absolute path to an input `.docx` file.

The output PDF is written to the same directory as the input file, with the same filename and a `.pdf` extension.
