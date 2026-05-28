# main

## WordTemplateToPdf

A .NET 8 console app that:

1. Loads a `.docx` Word template.
2. Replaces known placeholders with dummy data using `DocX`.
3. Saves the updated document as a `NEW_`-prefixed `.docx` file in the same directory.

### Run

```bash
cd /tmp/workspace/Ota85/POF_Generator/WordTemplateToPdf
dotnet run
```

When prompted, enter the absolute path to an input `.docx` file.

The output document is written to the same directory as the input file, with filename prefixed by `NEW_`.
