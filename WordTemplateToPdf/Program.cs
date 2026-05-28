using System.Diagnostics;
using System.Text;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace WordTemplateToPdf;

public static class Program
{
    private static readonly IReadOnlyDictionary<string, string> PlaceholderValues = new Dictionary<string, string>
    {
        ["{{LoanNumber}}"] = "LN-2026-0001",
        ["{{FirstName}}"] = "John",
        ["{{LastName}}"] = "Doe",
        ["{{PersonalIdentifierNumber}}"] = "8501011234",
        ["{{PermanentStreet}}"] = "Main Street 12",
        ["{{PermanentCity}}"] = "Bratislava",
        ["{{PermanentZipCode}}"] = "81101",
        ["{{ContactStreet}}"] = "Second Avenue 45",
        ["{{ContactCity}}"] = "Kosice",
        ["{{ContactZipCode}}"] = "04001",
        ["{{PhoneNumber}}"] = "+421900000000",
        ["{{Email}}"] = "john.doe@example.com",
        ["{{LoanAmount}}"] = "10000.00",
        ["{{FeeWithoutDiscount}}"] = "250.00",
        ["{{Fee}}"] = "200.00",
        ["{{VariableSymbol}}"] = "1234567890",
        ["{{HardDueDate}}"] = "2026-12-31",
        ["{{AmountToPayWithoutDiscount}}"] = "10250.00",
        ["{{AmountToPay}}"] = "10200.00",
        ["{{AprWithoutDiscount}}"] = "14.90%",
        ["{{Apr}}"] = "12.90%"
    };

    private static int Main()
    {
        Console.Write("Enter path to .docx template: ");
        var inputPath = Console.ReadLine();

        try
        {
            var validatedPath = ValidateInputPath(inputPath);
            Console.WriteLine("Input validated.");

            var outputPdfPath = ProcessDocument(validatedPath);
            Console.WriteLine($"PDF generated successfully: {outputPdfPath}");
            return 0;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            Console.Error.WriteLine($"I/O error: {ex.Message}");
            return 1;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
            return 1;
        }
    }

    public static string ValidateInputPath(string? inputPath)
    {
        if (string.IsNullOrWhiteSpace(inputPath))
        {
            throw new ArgumentException("Input file path is required.");
        }

        var fullPath = Path.GetFullPath(inputPath.Trim());

        if (!File.Exists(fullPath))
        {
            throw new FileNotFoundException("Input file not found.", fullPath);
        }

        var extension = Path.GetExtension(fullPath);
        if (extension.Equals(".doc", StringComparison.OrdinalIgnoreCase))
        {
            throw new NotSupportedException("Unsupported format, please convert to .docx first");
        }

        if (!extension.Equals(".docx", StringComparison.OrdinalIgnoreCase))
        {
            throw new NotSupportedException("Unsupported file extension. Only .docx is supported.");
        }

        return fullPath;
    }

    private static string ProcessDocument(string inputDocxPath)
    {
        var outputDirectory = Path.GetDirectoryName(inputDocxPath)
                              ?? throw new InvalidOperationException("Could not determine output directory.");
        var outputDocxPath = Path.Combine(outputDirectory, $"NEW_{Path.GetFileName(inputDocxPath)}");
        var outputPdfPath = Path.Combine(outputDirectory, $"{Path.GetFileNameWithoutExtension(outputDocxPath)}.pdf");

        Console.WriteLine("Preparing NEW_ document copy...");
        File.Copy(inputDocxPath, outputDocxPath, overwrite: true);

        Console.WriteLine("Replacing placeholders...");
        using (var document = DocX.Load(outputDocxPath))
        {
            foreach (var placeholder in PlaceholderValues)
            {
                document.ReplaceText(placeholder.Key, placeholder.Value);
            }

            document.Save();
        }

        Console.WriteLine($"Updated document saved: {outputDocxPath}");
        Console.WriteLine("Converting to PDF with LibreOffice Headless...");
        ConvertDocxToPdf(outputDocxPath, outputDirectory);

        if (!File.Exists(outputPdfPath))
        {
            throw new InvalidOperationException($"PDF conversion failed. Expected output file was not created: {outputPdfPath}");
        }

        return outputPdfPath;
    }

    private static void ConvertDocxToPdf(string docxPath, string outputDirectory)
    {
        var processResult = RunLibreOffice("soffice", docxPath, outputDirectory);
        if (processResult.ExitCode == 0)
        {
            return;
        }

        processResult = RunLibreOffice("libreoffice", docxPath, outputDirectory);
        if (processResult.ExitCode == 0)
        {
            return;
        }

        throw new InvalidOperationException(
            $"LibreOffice conversion failed using both 'soffice' and 'libreoffice'. " +
            $"Exit code: {processResult.ExitCode}. Error: {processResult.StandardError}");
    }

    private static ProcessResult RunLibreOffice(string executable, string docxPath, string outputDirectory)
    {
        try
        {
            using var process = new Process();
            process.StartInfo = new ProcessStartInfo
            {
                FileName = executable,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            process.StartInfo.ArgumentList.Add("--headless");
            process.StartInfo.ArgumentList.Add("--convert-to");
            process.StartInfo.ArgumentList.Add("pdf");
            process.StartInfo.ArgumentList.Add(docxPath);
            process.StartInfo.ArgumentList.Add("--outdir");
            process.StartInfo.ArgumentList.Add(outputDirectory);

            var standardOutput = new StringBuilder();
            var standardError = new StringBuilder();

            process.OutputDataReceived += (_, e) =>
            {
                if (e.Data is not null)
                {
                    standardOutput.AppendLine(e.Data);
                }
            };

            process.ErrorDataReceived += (_, e) =>
            {
                if (e.Data is not null)
                {
                    standardError.AppendLine(e.Data);
                }
            };

            if (!process.Start())
            {
                throw new InvalidOperationException($"Failed to start {executable} process.");
            }

            process.BeginOutputReadLine();
            process.BeginErrorReadLine();

            if (!process.WaitForExit(120_000))
            {
                process.Kill(entireProcessTree: true);
                throw new TimeoutException($"LibreOffice process '{executable}' timed out.");
            }

            return new ProcessResult(process.ExitCode, standardOutput.ToString(), standardError.ToString());
        }
        catch (System.ComponentModel.Win32Exception)
        {
            return new ProcessResult(-1, string.Empty, $"Executable '{executable}' not found in PATH.");
        }
    }

    private sealed record ProcessResult(int ExitCode, string StandardOutput, string StandardError);
}
