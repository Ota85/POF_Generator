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

            var outputDocxPath = ProcessDocument(validatedPath);
            Console.WriteLine($"Document generated successfully: {outputDocxPath}");
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
        return outputDocxPath;
    }
}
