using System.Diagnostics;
using System.Linq.Expressions;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using MiniSoftware;

namespace WordTemplateToPdf;

public static class Program
{
    private const string TemplateTagPrefix = "[[";
    private const string TemplateTagSuffix = "]]";
    private const string MiniWordTagPrefix = "{{";
    private const string MiniWordTagSuffix = "}}";
    private const string UseLegacyOpenXmlPreprocessingEnvVar = "WORD_TEMPLATE_USE_OPENXML_PREPROCESSING";

    private static readonly IReadOnlyDictionary<string, string> PlaceholderValues = new Dictionary<string, string>
    {
        ["[[LoanNumber]]"] = "LN-2026-0001",
        ["[[FirstName]]"] = "John",
        ["[[LastName]]"] = "Doe",
        ["[[PersonalIdentifierNumber]]"] = "8501011234",
        ["[[PermanentStreet]]"] = "Main Street 12",
        ["[[PermanentCity]]"] = "Bratislava",
        ["[[PermanentZipCode]]"] = "81101",
        ["[[ContactStreet]]"] = "Second Avenue 45",
        ["[[ContactCity]]"] = "Kosice",
        ["[[ContactZipCode]]"] = "04001",
        ["[[PhoneNumber]]"] = "+421900000000",
        ["[[Email]]"] = "john.doe@example.com",
        ["[[LoanAmount]]"] = "10000.00",
        ["[[FeeWithoutDiscount]]"] = "250.00",
        ["[[Fee]]"] = "200.00",
        ["[[VariableSymbol]]"] = "1234567890",
        ["[[HardDueDate]]"] = "2026-12-31",
        ["[[AmountToPayWithoutDiscount]]"] = "10250.00",
        ["[[AmountToPay]]"] = "10200.00",
        ["[[AprWithoutDiscount]]"] = "14.90%",
        ["[[Apr]]"] = "12.90%"
    };

    private static int Main()
    {
        Console.Write("Enter .docx template filename: ");    
        var inputPath = Console.ReadLine();
        inputPath = @"c:\Users\Ota\Desktop\LinkSoft\PDF_generator\" + inputPath + @".docx";

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
        Console.WriteLine("Replacing placeholders...");

        var templateData = CreateTemplateData();

        if (!UseLegacyOpenXmlTemplatePreprocessing() && TryConfigureMiniWordTags())
        {
            MiniWord.SaveAsByTemplate(outputDocxPath, inputDocxPath, templateData);
        }
        else
        {
            var miniWordTemplateBytes = PrepareMiniWordTemplate(inputDocxPath);
            MiniWord.SaveAsByTemplate(outputDocxPath, miniWordTemplateBytes, templateData);
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
        var processResult = RunLibreOffice(@"C:\Program Files\LibreOffice\program\soffice.exe", docxPath, outputDirectory);
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

    internal static IReadOnlyDictionary<string, object> CreateTemplateData()
    {
        return PlaceholderValues.ToDictionary(
            placeholder => NormalizeTemplateKey(placeholder.Key),
            placeholder => (object)placeholder.Value);
    }

    private static bool TryConfigureMiniWordTags()
    {
        var configureMethod = typeof(MiniWord).GetMethod(
            "Configure",
            System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);

        if (configureMethod is null)
        {
            return false;
        }

        var configureParameterType = configureMethod.GetParameters().SingleOrDefault()?.ParameterType;
        if (configureParameterType is null ||
            !configureParameterType.IsGenericType ||
            configureParameterType.GetGenericTypeDefinition() != typeof(Action<>))
        {
            return false;
        }

        var optionsType = configureParameterType.GetGenericArguments()[0];
        var configureTagsMethod = optionsType.GetMethod(
            "ConfigureTags",
            System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance,
            binder: null,
            types: [typeof(string), typeof(string)],
            modifiers: null);

        if (configureTagsMethod is null)
        {
            return false;
        }

        var optionsParameter = Expression.Parameter(optionsType, "options");
        var configureTagsCall = Expression.Call(
            optionsParameter,
            configureTagsMethod,
            Expression.Constant(TemplateTagPrefix),
            Expression.Constant(TemplateTagSuffix));
        var configurationDelegate = Expression
            .Lambda(configureParameterType, configureTagsCall, optionsParameter)
            .Compile();

        configureMethod.Invoke(null, [configurationDelegate]);
        return true;
    }

    internal static bool UseLegacyOpenXmlTemplatePreprocessing()
    {
        return ShouldUseLegacyOpenXmlPreprocessing(
            Environment.GetEnvironmentVariable(UseLegacyOpenXmlPreprocessingEnvVar));
    }

    internal static bool ShouldUseLegacyOpenXmlPreprocessing(string? configuredValue)
    {
        return string.Equals(configuredValue, "1", StringComparison.OrdinalIgnoreCase) ||
               string.Equals(configuredValue, "true", StringComparison.OrdinalIgnoreCase);
    }

    // Legacy fallback for split-run placeholder handling; disabled by default.
    private static byte[] PrepareMiniWordTemplate(string inputDocxPath)
    {
        var templateBytes = File.ReadAllBytes(inputDocxPath);
        using var stream = new MemoryStream();
        stream.Write(templateBytes, 0, templateBytes.Length);
        stream.Position = 0;

        using (var document = WordprocessingDocument.Open(stream, true))
        {
            ReplaceTemplateDelimiters(document.MainDocumentPart?.Document);

            foreach (var headerPart in document.MainDocumentPart?.HeaderParts ?? Enumerable.Empty<HeaderPart>())
            {
                ReplaceTemplateDelimiters(headerPart.Header);
            }

            foreach (var footerPart in document.MainDocumentPart?.FooterParts ?? Enumerable.Empty<FooterPart>())
            {
                ReplaceTemplateDelimiters(footerPart.Footer);
            }

            document.Save();
        }

        return stream.ToArray();
    }

    private static void ReplaceTemplateDelimiters(OpenXmlPartRootElement? rootElement)
    {
        if (rootElement is null)
        {
            return;
        }

        MergeSplitTemplateTags(rootElement);

        foreach (var text in rootElement.Descendants<Text>())
        {
            text.Text = text.Text
                .Replace(TemplateTagPrefix, MiniWordTagPrefix, StringComparison.Ordinal)
                .Replace(TemplateTagSuffix, MiniWordTagSuffix, StringComparison.Ordinal);
        }
    }

    private static void MergeSplitTemplateTags(OpenXmlElement xmlElement)
    {
        var textNodes = xmlElement.Descendants<Text>().ToList();
        var pendingTextNodes = new List<Text>();
        var mergedText = new StringBuilder();
        var insideTag = false;

        foreach (var textNode in textNodes)
        {
            var completedTag = false;
            if (textNode.InnerText.TrimStart().StartsWith("[", StringComparison.Ordinal))
            {
                insideTag = true;
            }

            if (insideTag)
            {
                mergedText.Append(textNode.InnerText);
                pendingTextNodes.Add(textNode);

                var candidate = mergedText.ToString().TrimStart();
                var foreachBalanced = candidate.Split(new[] { "[[foreach" }, StringSplitOptions.None).Length - 1 ==
                                      candidate.Split(new[] { "endforeach]]" }, StringSplitOptions.None).Length - 1;
                var ifBalanced = candidate.Split(new[] { "[[if" }, StringSplitOptions.None).Length - 1 ==
                                 candidate.Split(new[] { "endif]]" }, StringSplitOptions.None).Length - 1;
                var hasFullTag = candidate.StartsWith(TemplateTagPrefix, StringComparison.Ordinal) &&
                                 candidate.Contains(TemplateTagSuffix, StringComparison.Ordinal);

                if (foreachBalanced && ifBalanced && hasFullTag)
                {
                    if (mergedText.Length <= 1000)
                    {
                        var firstTextNode = pendingTextNodes[0];
                        var mergedNode = (Text)firstTextNode.CloneNode(true);
                        mergedNode.Text = candidate;
                        firstTextNode.Parent?.InsertBefore(mergedNode, firstTextNode);

                        foreach (var pendingTextNode in pendingTextNodes)
                        {
                            pendingTextNode.Text = string.Empty;
                        }
                    }

                    completedTag = true;
                }
            }

            if (completedTag)
            {
                mergedText.Clear();
                pendingTextNodes.Clear();
                insideTag = false;
            }
        }
    }

    internal static string NormalizeTemplateKey(string key)
    {
        var normalized = key.Trim();
        if (HasWrappingDelimiters(normalized, TemplateTagPrefix, TemplateTagSuffix) ||
            HasWrappingDelimiters(normalized, MiniWordTagPrefix, MiniWordTagSuffix))
        {
            return normalized[2..^2].Trim();
        }

        return normalized;
    }

    private static bool HasWrappingDelimiters(string value, string prefix, string suffix)
    {
        return value.StartsWith(prefix, StringComparison.Ordinal) &&
               value.EndsWith(suffix, StringComparison.Ordinal) &&
               value.Length > prefix.Length + suffix.Length;
    }

    private sealed record ProcessResult(int ExitCode, string StandardOutput, string StandardError);
}
