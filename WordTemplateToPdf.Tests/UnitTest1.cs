namespace WordTemplateToPdf.Tests;

public class UnitTest1
{
    [Fact]
    public void ValidateInputPath_WhenDocFile_ThrowsExplicitUnsupportedMessage()
    {
        var tempFilePath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.doc");
        File.WriteAllText(tempFilePath, "placeholder");

        try
        {
            var exception = Assert.Throws<NotSupportedException>(() => WordTemplateToPdf.Program.ValidateInputPath(tempFilePath));
            Assert.Equal("Unsupported format, please convert to .docx first", exception.Message);
        }
        finally
        {
            if (File.Exists(tempFilePath))
            {
                File.Delete(tempFilePath);
            }
        }
    }

    [Fact]
    public void ValidateInputPath_WhenDocxFileExists_ReturnsFullPath()
    {
        var tempFilePath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.docx");
        File.WriteAllText(tempFilePath, "placeholder");

        try
        {
            var result = WordTemplateToPdf.Program.ValidateInputPath(tempFilePath);
            Assert.Equal(Path.GetFullPath(tempFilePath), result);
        }
        finally
        {
            if (File.Exists(tempFilePath))
            {
                File.Delete(tempFilePath);
            }
        }
    }

    [Theory]
    [InlineData("[[LoanNumber]]", "LoanNumber")]
    [InlineData("{{LoanNumber}}", "LoanNumber")]
    [InlineData("LoanNumber", "LoanNumber")]
    public void NormalizeTemplateKey_SupportsSquareAndCurlyDelimiters(string key, string expected)
    {
        var result = WordTemplateToPdf.Program.NormalizeTemplateKey(key);

        Assert.Equal(expected, result);
    }

    [Fact]
    public void CreateTemplateData_UsesNormalizedKeysWithoutDelimiters()
    {
        var result = WordTemplateToPdf.Program.CreateTemplateData();

        Assert.Contains("LoanNumber", result.Keys);
        Assert.DoesNotContain("[[LoanNumber]]", result.Keys);
        Assert.Equal("LN-2026-0001", Assert.IsType<string>(result["LoanNumber"]));
    }
}
