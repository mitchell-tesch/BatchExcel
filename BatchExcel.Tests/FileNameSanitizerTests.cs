using BatchExcel.Services;

namespace BatchExcel.Tests;

public class FileNameSanitizerTests
{
    [Theory]
    [InlineData("simple", "simple")]
    [InlineData("with spaces.xlsx", "with spaces.xlsx")]
    [InlineData("with/slash", "with_slash")]
    [InlineData("with\\backslash", "with_backslash")]
    [InlineData("with:colon", "with_colon")]
    [InlineData("with*asterisk", "with_asterisk")]
    [InlineData("with?question", "with_question")]
    [InlineData("with\"quote", "with_quote")]
    [InlineData("with<less>greater", "with_less_greater")]
    [InlineData("with|pipe", "with_pipe")]
    [InlineData("1_Run Name_Calculation.xlsx", "1_Run Name_Calculation.xlsx")]
    [InlineData("Sect 3/Case 2", "Sect 3_Case 2")]
    public void Sanitize_ReplacesInvalidCharactersWithUnderscore(string input, string expected)
    {
        Assert.Equal(expected, FileNameSanitizer.Sanitize(input));
    }

    [Fact]
    public void Sanitize_EmptyString_ReturnsEmpty()
    {
        Assert.Equal("", FileNameSanitizer.Sanitize(""));
    }

    [Fact]
    public void Sanitize_PreservesLength()
    {
        const string input = "a/b\\c:d*e?f\"g<h>i|j";
        var result = FileNameSanitizer.Sanitize(input);
        Assert.Equal(input.Length, result.Length);
    }

    [Theory]
    [InlineData("short.xlsx", 20, "short.xlsx")]
    [InlineData("1_VeryLongRunTitle_calc.xlsx", 15, "1_VeryLong.xlsx")]
    [InlineData("a/b\\c_calc.xlsx", 12, "a_b_c_c.xlsx")]
    public void Sanitize_WithMaxLength_ClampsAndPreservesExtension(string input, int max, string expected)
    {
        var result = FileNameSanitizer.Sanitize(input, max);
        Assert.Equal(expected, result);
        Assert.True(result.Length <= max);
    }

    [Fact]
    public void Sanitize_WithMaxLength_NoExtension_TruncatesStem()
    {
        Assert.Equal("abcde", FileNameSanitizer.Sanitize("abcdefghij", 5));
    }

    [Fact]
    public void Sanitize_WithMaxLength_ExtensionLongerThanBudget_TruncatesWhole()
    {
        // ".verylongext" is 12 chars; budget 5 forces a hard right-truncate
        Assert.Equal("a.ver", FileNameSanitizer.Sanitize("a.verylongext", 5));
    }

    [Fact]
    public void Sanitize_WithMaxLength_NonPositive_ReturnsEmpty()
    {
        Assert.Equal("", FileNameSanitizer.Sanitize("anything.xlsx", 0));
        Assert.Equal("", FileNameSanitizer.Sanitize("anything.xlsx", -3));
    }
}

