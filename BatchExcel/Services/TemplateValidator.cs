using BatchExcel.Models;

namespace BatchExcel.Services;

public class ValidationException : Exception
{
    public ValidationException(string message) : base(message) { }
}

public static class TemplateValidator
{
    public static void Validate(string calculationPath, BatchConfig config)
    {
        // Stub for compilation
    }
}
