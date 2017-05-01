using System.Globalization;
using System.IO;
using System.Windows.Controls;

namespace VBASync.WPF
{
    public class FilePathValidationRule : ValidationRule
    {
        public override ValidationResult Validate(object value, CultureInfo cultureInfo)
            => new ValidationResult(string.IsNullOrEmpty(value as string) || File.Exists(value as string), null);
    }

    public class FolderPathValidationRule : ValidationRule
    {
        public override ValidationResult Validate(object value, CultureInfo cultureInfo)
            => new ValidationResult(string.IsNullOrEmpty(value as string) || Directory.Exists(value as string), null);
    }
}
