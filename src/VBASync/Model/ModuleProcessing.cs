using System;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using VBASync.Localization;

namespace VBASync.Model
{
    public static class ModuleProcessing
    {
        private static bool _inBlock;

        public static string ExtensionFromType(ModuleType type)
        {
            switch (type)
            {
            case ModuleType.Class:
                return ".cls";
            case ModuleType.Form:
                return ".frm";
            case ModuleType.Standard:
                return ".bas";
            case ModuleType.StaticClass:
                return ".cls";
            case ModuleType.Ini:
                return ".ini";
            case ModuleType.Licenses:
                return ".bin";
            default:
                throw new ApplicationException(VBASyncResources.ErrorUnrecognizedModuleType);
            }
        }

        public static string FixCase(string oldString, string newString)
        {
            var sb = new StringBuilder();
            var dr = VbaDiffer.CreateVbaDiffs(oldString, newString);
            var aPos = 0;
            var bPos = 0;
            foreach (var db in dr.DiffBlocks)
            {
                while (bPos < db.InsertStartB)
                {
                    // don't use AppendLine since we want to always have DOS line endings
                    if (sb.Length > 0)
                    {
                        sb.Append("\r\n");
                    }
                    sb.Append(dr.PiecesOld[aPos++]);
                    bPos++;
                }
                var i = 0;
                aPos += db.DeleteCountA;
                while (i < (db.DeleteCountA < db.InsertCountB ? db.DeleteCountA : db.InsertCountB))
                {
                    var isVbBaseLine = (dr.PiecesNew[i + db.InsertStartB].StartsWith("Attribute VB_Base =", StringComparison.InvariantCultureIgnoreCase)
                                        || dr.PiecesNew[i + db.InsertStartB].StartsWith("Attribute VB_Base=", StringComparison.InvariantCultureIgnoreCase))
                                       && !dr.PiecesNew[i + db.InsertStartB].EndsWith("_")
                                       && (dr.PiecesOld[i + db.DeleteStartA].StartsWith("Attribute VB_Base =", StringComparison.InvariantCultureIgnoreCase)
                                           || dr.PiecesOld[i + db.DeleteStartA].StartsWith("Attribute VB_Base=", StringComparison.InvariantCultureIgnoreCase))
                                       && !dr.PiecesOld[i + db.DeleteStartA].EndsWith("_");
                    // don't use AppendLine since we want to always have DOS line endings
                    if (sb.Length > 0)
                    {
                        sb.Append("\r\n");
                    }
                    sb.Append(isVbBaseLine ? dr.PiecesOld[i++ + db.DeleteStartA] : dr.PiecesNew[i++ + db.InsertStartB]);
                    bPos++;
                }
                if (db.InsertCountB > db.DeleteCountA)
                {
                    while (i < db.InsertCountB)
                    {
                        // don't use AppendLine since we want to always have DOS line endings
                        if (sb.Length > 0)
                        {
                            sb.Append("\r\n");
                        }
                        sb.Append(dr.PiecesNew[i++ + db.InsertStartB]);
                        bPos++;
                    }
                }
            }
            while (aPos < dr.PiecesOld.Length)
            {
                // don't use AppendLine since we want to always have DOS line endings
                if (sb.Length > 0)
                {
                    sb.Append("\r\n");
                }
                sb.Append(dr.PiecesOld[aPos++]);
            }
            return sb.ToString();
        }

        internal static bool HasCode(string moduleText)
        {
            _inBlock = false;
            return moduleText?.Split(new[] { "\r\n" }, StringSplitOptions.None).Any(LineIsCode) ?? false;
        }

        internal static string StubOut(string moduleText)
        {
            _inBlock = false;
            var lines = moduleText?.Split(new[] { "\r\n" }, StringSplitOptions.None) ?? new string[0];
            var sb = new StringBuilder();
            for (var i = 0; i < lines.Length; ++i)
            {
                if (LineIsCode(lines[i]))
                {
                    return sb.ToString();
                }
                sb.Append(lines[i]).Append("\r\n"); // don't use AppendLine() because we need \r\n line endings specifically
            }
            return sb.ToString();
        }

        internal static ModuleType TypeFromText(string moduleText)
        {
            if (Regex.IsMatch(moduleText, @"^\s*(\s_\r\n\s*)*Version\s*(\s_\r\n\s*)*5.00\s*(\s_\r\n\s*)*$",
                RegexOptions.Multiline | RegexOptions.IgnoreCase))
            {
                return ModuleType.Form;
            }
            if (!Regex.IsMatch(moduleText, @"^\s*(\s_\r\n\s*)*Version\s*(\s_\r\n\s*)*1.0\s*(\s_\r\n\s*)*Class\s*(\s_\r\n\s*)*$",
                RegexOptions.Multiline | RegexOptions.IgnoreCase))
            {
                // matches anything other than a class
                return ModuleType.Standard;
            }
            if (Regex.IsMatch(moduleText, @"^\s*(\s_\r\n\s*)*Attribute\s*(\s_\r\n\s*)*VB_PredeclaredID\s*(\s_\r\n\s*)*=\s*(\s_\r\n\s*)*True\s*(\s_\r\n\s*)*$",
                RegexOptions.Multiline | RegexOptions.IgnoreCase))
            {
                return ModuleType.StaticClass;
            }
            return ModuleType.Class;
        }

        private static bool LineIsCode(string line)
        {
            // must set _inBlock to false before using this
            var trimLine = line?.TrimStart();
            if (string.IsNullOrWhiteSpace(trimLine) || trimLine.StartsWith("Version", StringComparison.InvariantCultureIgnoreCase)
                || trimLine.StartsWith("Attribute", StringComparison.InvariantCultureIgnoreCase))
            {
                return false;
            }
            if (trimLine.StartsWith("Begin", StringComparison.InvariantCultureIgnoreCase))
            {
                _inBlock = true;
                return false;
            }
            if (_inBlock && trimLine.StartsWith("End", StringComparison.InvariantCultureIgnoreCase))
            {
                _inBlock = false;
                return false;
            }
            return !_inBlock;
        }
    }
}
