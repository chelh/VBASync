using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace VBASync.Model
{
    internal class ProjectIni : IniFile
    {
        private readonly List<string> _moduleStrings = new List<string>();

        public ProjectIni(string filePath, Encoding encoding = null) : base(filePath, encoding)
        {
        }

        public string GetConstantsString()
            => string.Join(" : ", Dict.Keys.Where(s => s.StartsWith("Constants\0")).Select(s => $"{s.Substring(10)} = {Dict[s]}"));

        public string GetProjectText()
        {
            var sb = new StringBuilder();
            sb.AppendLine($"ID=\"{GetString("General", "ID")}\"");
            if (GetString("General", "Package") != null)
            {
                sb.AppendLine($"Package={GetString("General", "Package")}");
            }
            foreach (var modLine in _moduleStrings)
            {
                sb.AppendLine(modLine);
            }
            if (GetString("General", "HelpFile") != null)
            {
                sb.AppendLine($"HelpFile=\"{GetString("General", "HelpFile")}\"");
            }
            if (GetString("General", "ExeName32") != null)
            {
                sb.AppendLine($"ExeName32=\"{GetString("General", "ExeName32")}\"");
            }
            sb.AppendLine($"Name=\"{GetString("General", "Name")}\"");
            sb.AppendLine($"HelpContextID=\"{GetInt("General", "HelpContextID") ?? 0}\"");
            if (GetString("General", "Description") != null)
            {
                sb.AppendLine($"Description=\"{GetString("General", "Description")}\"");
            }
            if (GetString("General", "VersionCompatible32") != null)
            {
                sb.AppendLine($"VersionCompatible32=\"{GetString("General", "VersionCompatible32")}\"");
            }
            sb.AppendLine($"CMG=\"{GetString("General", "CMG")}\"");
            sb.AppendLine($"DPB=\"{GetString("General", "DPB")}\"");
            sb.AppendLine($"GC=\"{GetString("General", "GC")}\"");
            sb.AppendLine("");
            sb.AppendLine("[Host Extender Info]");
            var i = 0;
            while (GetString("Host Extender Info", $"&H{(++i).ToString("X8")}") != null)
            {
                sb.AppendLine($"&H{i.ToString("X8")}={GetString("Host Extender Info", $"&H{i.ToString("X8")}")}");
            }
            return sb.ToString();
        }

        public IEnumerable<string> GetReferenceNames()
            => Dict.Keys.Where(s => s.StartsWith("Reference ", StringComparison.InvariantCultureIgnoreCase))
                    .Select(s => s.Split('\0')[0].Substring(10)).Distinct();

        public void RegisterModule(string name, ModuleType type, uint version = 0)
        {
            switch (type)
            {
            case ModuleType.StaticClass:
                _moduleStrings.Add($"Document={name}/&H{version.ToString("X8")}");
                break;
            case ModuleType.Class:
                _moduleStrings.Add($"Class={name}");
                break;
            case ModuleType.Form:
                _moduleStrings.Add($"BaseClass={name}");
                break;
            case ModuleType.Standard:
                _moduleStrings.Add($"Module={name}");
                break;
            default:
                throw new ArgumentException("Invalid type for this function.", nameof(type));
            }
        }
    }
}
