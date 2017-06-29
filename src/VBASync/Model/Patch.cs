using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using VBASync.Localization;

namespace VBASync.Model
{
    public enum ChangeType
    {
        AddFile = 0,
        MoveFile = 1,
        DeleteFile = 2,
        AddSub = 3,
        DeleteSub = 4,
        AddLines = 5,
        ChangeLines = 6,
        DeleteLines = 7,
        ReplaceId = 8,
        ChangeFormControls = 9,
        DeleteFrx = 10,
        ChangeFileType = 11,
        WholeFile = 12,
        MoveSub = 13,
        Project = 14,
        Licenses = 15
    }

    public enum ModuleType
    {
        Standard = 0,
        StaticClass = 1,
        Class = 2,
        Form = 3,
        Ini = 4,
        Licenses = 5
    }

    internal struct Chunk
    {
        public int NewStartLine { get; set; }
        public string NewText { get; set; }
        public int OldStartLine { get; set; }
        public string OldText { get; set; }
    }

    internal struct SideBySideArgs
    {
        public string Name { get; set; }
        public string NewText { get; set; }
        public ModuleType NewType { get; set; }
        public string OldText { get; set; }
        public ModuleType OldType { get; set; }
    }

    public class Patch
    {
        private readonly List<Chunk> _chunks;
        private bool _commit;

        private Patch(ModuleType moduleType, string moduleName, ChangeType changeType,
                      string description = "", bool commit = true, IEnumerable<Chunk> chunks = null,
                      string sideBySideNewText = "")
        {
            ModuleType = moduleType;
            ModuleName = moduleName;
            ChangeType = changeType;
            Description = description;
            Commit = commit;
            _chunks = chunks?.ToList() ?? new List<Chunk>();
            SideBySideNewText = sideBySideNewText;
        }

        public event EventHandler<PropertyChangedEventArgs> CommitChanged;

        public ChangeType ChangeType { get; }
        public bool Commit
        {
            get { return _commit; }
            set
            {
                _commit = value;
                CommitChanged?.Invoke(null, new PropertyChangedEventArgs(nameof(Commit)));
            }
        }

        public string Description { get; }
        public string ModuleName { get; }
        public ModuleType ModuleType { get; }

        internal string SideBySideNewText { get; }

        internal static IEnumerable<Patch> CompareSideBySide(SideBySideArgs args)
        {
            var lineDiff = CountStringLines(args.NewText) - CountStringLines(args.OldText);
            var patches = new List<Patch>();
            if (args.OldType == args.NewType)
            {
                patches.Add(new Patch(args.NewType, args.Name, ChangeType.WholeFile,
                        string.Format(VBASyncResources.CDWholeFile, lineDiff.ToString("+#;-#;—")),
                        chunks: new[] { new Chunk { OldStartLine = 1, NewStartLine = 1, OldText = args.OldText, NewText = args.NewText } },
                        sideBySideNewText: args.NewText));
            }
            else
            {
                patches.Add(new Patch(args.NewType, args.Name, ChangeType.ChangeFileType,
                        string.Format(VBASyncResources.CDChangeFileType, GetModuleTypeName(args.OldType), GetModuleTypeName(args.NewType), lineDiff.ToString("+#;-#;—")),
                        chunks: new[] { new Chunk { OldStartLine = 1, NewStartLine = 1, OldText = args.OldText, NewText = args.NewText } },
                        sideBySideNewText: args.NewText));
            }
            return patches;
        }

        internal static Patch MakeDeletion(string name, ModuleType type, string module)
                => new Patch(type, name, ChangeType.DeleteFile,
                        chunks: new[] { new Chunk { OldStartLine = 1, NewStartLine = 1, OldText = module } });

        internal static Patch MakeFrxChange(string name) => new Patch(ModuleType.Form, name, ChangeType.ChangeFormControls);

        internal static Patch MakeFrxDeletion(string name) => new Patch(ModuleType.Form, name, ChangeType.DeleteFrx);

        internal static Patch MakeInsertion(string name, ModuleType type, string module)
                => new Patch(type, name, ChangeType.AddFile,
                        chunks: new[] { new Chunk { OldStartLine = 1, NewStartLine = 1, NewText = module } });

        internal static Patch MakeLicensesChange(byte[] oldBytes, byte[] newBytes)
        {
            if (oldBytes.SequenceEqual(newBytes))
            {
                return null;
            }
            else
            {
                return new Patch(ModuleType.Licenses, "LicenseKeys", ChangeType.Licenses);
            }
        }

        internal static Patch MakeProjectChange(string oldText, string newText)
        {
            if (oldText == newText)
            {
                return null;
            }
            else
            {
                var lineDiff = CountStringLines(newText) - CountStringLines(oldText);
                return new Patch(ModuleType.Ini, "Project", ChangeType.Project,
                    string.Format(VBASyncResources.CDWholeFile, lineDiff.ToString("+#;-#;—")));
            }
        }

        private static int CountStringLines(string s)
        {
            // s is guaranteed to have \r\n line endings
            return string.IsNullOrEmpty(s) ? 0 : s.Length - s.Replace("\r", "").Length;
        }

        private static string GetModuleTypeName(ModuleType mt)
        {
            switch (mt)
            {
                case ModuleType.Standard:
                    return VBASyncResources.MTStandard;
                case ModuleType.StaticClass:
                    return VBASyncResources.MTStaticClass;
                case ModuleType.Class:
                    return VBASyncResources.MTClass;
                case ModuleType.Form:
                    return VBASyncResources.MTForm;
                case ModuleType.Ini:
                    return VBASyncResources.MTIni;
                default:
                    return null;
            }
        }
    }
}
