using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;

namespace VbaSync {
    public enum ChangeType {
        AddFile,
        MoveFile,
        DeleteFile,
        AddSub,
        DeleteSub,
        AddLines,
        ChangeLines,
        DeleteLines,
        ReplaceId,
        ChangeFormControls,
        DeleteFrx,
        ChangeFileType,
        WholeFile,
        MoveSub,
        Project
    }

    public enum ModuleType {
        Standard,
        StaticClass,
        Class,
        Form,
        Ini
    }

    internal struct Chunk {
        public int NewStartLine;
        public string NewText;
        public int OldStartLine;
        public string OldText;
    }

    internal struct SideBySideArgs {
        public string Name;
        public string NewText;
        public ModuleType NewType;
        public string OldText;
        public ModuleType OldType;
    }

    public class Patch {
        private readonly List<Chunk> _chunks;
        private bool _commit;

        private Patch(ModuleType moduleType, string moduleName, ChangeType changeType,
                      string description = "", bool commit = true, IEnumerable<Chunk> chunks = null) {
            ModuleType = moduleType;
            ModuleName = moduleName;
            ChangeType = changeType;
            Description = description;
            Commit = commit;
            _chunks = chunks?.ToList() ?? new List<Chunk>();
        }

        public event PropertyChangedEventHandler CommitChanged;

        public ChangeType ChangeType { get; }
        public bool Commit {
            get { return _commit; }
            set {
                _commit = value;
                CommitChanged?.Invoke(null, new PropertyChangedEventArgs(nameof(Commit)));
            }
        }

        public string Description { get; }
        public string ModuleName { get; }
        public ModuleType ModuleType { get; }

        internal static IEnumerable<Patch> CompareSideBySide(SideBySideArgs args) {
            var lineDiff = CountStringLines(args.NewText) - CountStringLines(args.OldText);
            var patches = new List<Patch>();
            if (args.OldType == args.NewType) {
                patches.Add(new Patch(args.NewType, args.Name, ChangeType.WholeFile,
                        $"File changed ({lineDiff.ToString("+#;-#;—")})",
                        chunks: new[] { new Chunk { OldStartLine = 1, NewStartLine = 1, OldText = args.OldText, NewText = args.NewText } }));
            } else {
                patches.Add(new Patch(args.NewType, args.Name, ChangeType.ChangeFileType,
                        $"{GetEnumDescription(args.OldType)} → {GetEnumDescription(args.NewType)} and file changed ({lineDiff.ToString("+#;-#;—")})",
                        chunks: new[] { new Chunk { OldStartLine = 1, NewStartLine = 1, OldText = args.OldText, NewText = args.NewText } }));
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

        internal static Patch MakeProjectChange(string oldText, string newText) {
            if (oldText == newText) {
                return null;
            } else {
                var lineDiff = CountStringLines(newText) - CountStringLines(oldText);
                return new Patch(ModuleType.Ini, "Project", ChangeType.Project, $"File changed ({lineDiff.ToString("+#;-#;—")})");
            }
        }

        private static int CountStringLines(string s) {
            int i = 0;
            using (var r = new StringReader(s)) {
                while (r.ReadLine() != null) {
                    i++;
                }
            }
            return i;
        }

        private static string GetEnumDescription(object v) {
            dynamic attr = v.GetType().GetMember(v.ToString())[0]
                    ?.GetCustomAttributes(typeof(DescriptionAttribute), false)[0];
            return attr?.Description ?? v.ToString();
        }
    }
}
