using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using OpenMcdf;
using VbaSync.FormObjects;

namespace VbaSync {
    public static class Lib {
        public static IList<KeyValuePair<string, Tuple<string, ModuleType>>> GetFolderModules(string folderPath) {
            var modulesText = new Dictionary<string, Tuple<string, ModuleType>>();
            var extensions = new[] { ".bas", ".cls", ".frm" };
            var projIni = new ProjectIni(Path.Combine(folderPath, "Project.INI"));
            if (File.Exists(Path.Combine(folderPath, "Project.INI.local"))) {
                projIni.AddFile(Path.Combine(folderPath, "Project.INI.local"));
            }
            var projEncoding = Encoding.GetEncoding(projIni.GetInt("General", "CodePage") ?? Encoding.Default.CodePage);
            foreach (var filePath in Directory.GetFiles(folderPath, "*.*").Where(s => extensions.Any(s.EndsWith)).Select(s => Path.Combine(folderPath, Path.GetFileName(s)))) {
                var moduleText = File.ReadAllText(filePath, projEncoding).TrimEnd('\r', '\n') + "\r\n";
                modulesText[Path.GetFileNameWithoutExtension(filePath)] = Tuple.Create(moduleText, ModuleProcessing.TypeFromText(moduleText));
            }
            return modulesText.ToList();
        }

        public static IEnumerable<Patch> GetModulePatches(ISession session, string vbaFolderPath,
            IList<KeyValuePair<string, Tuple<string, ModuleType>>> folderModules,
            IList<KeyValuePair<string, Tuple<string, ModuleType>>> fileModules) {
            IList<KeyValuePair<string, Tuple<string, ModuleType>>> oldModules;
            IList<KeyValuePair<string, Tuple<string, ModuleType>>> newModules;
            string oldFolder;
            string newFolder;
            if (session.Action == ActionType.Extract) {
                oldModules = folderModules;
                newModules = fileModules;
                oldFolder = session.FolderPath;
                newFolder = vbaFolderPath;
            } else {
                oldModules = fileModules;
                newModules = folderModules;
                oldFolder = vbaFolderPath;
                newFolder = session.FolderPath;
            }

            // find modules which aren't in both lists and record them as new/deleted
            var patches = new List<Patch>();
            patches.AddRange(GetDeletedModuleChanges(oldModules, newModules));
            patches.AddRange(GetNewModuleChanges(oldModules, newModules));

            // this also filters the new/deleted modules from the last step
            var sideBySide = (from o in oldModules
                              join n in newModules on o.Key equals n.Key
                              select new SideBySideArgs {
                                  Name = o.Key, OldType = o.Value.Item2, NewType = n.Value.Item2,
                                  OldText = o.Value.Item1, NewText = n.Value.Item1
                              }).ToArray();
            patches.AddRange(GetFrxChanges(oldFolder, newFolder, sideBySide.Where(x => x.NewType == ModuleType.Form).Select(x => x.Name)));
            sideBySide = sideBySide.Where(sxs => sxs.OldText != sxs.NewText).ToArray();
            foreach (var sxs in sideBySide) {
                patches.AddRange(Patch.CompareSideBySide(sxs));
            }

            return patches;
        }

        public static Patch GetProjectPatch(ISession session, string evfPath) {
            var folderIniPath = Path.Combine(session.FolderPath, "Project.ini");
            var fileIniPath = Path.Combine(evfPath, "Project.ini");
            var folderHasIni = File.Exists(folderIniPath);
            var fileHasIni = File.Exists(fileIniPath);
            if (folderHasIni && fileHasIni) {
                if (session.Action == ActionType.Extract) {
                    return Patch.MakeProjectChange(File.ReadAllText(folderIniPath), File.ReadAllText(fileIniPath));
                } else {
                    return Patch.MakeProjectChange(File.ReadAllText(fileIniPath), File.ReadAllText(folderIniPath));
                }
            } else if (!folderHasIni && !fileHasIni) {
                return null;
            } else if (session.Action == ActionType.Extract ? !folderHasIni : !fileHasIni) {
                return Patch.MakeInsertion("Project", ModuleType.Ini, File.ReadAllText(session.Action == ActionType.Extract ? fileIniPath : folderIniPath));
            } else {
                return null; // never suggest deleting Project.INI
            }
        }

        static IEnumerable<Patch> GetFrxChanges(string oldFolder, string newFolder, IEnumerable<string> frmModules) {
            foreach (var modName in frmModules) {
                var oldFrxPath = Path.Combine(oldFolder, modName + ".frx");
                var newFrxPath = Path.Combine(newFolder, modName + ".frx");
                string explain;
                if (!File.Exists(oldFrxPath) && !File.Exists(newFrxPath)) {
                } else if (File.Exists(oldFrxPath) && !File.Exists(newFrxPath)) {
                    yield return Patch.MakeFrxDeletion(modName);
                } else if (!File.Exists(oldFrxPath) && File.Exists(newFrxPath)) {
                    yield return Patch.MakeFrxChange(modName);
                } else if (FrxFilesAreDifferent(oldFrxPath, newFrxPath, out explain)) {
                    yield return Patch.MakeFrxChange(modName);
                }
            }
        }

        public static bool FrxFilesAreDifferent(string frxPath1, string frxPath2, out string explain) {
            var frxBytes1 = File.ReadAllBytes(frxPath1);
            var frxBytes2 = File.ReadAllBytes(frxPath2);
            using (var frxCfStream1 = new MemoryStream(frxBytes1, 24, frxBytes1.Length - 24, false))
            using (var frxCfStream2 = new MemoryStream(frxBytes2, 24, frxBytes2.Length - 24, false))
            using (var cf1 = new CompoundFile(frxCfStream1))
            using (var cf2 = new CompoundFile(frxCfStream2)) {
                return CfStoragesAreDifferent(cf1.RootStorage, cf2.RootStorage, out explain);
            }
        }

        static bool CfStoragesAreDifferent(CFStorage s1, CFStorage s2, out string explain) {
            var s1Names = new List<Tuple<string, bool>>();
            var s2Names = new List<Tuple<string, bool>>();
            s1.VisitEntries(i => s1Names.Add(Tuple.Create(i.Name, i.IsStorage)), false);
            s2.VisitEntries(i => s2Names.Add(Tuple.Create(i.Name, i.IsStorage)), false);
            s1Names.Sort();
            s2Names.Sort();
            if (!s1Names.SequenceEqual(s2Names)) {
                explain = $"Different file lists in storage '{s1.Name}'.\r\nFile 1: {{'{string.Join("', '", s1Names)}'}}\r\nFile 2: {{'{string.Join("'', '", s2Names)}'}}.";
                return true;
            }
            foreach (var t in s1Names) {
                if (t.Item2) {
                    if (CfStoragesAreDifferent(s1.GetStorage(t.Item1), s2.GetStorage(t.Item1), out explain)) {
                        return true;
                    }
                } else if (t.Item1 == "f") {
                    if (FStreamsAreDifferent(s1.GetStream(t.Item1).GetData(), s2.GetStream(t.Item1).GetData(), s1.Name, out explain)) {
                        return true;
                    }
                } else if (!s1.GetStream(t.Item1).GetData().SequenceEqual(s2.GetStream(t.Item1).GetData())) {
                    explain = $"Different contents of stream '{t.Item1}' in storage '{s1.Name}'.";
                    return true;
                }
            }
            explain = "No differences found.";
            return false;
        }

        static bool FStreamsAreDifferent(byte[] b1, byte[] b2, string storageName, out string explain) {
            explain = $"Different contents of stream 'f' in storage '{storageName}'.";
            if (b1.Length != b2.Length) {
                return true;
            }
            var fc1 = new FormControl(b1);
            var fc2 = new FormControl(b2);
            if (!fc1.Equals(fc2)) {
                return true;
            }
            explain = "No differences found.";
            return false;
        }

        static IEnumerable<Patch> GetDeletedModuleChanges(
                IList<KeyValuePair<string, Tuple<string, ModuleType>>> oldModules,
                IList<KeyValuePair<string, Tuple<string, ModuleType>>> newModules) {
            var deleted = oldModules.Select(kvp => kvp.Key).Except(newModules.Select(kvp => kvp.Key));
            return oldModules.Where(kvp => deleted.Contains(kvp.Key)).Select(m => Patch.MakeDeletion(m.Key, m.Value.Item2, m.Value.Item1));
        }

        static IEnumerable<Patch> GetNewModuleChanges(
                IList<KeyValuePair<string, Tuple<string, ModuleType>>> oldModules,
                IList<KeyValuePair<string, Tuple<string, ModuleType>>> newModules) {
            var inserted = newModules.Select(kvp => kvp.Key).Except(oldModules.Select(kvp => kvp.Key));
            return newModules.Where(kvp => inserted.Contains(kvp.Key)).Select(m => Patch.MakeInsertion(m.Key, m.Value.Item2, m.Value.Item1));
        }
    }
}