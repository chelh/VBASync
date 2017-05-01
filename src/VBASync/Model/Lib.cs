using OpenMcdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using VBASync.Localization;
using VBASync.Model.FrxObjects;

namespace VBASync.Model
{
    public static class Lib
    {
        public static bool FrxFilesAreDifferent(string frxPath1, string frxPath2, out string explain)
        {
            var frxBytes1 = File.ReadAllBytes(frxPath1);
            var frxBytes2 = File.ReadAllBytes(frxPath2);
            using (var frxCfStream1 = new MemoryStream(frxBytes1, 24, frxBytes1.Length - 24, false))
            using (var frxCfStream2 = new MemoryStream(frxBytes2, 24, frxBytes2.Length - 24, false))
            using (var cf1 = new CompoundFile(frxCfStream1))
            using (var cf2 = new CompoundFile(frxCfStream2))
            {
                return CfStoragesAreDifferent(cf1.RootStorage, cf2.RootStorage, out explain);
            }
        }

        public static IList<KeyValuePair<string, Tuple<string, ModuleType>>> GetFolderModules(string folderPath)
        {
            var modulesText = new Dictionary<string, Tuple<string, ModuleType>>();
            var extensions = new[] { ".bas", ".cls", ".frm" };
            var projIni = new ProjectIni(Path.Combine(folderPath, "Project.INI"));
            if (File.Exists(Path.Combine(folderPath, "Project.INI.local")))
            {
                projIni.AddFile(Path.Combine(folderPath, "Project.INI.local"));
            }
            var projEncoding = Encoding.GetEncoding(projIni.GetInt("General", "CodePage") ?? Encoding.Default.CodePage);
            foreach (var filePath in Directory.GetFiles(folderPath, "*.*").Where(s => extensions.Any(s.EndsWith)).Select(s => Path.Combine(folderPath, Path.GetFileName(s))))
            {
                var moduleText = File.ReadAllText(filePath, projEncoding).TrimEnd('\r', '\n') + "\r\n";
                modulesText[Path.GetFileNameWithoutExtension(filePath)] = Tuple.Create(moduleText, ModuleProcessing.TypeFromText(moduleText));
            }
            return modulesText.ToList();
        }

        public static Patch GetLicensesPatch(ISession session, string evfPath)
        {
            var folderLicensesPath = Path.Combine(session.FolderPath, "LicenseKeys.bin");
            var fileLicensesPath = Path.Combine(evfPath, "LicenseKeys.bin");
            var folderHasLicenses = File.Exists(folderLicensesPath);
            var fileHasLicenses = File.Exists(fileLicensesPath);
            if (folderHasLicenses && fileHasLicenses)
            {
                if (session.Action == ActionType.Extract)
                {
                    return Patch.MakeLicensesChange(File.ReadAllBytes(folderLicensesPath), File.ReadAllBytes(fileLicensesPath));
                }
                else
                {
                    return Patch.MakeLicensesChange(File.ReadAllBytes(fileLicensesPath), File.ReadAllBytes(folderLicensesPath));
                }
            }
            else if (!folderHasLicenses && !fileHasLicenses)
            {
                return null;
            }
            else if (session.Action == ActionType.Extract ? !folderHasLicenses : !fileHasLicenses)
            {
                return Patch.MakeLicensesChange(new byte[0], File.ReadAllBytes(session.Action == ActionType.Extract ? fileLicensesPath : folderLicensesPath));
            }
            else
            {
                return Patch.MakeLicensesChange(File.ReadAllBytes(session.Action == ActionType.Extract ? folderLicensesPath : fileLicensesPath), new byte[0]);
            }
        }

        public static IEnumerable<Patch> GetModulePatches(ISession session, ISessionSettings sessionSettings,
            string vbaFolderPath, IList<KeyValuePair<string, Tuple<string, ModuleType>>> folderModules,
            IList<KeyValuePair<string, Tuple<string, ModuleType>>> fileModules)
        {
            IList<KeyValuePair<string, Tuple<string, ModuleType>>> oldModules;
            IList<KeyValuePair<string, Tuple<string, ModuleType>>> newModules;
            string oldFolder;
            string newFolder;
            if (session.Action == ActionType.Extract)
            {
                oldModules = folderModules;
                newModules = fileModules;
                oldFolder = session.FolderPath;
                newFolder = vbaFolderPath;
            }
            else
            {
                oldModules = fileModules;
                newModules = folderModules;
                oldFolder = vbaFolderPath;
                newFolder = session.FolderPath;
            }

            // find modules which aren't in both lists and record them as new/deleted
            var patches = new List<Patch>();
            patches.AddRange(GetDeletedModuleChanges(oldModules, newModules));
            patches.AddRange(GetNewModuleChanges(oldModules, newModules, session.Action, sessionSettings.AddNewDocumentsToFile));

            // this also filters the new/deleted modules from the last step
            var sideBySide = (from o in oldModules
                              join n in newModules on o.Key equals n.Key
                              select new SideBySideArgs {
                                  Name = o.Key,
                                  OldType = o.Value.Item2,
                                  NewType = n.Value.Item2,
                                  OldText = o.Value.Item1,
                                  NewText = n.Value.Item1
                              }).ToArray();
            patches.AddRange(GetFrxChanges(oldFolder, newFolder, sideBySide.Where(x => x.NewType == ModuleType.Form).Select(x => x.Name)));
            sideBySide = sideBySide.Where(sxs => sxs.OldText != sxs.NewText).ToArray();
            foreach (var sxs in sideBySide)
            {
                patches.AddRange(Patch.CompareSideBySide(sxs));
            }

            return patches;
        }

        public static Patch GetProjectPatch(ISession session, string evfPath)
        {
            var folderIniPath = Path.Combine(session.FolderPath, "Project.ini");
            var fileIniPath = Path.Combine(evfPath, "Project.ini");
            var folderHasIni = File.Exists(folderIniPath);
            var fileHasIni = File.Exists(fileIniPath);
            if (folderHasIni && fileHasIni)
            {
                if (session.Action == ActionType.Extract)
                {
                    return Patch.MakeProjectChange(File.ReadAllText(folderIniPath), File.ReadAllText(fileIniPath));
                }
                else
                {
                    return Patch.MakeProjectChange(File.ReadAllText(fileIniPath), File.ReadAllText(folderIniPath));
                }
            }
            else if (!folderHasIni && !fileHasIni)
            {
                return null;
            }
            else if (session.Action == ActionType.Extract ? !folderHasIni : !fileHasIni)
            {
                return Patch.MakeInsertion("Project", ModuleType.Ini, File.ReadAllText(session.Action == ActionType.Extract ? fileIniPath : folderIniPath));
            }
            else
            {
                return null; // never suggest deleting Project.INI
            }
        }

        private static bool CfStoragesAreDifferent(CFStorage s1, CFStorage s2, out string explain)
        {
            var s1Names = new List<Tuple<string, bool>>();
            var s2Names = new List<Tuple<string, bool>>();
            s1.VisitEntries(i => s1Names.Add(Tuple.Create(i.Name, i.IsStorage)), false);
            s2.VisitEntries(i => s2Names.Add(Tuple.Create(i.Name, i.IsStorage)), false);
            s1Names.Sort();
            s2Names.Sort();
            if (!s1Names.SequenceEqual(s2Names))
            {
                explain = string.Format(VBASyncResources.ExplainFrxDifferentFileLists, s1.Name, string.Join("', '", s1Names), string.Join("', '", s2Names));
                return true;
            }
            FormControl fc1 = null;
            foreach (var t in s1Names)
            {
                if (t.Item2)
                {
                    if (CfStoragesAreDifferent(s1.GetStorage(t.Item1), s2.GetStorage(t.Item1), out explain))
                    {
                        return true;
                    }
                }
                else if (t.Item1 == "f")
                {
                    fc1 = new FormControl(s1.GetStream("f").GetData());
                    var fc2 = new FormControl(s2.GetStream("f").GetData());
                    if (!fc1.Equals(fc2))
                    {
                        explain = string.Format(VBASyncResources.ExplainFrxGeneralStreamDifference, "f", s1.Name);
                        return true;
                    }
                }
                else if (t.Item1 == "o" && fc1 != null)
                {
                    var o1 = s1.GetStream("o").GetData();
                    var o2 = s2.GetStream("o").GetData();
                    uint idx = 0;
                    foreach (var site in fc1.Sites)
                    {
                        explain = string.Format(VBASyncResources.ExplainFrxOStreamDifference, site.Name, s1.Name);
                        var o1Range = o1.Range(idx, site.ObjectStreamSize);
                        var o2Range = o2.Range(idx, site.ObjectStreamSize);
                        switch (site.ClsidCacheIndex)
                        {
                        case 15: // MorphData
                        case 26: // CheckBox
                        case 25: // ComboBox
                        case 24: // ListBox
                        case 27: // OptionButton
                        case 23: // TextBox
                        case 28: // ToggleButton
                            if (!new MorphDataControl(o1Range).Equals(new MorphDataControl(o2Range)))
                            {
                                return true;
                            }
                            break;
                        case 17: // CommandButton
                            if (!new CommandButtonControl(o1Range).Equals(new CommandButtonControl(o2Range)))
                            {
                                return true;
                            }
                            break;
                        case 18: // TabStrip
                            if (!new TabStripControl(o1Range).Equals(new TabStripControl(o2Range)))
                            {
                                return true;
                            }
                            break;
                        case 21: // Label
                            if (!new LabelControl(o1Range).Equals(new LabelControl(o2Range)))
                            {
                                return true;
                            }
                            break;
                        default:
                            if (!o1Range.SequenceEqual(o2Range))
                            {
                                return true;
                            }
                            break;
                        }
                        idx += site.ObjectStreamSize;
                    }
                }
                else if (!s1.GetStream(t.Item1).GetData().SequenceEqual(s2.GetStream(t.Item1).GetData()))
                {
                    explain = string.Format(VBASyncResources.ExplainFrxGeneralStreamDifference, t.Item1, s1.Name);
                    return true;
                }
            }
            explain = "No differences found.";
            return false;
        }

        private static IEnumerable<Patch> GetDeletedModuleChanges(
                IList<KeyValuePair<string, Tuple<string, ModuleType>>> oldModules,
                IList<KeyValuePair<string, Tuple<string, ModuleType>>> newModules)
        {
            var deleted = oldModules.Select(kvp => kvp.Key).Except(newModules.Select(kvp => kvp.Key));
            return oldModules.Where(kvp => deleted.Contains(kvp.Key)).Select(m => Patch.MakeDeletion(m.Key, m.Value.Item2, m.Value.Item1));
        }

        private static IEnumerable<Patch> GetFrxChanges(string oldFolder, string newFolder, IEnumerable<string> frmModules)
        {
            foreach (var modName in frmModules)
            {
                var oldFrxPath = Path.Combine(oldFolder, modName + ".frx");
                var newFrxPath = Path.Combine(newFolder, modName + ".frx");
                if (!File.Exists(oldFrxPath) && !File.Exists(newFrxPath))
                {
                }
                else if (File.Exists(oldFrxPath) && !File.Exists(newFrxPath))
                {
                    yield return Patch.MakeFrxDeletion(modName);
                }
                else if (!File.Exists(oldFrxPath) && File.Exists(newFrxPath))
                {
                    yield return Patch.MakeFrxChange(modName);
                }
                else if (FrxFilesAreDifferent(oldFrxPath, newFrxPath, out var explain))
                {
                    yield return Patch.MakeFrxChange(modName);
                }
            }
        }

        private static IEnumerable<Patch> GetNewModuleChanges(
                IList<KeyValuePair<string, Tuple<string, ModuleType>>> oldModules,
                IList<KeyValuePair<string, Tuple<string, ModuleType>>> newModules,
                ActionType action, bool addNewDocumentsToFile)
        {
            var inserted = newModules.Select(kvp => kvp.Key).Except(oldModules.Select(kvp => kvp.Key));
            if (action == ActionType.Extract || addNewDocumentsToFile)
            {
                return newModules.Where(kvp => inserted.Contains(kvp.Key))
                    .Select(m => Patch.MakeInsertion(m.Key, m.Value.Item2, m.Value.Item1));
            }
            return newModules.Where(kvp => inserted.Contains(kvp.Key) && kvp.Value.Item2 != ModuleType.StaticClass)
                .Select(m => Patch.MakeInsertion(m.Key, m.Value.Item2, m.Value.Item1));
        }
    }
}
