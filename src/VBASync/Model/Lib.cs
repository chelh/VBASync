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
        // This is used from this UI layer
        public static bool FrxFilesAreDifferent(string frxPath1, string frxPath2, out string explain)
            => FrxFilesAreDifferent(new RealSystemOperations(), frxPath1, frxPath2, out explain);

        internal static bool FrxFilesAreDifferent(ISystemOperations so, string frxPath1, string frxPath2, out string explain)
        {
            try
            {
                var frxBytes1 = so.FileReadAllBytes(frxPath1);
                var frxBytes2 = so.FileReadAllBytes(frxPath2);
                using (var frxCfStream1 = new MemoryStream(frxBytes1, 24, frxBytes1.Length - 24, false))
                using (var frxCfStream2 = new MemoryStream(frxBytes2, 24, frxBytes2.Length - 24, false))
                using (var cf1 = new CompoundFile(frxCfStream1))
                using (var cf2 = new CompoundFile(frxCfStream2))
                {
                    return CfStoragesAreDifferent(cf1.RootStorage, cf2.RootStorage, out explain);
                }
            }
            catch (Exception ex)
            {
                explain = ex.Message;
                return true;
            }
        }

        internal static Patch GetLicensesPatch(ISystemOperations so, ISession session, string evfPath)
        {
            var folderLicensesPath = so.PathCombine(session.FolderPath, "LicenseKeys.bin");
            var fileLicensesPath = so.PathCombine(evfPath, "LicenseKeys.bin");
            var folderHasLicenses = so.FileExists(folderLicensesPath);
            var fileHasLicenses = so.FileExists(fileLicensesPath);
            if (folderHasLicenses && fileHasLicenses)
            {
                if (session.Action == ActionType.Extract)
                {
                    return Patch.MakeLicensesChange(so.FileReadAllBytes(folderLicensesPath), so.FileReadAllBytes(fileLicensesPath));
                }
                else
                {
                    return Patch.MakeLicensesChange(so.FileReadAllBytes(fileLicensesPath), so.FileReadAllBytes(folderLicensesPath));
                }
            }
            else if (!folderHasLicenses && !fileHasLicenses)
            {
                return null;
            }
            else if (session.Action == ActionType.Extract ? !folderHasLicenses : !fileHasLicenses)
            {
                return Patch.MakeLicensesChange(new byte[0], so.FileReadAllBytes(session.Action == ActionType.Extract ? fileLicensesPath : folderLicensesPath));
            }
            else
            {
                return Patch.MakeLicensesChange(so.FileReadAllBytes(session.Action == ActionType.Extract ? folderLicensesPath : fileLicensesPath), new byte[0]);
            }
        }

        internal static IEnumerable<Patch> GetModulePatches(ISystemOperations so, ISession session, ISessionSettings sessionSettings,
            ILocateModules folderModuleLocator, IList<KeyValuePair<string, Tuple<string, ModuleType>>> folderModules,
            ILocateModules fileModuleLocator, IList<KeyValuePair<string, Tuple<string, ModuleType>>> fileModules)
        {
            IList<KeyValuePair<string, Tuple<string, ModuleType>>> oldModules;
            IList<KeyValuePair<string, Tuple<string, ModuleType>>> newModules;
            ILocateModules oldModuleLocator;
            ILocateModules newModuleLocator;
            if (session.Action == ActionType.Extract)
            {
                oldModules = folderModules;
                newModules = fileModules;
                oldModuleLocator = folderModuleLocator;
                newModuleLocator = fileModuleLocator;
            }
            else
            {
                oldModules = fileModules;
                newModules = folderModules;
                oldModuleLocator = fileModuleLocator;
                newModuleLocator = folderModuleLocator;
            }

            if (sessionSettings.IgnoreEmpty)
            {
                oldModules = oldModules.Where(kvp => ModuleProcessing.HasCode(kvp.Value.Item1)).ToList();
                newModules = newModules.Where(kvp => ModuleProcessing.HasCode(kvp.Value.Item1)).ToList();
            }

            // find modules which aren't in both lists and record them as new/deleted
            var patches = new List<Patch>();
            patches.AddRange(GetDeletedModuleChanges(oldModules, newModules, session.Action, sessionSettings.DeleteDocumentsFromFile));
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
            patches.AddRange(GetFrxChanges(so, oldModuleLocator, newModuleLocator, sideBySide.Where(x => x.NewType == ModuleType.Form).Select(x => x.Name)));
            sideBySide = sideBySide.Where(sxs => sxs.OldText != sxs.NewText).ToArray();
            foreach (var sxs in sideBySide)
            {
                patches.AddRange(Patch.CompareSideBySide(sxs));
            }

            return patches;
        }

        internal static Patch GetProjectPatch(ISystemOperations so, ISession session, string evfPath)
        {
            var folderIniPath = so.PathCombine(session.FolderPath, "Project.ini");
            var fileIniPath = so.PathCombine(evfPath, "Project.ini");
            var folderHasIni = so.FileExists(folderIniPath);
            var fileHasIni = so.FileExists(fileIniPath);
            if (folderHasIni && fileHasIni)
            {
                if (session.Action == ActionType.Extract)
                {
                    return Patch.MakeProjectChange(so.FileReadAllText(folderIniPath, Encoding.UTF8), so.FileReadAllText(fileIniPath, Encoding.UTF8));
                }
                else
                {
                    return Patch.MakeProjectChange(so.FileReadAllText(fileIniPath, Encoding.UTF8), so.FileReadAllText(folderIniPath, Encoding.UTF8));
                }
            }
            else if (!folderHasIni && !fileHasIni)
            {
                return null;
            }
            else if (session.Action == ActionType.Extract ? !folderHasIni : !fileHasIni)
            {
                return Patch.MakeInsertion("Project", ModuleType.Ini, so.FileReadAllText(session.Action == ActionType.Extract ? fileIniPath : folderIniPath, Encoding.UTF8));
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
                explain = string.Format(VBASyncResources.ExplainFrxDifferentFileLists, s1.Name,
                    string.Join("', '", s1Names.Select(t => t.Item1)),
                    string.Join("', '", s2Names.Select(t => t.Item1)));
                return true;
            }
            FormControl fc1 = null;
            FormControl fc2 = null;
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
                    fc2 = new FormControl(s2.GetStream("f").GetData());
                    if (!fc1.Equals(fc2))
                    {
                        explain = string.Format(VBASyncResources.ExplainFrxGeneralStreamDifference, "f", s1.Name);
                        return true;
                    }
                }
                else if (t.Item1 == "o" && fc1 != null && fc2 != null)
                {
                    var fc2SitesList = fc2.Sites.ToList();
                    var o1Controls = DecomposeOStream(fc1.Sites, s1.GetStream("o").GetData());
                    var o2Controls = DecomposeOStream(fc2.Sites, s2.GetStream("o").GetData());
                    for (var siteIdx1 = 0; siteIdx1 < fc1.Sites.Length; ++siteIdx1)
                    {
                        var siteIdx2 = fc2SitesList.FindIndex(s => s.Id == fc1.Sites[siteIdx1].Id);
                        if (!Equals(o1Controls[siteIdx1], o2Controls[siteIdx2]))
                        {
                            explain = string.Format(VBASyncResources.ExplainFrxOStreamDifference,
                                fc1.Sites[siteIdx1].Name, s1.Name);
                            return true;
                        }
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

        private static object[] DecomposeOStream(OleSiteConcreteControl[] sites, byte[] content)
        {
            var ret = new object[sites.Length];
            uint contentIdx = 0;
            for (var siteIdx = 0; siteIdx < ret.Length; ++siteIdx)
            {
                var range = content.Range(contentIdx, sites[siteIdx].ObjectStreamSize);
                switch (sites[siteIdx].ClsidCacheIndex)
                {
                    case 15: // MorphData
                    case 26: // CheckBox
                    case 25: // ComboBox
                    case 24: // ListBox
                    case 27: // OptionButton
                    case 23: // TextBox
                    case 28: // ToggleButton
                        ret[siteIdx] = new MorphDataControl(range);
                        break;
                    case 17: // CommandButton
                        ret[siteIdx] = new CommandButtonControl(range);
                        break;
                    case 18: // TabStrip
                        ret[siteIdx] = new TabStripControl(range);
                        break;
                    case 21: // Label
                        ret[siteIdx] = new LabelControl(range);
                        break;
                    default: // some other control – treat it as a raw byte sequence
                        ret[siteIdx] = new RawControl(range);
                        break;
                }
                contentIdx += sites[siteIdx].ObjectStreamSize;
            }
            return ret;
        }

        private static IEnumerable<Patch> GetDeletedModuleChanges(
                IList<KeyValuePair<string, Tuple<string, ModuleType>>> oldModules,
                IList<KeyValuePair<string, Tuple<string, ModuleType>>> newModules,
                ActionType action, bool deleteDocumentsFromFile)
        {
            var deleted = new HashSet<string>(oldModules.Select(kvp => kvp.Key).Except(newModules.Select(kvp => kvp.Key)));
            var deletedModulesKvp = oldModules.Where(kvp => deleted.Contains(kvp.Key)).ToList();
            if (action == ActionType.Extract || deleteDocumentsFromFile)
            {
                return deletedModulesKvp.Select(kvp => Patch.MakeDeletion(kvp.Key, kvp.Value.Item2, kvp.Value.Item1));
            }
            return deletedModulesKvp.Where(kvp => kvp.Value.Item2 != ModuleType.StaticClass)
                .Select(kvp => Patch.MakeDeletion(kvp.Key, kvp.Value.Item2, kvp.Value.Item1))
                .Concat(deletedModulesKvp.Where(kvp => kvp.Value.Item2 == ModuleType.StaticClass).SelectMany(GetStubModulePatches));

            IEnumerable<Patch> GetStubModulePatches(KeyValuePair<string, Tuple<string, ModuleType>> kvp)
            {
                var oldText = kvp.Value.Item1;
                var newText = ModuleProcessing.StubOut(oldText);
                if (oldText?.TrimEnd('\r', '\n') == newText?.TrimEnd('\r', '\n'))
                {
                    return new Patch[0];
                }
                return Patch.CompareSideBySide(new SideBySideArgs
                {
                    Name = kvp.Key,
                    NewText = newText,
                    NewType = kvp.Value.Item2,
                    OldText = oldText,
                    OldType = kvp.Value.Item2
                });
            }
        }

        private static IEnumerable<Patch> GetFrxChanges(ISystemOperations fo, ILocateModules oldModuleLocator,
            ILocateModules newModuleLocator, IEnumerable<string> frmModules)
        {
            foreach (var modName in frmModules)
            {
                var oldFrxPath = oldModuleLocator.GetFrxPath(modName);
                var newFrxPath = newModuleLocator.GetFrxPath(modName);
                if (string.IsNullOrEmpty(oldFrxPath) && string.IsNullOrEmpty(newFrxPath))
                {
                }
                else if (!string.IsNullOrEmpty(oldFrxPath) && string.IsNullOrEmpty(newFrxPath))
                {
                    yield return Patch.MakeFrxDeletion(modName);
                }
                else if (string.IsNullOrEmpty(oldFrxPath) && !string.IsNullOrEmpty(newFrxPath))
                {
                    yield return Patch.MakeFrxChange(modName);
                }
                else if (FrxFilesAreDifferent(fo, oldFrxPath, newFrxPath, out var explain))
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
