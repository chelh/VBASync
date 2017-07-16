using ICSharpCode.SharpZipLib.Core;
using ICSharpCode.SharpZipLib.Zip;
using OfficeOpenXml.Utils;
using OpenMcdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Globalization;
using System.Linq;
using System.Text;
using VBASync.Localization;

namespace VBASync.Model
{
    internal class VbaTemporaryFolder : VbaFolder, IDisposable
    {
        private readonly ISystemOperations _so;

        private bool _disposed;
        private List<Module> _modules;

        internal VbaTemporaryFolder() : this(new RealSystemOperations())
        {
        }

        internal VbaTemporaryFolder(ISystemOperations so) : base(so, GetTemporaryDirectoryPath(so))
        {
            _so = so;
            so.DirectoryCreateDirectory(FolderPath);
        }

        ~VbaTemporaryFolder()
        {
            Dispose(false);
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        public void FixCase(IDictionary<string, string> compareModules)
        {
            foreach (var m in _modules)
            {
                if (compareModules.ContainsKey(m.Name))
                {
                    var path = _so.PathCombine(FolderPath, m.FileName);
                    _so.FileWriteAllText(path, ModuleProcessing.FixCase(compareModules[m.Name],
                       _so.FileReadAllText(path, ProjectEncoding)), ProjectEncoding);
                }
            }
        }

        public void Read(string path)
        {
            if (string.IsNullOrEmpty(path))
            {
                throw new ArgumentException("Path to file cannot be null or empty.", nameof(path));
            }

            Stream fs = null;
            try
            {
                try
                {
                    fs = _so.OpenFileForRead(path);
                }
                catch (IOException)
                {
                    // most often we get this because Office has the file locked. move the file to another location and retry.
                    var dest = _so.PathCombine(FolderPath, "FileToExtract" + _so.PathGetExtension(path));
                    _so.FileCopy(path, dest);
                    path = dest;
                    fs = _so.OpenFileForRead(path);
                }

                CFStorage vbaProject;
                if (_so.PathGetFileName(path).Equals("vbaProject.bin", StringComparison.InvariantCultureIgnoreCase))
                {
                    vbaProject = new CompoundFile(fs).RootStorage;
                }
                else
                {
                    var sig = new byte[4];
                    fs.Read(sig, 0, 4);
                    if (sig.SequenceEqual(new byte[] { 0x50, 0x4b, 0x03, 0x04 }))
                    {
                        var zipFile = new ZipFile(fs);
                        var zipEntry = zipFile.Cast<ZipEntry>().FirstOrDefault(e => e.Name.EndsWith("vbaProject.bin", StringComparison.InvariantCultureIgnoreCase));
                        if (zipEntry == null)
                        {
                            throw new ApplicationException("Cannot find 'vbaProject.bin' in ZIP archive.");
                        }
                        using (var sw = _so.CreateNewFile(_so.PathCombine(FolderPath, "vbaProject.bin")))
                        {
                            StreamUtils.Copy(zipFile.GetInputStream(zipEntry), sw, new byte[4096]);
                        }
                        fs.Dispose();
                        fs = _so.OpenFileForRead(_so.PathCombine(FolderPath, "vbaProject.bin"));
                        vbaProject = new CompoundFile(fs).RootStorage;
                    }
                    else
                    {
                        fs.Seek(0, SeekOrigin.Begin);
                        vbaProject = new CompoundFile(fs).GetAllNamedEntries("_VBA_PROJECT_CUR").FirstOrDefault() as CFStorage
                            ?? new CompoundFile(fs).GetAllNamedEntries("OutlookVbaData").FirstOrDefault() as CFStorage
                            ?? new CompoundFile(fs).GetAllNamedEntries("Macros").FirstOrDefault() as CFStorage;
                    }
                }
                if (vbaProject == null)
                {
                    throw new ApplicationException("Cannot find VBA project storage in file.");
                }

                var projectLk = vbaProject.TryGetStream("PROJECTlk");
                if (projectLk != null)
                {
                    _so.FileWriteAllBytes(_so.PathCombine(FolderPath, "LicenseKeys.bin"), projectLk.GetData());
                }

                var projEncoding = Encoding.Default;
                uint projSysKind = 1;
                var projVersion = new Version(1, 1);
                var projConstants = new List<string>();
                string currentRefName = null;
                var references = new List<Reference>();
                var originalLibIds = new Dictionary<string, string>(StringComparer.InvariantCultureIgnoreCase);
                _modules = new List<Module>();
                Module currentModule = null;
                var projStrings = new List<string>();
                using (var br = new BinaryReader(new MemoryStream(DecompressStream(vbaProject.GetStorage("VBA").GetStream("dir")))))
                {
                    Action<int> seek = i => br.BaseStream.Seek(i, SeekOrigin.Current);
                    while (br.BaseStream.Position < br.BaseStream.Length)
                    {
                        switch (br.ReadUInt16())
                        {
                            case 0x0001:
                                // PROJECTSYSKIND
                                seek(4); // seek past size (always 4)
                                projSysKind = br.ReadUInt32();
                                break;
                            case 0x0002:
                            case 0x0014:
                            case 0x0008:
                            case 0x0007:
                                // PROJECTLCID, PROJECTLCIDINVOKE, PROJECTLIBFLAGS, and PROJECTHELPCONTEXT
                                seek(8); // seek past whole record (always 8 bytes long)
                                break;
                            case 0x0003:
                                // PROJECTCODEPAGE
                                seek(4); // seek past size (always 4)
                                projEncoding = Encoding.GetEncoding(br.ReadInt16());
                                break;
                            case 0x0004:
                                // PROJECTNAME
                                seek(br.ReadInt32()); // seek past whole record, since its contents are already in PROJECT
                                break;
                            case 0x0005:
                            case 0x0006:
                                // PROJECTDOCSTRING and PROJECTHELPFILEPATH
                                seek(br.ReadInt32() + 2); // seek past whole record, since its contents are already in PROJECT
                                seek(br.ReadInt32());
                                break;
                            case 0x0009:
                                // PROJECTVERSION
                                seek(4); // seek past Reserved
                                projVersion = new Version(br.ReadInt32(), br.ReadInt16());
                                break;
                            case 0x000c:
                                // PROJECTCONSTANTS
                                seek(br.ReadInt32() + 2); // seek past Constants and Reserved
                                projConstants.AddRange(Encoding.Unicode.GetString(br.ReadBytes(br.ReadInt32())).Split(':').Select(s => s.Trim()));
                                if (projConstants.Count == 1 && string.IsNullOrEmpty(projConstants[0]))
                                {
                                    projConstants.RemoveAt(0);
                                }
                                break;
                            case 0x0016:
                                // REFERENCENAME
                                seek(br.ReadInt32() + 2); // seek past Name and Reserved
                                currentRefName = Encoding.Unicode.GetString(br.ReadBytes(br.ReadInt32()));
                                break;
                            case 0x0033:
                                // REFERENCEORIGINAL
                                originalLibIds.Add(currentRefName, projEncoding.GetString(br.ReadBytes(br.ReadInt32())));
                                break;
                            case 0x002f:
                                // REFERENCECONTROL (after optional REFERENCEORIGINAL)
                                seek(4); // seek past SizeTwiddled
                                var libIdTwiddled = projEncoding.GetString(br.ReadBytes(br.ReadInt32()));
                                seek(6); // seek past Reserved1 and Reserved2
                                string nameRecordExtended = null;
                                if (br.PeekChar() == 0x16)
                                {
                                    // an optional REFERENCENAME record
                                    seek(2); // seek past Id
                                    seek(br.ReadInt32() + 2); // seek past Name and Reserved
                                    nameRecordExtended = Encoding.Unicode.GetString(br.ReadBytes(br.ReadInt32()));
                                }
                                seek(6); // seek past Reserved3 and SizeExtended
                                var libIdExtended = projEncoding.GetString(br.ReadBytes(br.ReadInt32()));
                                seek(6); // seek past Reserved4 and Reserved5
                                var originalTypeLib = new Guid(br.ReadBytes(16));
                                var cookie = br.ReadUInt32();
                                var refCtl = new ReferenceControl
                                {
                                    Name = currentRefName,
                                    Cookie = cookie,
                                    LibIdExtended = libIdExtended,
                                    LibIdTwiddled = libIdTwiddled,
                                    NameRecordExtended = nameRecordExtended,
                                    OriginalTypeLib = originalTypeLib
                                };
                                if (originalLibIds.ContainsKey(currentRefName))
                                {
                                    refCtl.OriginalLibId = originalLibIds[currentRefName];
                                }
                                references.Add(refCtl);
                                break;
                            case 0x000d:
                                // REFERENCEREGISTERED
                                seek(4); // seek past Size
                                references.Add(new ReferenceRegistered
                                {
                                    Name = currentRefName,
                                    LibId = projEncoding.GetString(br.ReadBytes(br.ReadInt32()))
                                });
                                seek(6); // seek past Reserved1 and Reserved2
                                break;
                            case 0x000e:
                                // REFERENCEPROJECT
                                seek(4); // seek past Size
                                references.Add(new ReferenceProject
                                {
                                    Name = currentRefName,
                                    LibIdAbsolute = projEncoding.GetString(br.ReadBytes(br.ReadInt32())),
                                    LibIdRelative = projEncoding.GetString(br.ReadBytes(br.ReadInt32())),
                                    Version = new Version(br.ReadInt32(), br.ReadInt16())
                                });
                                break;
                            case 0x000f:
                                // PROJECTMODULES
                                seek(6); // ignore entire record
                                break;
                            case 0x0013:
                                // PROJECTCOOKIE
                                seek(6); // ignore entire record
                                break;
                            case 0x0019:
                                // MODULENAME
                                seek(br.ReadInt32()); // ignore entire record
                                break;
                            case 0x0047:
                                // MODULENAMEUNICODE
                                _modules.Add(currentModule = new Module
                                {
                                    Name = Encoding.Unicode.GetString(br.ReadBytes(br.ReadInt32()))
                                });
                                break;
                            case 0x001a:
                                // MODULESTREAMNAME
                                seek(br.ReadInt32() + 2); // seek past StreamName and Reserved
                                currentModule.StreamName = Encoding.Unicode.GetString(br.ReadBytes(br.ReadInt32()));
                                break;
                            case 0x001c:
                                // MODULEDOCSTRING - ignore since this info is already in the module itself
                                seek(br.ReadInt32() + 2); // seek past DocString and Reserved
                                seek(br.ReadInt32()); // seek past DocStringUnicode
                                break;
                            case 0x0031:
                                // MODULEOFFSET
                                seek(4); // seek past size (always 4)
                                currentModule.Offset = br.ReadUInt32();
                                break;
                            case 0x001e:
                                // MODULEHELPCONTEXT
                                seek(8); // seek past entire record - this information is in the module itself as well
                                break;
                            case 0x002c:
                                // MODULECOOKIE
                                seek(6); // ignore entire record
                                break;
                            case 0x0021:
                                // MODULETYPE - procedural flag
                                seek(4); // ignore entire record since we get information about this from the PROJECT stream
                                break;
                            case 0x0022:
                                // MODULETYPE - document, class, or designer flag
                                seek(4); // ignore entire record since we get information about this from the PROJECT stream
                                break;
                            case 0x0025:
                                // MODULEREADONLY
                                currentModule.ReadOnly = true;
                                seek(4); // seek past Reserved
                                break;
                            case 0x0028:
                                // MODULEPRIVATE
                                currentModule.Private = true;
                                seek(4); // seek past Reserved
                                break;
                            case 0x002b:
                                // module terminator
                                currentModule = null;
                                seek(4);
                                break;
                            case 0x0010:
                                // global terminator
                                seek(4);
                                break;
                            default:
                                seek(-2);
                                throw new ApplicationException($"Unknown record id '0x{br.ReadInt16().ToString("X4")}'.");
                        }
                    }
                }
                ProjectEncoding = projEncoding;
                using (var sr = new StreamReader(new MemoryStream(vbaProject.GetStream("PROJECT").GetData()), projEncoding))
                {
                    string line;
                    while ((line = sr.ReadLine()) != null)
                    {
                        var split = line.Split('=');
                        var breakLoop = false;
                        switch (split[0]?.ToUpperInvariant())
                        {
                            case "MODULE":
                                _modules.First(m => string.Equals(m.Name, split[1], StringComparison.InvariantCultureIgnoreCase))
                                    .Type = ModuleType.Standard;
                                break;
                            case "DOCUMENT":
                                var split2 = split[1].Split('/');
                                var mod = _modules.First(m => string.Equals(m.Name, split2[0], StringComparison.InvariantCultureIgnoreCase));
                                mod.Type = ModuleType.StaticClass;
                                mod.Version = uint.Parse(split2[1].Substring(2), NumberStyles.HexNumber);
                                break;
                            case "CLASS":
                                _modules.First(m => string.Equals(m.Name, split[1], StringComparison.InvariantCultureIgnoreCase))
                                    .Type = ModuleType.Class;
                                break;
                            case "BASECLASS":
                                _modules.First(m => string.Equals(m.Name, split[1], StringComparison.InvariantCultureIgnoreCase))
                                    .Type = ModuleType.Form;
                                break;
                            default:
                                if (line.Equals("[Workspace]", StringComparison.InvariantCultureIgnoreCase))
                                {
                                    breakLoop = true; // don't output all the cruft after [Workspace]
                                }
                                else
                                {
                                    projStrings.Add(line);
                                }
                                break;
                        }
                        if (breakLoop)
                        {
                            break;
                        }
                    }
                }
                DeleteBlankLinesFromEnd(projStrings);
                projStrings.Insert(0, $"Version={projVersion}");
                projStrings.Insert(0, $"SysKind={projSysKind.ToString(CultureInfo.InvariantCulture)}");
                projStrings.Insert(0, $"CodePage={projEncoding.CodePage.ToString(CultureInfo.InvariantCulture)}");
                projStrings.Add("");
                projStrings.Add("[Constants]");
                projStrings.AddRange(projConstants.Select(s => string.Join("=", s.Split('=').Select(t => t.Trim()))));
                if (_modules.Any(m => m.Version > 0))
                {
                    projStrings.Add("");
                    projStrings.Add("[DocTLibVersions]");
                    projStrings.AddRange(_modules.Where(m => m.Version > 0).Select(m => $"{m.Name}={m.Version.ToString(CultureInfo.InvariantCulture)}"));
                }
                foreach (var refer in references)
                {
                    projStrings.Add("");
                    projStrings.Add($"[Reference {refer.Name}]");
                    projStrings.AddRange(refer.GetConfigStrings());
                }

                _so.FileWriteAllLines(_so.PathCombine(FolderPath, "Project.ini"), projStrings, Encoding.UTF8); // write using system line ending

                foreach (var m in _modules)
                {
                    var moduleText = projEncoding.GetString(DecompressStream(vbaProject.GetStorage("VBA").GetStream(m.StreamName), m.Offset));
                    moduleText = (GetPrepend(m, vbaProject, projEncoding) + moduleText).TrimEnd('\r', '\n') + "\r\n";
                    _so.FileWriteAllText(_so.PathCombine(FolderPath, m.Name + ModuleProcessing.ExtensionFromType(m.Type)), moduleText, projEncoding);
                }

                foreach (var m in _modules.Where(mod => mod.Type == ModuleType.Form))
                {
                    var cf = new CompoundFile();
                    CopyCfStreamsExcept(vbaProject.GetStorage(m.StreamName), cf.RootStorage, "\x0003VBFrame");
                    cf.RootStorage.CLSID = new Guid("c62a69f0-16dc-11ce-9e98-00aa00574a4f");
                    var frxPath = _so.PathCombine(FolderPath, m.Name + ".frx");
                    cf.Save(frxPath);
                    var bytes = _so.FileReadAllBytes(frxPath);
                    var size = bytes.Length;
                    using (var bw = new BinaryWriter(_so.OpenFileForWrite(frxPath)))
                    {
                        bw.Write((short)0x424c);
                        bw.Write(new byte[3]);
                        bw.Write(size >> 8);
                        bw.Write(new byte[15]);
                        bw.Write(bytes);
                    }
                }
            }
            finally
            {
                fs.Dispose();
            }
        }

        public void Write(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                throw new ArgumentException(VBASyncResources.ErrorPathCannotbeNullOrEmpty, nameof(filePath));
            }

            _so.FileDelete(_so.PathCombine(FolderPath, "vbaProject.bin"));
            var cf = new CompoundFile();
            var vbaProject = cf.RootStorage;

            var projIni = new ProjectIni(_so.PathCombine(FolderPath, "Project.ini"));
            projIni.AddFile(_so.PathCombine(FolderPath, "Project.ini.local"));
            var projEncoding = Encoding.GetEncoding(projIni.GetInt("General", "CodePage") ?? Encoding.Default.CodePage);
            if (!projEncoding.Equals(Encoding.Default))
            {
                projIni = new ProjectIni(_so.PathCombine(FolderPath, "Project.ini"), projEncoding);
                projIni.AddFile(_so.PathCombine(FolderPath, "Project.ini.local"));
            }
            var projSysKind = (uint)(projIni.GetInt("General", "SysKind") ?? 1);
            var projVersion = projIni.GetVersion("General", "Version") ?? new Version(1, 1);

            var refs = new List<Reference>();
            foreach (var refName in projIni.GetReferenceNames())
            {
                var refSubj = $"Reference {refName}";
                if (projIni.GetString(refSubj, "LibIdTwiddled") != null)
                {
                    refs.Add(new ReferenceControl
                    {
                        Cookie = (uint)(projIni.GetInt(refSubj, "Cookie") ?? 0),
                        LibIdExtended = projIni.GetString(refSubj, "LibIdExtended"),
                        LibIdTwiddled = projIni.GetString(refSubj, "LibIdTwiddled"),
                        Name = refName,
                        NameRecordExtended = projIni.GetString(refSubj, "NameRecordExtended"),
                        OriginalLibId = projIni.GetString(refSubj, "OriginalLibId"),
                        OriginalTypeLib = projIni.GetGuid(refSubj, "OriginalTypeLib") ?? Guid.Empty
                    });
                }
                else if (projIni.GetString(refSubj, "LibIdAbsolute") != null)
                {
                    refs.Add(new ReferenceProject
                    {
                        LibIdAbsolute = projIni.GetString(refSubj, "LibIdAbsolute"),
                        LibIdRelative = projIni.GetString(refSubj, "LibIdRelative"),
                        Name = refName,
                        Version = projIni.GetVersion(refSubj, "Version")
                    });
                }
                else
                {
                    refs.Add(new ReferenceRegistered
                    {
                        LibId = projIni.GetString(refSubj, "LibId"),
                        Name = refName
                    });
                }
            }

            var mods = new List<Module>();
            var modExts = new[] { ".bas", ".cls", ".frm" };
            foreach (var modPath in _so.DirectoryGetFiles(FolderPath).Where(s => modExts.Contains(_so.PathGetExtension(s) ?? "", StringComparer.InvariantCultureIgnoreCase)))
            {
                var modName = _so.PathGetFileNameWithoutExtension(modPath);
                var modText = EnsureCrLfEndings(_so.FileReadAllText(modPath, projEncoding));
                var modType = ModuleProcessing.TypeFromText(modText);
                projIni.RegisterModule(modName, modType, (uint)(projIni.GetInt("DocTLibVersions", modName) ?? 0));
                mods.Add(new Module
                {
                    DocString = GetAttribute(modText, "VB_Description"),
                    HelpContext = GetAttributeUInt(modText, "VB_HelpID"),
                    Name = modName,
                    StreamName = modName,
                    Type = modType
                });
            }

            vbaProject.AddStream("PROJECT");
            vbaProject.GetStream("PROJECT").SetData(projEncoding.GetBytes(projIni.GetProjectText()));

            if (_so.FileExists(_so.PathCombine(FolderPath, "LicenseKeys.bin")))
            {
                vbaProject.AddStream("PROJECTlk");
                vbaProject.GetStream("PROJECTlk").SetData(_so.FileReadAllBytes(_so.PathCombine(FolderPath, "LicenseKeys.bin")));
            }

            var projectWm = new List<byte>();
            foreach (var modName in mods.Select(m => m.Name))
            {
                projectWm.AddRange(projEncoding.GetBytes(modName));
                projectWm.Add(0x00);
                projectWm.AddRange(Encoding.Unicode.GetBytes(modName));
                projectWm.Add(0x00);
                projectWm.Add(0x00);
            }
            projectWm.AddRange(new byte[] { 0x00, 0x00 });
            vbaProject.AddStream("PROJECTwm");
            vbaProject.GetStream("PROJECTwm").SetData(projectWm.ToArray());

            vbaProject.AddStorage("VBA");
            var vbaProjectVba = vbaProject.GetStorage("VBA");

            vbaProjectVba.AddStream("_VBA_PROJECT");
            vbaProjectVba.GetStream("_VBA_PROJECT").SetData(new byte[] { 0xCC, 0x61, 0xFF, 0xFF, 0x00, 0x00, 0x00 });

            var dirUc = new List<byte>(); // uncompressed dir stream
            dirUc.AddRange(BitConverter.GetBytes((short)0x0001));
            dirUc.AddRange(BitConverter.GetBytes(4));
            dirUc.AddRange(BitConverter.GetBytes(projSysKind));
            dirUc.AddRange(BitConverter.GetBytes((short)0x0002));
            dirUc.AddRange(BitConverter.GetBytes(4));
            dirUc.AddRange(BitConverter.GetBytes(0x00000409)); // LCID
            dirUc.AddRange(BitConverter.GetBytes((short)0x0014));
            dirUc.AddRange(BitConverter.GetBytes(4));
            dirUc.AddRange(BitConverter.GetBytes(0x00000409)); // LCIDINVOKE
            dirUc.AddRange(BitConverter.GetBytes((short)0x0003));
            dirUc.AddRange(BitConverter.GetBytes(2));
            dirUc.AddRange(BitConverter.GetBytes((short)projEncoding.CodePage));
            var projNameBytes = projEncoding.GetBytes(projIni.GetString("General", "Name") ?? "");
            dirUc.AddRange(BitConverter.GetBytes((short)0x0004));
            dirUc.AddRange(BitConverter.GetBytes(projNameBytes.Length));
            dirUc.AddRange(projNameBytes);
            var projDocStringBytes = projEncoding.GetBytes(projIni.GetString("General", "Description") ?? "");
            var projDocStringUtfBytes = Encoding.Unicode.GetBytes(projIni.GetString("General", "Description") ?? "");
            dirUc.AddRange(BitConverter.GetBytes((short)0x0005));
            dirUc.AddRange(BitConverter.GetBytes(projDocStringBytes.Length));
            dirUc.AddRange(projDocStringBytes);
            dirUc.AddRange(BitConverter.GetBytes((short)0x0040));
            dirUc.AddRange(BitConverter.GetBytes(projDocStringUtfBytes.Length));
            dirUc.AddRange(projDocStringUtfBytes);
            var projHelpFileBytes = projEncoding.GetBytes(projIni.GetString("General", "HelpFile") ?? "");
            dirUc.AddRange(BitConverter.GetBytes((short)0x0006));
            dirUc.AddRange(BitConverter.GetBytes(projHelpFileBytes.Length));
            dirUc.AddRange(projHelpFileBytes);
            dirUc.AddRange(BitConverter.GetBytes((short)0x003D));
            dirUc.AddRange(BitConverter.GetBytes(projHelpFileBytes.Length));
            dirUc.AddRange(projHelpFileBytes);
            dirUc.AddRange(BitConverter.GetBytes((short)0x0007));
            dirUc.AddRange(BitConverter.GetBytes(4));
            dirUc.AddRange(BitConverter.GetBytes(projIni.GetInt("General", "HelpContextID") ?? 0));
            dirUc.AddRange(BitConverter.GetBytes((short)0x0008));
            dirUc.AddRange(BitConverter.GetBytes(4));
            dirUc.AddRange(BitConverter.GetBytes(0)); // LIBFLAGS
            dirUc.AddRange(BitConverter.GetBytes((short)0x0009));
            dirUc.AddRange(BitConverter.GetBytes(4));
            dirUc.AddRange(BitConverter.GetBytes(projVersion.Major));
            dirUc.AddRange(BitConverter.GetBytes((short)projVersion.Minor));
            var projConstantsBytes = projEncoding.GetBytes(projIni.GetConstantsString());
            var projConstantsUtfBytes = Encoding.Unicode.GetBytes(projIni.GetConstantsString());
            dirUc.AddRange(BitConverter.GetBytes((short)0x000C));
            dirUc.AddRange(BitConverter.GetBytes(projConstantsBytes.Length));
            dirUc.AddRange(projConstantsBytes);
            dirUc.AddRange(BitConverter.GetBytes((short)0x003C));
            dirUc.AddRange(BitConverter.GetBytes(projConstantsUtfBytes.Length));
            dirUc.AddRange(projConstantsUtfBytes);
            foreach (var rfc in refs)
            {
                dirUc.AddRange(rfc.GetBytes(projEncoding));
            }
            dirUc.AddRange(BitConverter.GetBytes((short)0x000F));
            dirUc.AddRange(BitConverter.GetBytes(2));
            dirUc.AddRange(BitConverter.GetBytes((short)mods.Count));
            dirUc.AddRange(BitConverter.GetBytes((short)0x0013));
            dirUc.AddRange(BitConverter.GetBytes(2));
            dirUc.AddRange(BitConverter.GetBytes((short)-1)); // PROJECTCOOKIE
            foreach (var mod in mods)
            {
                dirUc.AddRange(mod.GetBytes(projEncoding));
            }
            dirUc.AddRange(BitConverter.GetBytes((short)0x0010));
            dirUc.AddRange(BitConverter.GetBytes(0)); // global terminator
            vbaProjectVba.AddStream("dir");
            vbaProjectVba.GetStream("dir").SetData(CompoundDocumentCompression.CompressPart(dirUc.ToArray()));

            foreach (var mod in mods)
            {
                switch (mod.Type)
                {
                    case ModuleType.Class:
                    case ModuleType.StaticClass:
                        vbaProjectVba.AddStream(mod.StreamName);
                        var fileText = _so.FileReadAllText(_so.PathCombine(FolderPath, mod.Name + ModuleProcessing.ExtensionFromType(mod.Type)), projEncoding);
                        vbaProjectVba.GetStream(mod.StreamName).SetData(CompoundDocumentCompression.CompressPart(projEncoding.GetBytes(SeparateClassCode(fileText))));
                        break;
                    case ModuleType.Form:
                        string vbFrame;
                        vbaProjectVba.AddStream(mod.StreamName);
                        vbaProjectVba.GetStream(mod.StreamName).SetData(CompoundDocumentCompression.CompressPart(
                            projEncoding.GetBytes(SeparateVbFrame(_so.FileReadAllText(_so.PathCombine(FolderPath, mod.Name + ".frm"), projEncoding), out vbFrame))));
                        vbaProject.AddStorage(mod.StreamName);
                        var frmStorage = vbaProject.GetStorage(mod.StreamName);
                        frmStorage.AddStream("\x03VBFrame");
                        frmStorage.GetStream("\x03VBFrame").SetData(projEncoding.GetBytes(vbFrame));
                        var b = _so.FileReadAllBytes(_so.PathCombine(FolderPath, mod.Name + ".frx"));
                        using (var ms = new MemoryStream(b, 24, b.Length - 24))
                        {
                            var frx = new CompoundFile(ms);
                            CopyCfStreamsExcept(frx.RootStorage, frmStorage, "\x03VBFrame");
                        }
                        break;
                    default:
                        vbaProjectVba.AddStream(mod.StreamName);
                        vbaProjectVba.GetStream(mod.StreamName).SetData(CompoundDocumentCompression.CompressPart(
                            _so.FileReadAllBytes(_so.PathCombine(FolderPath, mod.Name + ModuleProcessing.ExtensionFromType(mod.Type)))));
                        break;
                }
            }

            using (var projSm = _so.CreateNewFile(_so.PathCombine(FolderPath, "vbaProject.bin")))
            {
                cf.Save(projSm);
            }

            if (_so.PathGetFileName(filePath).Equals("vbaProject.bin", StringComparison.InvariantCultureIgnoreCase))
            {
                _so.FileCopy(_so.PathCombine(FolderPath, "vbaProject.bin"), filePath, true);
            }
            else
            {
                Stream fs = null;
                try
                {
                    fs = _so.OpenFileForWrite(filePath);
                    var sig = new byte[4];
                    fs.Seek(0, SeekOrigin.Begin);
                    fs.Read(sig, 0, 4);
                    fs.Seek(0, SeekOrigin.Begin);
                    if (sig.SequenceEqual(new byte[] { 0x50, 0x4b, 0x03, 0x04 }))
                    {
                        var zipFile = new ZipFile(fs);
                        var zipEntry = zipFile.Cast<ZipEntry>().FirstOrDefault(e => e.Name.EndsWith("vbaProject.bin", StringComparison.InvariantCultureIgnoreCase));
                        if (zipEntry == null)
                        {
                            throw new ApplicationException(VBASyncResources.ErrorCannotFindVbaProject);
                        }
                        zipFile.BeginUpdate();
                        using (var st = _so.OpenFileForRead(_so.PathCombine(FolderPath, "vbaProject.bin")))
                        {
                            zipFile.Add(new StreamStaticDataSource(st), zipEntry.Name);
                            zipFile.CommitUpdate();
                        }
                        zipFile.Close();
                    }
                    else
                    {
                        var destCf = new CompoundFile(fs, CFSUpdateMode.Update, CFSConfiguration.Default);
                        var destVbaProject = destCf.GetAllNamedEntries("_VBA_PROJECT_CUR").FirstOrDefault() as CFStorage
                            ?? destCf.GetAllNamedEntries("OutlookVbaData").FirstOrDefault() as CFStorage
                            ?? destCf.GetAllNamedEntries("Macros").FirstOrDefault() as CFStorage;
                        var ls = new List<string>();
                        destVbaProject.VisitEntries(i => ls.Add(i.Name), true);
                        while (ls.Count > 0)
                        {
                            DeleteStorageChildren(destVbaProject);
                            ls.Clear();
                            destVbaProject.VisitEntries(i => ls.Add(i.Name), true);
                        }
                        CopyCfStreamsExcept(cf.RootStorage, destVbaProject, null);
                        destCf.Commit();
                        //destCf.Save(fs);
                        destCf.Close();
                    }
                }
                finally
                {
                    fs?.Dispose();
                }
            }
        }

        private static void CopyCfStreamsExcept(CFStorage src, CFStorage dest, string excludeName)
        {
            src.VisitEntries(i =>
            {
                if (i.Name?.Equals(excludeName, StringComparison.InvariantCultureIgnoreCase) ?? false)
                {
                    return;
                }
                if (i.IsStorage)
                {
                    dest.AddStorage(i.Name);
                    CopyCfStreamsExcept((CFStorage)i, dest.GetStorage(i.Name), null);
                }
                else
                {
                    dest.AddStream(i.Name);
                    dest.GetStream(i.Name).SetData(((CFStream)i).GetData());
                }
            }, false);
        }

        private static byte[] DecompressStream(CFStream sm, uint offset = 0)
        {
            try
            {
                var cd = new byte[sm.Size - offset];
                sm.Read(cd, offset, cd.Length);
                return CompoundDocumentCompression.DecompressPart(cd);
            }
            catch (Exception ex)
            {
                throw new ApplicationException(string.Format(VBASyncResources.ErrorDecompressingStream, sm.Name, ex.Message), ex);
            }
        }

        private static void DeleteBlankLinesFromEnd(IList<string> ls)
        {
            for (var i = ls.Count - 1; i >= 0; i--)
            {
                if (string.IsNullOrEmpty(ls[i]))
                {
                    ls.RemoveAt(i);
                }
                else
                {
                    return;
                }
            }
        }

        private static void DeleteStorageChildren(CFStorage target)
        {
            target.VisitEntries(i =>
            {
                if (i.IsStorage)
                {
                    DeleteStorageChildren((CFStorage)i);
                }
                target.Delete(i.Name);
            }, false);
        }

        private static string EnsureCrLfEndings(string text)
        {
            return string.Join("\r\n", text?.Split('\n').Select(s => s.Trim('\r')) ?? new string[0]);
        }

        private static string GetAttribute(string moduleText, string attribName)
        {
            using (var sr = new StringReader(moduleText))
            {
                string line;
                while ((line = sr.ReadLine()?.TrimEnd('\r')) != null)
                {
                    var start = line.TrimStart();
                    if (start.Length < $"Attribute {attribName} ".Length)
                    {
                        continue;
                    }
                    start = start.Substring(0, $"Attribute {attribName} ".Length);
                    if (start.Equals($"Attribute {attribName} ", StringComparison.InvariantCultureIgnoreCase)
                        || start.Equals($"Attribute {attribName}=", StringComparison.InvariantCultureIgnoreCase))
                    {
                        var split = line.Split(new[] { '=' }, 2);
                        if (split.Length == 2)
                        {
                            return split[1].Trim();
                        }
                    }
                }
            }
            return null;
        }

        private static uint? GetAttributeUInt(string moduleText, string attribName)
        {
            if (uint.TryParse(GetAttribute(moduleText, attribName), out var i))
            {
                return i;
            }
            return null;
        }

        private static string GetTemporaryDirectoryPath(ISystemOperations so)
            => $"{so.PathGetTempPath()}VBASync-{Guid.NewGuid().ToString()}";

        private static string GetPrepend(Module mod, CFStorage vbaProject, Encoding enc)
        {
            switch (mod.Type)
            {
            case ModuleType.Class:
            case ModuleType.StaticClass:
                return "VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\n";
            case ModuleType.Form:
                var vbFrameLines = enc.GetString(vbaProject.GetStorage(mod.StreamName).GetStream("\x0003VBFrame").GetData())
                    .Split('\n').Select(s => s.TrimEnd('\r')).ToList();
                    DeleteBlankLinesFromEnd(vbFrameLines);
                vbFrameLines.Insert(2, $"   OleObjectBlob   =   \"{mod.Name}.frx\":0000");
                return string.Join("\r\n", vbFrameLines) + "\r\n";
            case ModuleType.Standard:
                return "";
            default:
                throw new ApplicationException("Unrecognized module type");
            }
        }

        private static string SeparateClassCode(string wholeFile)
        {
            using (var sr = new StringReader(wholeFile))
            {
                var firstLine = true;
                string line;
                while ((line = sr.ReadLine()) != null)
                {
                    if (firstLine && !line.TrimStart().StartsWith("Version", StringComparison.InvariantCultureIgnoreCase))
                    {
                        return wholeFile;
                    }
                    firstLine = false;
                    if (line.TrimStart().StartsWith("End", StringComparison.InvariantCultureIgnoreCase))
                    {
                        break;
                    }
                }
                return sr.ReadToEnd();
            }
        }

        private static string SeparateVbFrame(string wholeFile, out string vbFrame)
        {
            var frmB = new StringBuilder();
            using (var sr = new StringReader(wholeFile))
            {
                string line;
                while ((line = sr.ReadLine()) != null)
                {
                    if (!line.TrimStart().StartsWith("OleObjectBlob", StringComparison.InvariantCultureIgnoreCase))
                    {
                        frmB.AppendLine(line);
                    }
                    if (line.TrimStart().StartsWith("End", StringComparison.InvariantCultureIgnoreCase))
                    {
                        break;
                    }
                }
                vbFrame = frmB.ToString();
                return sr.ReadToEnd();
            }
        }

        private void Dispose(bool explicitCall)
        {
            if (_disposed)
            {
                return;
            }

            if (explicitCall)
            {
                //_fs?.Dispose();
            }

            _so.DirectoryDelete(FolderPath, true);

            _disposed = true;
        }

        internal class StreamStaticDataSource : IStaticDataSource
        {
            private readonly Stream _source;

            internal StreamStaticDataSource(Stream source)
            {
                _source = source;
            }

            public Stream GetSource() => _source;
        }

        private class Module
        {
            public string DocString { get; set; }
            public string FileName => Name + GetExtension();
            public uint? HelpContext { get; set; }
            public string Name { get; set; }
            public uint Offset { get; set; }
            public bool Private { get; set; }
            public bool ReadOnly { get; set; }
            public string StreamName { get; set; }
            public ModuleType Type { get; set; }
            public uint Version { get; set; }

            public IList<byte> GetBytes(Encoding projEncoding)
            {
                var ret = new List<byte>();
                var nameBytes = projEncoding.GetBytes(Name);
                ret.AddRange(BitConverter.GetBytes((short)0x0019));
                ret.AddRange(BitConverter.GetBytes(nameBytes.Length));
                ret.AddRange(nameBytes);
                var nameUtfBytes = Encoding.Unicode.GetBytes(Name);
                ret.AddRange(BitConverter.GetBytes((short)0x0047));
                ret.AddRange(BitConverter.GetBytes(nameUtfBytes.Length));
                ret.AddRange(nameUtfBytes);
                ret.AddRange(BitConverter.GetBytes((short)0x001A));
                ret.AddRange(BitConverter.GetBytes(nameBytes.Length));
                ret.AddRange(nameBytes); // StreamName
                ret.AddRange(BitConverter.GetBytes((short)0x0032));
                ret.AddRange(BitConverter.GetBytes(nameUtfBytes.Length));
                ret.AddRange(nameUtfBytes); // StreamNameUnicode
                var docStringBytes = projEncoding.GetBytes(DocString ?? "");
                var docStringUtfBytes = Encoding.Unicode.GetBytes(DocString ?? "");
                ret.AddRange(BitConverter.GetBytes((short)0x001C));
                ret.AddRange(BitConverter.GetBytes(docStringBytes.Length));
                ret.AddRange(docStringBytes);
                ret.AddRange(BitConverter.GetBytes((short)0x0048));
                ret.AddRange(BitConverter.GetBytes(docStringUtfBytes.Length));
                ret.AddRange(docStringUtfBytes);
                ret.AddRange(BitConverter.GetBytes((short)0x0031));
                ret.AddRange(BitConverter.GetBytes(4));
                ret.AddRange(BitConverter.GetBytes(0)); // MODULEOFFSET
                ret.AddRange(BitConverter.GetBytes((short)0x001E));
                ret.AddRange(BitConverter.GetBytes(4));
                ret.AddRange(BitConverter.GetBytes(HelpContext ?? 0));
                ret.AddRange(BitConverter.GetBytes(Type == ModuleType.Standard ? (short)0x0021 : (short)0x0022));
                ret.AddRange(BitConverter.GetBytes(0));
                if (ReadOnly)
                {
                    ret.AddRange(BitConverter.GetBytes((short)0x0025));
                    ret.AddRange(BitConverter.GetBytes(0));
                }
                if (Private)
                {
                    ret.AddRange(BitConverter.GetBytes((short)0x0028));
                    ret.AddRange(BitConverter.GetBytes(0));
                }
                ret.AddRange(BitConverter.GetBytes((short)0x002B));
                ret.AddRange(BitConverter.GetBytes(0)); // module terminator
                return ret;
            }

            public override string ToString() => Name;

            private string GetExtension()
            {
                switch (Type)
                {
                    case ModuleType.Standard:
                        return ".bas";
                    case ModuleType.StaticClass:
                    case ModuleType.Class:
                        return ".cls";
                    case ModuleType.Form:
                        return ".frm";
                    case ModuleType.Ini:
                        return ".ini";
                    case ModuleType.Licenses:
                        return ".bin";
                    default:
                        throw new ArgumentOutOfRangeException(nameof(Type));
                }
            }
        }

        private abstract class Reference
        {
            public string Name { get; set; }

            public virtual IList<byte> GetBytes(Encoding projEncoding)
            {
                var ret = new List<byte>();
                var nameBytes = projEncoding.GetBytes(Name);
                var nameUtfBytes = Encoding.Unicode.GetBytes(Name);
                ret.AddRange(BitConverter.GetBytes((short)0x0016));
                ret.AddRange(BitConverter.GetBytes(nameBytes.Length));
                ret.AddRange(nameBytes);
                ret.AddRange(BitConverter.GetBytes((short)0x003E));
                ret.AddRange(BitConverter.GetBytes(nameUtfBytes.Length));
                ret.AddRange(nameUtfBytes);
                return ret;
            }

            public abstract List<string> GetConfigStrings();
            public override string ToString() => Name;
        }

        private class ReferenceControl : Reference
        {
            public uint Cookie { get; set; }
            public string LibIdExtended { get; set; }
            public string LibIdTwiddled { get; set; }
            public string NameRecordExtended { get; set; }
            public string OriginalLibId { get; set; }
            public Guid OriginalTypeLib { get; set; }

            public override IList<byte> GetBytes(Encoding projEncoding)
            {
                var ret = new List<byte>();
                ret.AddRange(base.GetBytes(projEncoding));
                if (!string.IsNullOrEmpty(OriginalLibId))
                {
                    var originalLibIdBytes = projEncoding.GetBytes(OriginalLibId);
                    ret.AddRange(BitConverter.GetBytes((short)0x0033));
                    ret.AddRange(BitConverter.GetBytes(originalLibIdBytes.Length));
                    ret.AddRange(originalLibIdBytes);
                }
                var twiddledLibIdBytes = projEncoding.GetBytes(LibIdTwiddled);
                ret.AddRange(BitConverter.GetBytes((short)0x002F));
                ret.AddRange(BitConverter.GetBytes(twiddledLibIdBytes.Length + 10));
                ret.AddRange(BitConverter.GetBytes(twiddledLibIdBytes.Length));
                ret.AddRange(twiddledLibIdBytes);
                ret.AddRange(BitConverter.GetBytes(0));
                ret.AddRange(BitConverter.GetBytes((short)0));
                if (!string.IsNullOrEmpty(NameRecordExtended))
                {
                    var nameRecordExtendedBytes = projEncoding.GetBytes(NameRecordExtended);
                    var nameRecordExtendedUnicodeBytes = Encoding.Unicode.GetBytes(NameRecordExtended);
                    ret.AddRange(BitConverter.GetBytes((short)0x0016));
                    ret.AddRange(BitConverter.GetBytes(nameRecordExtendedBytes.Length));
                    ret.AddRange(nameRecordExtendedBytes);
                    ret.AddRange(BitConverter.GetBytes((short)0x003e));
                    ret.AddRange(BitConverter.GetBytes(nameRecordExtendedUnicodeBytes.Length));
                    ret.AddRange(nameRecordExtendedUnicodeBytes);
                }
                ret.AddRange(BitConverter.GetBytes((short)0x0030));
                var extendedLibIdBytes = projEncoding.GetBytes(LibIdExtended);
                ret.AddRange(BitConverter.GetBytes(extendedLibIdBytes.Length + 30));
                ret.AddRange(BitConverter.GetBytes(extendedLibIdBytes.Length));
                ret.AddRange(extendedLibIdBytes);
                ret.AddRange(BitConverter.GetBytes(0));
                ret.AddRange(BitConverter.GetBytes((short)0));
                ret.AddRange(OriginalTypeLib.ToByteArray());
                ret.AddRange(BitConverter.GetBytes(Cookie));
                return ret;
            }

            public override List<string> GetConfigStrings()
            {
                var ls = new List<string> {
                    $"LibIdTwiddled={LibIdTwiddled}",
                    $"LibIdExtended={LibIdExtended}"
                };
                if (!string.IsNullOrEmpty(NameRecordExtended))
                {
                    ls.Add($"NameRecordExtended={NameRecordExtended}");
                }
                ls.Add($"OriginalLibId={OriginalLibId}");
                if (OriginalTypeLib != Guid.Empty)
                {
                    ls.Add($"OriginalTypeLib={OriginalTypeLib.ToString("B")}");
                }
                if (Cookie > 0)
                {
                    ls.Add($"Cookie={Cookie.ToString(CultureInfo.InvariantCulture)}");
                }
                return ls;
            }
        }

        private class ReferenceProject : Reference
        {
            public string LibIdAbsolute { get; set; }
            public string LibIdRelative { get; set; }
            public Version Version { get; set; }

            public override IList<byte> GetBytes(Encoding projEncoding)
            {
                var ret = new List<byte>();
                ret.AddRange(base.GetBytes(projEncoding));
                var libIdAbsoluteBytes = projEncoding.GetBytes(LibIdAbsolute);
                var libIdRelativeBytes = projEncoding.GetBytes(LibIdRelative);
                ret.AddRange(BitConverter.GetBytes((short)0x000E));
                ret.AddRange(BitConverter.GetBytes(libIdAbsoluteBytes.Length + libIdRelativeBytes.Length + 14));
                ret.AddRange(BitConverter.GetBytes(libIdAbsoluteBytes.Length));
                ret.AddRange(libIdAbsoluteBytes);
                ret.AddRange(BitConverter.GetBytes(libIdRelativeBytes.Length));
                ret.AddRange(libIdRelativeBytes);
                ret.AddRange(BitConverter.GetBytes(Version.Major));
                ret.AddRange(BitConverter.GetBytes((short)Version.Minor));
                return ret;
            }

            public override List<string> GetConfigStrings()
            {
                return new List<string> {
                    $"LibIdAbsolute={LibIdAbsolute}",
                    $"LibIdRelative={LibIdRelative}",
                    $"Version={Version}"
                };
            }
        }

        private class ReferenceRegistered : Reference
        {
            public string LibId { get; set; }

            public override IList<byte> GetBytes(Encoding projEncoding)
            {
                var ret = new List<byte>();
                ret.AddRange(base.GetBytes(projEncoding));
                var libIdBytes = projEncoding.GetBytes(LibId);
                ret.AddRange(BitConverter.GetBytes((short)0x000D));
                ret.AddRange(BitConverter.GetBytes(libIdBytes.Length + 10));
                ret.AddRange(BitConverter.GetBytes(libIdBytes.Length));
                ret.AddRange(libIdBytes);
                ret.AddRange(BitConverter.GetBytes(0));
                ret.AddRange(BitConverter.GetBytes((short)0));
                return ret;
            }

            public override List<string> GetConfigStrings()
            {
                return new List<string> {
                    $"LibId={LibId}"
                };
            }
        }
    }
}
