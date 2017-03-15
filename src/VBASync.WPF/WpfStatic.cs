using System.Windows.Data;
using VBASync.Localization;
using VBASync.Model;

namespace VBASync.WPF
{
    public static class WpfStatic
    {
        public static readonly WpfConverter ChangeTypeToDescriptionOneWay = new WpfConverter(
            v => {
                switch ((ChangeType)v)
                {
                case ChangeType.AddFile:
                    return VBASyncResources.CTAddFile;
                case ChangeType.MoveFile:
                    return VBASyncResources.CTMoveFile;
                case ChangeType.DeleteFile:
                    return VBASyncResources.CTDeleteFile;
                case ChangeType.AddSub:
                    return VBASyncResources.CTAddSub;
                case ChangeType.DeleteSub:
                    return VBASyncResources.CTDeleteSub;
                case ChangeType.AddLines:
                    return VBASyncResources.CTAddLines;
                case ChangeType.ChangeLines:
                    return VBASyncResources.CTChangeLines;
                case ChangeType.DeleteLines:
                    return VBASyncResources.CTDeleteLines;
                case ChangeType.ReplaceId:
                    return VBASyncResources.CTReplaceId;
                case ChangeType.ChangeFormControls:
                    return VBASyncResources.CTChangeFormControls;
                case ChangeType.DeleteFrx:
                    return VBASyncResources.CTDeleteFrx;
                case ChangeType.ChangeFileType:
                    return VBASyncResources.CTChangeFileType;
                case ChangeType.WholeFile:
                    return VBASyncResources.CTWholeFile;
                case ChangeType.MoveSub:
                    return VBASyncResources.CTMoveSub;
                case ChangeType.Project:
                    return VBASyncResources.CTProject;
                default:
                    return null;
                }
            });

        public static readonly WpfConverter EnumToBoolean = new WpfConverter(
                (v, p) => v.Equals(p),
                (v, p) => v.Equals(true) ? p : Binding.DoNothing
                );

        public static readonly WpfConverter ModuleTypeToDescriptionOneWay = new WpfConverter(
            v => {
                switch ((ModuleType)v)
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
            });
    }
}
