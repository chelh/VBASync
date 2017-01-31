using System.Windows.Data;

namespace VbaSync {
    public static class WpfStatic {
        public static readonly WpfConverter ChangeTypeToDescriptionOneWay = new WpfConverter(
            v => {
                switch ((ChangeType)v) {
                    case ChangeType.AddFile:
                        return Properties.Resources.CTAddFile;
                    case ChangeType.MoveFile:
                        return Properties.Resources.CTMoveFile;
                    case ChangeType.DeleteFile:
                        return Properties.Resources.CTDeleteFile;
                    case ChangeType.AddSub:
                        return Properties.Resources.CTAddSub;
                    case ChangeType.DeleteSub:
                        return Properties.Resources.CTDeleteSub;
                    case ChangeType.AddLines:
                        return Properties.Resources.CTAddLines;
                    case ChangeType.ChangeLines:
                        return Properties.Resources.CTChangeLines;
                    case ChangeType.DeleteLines:
                        return Properties.Resources.CTDeleteLines;
                    case ChangeType.ReplaceId:
                        return Properties.Resources.CTReplaceId;
                    case ChangeType.ChangeFormControls:
                        return Properties.Resources.CTChangeFormControls;
                    case ChangeType.DeleteFrx:
                        return Properties.Resources.CTDeleteFrx;
                    case ChangeType.ChangeFileType:
                        return Properties.Resources.CTChangeFileType;
                    case ChangeType.WholeFile:
                        return Properties.Resources.CTWholeFile;
                    case ChangeType.MoveSub:
                        return Properties.Resources.CTMoveSub;
                    case ChangeType.Project:
                        return Properties.Resources.CTProject;
                    default:
                        return null;
                }
            });

        public static readonly WpfConverter EnumToBoolean = new WpfConverter(
                (v, p) => v.Equals(p),
                (v, p) => v.Equals(true)? p : Binding.DoNothing
                );

        public static readonly WpfConverter ModuleTypeToDescriptionOneWay = new WpfConverter(
            v => {
                switch ((ModuleType)v) {
                    case ModuleType.Standard:
                        return Properties.Resources.MTStandard;
                    case ModuleType.StaticClass:
                        return Properties.Resources.MTStaticClass;
                    case ModuleType.Class:
                        return Properties.Resources.MTClass;
                    case ModuleType.Form:
                        return Properties.Resources.MTForm;
                    case ModuleType.Ini:
                        return Properties.Resources.MTIni;
                    default:
                        return null;
                }
            });
    }
}
