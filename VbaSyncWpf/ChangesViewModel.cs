using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows;
using System.Windows.Data;
using System.Windows.Media;

namespace VbaSync {
    public class ChangesViewModel : ObservableCollection<Patch> {
        public ChangesViewModel(IEnumerable<Patch> a) : base(a) { }
        public ChangesViewModel() { }
    }

    public class ModuleTypeToIconConverter : WpfConverter {
        public ModuleTypeToIconConverter() : base(
                (v, t, p, c) => {
                    switch ((ModuleType)v) {
                        case ModuleType.Class:
                            return "pack://application:,,,/Icons/ClassIcon.png";
                        case ModuleType.Standard:
                            return "pack://application:,,,/Icons/ModuleIcon.png";
                        case ModuleType.Form:
                            return "pack://application:,,,/Icons/FormIcon.png";
                        default:
                            return "pack://application:,,,/Icons/DocIcon.png";
                    }
                }) { }
    }

    public class ChangeTypeToBrushConverter : WpfConverter {
        public ChangeTypeToBrushConverter() : base(
                (v, t, p, c) => {
                    switch ((ChangeType)v) {
                        case ChangeType.DeleteFile:
                            return Brushes.Red;

                        case ChangeType.ChangeFileType:
                        case ChangeType.MoveFile:
                            return Brushes.OrangeRed;

                        case ChangeType.AddLines:
                        case ChangeType.AddSub:
                            return Brushes.DarkSlateGray;

                        case ChangeType.AddFile:
                            return Brushes.Magenta;

                        case ChangeType.DeleteLines:
                        case ChangeType.DeleteSub:
                            return Brushes.DarkMagenta;

                        default:
                            return Brushes.DarkBlue;
                    }
                }) { }
    }
}