using System;
using System.Globalization;
using System.Windows.Data;

namespace VBASync.WPF
{
    public class WpfConverter : IValueConverter
    {
        private readonly Func<object, Type, object, CultureInfo, object> _forward;
        private readonly Func<object, Type, object, CultureInfo, object> _reverse;

        public WpfConverter(Func<dynamic, dynamic> forward,
                            Func<dynamic, dynamic> reverse = null)
        {
            _forward = (v, t, p, c) => forward(v);
            _reverse = (v, t, p, c) => (reverse ?? (v2 => Binding.DoNothing))(v);
        }

        public WpfConverter(Func<dynamic, dynamic, dynamic> forwardVp,
                            Func<dynamic, dynamic, dynamic> reverseVp = null)
        {
            _forward = (v, t, p, c) => forwardVp(v, p);
            _reverse = (v, t, p, c) => (reverseVp ?? ((v2, p2) => Binding.DoNothing))(v, p);
        }

        public WpfConverter(Func<dynamic, Type, dynamic, CultureInfo, dynamic> forwardVtpc,
                            Func<dynamic, Type, dynamic, CultureInfo, dynamic> reverseVtpc = null)
        {
            _forward = forwardVtpc;
            _reverse = reverseVtpc ?? ((v, t, p, c) => Binding.DoNothing);
        }

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
            => _forward(value, targetType, parameter, culture);

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => _reverse(value, targetType, parameter, culture);
    }
}
