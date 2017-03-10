using System;
using System.Globalization;
using System.Windows.Data;

namespace VBASync.WPF {
    public class WpfConverter : IValueConverter {
        private readonly Func<object, Type, object, CultureInfo, object> _forward;
        private readonly Func<object, Type, object, CultureInfo, object> _reverse;

        public WpfConverter(Func<dynamic, dynamic> forward,
                            Func<dynamic, dynamic> reverse = null) {
            _forward = (v, t, p, c) => forward(v);
            _reverse = (v, t, p, c) => (reverse ?? DefaultConvert)(v);
        }

        public WpfConverter(Func<dynamic, dynamic, dynamic> forwardVp,
                            Func<dynamic, dynamic, dynamic> reverseVp = null) {
            _forward = (v, t, p, c) => forwardVp(v, p);
            _reverse = (v, t, p, c) => (reverseVp ?? DefaultConvertVp)(v, p);
        }

        public WpfConverter(Func<dynamic, Type, dynamic, CultureInfo, dynamic> forwardVtpc,
                            Func<dynamic, Type, dynamic, CultureInfo, dynamic> reverseVtpc = null) {
            _forward = forwardVtpc;
            _reverse = reverseVtpc ?? DefaultConvertVtpc;
        }

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
            => _forward(value, targetType, parameter, culture);

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => _reverse(value, targetType, parameter, culture);

        private static object DefaultConvert(object value) => Binding.DoNothing;
        private static object DefaultConvertVp(object value, object parameter) => Binding.DoNothing;

        private static object DefaultConvertVtpc(object value, Type targetType, object parameter, CultureInfo culture)
            => Binding.DoNothing;
    }
}
