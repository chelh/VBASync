using System.IO;

namespace VbaSync.FrxObjects {
    static class AlignmentHelpers {
        public static void AlignTo(ushort alignment, Stream st, ref ushort bytesAccum) {
            if (bytesAccum % alignment == 0)
                return;
            st.Seek(alignment - bytesAccum % alignment, SeekOrigin.Current);
            bytesAccum += (ushort)(alignment - bytesAccum % alignment);
        }

        public static void IgnoreNext(ushort bytes, Stream st, ref ushort bytesAccum) {
            st.Seek(bytes, SeekOrigin.Current);
            bytesAccum += bytes;
        }

        public static void AlignTo(uint alignment, Stream st, ref uint bytesAccum) {
            if (bytesAccum % alignment == 0)
                return;
            st.Seek(alignment - bytesAccum % alignment, SeekOrigin.Current);
            bytesAccum += alignment - bytesAccum % alignment;
        }

        public static void IgnoreNext(uint bytes, Stream st, ref uint bytesAccum) {
            st.Seek(bytes, SeekOrigin.Current);
            bytesAccum += bytes;
        }
    }
}
