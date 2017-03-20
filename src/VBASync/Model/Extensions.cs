using System;

namespace VBASync.Model
{
    internal static class Extensions
    {
        public static byte[] Range(this byte[] src, uint start, uint size)
        {
            var ret = new byte[size];
            if (size > 0)
            {
                Array.Copy(src, start, ret, 0, size);
            }
            return ret;
        }
    }
}
