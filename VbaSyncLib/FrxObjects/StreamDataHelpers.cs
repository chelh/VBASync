using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace VbaSync.FrxObjects {
    static class StreamDataHelpers {
        public static int CcbToLength(int ccb, out bool compressed) {
            if (ccb < 0) {
                compressed = true;
                return unchecked((int)(ccb ^ 0x80000000));
            }
            compressed = false;
            return ccb;
        }

        public static byte[] ReadGuidAndPicture(BinaryReader r) {
            r.BaseStream.Seek(20, SeekOrigin.Current); // skip GUID and Preamble
            return r.ReadBytes(r.ReadInt32());
        }

        public static bool GetFontIsStdFont(BinaryReader r)
            => new Guid(r.ReadBytes(16)) == new Guid(0x0BE35203, 0x8F91, 0x11CE, 0x9D, 0xE3, 0x00, 0xAA, 0x00, 0x4B, 0xB8, 0x51);

        public static Tuple<short, byte, short, uint, string> ReadStdFont(BinaryReader r) {
            r.BaseStream.Seek(1, SeekOrigin.Current); // skip Version
            return Tuple.Create(r.ReadInt16(), r.ReadByte(), r.ReadInt16(), r.ReadUInt32(), Encoding.ASCII.GetString(r.ReadBytes(r.ReadByte())));
        }

        public static byte[] ReadTextProps(BinaryReader r) {
            r.BaseStream.Seek(2, SeekOrigin.Current); // skip MinorVersion and MajorVersion
            return r.ReadBytes(r.ReadUInt16());
        }
    }
}
