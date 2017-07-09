using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using VBASync.Localization;

namespace VBASync.Model.FrxObjects
{
    internal class FrxReader : IDisposable
    {
        private static readonly Guid StdFontGuid = new Guid(0x0BE35203, 0x8F91, 0x11CE,
            0x9D, 0xE3, 0x00, 0xAA, 0x00, 0x4B, 0xB8, 0x51);

        public FrxReader(Stream st)
        {
            Unaligned = new BinaryReader(st);
        }

        public Stream BaseStream => Unaligned.BaseStream;
        public BinaryReader Unaligned { get; }

        public void AlignTo(ushort alignment)
        {
            var rem = BaseStream.Position % alignment;
            if (rem == 0)
            {
                return;
            }

            BaseStream.Seek(alignment - rem, SeekOrigin.Current);
        }

        public void Dispose() => Unaligned.Dispose();
        public bool GetFontIsStdFont() => new Guid(Unaligned.ReadBytes(16)) == StdFontGuid;

        public string[] ReadArrayStrings(uint cbArrayString)
        {
            if (cbArrayString == 0)
            {
                return new string[0];
            }

            var startPos = Unaligned.BaseStream.Position;
            var endPos = startPos + cbArrayString;
            var ret = new List<string>();

            while (Unaligned.BaseStream.Position < endPos)
            {
                ret.Add(ReadStringFromCcb(ReadCcbWithDoubleSizeIfUncompressed()));
            }

            if (Unaligned.BaseStream.Position != endPos)
            {
                throw new ApplicationException(string.Format(VBASyncResources.ErrorFrxStreamSizeMismatch,
                    "o", "cbArrayString", BaseStream.Position - startPos, cbArrayString));
            }

            return ret.ToArray();
        }

        public BorderStyle ReadBorderStyle() => (BorderStyle)ReadByte();
        public byte ReadByte() => Unaligned.ReadByte();

        public Tuple<int, bool> ReadCcb()
        {
            AlignTo(4);
            var i = Unaligned.ReadInt32();
            return i < 0 ? Tuple.Create(unchecked((int)(i ^ 0x80000000)), true) : Tuple.Create(i, false);
        }

        public Tuple<int, bool> ReadCcbWithDoubleSizeIfUncompressed()
        {
            AlignTo(4);
            var i = Unaligned.ReadInt32();
            return i < 0 ? Tuple.Create(unchecked((int)(i ^ 0x80000000)), true) : Tuple.Create(i * 2, false);
        }

        public Tuple<int, int> ReadCoords()
        {
            AlignTo(4);
            return Tuple.Create(Unaligned.ReadInt32(), Unaligned.ReadInt32());
        }

        public Cycle ReadCycle() => (Cycle)ReadByte();

        public byte[] ReadGuidAndPicture()
        {
            BaseStream.Seek(20, SeekOrigin.Current); // skip GUID and Preamble
            return Unaligned.ReadBytes(Unaligned.ReadInt32());
        }

        public short ReadInt16()
        {
            AlignTo(2);
            return Unaligned.ReadInt16();
        }

        public int ReadInt32()
        {
            AlignTo(4);
            return Unaligned.ReadInt32();
        }

        public MousePointer ReadMousePointer() => (MousePointer)ReadByte();

        public OleColor ReadOleColor()
        {
            AlignTo(4);
            return new OleColor(Unaligned.ReadBytes(4));
        }

        public PictureAlignment ReadPictureAlignment() => (PictureAlignment)ReadByte();
        public PicturePosition ReadPicturePosition() => (PicturePosition)ReadUInt32();
        public PictureSizeMode ReadPictureSizeMode() => (PictureSizeMode)ReadByte();
        public SpecialEffect ReadSpecialEffect() => (SpecialEffect)ReadByte();
        public SpecialEffect ReadSpecialEffect2() => (SpecialEffect)ReadUInt16();
        public SpecialEffect ReadSpecialEffect4() => (SpecialEffect)ReadUInt32();

        public Tuple<short, byte, short, uint, string> ReadStdFont()
        {
            BaseStream.Seek(1, SeekOrigin.Current); // skip Version
            return Tuple.Create(Unaligned.ReadInt16(), Unaligned.ReadByte(), Unaligned.ReadInt16(), Unaligned.ReadUInt32(),
                Encoding.ASCII.GetString(Unaligned.ReadBytes(Unaligned.ReadByte())));
        }

        public string ReadStringFromCcb(Tuple<int, bool> ccb)
        {
            if (ccb.Item1 == 0)
            {
                return "";
            }

            AlignTo(4);
            var s = (ccb.Item2 ? Encoding.UTF8 : Encoding.Unicode).GetString(Unaligned.ReadBytes(ccb.Item1));
            AlignTo(4);
            return s;
        }

        public TextProps ReadTextProps()
        {
            BaseStream.Seek(2, SeekOrigin.Current); // skip MinorVersion and MajorVersion
            var cbTextProps = Unaligned.ReadUInt16();
            BaseStream.Seek(-4, SeekOrigin.Current); // reset to beginning of TextProps
            return cbTextProps > 0 ? new TextProps(Unaligned.ReadBytes(cbTextProps + 4)) : null;
        }

        public ushort ReadUInt16()
        {
            AlignTo(2);
            return Unaligned.ReadUInt16();
        }

        public uint ReadUInt32()
        {
            AlignTo(4);
            return Unaligned.ReadUInt32();
        }

        public ulong ReadUInt64()
        {
            AlignTo(4);
            return Unaligned.ReadUInt64();
        }

        public string ReadWChar()
        {
            AlignTo(2);
            return Encoding.Unicode.GetString(Unaligned.ReadBytes(2));
        }

        public void Skip2Bytes()
        {
            AlignTo(2);
            BaseStream.Seek(2, SeekOrigin.Current);
        }
    }
}
