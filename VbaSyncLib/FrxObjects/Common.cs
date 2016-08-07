using System;
using System.IO;
using System.Text;

namespace VbaSync.FrxObjects {
    abstract class FrxCommon {
        ushort _dBytes;
        bool _inD;
        bool _inX;
        ushort _xBytes;

        protected ushort DataBlockBytes => _dBytes;
        protected ushort ExtraDataBlockBytes => _xBytes;

        protected void BeginDataBlock() {
            _dBytes = 0;
            _inD = true;
        }

        protected void BeginExtraDataBlock() {
            _xBytes = 0;
            _inX = true;
        }

        protected void EndDataBlock(BinaryReader r) {
            AlignTo(4, r.BaseStream);
            _inD = false;
        }

        protected void EndExtraDataBlock(BinaryReader r) {
            AlignTo(4, r.BaseStream);
            _inX = false;
        }

        protected byte ReadByteIf(bool b, BinaryReader r, byte ifNot = 0) {
            if (!b) return ifNot;
            AboutToRead(1);
            return r.ReadByte();
        }

        protected void Ignore2AlignedBytesIf(bool b, BinaryReader r) {
            if (!b) return;
            AlignTo(2, r.BaseStream);
            IgnoreNext(2, r.BaseStream);
        }

        protected short ReadAlignedInt16If(bool b, BinaryReader r, short ifNot = 0) {
            if (!b) return ifNot;
            AlignTo(2, r.BaseStream);
            AboutToRead(2);
            return r.ReadInt16();
        }

        protected ushort ReadAlignedUInt16If(bool b, BinaryReader r, ushort ifNot = 0) {
            if (!b) return ifNot;
            AlignTo(2, r.BaseStream);
            AboutToRead(2);
            return r.ReadUInt16();
        }

        protected int ReadAlignedInt32If(bool b, BinaryReader r, int ifNot = 0) {
            if (!b) return ifNot;
            AlignTo(4, r.BaseStream);
            AboutToRead(4);
            return r.ReadInt32();
        }

        protected uint ReadAlignedUInt32If(bool b, BinaryReader r, uint ifNot = 0) {
            if (!b) return ifNot;
            AlignTo(4, r.BaseStream);
            AboutToRead(4);
            return r.ReadUInt32();
        }

        protected Tuple<int, int> ReadAlignedCoordsIf(bool b, BinaryReader r) {
            if (!b) return Tuple.Create(0, 0);
            AlignTo(4, r.BaseStream);
            AboutToRead(8);
            return Tuple.Create(r.ReadInt32(), r.ReadInt32());
        }

        protected Tuple<int, bool> ReadAlignedCcbIf(bool b, BinaryReader r) {
            if (!b) return Tuple.Create(0, false);
            AlignTo(4, r.BaseStream);
            AboutToRead(4);
            var i = r.ReadInt32();
            return i < 0 ? Tuple.Create(unchecked((int)(i ^ 0x80000000)), true) : Tuple.Create(i, false);
        }

        protected OleColor ReadAlignedOleColorIf(bool b, BinaryReader r) {
            if (!b) return null;
            AlignTo(4, r.BaseStream);
            AboutToRead(4);
            return new OleColor(r.ReadBytes(4));
        }

        protected string ReadAlignedWCharIf(bool b, BinaryReader r) {
            if (!b) return "";
            AlignTo(2, r.BaseStream);
            AboutToRead(2);
            return Encoding.Unicode.GetString(r.ReadBytes(2));
        }

        protected string ReadStringFromCcb(Tuple<int, bool> ccb, BinaryReader r) {
            if (ccb.Item1 == 0) return "";
            AboutToRead((ushort)ccb.Item1);
            var s = (ccb.Item2 ? Encoding.UTF8 : Encoding.Unicode).GetString(r.ReadBytes(ccb.Item1));
            AlignTo(4, r.BaseStream);
            return s;
        }

        void AboutToRead(ushort numBytes) {
            if (_inD) {
                _dBytes += numBytes;
            } else if (_inX) {
                _xBytes += numBytes;
            }
        }

        void AlignTo(ushort alignment, Stream st) {
            if (_inD) {
                if (_dBytes%alignment == 0) return;
                st.Seek(alignment - _dBytes%alignment, SeekOrigin.Current);
                _dBytes += (ushort)(alignment - _dBytes%alignment);
            } else if (_inX) {
                if (_xBytes%alignment == 0) return;
                st.Seek(alignment - _xBytes%alignment, SeekOrigin.Current);
                _xBytes += (ushort)(alignment - _xBytes%alignment);
            }
        }

        void IgnoreNext(ushort bytes, Stream st) {
            st.Seek(bytes, SeekOrigin.Current);
            if (_inD) {
                _dBytes += bytes;
            } else if (_inX) {
                _xBytes += bytes;
            }
        }
    }

    class OleColor {
        public OleColorType ColorType { get; }
        public ushort PaletteIndex { get; }
        public byte Red { get; }
        public byte Blue { get; }
        public byte Green { get; }

        public OleColor(byte[] b) {
            if (b.Length != 4)
                throw new ArgumentException($"Error creating {nameof(OleColor)}. Expected 4 bytes but got {b.Length}.", nameof(b));
            Blue = b[0];
            Green = b[1];
            Red = b[2];
            PaletteIndex = unchecked((ushort)((Green << 8) | Blue));
            ColorType = (OleColorType)b[3];
        }

        public override bool Equals(object o) {
            var other = o as OleColor;
            return other != null && ColorType == other.ColorType && Red == other.Red && Blue == other.Blue && Green == other.Green;
        }

        public override int GetHashCode() {
            unchecked {
                var hashCode = (int)ColorType;
                hashCode = (hashCode * 397) ^ Red.GetHashCode();
                hashCode = (hashCode * 397) ^ Blue.GetHashCode();
                hashCode = (hashCode * 397) ^ Green.GetHashCode();
                return hashCode;
            }
        }
    }

    class TextProps : FrxCommon {
        public byte MinorVersion { get; }
        public byte MajorVersion { get; }
        public TextPropsPropMask PropMask { get; }
        public uint FontEffects { get; }
        public uint FontHeight { get; }
        public byte FontCharSet { get; }
        public byte FontPitchAndFamily { get; }
        public byte ParagraphAlign { get; }
        public ushort FontWeight { get; }
        public string FontName { get; }

        public TextProps(byte[] b) {
            using (var st = new MemoryStream(b))
            using (var r = new BinaryReader(st)) {
                MinorVersion = r.ReadByte();
                MajorVersion = r.ReadByte();

                var cbTextProps = r.ReadUInt16();
                PropMask = new TextPropsPropMask(r.ReadUInt32());

                BeginDataBlock();
                var fontNameCcb = ReadAlignedCcbIf(PropMask.HasFontName, r);
                FontEffects = ReadAlignedUInt32If(PropMask.HasFontEffects, r);
                FontHeight = ReadAlignedUInt32If(PropMask.HasFontHeight, r);
                FontCharSet = ReadByteIf(PropMask.HasFontCharSet, r);
                FontPitchAndFamily = ReadByteIf(PropMask.HasFontPitchAndFamily, r);
                ParagraphAlign = ReadByteIf(PropMask.HasParagraphAlign, r);
                FontWeight = ReadAlignedUInt16If(PropMask.HasFontWeight, r);
                EndDataBlock(r);

                BeginExtraDataBlock();
                FontName = ReadStringFromCcb(fontNameCcb, r);
                EndExtraDataBlock(r);
            }
        }

        public override bool Equals(object obj) {
            if (ReferenceEquals(null, obj)) {
                return false;
            }
            if (ReferenceEquals(this, obj)) {
                return true;
            }
            if (obj.GetType() != this.GetType()) {
                return false;
            }
            return Equals((TextProps)obj);
        }

        protected bool Equals(TextProps other) {
            return MinorVersion == other.MinorVersion && MajorVersion == other.MajorVersion && Equals(PropMask, other.PropMask) &&
                   FontEffects == other.FontEffects && FontHeight == other.FontHeight && FontCharSet == other.FontCharSet &&
                   FontPitchAndFamily == other.FontPitchAndFamily && ParagraphAlign == other.ParagraphAlign && FontWeight == other.FontWeight &&
                   string.Equals(FontName, other.FontName);
        }

        public override int GetHashCode() {
            unchecked {
                var hashCode = MinorVersion.GetHashCode();
                hashCode = (hashCode*397) ^ MajorVersion.GetHashCode();
                hashCode = (hashCode*397) ^ (PropMask?.GetHashCode() ?? 0);
                hashCode = (hashCode*397) ^ (int)FontEffects;
                hashCode = (hashCode*397) ^ (int)FontHeight;
                hashCode = (hashCode*397) ^ FontCharSet.GetHashCode();
                hashCode = (hashCode*397) ^ FontPitchAndFamily.GetHashCode();
                hashCode = (hashCode*397) ^ ParagraphAlign.GetHashCode();
                hashCode = (hashCode*397) ^ FontWeight.GetHashCode();
                hashCode = (hashCode*397) ^ (FontName?.GetHashCode() ?? 0);
                return hashCode;
            }
        }
    }

    class TextPropsPropMask {
        public bool HasFontName { get; }
        public bool HasFontEffects { get; }
        public bool HasFontHeight { get; }
        public bool HasFontCharSet { get; }
        public bool HasFontPitchAndFamily { get; }
        public bool HasParagraphAlign { get; }
        public bool HasFontWeight { get; }

        public TextPropsPropMask(uint i) {
            Func<int, bool> bit = j => (i & ((uint)1 << j)) != 0;
            HasFontName = bit(0);
            HasFontEffects = bit(1);
            HasFontHeight = bit(2);
            HasFontCharSet = bit(4);
            HasFontPitchAndFamily = bit(5);
            HasParagraphAlign = bit(6);
            HasFontWeight = bit(7);
        }

        public override bool Equals(object obj) {
            if (ReferenceEquals(null, obj)) {
                return false;
            }
            if (ReferenceEquals(this, obj)) {
                return true;
            }
            if (obj.GetType() != this.GetType()) {
                return false;
            }
            return Equals((TextPropsPropMask)obj);
        }

        protected bool Equals(TextPropsPropMask other) {
            return HasFontName == other.HasFontName && HasFontEffects == other.HasFontEffects && HasFontHeight == other.HasFontHeight &&
                   HasFontCharSet == other.HasFontCharSet && HasFontPitchAndFamily == other.HasFontPitchAndFamily &&
                   HasParagraphAlign == other.HasParagraphAlign && HasFontWeight == other.HasFontWeight;
        }

        public override int GetHashCode() {
            unchecked {
                var hashCode = HasFontName.GetHashCode();
                hashCode = (hashCode*397) ^ HasFontEffects.GetHashCode();
                hashCode = (hashCode*397) ^ HasFontHeight.GetHashCode();
                hashCode = (hashCode*397) ^ HasFontCharSet.GetHashCode();
                hashCode = (hashCode*397) ^ HasFontPitchAndFamily.GetHashCode();
                hashCode = (hashCode*397) ^ HasParagraphAlign.GetHashCode();
                hashCode = (hashCode*397) ^ HasFontWeight.GetHashCode();
                return hashCode;
            }
        }
    }

    enum BorderStyle {
        None = 0x00,
        Single = 0x01
    }

    enum OleColorType {
        Default = 0x00,
        PaletteEntry = 0x01,
        RgbColor = 0x02,
        SystemPalette = 0x80
    }

    enum MousePointer {
        Default = 0x00,
        Arrow = 0x01,
        Cross = 0x02,
        IBeam = 0x03,
        SizeNesw = 0x06,
        SizeNs = 0x07,
        SizeNwse = 0x08,
        SizeWe = 0x09,
        UpArrow = 0x0A,
        HourGlass = 0x0B,
        NoDrop = 0x0C,
        AppStarting = 0x0D,
        Help = 0x0E,
        SizeAll = 0x0F,
        Custom = 0x63
    }

    enum SpecialEffect {
        Flat = 0x00,
        Raised = 0x01,
        Sunken = 0x02,
        Etched = 0x03,
        Bump = 0x06
    }

    enum PictureAlignment {
        TopLeft = 0x00,
        TopRight = 0x01,
        Center = 0x02,
        BottomLeft = 0x03,
        BottomRight = 0x04
    }

    enum PictureSizeMode {
        Clip = 0x00,
        Stretch = 0x01,
        Zoom = 0x03
    }

    enum PicturePosition {
        LeftTop = 0x00020000,
        LeftCenter = 0x00050003,
        LeftBottom = 0x00080006,
        RightTop = 0x00000002,
        RightCenter = 0x00030005,
        RightBottom = 0x00060008,
        AboveLeft = 0x00060000,
        AboveCenter = 0x00070001,
        AboveRight = 0x00080002,
        BelowLeft = 0x00000006,
        BelowCenter = 0x00010007,
        BelowRight = 0x00020008,
        Center = 0x00040004
    }
}
