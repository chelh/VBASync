using System;
using System.IO;
using System.Text;
using static VbaSync.FrxObjects.AlignmentHelpers;
using static VbaSync.FrxObjects.StreamDataHelpers;

namespace VbaSync.FrxObjects {
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

    class TextProps {
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

                // DataBlock
                ushort dataBlockBytes = 0;
                var fontNameLength = 0;
                var fontNameCompressed = false;
                if (PropMask.HasFontName) {
                    fontNameLength = CcbToLength(r.ReadInt32(), out fontNameCompressed);
                    dataBlockBytes += 4;
                }
                if (PropMask.HasFontEffects) {
                    FontEffects = r.ReadUInt32();
                    dataBlockBytes += 4;
                }
                if (PropMask.HasFontHeight) {
                    FontHeight = r.ReadUInt32();
                    dataBlockBytes += 4;
                }
                if (PropMask.HasFontCharSet) {
                    FontCharSet = r.ReadByte();
                    dataBlockBytes += 4;
                }
                if (PropMask.HasFontPitchAndFamily) {
                    FontPitchAndFamily = r.ReadByte();
                    dataBlockBytes += 1;
                }
                if (PropMask.HasParagraphAlign) {
                    ParagraphAlign = r.ReadByte();
                    dataBlockBytes += 1;
                }
                AlignTo(4, st, ref dataBlockBytes);
                if (PropMask.HasFontWeight) {
                    FontWeight = r.ReadUInt16();
                }
                AlignTo(4, st, ref dataBlockBytes);

                // ExtraDataBlock
                if (fontNameLength > 0) {
                    FontName = (fontNameCompressed ? Encoding.UTF8 : Encoding.Unicode).GetString(r.ReadBytes(fontNameLength));
                }
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
