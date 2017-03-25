using System;
using System.IO;
using VBASync.Localization;

namespace VBASync.Model.FrxObjects
{
    internal enum BorderStyle
    {
        None = 0x00,
        Single = 0x01
    }

    internal enum MousePointer
    {
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

    internal enum OleColorType
    {
        Default = 0x00,
        PaletteEntry = 0x01,
        RgbColor = 0x02,
        SystemPalette = 0x80
    }

    internal enum PictureAlignment
    {
        TopLeft = 0x00,
        TopRight = 0x01,
        Center = 0x02,
        BottomLeft = 0x03,
        BottomRight = 0x04
    }

    internal enum PicturePosition
    {
        RightTop = 0x00000002,
        BelowLeft = 0x00000006,
        BelowCenter = 0x00010007,
        LeftTop = 0x00020000,
        BelowRight = 0x00020008,
        RightCenter = 0x00030005,
        Center = 0x00040004,
        LeftCenter = 0x00050003,
        AboveLeft = 0x00060000,
        RightBottom = 0x00060008,
        AboveCenter = 0x00070001,
        AboveRight = 0x00080002,
        LeftBottom = 0x00080006
    }

    internal enum PictureSizeMode
    {
        Clip = 0x00,
        Stretch = 0x01,
        Zoom = 0x03
    }

    internal enum SpecialEffect
    {
        Flat = 0x00,
        Raised = 0x01,
        Sunken = 0x02,
        Etched = 0x03,
        Bump = 0x06
    }

    internal class OleColor
    {
        public OleColor(byte[] b)
        {
            if (b.Length != 4)
            {
                throw new ApplicationException(string.Format(VBASyncResources.ErrorFrxColorSizeMismatch, b.Length));
            }

            Blue = b[0];
            Green = b[1];
            Red = b[2];
            PaletteIndex = unchecked((ushort)((Green << 8) | Blue));
            ColorType = (OleColorType)b[3];
        }

        public byte Blue { get; }
        public OleColorType ColorType { get; }
        public byte Green { get; }
        public ushort PaletteIndex { get; }
        public byte Red { get; }
        public override bool Equals(object obj)
        {
            var other = obj as OleColor;
            return other != null && ColorType == other.ColorType && Red == other.Red && Blue == other.Blue && Green == other.Green;
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = (int)ColorType;
                hashCode = (hashCode * 397) ^ Red.GetHashCode();
                hashCode = (hashCode * 397) ^ Blue.GetHashCode();
                hashCode = (hashCode * 397) ^ Green.GetHashCode();
                return hashCode;
            }
        }
    }

    internal class TextProps
    {
        public TextProps(byte[] b)
        {
            using (var st = new MemoryStream(b))
            using (var r = new FrxReader(st))
            {
                MinorVersion = r.ReadByte();
                MajorVersion = r.ReadByte();

                var cbTextProps = r.ReadUInt16();
                PropMask = new TextPropsPropMask(r.ReadUInt32());

                // DataBlock
                var fontNameCcb = PropMask.HasFontName ? r.ReadCcb() : Tuple.Create(0, false);
                FontEffects = PropMask.HasFontEffects ? r.ReadUInt32() : 0;
                FontHeight = PropMask.HasFontHeight ? r.ReadUInt32() : 0;
                FontCharSet = PropMask.HasFontCharSet ? r.ReadByte() : (byte)0;
                FontPitchAndFamily = PropMask.HasFontPitchAndFamily ? r.ReadByte() : (byte)0;
                ParagraphAlign = PropMask.HasParagraphAlign ? r.ReadByte() : (byte)0;
                FontWeight = PropMask.HasFontWeight ? r.ReadUInt16() : (ushort)0;

                // ExtraDataBlock
                FontName = r.ReadStringFromCcb(fontNameCcb);
            }
        }

        public byte FontCharSet { get; }
        public uint FontEffects { get; }
        public uint FontHeight { get; }
        public string FontName { get; }
        public byte FontPitchAndFamily { get; }
        public ushort FontWeight { get; }
        public byte MajorVersion { get; }
        public byte MinorVersion { get; }
        public byte ParagraphAlign { get; }
        public TextPropsPropMask PropMask { get; }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj))
            {
                return false;
            }
            if (ReferenceEquals(this, obj))
            {
                return true;
            }
            if (obj.GetType() != GetType())
            {
                return false;
            }
            return Equals((TextProps)obj);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = MinorVersion.GetHashCode();
                hashCode = (hashCode * 397) ^ MajorVersion.GetHashCode();
                hashCode = (hashCode * 397) ^ (PropMask?.GetHashCode() ?? 0);
                hashCode = (hashCode * 397) ^ (int)FontEffects;
                hashCode = (hashCode * 397) ^ (int)FontHeight;
                hashCode = (hashCode * 397) ^ FontCharSet.GetHashCode();
                hashCode = (hashCode * 397) ^ FontPitchAndFamily.GetHashCode();
                hashCode = (hashCode * 397) ^ ParagraphAlign.GetHashCode();
                hashCode = (hashCode * 397) ^ FontWeight.GetHashCode();
                hashCode = (hashCode * 397) ^ (FontName?.GetHashCode() ?? 0);
                return hashCode;
            }
        }

        protected bool Equals(TextProps other)
        {
            return MinorVersion == other.MinorVersion && MajorVersion == other.MajorVersion && Equals(PropMask, other.PropMask)
                   && FontEffects == other.FontEffects && FontHeight == other.FontHeight && FontCharSet == other.FontCharSet
                   && FontPitchAndFamily == other.FontPitchAndFamily && ParagraphAlign == other.ParagraphAlign && FontWeight == other.FontWeight
                   && string.Equals(FontName, other.FontName);
        }
    }

    internal class TextPropsPropMask
    {
        public TextPropsPropMask(uint i)
        {
            Func<int, bool> bit = j => (i & ((uint)1 << j)) != 0;
            HasFontName = bit(0);
            HasFontEffects = bit(1);
            HasFontHeight = bit(2);
            HasFontCharSet = bit(4);
            HasFontPitchAndFamily = bit(5);
            HasParagraphAlign = bit(6);
            HasFontWeight = bit(7);
        }

        public bool HasFontCharSet { get; }
        public bool HasFontEffects { get; }
        public bool HasFontHeight { get; }
        public bool HasFontName { get; }
        public bool HasFontPitchAndFamily { get; }
        public bool HasFontWeight { get; }
        public bool HasParagraphAlign { get; }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj))
            {
                return false;
            }
            if (ReferenceEquals(this, obj))
            {
                return true;
            }
            if (obj.GetType() != GetType())
            {
                return false;
            }
            return Equals((TextPropsPropMask)obj);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = HasFontName.GetHashCode();
                hashCode = (hashCode * 397) ^ HasFontEffects.GetHashCode();
                hashCode = (hashCode * 397) ^ HasFontHeight.GetHashCode();
                hashCode = (hashCode * 397) ^ HasFontCharSet.GetHashCode();
                hashCode = (hashCode * 397) ^ HasFontPitchAndFamily.GetHashCode();
                hashCode = (hashCode * 397) ^ HasParagraphAlign.GetHashCode();
                hashCode = (hashCode * 397) ^ HasFontWeight.GetHashCode();
                return hashCode;
            }
        }

        protected bool Equals(TextPropsPropMask other)
        {
            return HasFontName == other.HasFontName && HasFontEffects == other.HasFontEffects && HasFontHeight == other.HasFontHeight
                   && HasFontCharSet == other.HasFontCharSet && HasFontPitchAndFamily == other.HasFontPitchAndFamily
                   && HasParagraphAlign == other.HasParagraphAlign && HasFontWeight == other.HasFontWeight;
        }
    }
}
