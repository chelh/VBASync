using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

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
