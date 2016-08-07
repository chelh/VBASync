using System;
using System.IO;
using System.Linq;
using System.Text;
using static VbaSync.FrxObjects.AlignmentHelpers;
using static VbaSync.FrxObjects.StreamDataHelpers;

namespace VbaSync.FrxObjects {
    class LabelControl {
        public byte MinorVersion { get; }
        public byte MajorVersion { get; }
        public LabelPropMask PropMask { get; }
        public OleColor ForeColor { get; }
        public OleColor BackColor { get; }
        public uint VariousPropertyBits { get; }
        public string Caption { get; }
        public PicturePosition PicturePosition { get; }
        public MousePointer MousePointer { get; }
        public OleColor BorderColor { get; }
        public BorderStyle BorderStyle { get; }
        public SpecialEffect SpecialEffect { get; }
        public string Accelerator { get; }
        public Tuple<int, int> Size { get; }
        public byte[] Picture { get; } = new byte[0];
        public byte[] MouseIcon { get; } = new byte[0];
        
        public LabelControl(byte[] b) {
            using (var st = new MemoryStream(b))
            using (var r = new BinaryReader(st)) {
                MinorVersion = r.ReadByte();
                MajorVersion = r.ReadByte();

                var cbLabel = r.ReadUInt16();
                PropMask = new LabelPropMask(r.ReadUInt32());

                // DataBlock
                ushort dataBlockBytes = 0;
                if (PropMask.HasForeColor) {
                    ForeColor = new OleColor(r.ReadBytes(4));
                    dataBlockBytes += 4;
                }
                if (PropMask.HasBackColor) {
                    BackColor = new OleColor(r.ReadBytes(4));
                    dataBlockBytes += 4;
                }
                if (PropMask.HasVariousPropertyBits) {
                    VariousPropertyBits = r.ReadUInt32();
                    dataBlockBytes += 4;
                }
                var captionLength = 0;
                var captionCompressed = false;
                if (PropMask.HasCaption) {
                    captionLength = CcbToLength(r.ReadInt32(), out captionCompressed);
                    dataBlockBytes += 4;
                }
                if (PropMask.HasPicturePosition) {
                    PicturePosition = (PicturePosition)r.ReadUInt32();
                    dataBlockBytes += 4;
                }
                if (PropMask.HasMousePointer) {
                    MousePointer = (MousePointer)r.ReadByte();
                    dataBlockBytes += 1;
                }
                AlignTo(4, st, ref dataBlockBytes);
                if (PropMask.HasBorderColor) {
                    BorderColor = new OleColor(r.ReadBytes(4));
                    dataBlockBytes += 4;
                }
                if (PropMask.HasBorderStyle) {
                    BorderStyle = (BorderStyle)r.ReadInt16();
                    dataBlockBytes += 2;
                }
                AlignTo(4, st, ref dataBlockBytes);
                if (PropMask.HasSpecialEffect) {
                    SpecialEffect = (SpecialEffect)r.ReadInt16();
                    dataBlockBytes += 2;
                }
                AlignTo(4, st, ref dataBlockBytes);
                if (PropMask.HasPicture) IgnoreNext(2, st, ref dataBlockBytes);
                AlignTo(4, st, ref dataBlockBytes);
                if (PropMask.HasAccelerator) {
                    Accelerator = Encoding.Unicode.GetString(r.ReadBytes(2));
                    dataBlockBytes += 2;
                }
                AlignTo(4, st, ref dataBlockBytes);
                if (PropMask.HasMouseIcon) IgnoreNext(2, st, ref dataBlockBytes);
                AlignTo(4, st, ref dataBlockBytes);

                // ExtraDataBlock
                ushort extraDataBlockBytes = 0;
                if (captionLength > 0) {
                    Caption = (captionCompressed ? Encoding.UTF8 : Encoding.Unicode).GetString(r.ReadBytes(captionLength));
                    extraDataBlockBytes += (ushort)captionLength;
                }
                AlignTo(4, st, ref extraDataBlockBytes);
                if (PropMask.HasSize) {
                    Size = Tuple.Create(r.ReadInt32(), r.ReadInt32());
                    extraDataBlockBytes += 4;
                }

                // StreamData
                if (PropMask.HasPicture) {
                    Picture = ReadGuidAndPicture(r);
                }
                if (PropMask.HasMouseIcon) {
                    MouseIcon = ReadGuidAndPicture(r);
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
            return Equals((LabelControl)obj);
        }

        protected bool Equals(LabelControl other) {
            return Picture.SequenceEqual(other.Picture) && MouseIcon.SequenceEqual(other.MouseIcon) && MinorVersion == other.MinorVersion &&
                   MajorVersion == other.MajorVersion && Equals(PropMask, other.PropMask) && Equals(ForeColor, other.ForeColor) &&
                   Equals(BackColor, other.BackColor) && VariousPropertyBits == other.VariousPropertyBits && string.Equals(Caption, other.Caption) &&
                   PicturePosition == other.PicturePosition && MousePointer == other.MousePointer && Equals(BorderColor, other.BorderColor) &&
                   BorderStyle == other.BorderStyle && SpecialEffect == other.SpecialEffect && string.Equals(Accelerator, other.Accelerator) &&
                   Equals(Size, other.Size);
        }

        public override int GetHashCode() {
            unchecked {
                var hashCode = Picture?.Length.GetHashCode() ?? 0;
                hashCode = (hashCode*397) ^ (MouseIcon?.Length.GetHashCode() ?? 0);
                hashCode = (hashCode*397) ^ MinorVersion.GetHashCode();
                hashCode = (hashCode*397) ^ MajorVersion.GetHashCode();
                hashCode = (hashCode*397) ^ (PropMask?.GetHashCode() ?? 0);
                hashCode = (hashCode*397) ^ (ForeColor?.GetHashCode() ?? 0);
                hashCode = (hashCode*397) ^ (BackColor?.GetHashCode() ?? 0);
                hashCode = (hashCode*397) ^ (int)VariousPropertyBits;
                hashCode = (hashCode*397) ^ (Caption?.GetHashCode() ?? 0);
                hashCode = (hashCode*397) ^ (int)PicturePosition;
                hashCode = (hashCode*397) ^ (int)MousePointer;
                hashCode = (hashCode*397) ^ (BorderColor?.GetHashCode() ?? 0);
                hashCode = (hashCode*397) ^ (int)BorderStyle;
                hashCode = (hashCode*397) ^ (int)SpecialEffect;
                hashCode = (hashCode*397) ^ (Accelerator?.GetHashCode() ?? 0);
                hashCode = (hashCode*397) ^ (Size?.GetHashCode() ?? 0);
                return hashCode;
            }
        }
    }

    class LabelPropMask {
        public bool HasForeColor { get; }
        public bool HasBackColor { get; }
        public bool HasVariousPropertyBits { get; }
        public bool HasCaption { get; }
        public bool HasPicturePosition { get; }
        public bool HasSize { get; }
        public bool HasMousePointer { get; }
        public bool HasBorderColor { get; }
        public bool HasBorderStyle { get; }
        public bool HasSpecialEffect { get; }
        public bool HasPicture { get; }
        public bool HasAccelerator { get; }
        public bool HasMouseIcon { get; }

        public LabelPropMask(uint i) {
            Func<int, bool> bit = j => (i & ((uint)1 << j)) != 0;
            HasForeColor = bit(0);
            HasBackColor = bit(1);
            HasVariousPropertyBits = bit(2);
            HasCaption = bit(3);
            HasPicturePosition = bit(4);
            HasSize = bit(5);
            HasMousePointer = bit(6);
            HasBorderColor = bit(7);
            HasBorderStyle = bit(8);
            HasSpecialEffect = bit(9);
            HasPicture = bit(10);
            HasAccelerator = bit(11);
            HasMouseIcon = bit(12);
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
            return Equals((LabelPropMask)obj);
        }

        protected bool Equals(LabelPropMask other) {
            return HasForeColor == other.HasForeColor && HasBackColor == other.HasBackColor && HasVariousPropertyBits == other.HasVariousPropertyBits &&
                   HasCaption == other.HasCaption && HasPicturePosition == other.HasPicturePosition && HasSize == other.HasSize &&
                   HasMousePointer == other.HasMousePointer && HasBorderColor == other.HasBorderColor && HasBorderStyle == other.HasBorderStyle &&
                   HasSpecialEffect == other.HasSpecialEffect && HasPicture == other.HasPicture && HasAccelerator == other.HasAccelerator &&
                   HasMouseIcon == other.HasMouseIcon;
        }

        public override int GetHashCode() {
            unchecked {
                var hashCode = HasForeColor.GetHashCode();
                hashCode = (hashCode*397) ^ HasBackColor.GetHashCode();
                hashCode = (hashCode*397) ^ HasVariousPropertyBits.GetHashCode();
                hashCode = (hashCode*397) ^ HasCaption.GetHashCode();
                hashCode = (hashCode*397) ^ HasPicturePosition.GetHashCode();
                hashCode = (hashCode*397) ^ HasSize.GetHashCode();
                hashCode = (hashCode*397) ^ HasMousePointer.GetHashCode();
                hashCode = (hashCode*397) ^ HasBorderColor.GetHashCode();
                hashCode = (hashCode*397) ^ HasBorderStyle.GetHashCode();
                hashCode = (hashCode*397) ^ HasSpecialEffect.GetHashCode();
                hashCode = (hashCode*397) ^ HasPicture.GetHashCode();
                hashCode = (hashCode*397) ^ HasAccelerator.GetHashCode();
                hashCode = (hashCode*397) ^ HasMouseIcon.GetHashCode();
                return hashCode;
            }
        }
    }
}
