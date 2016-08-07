using System;
using System.IO;
using System.Linq;
using static VbaSync.FrxObjects.StreamDataHelpers;

namespace VbaSync.FrxObjects {
    class CommandButtonControl : FrxCommon {
        public byte MinorVersion { get; }
        public byte MajorVersion { get; }
        public CommandButtonPropMask PropMask { get; }
        public OleColor ForeColor { get; }
        public OleColor BackColor { get; }
        public uint VariousPropertyBits { get; }
        public string Caption { get; }
        public PicturePosition PicturePosition { get; }
        public MousePointer MousePointer { get; }
        public string Accelerator { get; }
        public Tuple<int, int> Size { get; }
        public byte[] Picture { get; } = new byte[0];
        public byte[] MouseIcon { get; } = new byte[0];
        public TextProps TextProps { get; }

        public CommandButtonControl(byte[] b) {
            using (var st = new MemoryStream(b))
            using (var r = new BinaryReader(st)) {
                MinorVersion = r.ReadByte();
                MajorVersion = r.ReadByte();

                var cbCommandButton = r.ReadUInt16();
                PropMask = new CommandButtonPropMask(r.ReadUInt32());

                BeginDataBlock();
                ForeColor = ReadAlignedOleColorIf(PropMask.HasForeColor, r);
                BackColor = ReadAlignedOleColorIf(PropMask.HasBackColor, r);
                VariousPropertyBits = ReadAlignedUInt32If(PropMask.HasVariousPropertyBits, r);
                var captionCcb = ReadAlignedCcbIf(PropMask.HasCaption, r);
                PicturePosition = (PicturePosition)ReadAlignedUInt32If(PropMask.HasPicturePosition, r);
                MousePointer = (MousePointer)ReadByteIf(PropMask.HasMousePointer, r);
                Ignore2AlignedBytesIf(PropMask.HasPicture, r);
                Accelerator = ReadAlignedWCharIf(PropMask.HasAccelerator, r);
                Ignore2AlignedBytesIf(PropMask.HasMouseIcon, r);
                EndDataBlock(r);

                BeginExtraDataBlock();
                Caption = ReadStringFromCcb(captionCcb, r);
                Size = ReadAlignedCoordsIf(PropMask.HasSize, r);
                EndExtraDataBlock(r);

                if (cbCommandButton != 4 + DataBlockBytes + ExtraDataBlockBytes)
                    throw new ApplicationException("Error reading 'o' stream in .frx data: expected cbCommandButton size "
                                                   + $"{4 + DataBlockBytes + ExtraDataBlockBytes}, but actual size was {cbCommandButton}.");
                
                // StreamData
                if (PropMask.HasPicture) {
                    Picture = ReadGuidAndPicture(r);
                }
                if (PropMask.HasMouseIcon) {
                    MouseIcon = ReadGuidAndPicture(r);
                }

                TextProps = ReadTextProps(r);
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
            return Equals((CommandButtonControl)obj);
        }

        protected bool Equals(CommandButtonControl other) {
            return Picture.SequenceEqual(other.Picture) && MouseIcon.SequenceEqual(other.MouseIcon) && MinorVersion == other.MinorVersion &&
                   MajorVersion == other.MajorVersion && Equals(PropMask, other.PropMask) && Equals(ForeColor, other.ForeColor) &&
                   Equals(BackColor, other.BackColor) && VariousPropertyBits == other.VariousPropertyBits && string.Equals(Caption, other.Caption) &&
                   PicturePosition == other.PicturePosition && MousePointer == other.MousePointer && string.Equals(Accelerator, other.Accelerator) &&
                   Equals(Size, other.Size) && Equals(TextProps, other.TextProps);
        }

        public override int GetHashCode() {
            unchecked {
                var hashCode = Picture?.Length.GetHashCode() ?? 0;
                hashCode = (hashCode * 397) ^ (MouseIcon?.Length.GetHashCode() ?? 0);
                hashCode = (hashCode * 397) ^ (TextProps?.GetHashCode() ?? 0);
                hashCode = (hashCode * 397) ^ MinorVersion.GetHashCode();
                hashCode = (hashCode * 397) ^ MajorVersion.GetHashCode();
                hashCode = (hashCode * 397) ^ (PropMask?.GetHashCode() ?? 0);
                hashCode = (hashCode * 397) ^ (ForeColor?.GetHashCode() ?? 0);
                hashCode = (hashCode * 397) ^ (BackColor?.GetHashCode() ?? 0);
                hashCode = (hashCode * 397) ^ (int)VariousPropertyBits;
                hashCode = (hashCode * 397) ^ (Caption?.GetHashCode() ?? 0);
                hashCode = (hashCode * 397) ^ (int)PicturePosition;
                hashCode = (hashCode * 397) ^ (int)MousePointer;
                hashCode = (hashCode * 397) ^ (Accelerator?.GetHashCode() ?? 0);
                hashCode = (hashCode * 397) ^ (Size?.GetHashCode() ?? 0);
                return hashCode;
            }
        }
    }

    class CommandButtonPropMask {
        public bool HasForeColor { get; }
        public bool HasBackColor { get; }
        public bool HasVariousPropertyBits { get; }
        public bool HasCaption { get; }
        public bool HasPicturePosition { get; }
        public bool HasSize { get; }
        public bool HasMousePointer { get; }
        public bool HasPicture { get; }
        public bool HasAccelerator { get; }
        public bool TakeFocusOnClick { get; }
        public bool HasMouseIcon { get; }

        public CommandButtonPropMask(uint i) {
            Func<int, bool> bit = j => (i & ((uint)1 << j)) != 0;
            HasForeColor = bit(0);
            HasBackColor = bit(1);
            HasVariousPropertyBits = bit(2);
            HasCaption = bit(3);
            HasPicturePosition = bit(4);
            HasSize = bit(5);
            HasMousePointer = bit(6);
            HasPicture = bit(7);
            HasAccelerator = bit(8);
            TakeFocusOnClick = !bit(9);
            HasMouseIcon = bit(10);
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
            return Equals((CommandButtonPropMask)obj);
        }

        protected bool Equals(CommandButtonPropMask other) {
            return HasForeColor == other.HasForeColor && HasBackColor == other.HasBackColor && HasVariousPropertyBits == other.HasVariousPropertyBits &&
                   HasCaption == other.HasCaption && HasPicturePosition == other.HasPicturePosition && HasSize == other.HasSize &&
                   HasMousePointer == other.HasMousePointer && HasPicture == other.HasPicture && HasAccelerator == other.HasAccelerator &&
                   TakeFocusOnClick == other.TakeFocusOnClick && HasMouseIcon == other.HasMouseIcon;
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
                hashCode = (hashCode*397) ^ HasPicture.GetHashCode();
                hashCode = (hashCode*397) ^ HasAccelerator.GetHashCode();
                hashCode = (hashCode*397) ^ TakeFocusOnClick.GetHashCode();
                hashCode = (hashCode*397) ^ HasMouseIcon.GetHashCode();
                return hashCode;
            }
        }
    }
}
