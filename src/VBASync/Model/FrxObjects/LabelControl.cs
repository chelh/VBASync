using System;
using System.IO;
using System.Linq;
using VBASync.Localization;

namespace VBASync.Model.FrxObjects
{
    internal class LabelControl
    {
        public LabelControl(byte[] b)
        {
            using (var st = new MemoryStream(b))
            using (var r = new FrxReader(st))
            {
                MinorVersion = r.ReadByte();
                MajorVersion = r.ReadByte();

                var cbLabel = r.ReadUInt16();
                PropMask = new LabelPropMask(r.ReadUInt32());

                // DataBlock
                ForeColor = PropMask.HasForeColor ? r.ReadOleColor() : null;
                BackColor = PropMask.HasBackColor ? r.ReadOleColor() : null;
                VariousPropertyBits = PropMask.HasVariousPropertyBits ? r.ReadUInt32() : 0;
                var captionCcb = PropMask.HasCaption ? r.ReadCcb() : Tuple.Create(0, false);
                PicturePosition = PropMask.HasPicturePosition ? r.ReadPicturePosition() : PicturePosition.RightTop;
                MousePointer = PropMask.HasMousePointer ? r.ReadMousePointer() : MousePointer.Default;
                BorderColor = PropMask.HasBorderColor ? r.ReadOleColor() : null;
                BorderStyle = PropMask.HasBorderStyle ? r.ReadBorderStyle() : BorderStyle.None;
                SpecialEffect = PropMask.HasSpecialEffect ? r.ReadSpecialEffect2() : SpecialEffect.Flat;
                if (PropMask.HasPicture)
                {
                    r.Skip2Bytes();
                }

                Accelerator = PropMask.HasAccelerator ? r.ReadWChar() : "";
                if (PropMask.HasMouseIcon)
                {
                    r.Skip2Bytes();
                }

                // ExtraDataBlock
                Caption = r.ReadStringFromCcb(captionCcb);
                Size = PropMask.HasSize ? r.ReadCoords() : Tuple.Create(0, 0);

                r.AlignTo(4);
                if (cbLabel != r.BaseStream.Position - 4)
                {
                    throw new ApplicationException(string.Format(VBASyncResources.ErrorFrxStreamSizeMismatch,
                        "o", "cbLabel", r.BaseStream.Position - 4, cbLabel));
                }

                // StreamData
                Picture = PropMask.HasPicture ? r.ReadGuidAndPicture() : new byte[0];
                MouseIcon = PropMask.HasMouseIcon ? r.ReadGuidAndPicture() : new byte[0];
            }
        }

        public string Accelerator { get; }
        public OleColor BackColor { get; }
        public OleColor BorderColor { get; }
        public BorderStyle BorderStyle { get; }
        public string Caption { get; }
        public OleColor ForeColor { get; }
        public byte MajorVersion { get; }
        public byte MinorVersion { get; }
        public byte[] MouseIcon { get; }
        public MousePointer MousePointer { get; }
        public byte[] Picture { get; }
        public PicturePosition PicturePosition { get; }
        public LabelPropMask PropMask { get; }
        public Tuple<int, int> Size { get; }
        public SpecialEffect SpecialEffect { get; }
        public uint VariousPropertyBits { get; }

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
            return Equals((LabelControl)obj);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = Picture?.Length.GetHashCode() ?? 0;
                hashCode = (hashCode * 397) ^ (MouseIcon?.Length.GetHashCode() ?? 0);
                hashCode = (hashCode * 397) ^ MinorVersion.GetHashCode();
                hashCode = (hashCode * 397) ^ MajorVersion.GetHashCode();
                hashCode = (hashCode * 397) ^ (PropMask?.GetHashCode() ?? 0);
                hashCode = (hashCode * 397) ^ (ForeColor?.GetHashCode() ?? 0);
                hashCode = (hashCode * 397) ^ (BackColor?.GetHashCode() ?? 0);
                hashCode = (hashCode * 397) ^ (int)VariousPropertyBits;
                hashCode = (hashCode * 397) ^ (Caption?.GetHashCode() ?? 0);
                hashCode = (hashCode * 397) ^ (int)PicturePosition;
                hashCode = (hashCode * 397) ^ (int)MousePointer;
                hashCode = (hashCode * 397) ^ (BorderColor?.GetHashCode() ?? 0);
                hashCode = (hashCode * 397) ^ (int)BorderStyle;
                hashCode = (hashCode * 397) ^ (int)SpecialEffect;
                hashCode = (hashCode * 397) ^ (Accelerator?.GetHashCode() ?? 0);
                hashCode = (hashCode * 397) ^ (Size?.GetHashCode() ?? 0);
                return hashCode;
            }
        }

        protected bool Equals(LabelControl other)
        {
            return Picture.SequenceEqual(other.Picture) && MouseIcon.SequenceEqual(other.MouseIcon) && MinorVersion == other.MinorVersion
                   && MajorVersion == other.MajorVersion && Equals(PropMask, other.PropMask) && Equals(ForeColor, other.ForeColor)
                   && Equals(BackColor, other.BackColor) && VariousPropertyBits == other.VariousPropertyBits && string.Equals(Caption, other.Caption)
                   && PicturePosition == other.PicturePosition && MousePointer == other.MousePointer && Equals(BorderColor, other.BorderColor)
                   && BorderStyle == other.BorderStyle && SpecialEffect == other.SpecialEffect && string.Equals(Accelerator, other.Accelerator)
                   && Equals(Size, other.Size);
        }
    }

    internal class LabelPropMask
    {
        public LabelPropMask(uint i)
        {
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

        public bool HasAccelerator { get; }
        public bool HasBackColor { get; }
        public bool HasBorderColor { get; }
        public bool HasBorderStyle { get; }
        public bool HasCaption { get; }
        public bool HasForeColor { get; }
        public bool HasMouseIcon { get; }
        public bool HasMousePointer { get; }
        public bool HasPicture { get; }
        public bool HasPicturePosition { get; }
        public bool HasSize { get; }
        public bool HasSpecialEffect { get; }
        public bool HasVariousPropertyBits { get; }

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
            return Equals((LabelPropMask)obj);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = HasForeColor.GetHashCode();
                hashCode = (hashCode * 397) ^ HasBackColor.GetHashCode();
                hashCode = (hashCode * 397) ^ HasVariousPropertyBits.GetHashCode();
                hashCode = (hashCode * 397) ^ HasCaption.GetHashCode();
                hashCode = (hashCode * 397) ^ HasPicturePosition.GetHashCode();
                hashCode = (hashCode * 397) ^ HasSize.GetHashCode();
                hashCode = (hashCode * 397) ^ HasMousePointer.GetHashCode();
                hashCode = (hashCode * 397) ^ HasBorderColor.GetHashCode();
                hashCode = (hashCode * 397) ^ HasBorderStyle.GetHashCode();
                hashCode = (hashCode * 397) ^ HasSpecialEffect.GetHashCode();
                hashCode = (hashCode * 397) ^ HasPicture.GetHashCode();
                hashCode = (hashCode * 397) ^ HasAccelerator.GetHashCode();
                hashCode = (hashCode * 397) ^ HasMouseIcon.GetHashCode();
                return hashCode;
            }
        }

        protected bool Equals(LabelPropMask other)
        {
            return HasForeColor == other.HasForeColor && HasBackColor == other.HasBackColor && HasVariousPropertyBits == other.HasVariousPropertyBits
                   && HasCaption == other.HasCaption && HasPicturePosition == other.HasPicturePosition && HasSize == other.HasSize
                   && HasMousePointer == other.HasMousePointer && HasBorderColor == other.HasBorderColor && HasBorderStyle == other.HasBorderStyle
                   && HasSpecialEffect == other.HasSpecialEffect && HasPicture == other.HasPicture && HasAccelerator == other.HasAccelerator
                   && HasMouseIcon == other.HasMouseIcon;
        }
    }
}
