using System;
using System.IO;
using System.Linq;
using VBASync.Localization;

namespace VBASync.Model.FrxObjects
{
    internal class MorphDataControl
    {
        public MorphDataControl(byte[] b)
        {
            using (var st = new MemoryStream(b))
            using (var r = new FrxReader(st))
            {
                MinorVersion = r.ReadByte();
                MajorVersion = r.ReadByte();

                var cbMorphData = r.ReadUInt16();
                PropMask = new MorphDataPropMask(r.ReadUInt64());

                // DataBlock
                VariousPropertyBits = PropMask.HasVariousPropertyBits ? r.ReadUInt32() : 0;
                BackColor = PropMask.HasBackColor ? r.ReadOleColor() : null;
                ForeColor = PropMask.HasForeColor ? r.ReadOleColor() : null;
                MaxLength = PropMask.HasMaxLength ? r.ReadUInt32() : 0;
                BorderStyle = PropMask.HasBorderStyle ? r.ReadBorderStyle() : BorderStyle.None;
                ScrollBars = new FormScrollBarFlags(PropMask.HasScrollBars ? r.ReadByte() : (byte)0);
                DisplayStyle = PropMask.HasDisplayStyle ? r.ReadByte() : (byte)0;
                MousePointer = PropMask.HasMousePointer ? r.ReadMousePointer() : MousePointer.Default;
                PasswordChar = PropMask.HasPasswordChar ? r.ReadWChar() : "";
                ListWidth = PropMask.HasListWidth ? r.ReadUInt32() : 0;
                BoundColumn = PropMask.HasBoundColumn ? r.ReadUInt16() : (ushort)0;
                TextColumn = PropMask.HasTextColumn ? r.ReadInt16() : (short)0;
                ColumnCount = PropMask.HasColumnCount ? r.ReadInt16() : (short)0;
                ListRows = PropMask.HasListRows ? r.ReadUInt16() : (ushort)0;
                ColumnInfoCount = PropMask.HasColumnInfoCount ? r.ReadUInt16() : (ushort)0;
                MatchEntry = PropMask.HasMatchEntry ? r.ReadByte() : (byte)0;
                ListStyle = PropMask.HasListStyle ? r.ReadByte() : (byte)0;
                ShowDropButtonWhen = PropMask.HasShowDropDownWhen ? r.ReadByte() : (byte)0;
                DropButtonStyle = PropMask.HasDropButtonStyle ? r.ReadByte() : (byte)0;
                MultiSelect = PropMask.HasMultiSelect ? r.ReadByte() : (byte)0;
                var valueCcb = PropMask.HasValue ? r.ReadCcb() : Tuple.Create(0, false);
                var captionCcb = PropMask.HasCaption ? r.ReadCcb() : Tuple.Create(0, false);
                PicturePosition = PropMask.HasPicturePosition ? r.ReadPicturePosition() : PicturePosition.RightTop;
                BorderColor = PropMask.HasBorderColor ? r.ReadOleColor() : null;
                SpecialEffect = PropMask.HasSpecialEffect ? r.ReadSpecialEffect4() : SpecialEffect.Flat;
                if (PropMask.HasMouseIcon)
                {
                    r.Skip2Bytes();
                }

                if (PropMask.HasPicture)
                {
                    r.Skip2Bytes();
                }

                Accelerator = PropMask.HasAccelerator ? r.ReadWChar() : "";
                var groupNameCcb = PropMask.HasGroupName ? r.ReadCcb() : Tuple.Create(0, false);

                // ExtraDataBlock
                Size = PropMask.HasSize ? r.ReadCoords() : Tuple.Create(0, 0);
                Value = r.ReadStringFromCcb(valueCcb);
                Caption = r.ReadStringFromCcb(captionCcb);
                GroupName = r.ReadStringFromCcb(groupNameCcb);

                r.AlignTo(4);
                if (cbMorphData != r.BaseStream.Position - 4)
                {
                    throw new ApplicationException(string.Format(VBASyncResources.ErrorFrxStreamSizeMismatch,
                        "o", "cbMorphData", r.BaseStream.Position - 4, cbMorphData));
                }

                // StreamData
                MouseIcon = PropMask.HasMouseIcon ? r.ReadGuidAndPicture() : new byte[0];
                Picture = PropMask.HasPicture ? r.ReadGuidAndPicture() : new byte[0];

                TextProps = r.ReadTextProps();

                Remainder = st.Position < st.Length ? r.Unaligned.ReadBytes((int)(st.Length - st.Position)) : new byte[0];
            }
        }

        public string Accelerator { get; }
        public OleColor BackColor { get; }
        public OleColor BorderColor { get; }
        public BorderStyle BorderStyle { get; }
        public ushort BoundColumn { get; }
        public string Caption { get; }
        public short ColumnCount { get; }
        public ushort ColumnInfoCount { get; }
        public byte DisplayStyle { get; }
        public byte DropButtonStyle { get; }
        public OleColor ForeColor { get; }
        public string GroupName { get; }
        public ushort ListRows { get; }
        public byte ListStyle { get; }
        public uint ListWidth { get; }
        public byte MajorVersion { get; }
        public byte MatchEntry { get; }
        public uint MaxLength { get; }
        public byte MinorVersion { get; }
        public byte[] MouseIcon { get; }
        public MousePointer MousePointer { get; }
        public byte MultiSelect { get; }
        public string PasswordChar { get; }
        public byte[] Picture { get; }
        public PicturePosition PicturePosition { get; }
        public MorphDataPropMask PropMask { get; }
        public byte[] Remainder { get; }
        public FormScrollBarFlags ScrollBars { get; }
        public byte ShowDropButtonWhen { get; }
        public Tuple<int, int> Size { get; }
        public SpecialEffect SpecialEffect { get; }
        public short TextColumn { get; }
        public TextProps TextProps { get; }
        public string Value { get; }
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
            return Equals((MorphDataControl)obj);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = MinorVersion.GetHashCode();
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
                hashCode = (hashCode * 397) ^ (Picture?.Length.GetHashCode() ?? 0);
                hashCode = (hashCode * 397) ^ (MouseIcon?.Length.GetHashCode() ?? 0);
                hashCode = (hashCode * 397) ^ (TextProps?.GetHashCode() ?? 0);
                hashCode = (hashCode * 397) ^ (int)MaxLength;
                hashCode = (hashCode * 397) ^ (int)BorderStyle;
                hashCode = (hashCode * 397) ^ (ScrollBars?.GetHashCode() ?? 0);
                hashCode = (hashCode * 397) ^ DisplayStyle.GetHashCode();
                hashCode = (hashCode * 397) ^ (PasswordChar?.GetHashCode() ?? 0);
                hashCode = (hashCode * 397) ^ (int)ListWidth;
                hashCode = (hashCode * 397) ^ BoundColumn.GetHashCode();
                hashCode = (hashCode * 397) ^ TextColumn.GetHashCode();
                hashCode = (hashCode * 397) ^ ColumnCount.GetHashCode();
                hashCode = (hashCode * 397) ^ ListRows.GetHashCode();
                hashCode = (hashCode * 397) ^ ColumnInfoCount.GetHashCode();
                hashCode = (hashCode * 397) ^ MatchEntry.GetHashCode();
                hashCode = (hashCode * 397) ^ ListStyle.GetHashCode();
                hashCode = (hashCode * 397) ^ ShowDropButtonWhen.GetHashCode();
                hashCode = (hashCode * 397) ^ DropButtonStyle.GetHashCode();
                hashCode = (hashCode * 397) ^ MultiSelect.GetHashCode();
                hashCode = (hashCode * 397) ^ (BorderColor?.GetHashCode() ?? 0);
                hashCode = (hashCode * 397) ^ (int)SpecialEffect;
                hashCode = (hashCode * 397) ^ (Value?.GetHashCode() ?? 0);
                hashCode = (hashCode * 397) ^ (GroupName?.GetHashCode() ?? 0);
                hashCode = (hashCode * 397) ^ (Remainder?.Length.GetHashCode() ?? 0);
                return hashCode;
            }
        }

        protected bool Equals(MorphDataControl other)
        {
            return MinorVersion == other.MinorVersion && MajorVersion == other.MajorVersion && Equals(PropMask, other.PropMask)
                   && Equals(ForeColor, other.ForeColor) && Equals(BackColor, other.BackColor) && VariousPropertyBits == other.VariousPropertyBits
                   && string.Equals(Caption, other.Caption) && PicturePosition == other.PicturePosition && MousePointer == other.MousePointer
                   && string.Equals(Accelerator, other.Accelerator) && Equals(Size, other.Size) && Picture.SequenceEqual(other.Picture)
                   && MouseIcon.SequenceEqual(other.MouseIcon) && Equals(TextProps, other.TextProps) && MaxLength == other.MaxLength && BorderStyle == other.BorderStyle
                   && Equals(ScrollBars, other.ScrollBars) && DisplayStyle == other.DisplayStyle && string.Equals(PasswordChar, other.PasswordChar)
                   && ListWidth == other.ListWidth && BoundColumn == other.BoundColumn && TextColumn == other.TextColumn && ColumnCount == other.ColumnCount
                   && ListRows == other.ListRows && ColumnInfoCount == other.ColumnInfoCount && MatchEntry == other.MatchEntry && ListStyle == other.ListStyle
                   && ShowDropButtonWhen == other.ShowDropButtonWhen && DropButtonStyle == other.DropButtonStyle && MultiSelect == other.MultiSelect
                   && Equals(BorderColor, other.BorderColor) && SpecialEffect == other.SpecialEffect && string.Equals(Value, other.Value)
                   && string.Equals(GroupName, other.GroupName) && Remainder.SequenceEqual(other.Remainder);
        }
    }

    internal class MorphDataPropMask
    {
        public MorphDataPropMask(ulong i)
        {
            Func<int, bool> bit = j => (i & ((ulong)1 << j)) != 0;
            HasVariousPropertyBits = bit(0);
            HasBackColor = bit(1);
            HasForeColor = bit(2);
            HasMaxLength = bit(3);
            HasBorderStyle = bit(4);
            HasScrollBars = bit(5);
            HasDisplayStyle = bit(6);
            HasMousePointer = bit(7);
            HasSize = bit(8);
            HasPasswordChar = bit(9);
            HasListWidth = bit(10);
            HasBoundColumn = bit(11);
            HasTextColumn = bit(12);
            HasColumnCount = bit(13);
            HasListRows = bit(14);
            HasColumnInfoCount = bit(15);
            HasMatchEntry = bit(16);
            HasListStyle = bit(17);
            HasShowDropDownWhen = bit(18);
            HasDropButtonStyle = bit(20);
            HasMultiSelect = bit(21);
            HasValue = bit(22);
            HasCaption = bit(23);
            HasPicturePosition = bit(24);
            HasBorderColor = bit(25);
            HasSpecialEffect = bit(26);
            HasMouseIcon = bit(27);
            HasPicture = bit(28);
            HasAccelerator = bit(29);
            HasGroupName = bit(32);
        }

        public bool HasAccelerator { get; }
        public bool HasBackColor { get; }
        public bool HasBorderColor { get; }
        public bool HasBorderStyle { get; }
        public bool HasBoundColumn { get; }
        public bool HasCaption { get; }
        public bool HasColumnCount { get; }
        public bool HasColumnInfoCount { get; }
        public bool HasDisplayStyle { get; }
        public bool HasDropButtonStyle { get; }
        public bool HasForeColor { get; }
        public bool HasGroupName { get; }
        public bool HasListRows { get; }
        public bool HasListStyle { get; }
        public bool HasListWidth { get; }
        public bool HasMatchEntry { get; }
        public bool HasMaxLength { get; }
        public bool HasMouseIcon { get; }
        public bool HasMousePointer { get; }
        public bool HasMultiSelect { get; }
        public bool HasPasswordChar { get; }
        public bool HasPicture { get; }
        public bool HasPicturePosition { get; }
        public bool HasScrollBars { get; }
        public bool HasShowDropDownWhen { get; }
        public bool HasSize { get; }
        public bool HasSpecialEffect { get; }
        public bool HasTextColumn { get; }
        public bool HasValue { get; }
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
            return Equals((MorphDataPropMask)obj);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = HasVariousPropertyBits.GetHashCode();
                hashCode = (hashCode * 397) ^ HasBackColor.GetHashCode();
                hashCode = (hashCode * 397) ^ HasForeColor.GetHashCode();
                hashCode = (hashCode * 397) ^ HasMaxLength.GetHashCode();
                hashCode = (hashCode * 397) ^ HasBorderStyle.GetHashCode();
                hashCode = (hashCode * 397) ^ HasScrollBars.GetHashCode();
                hashCode = (hashCode * 397) ^ HasDisplayStyle.GetHashCode();
                hashCode = (hashCode * 397) ^ HasMousePointer.GetHashCode();
                hashCode = (hashCode * 397) ^ HasSize.GetHashCode();
                hashCode = (hashCode * 397) ^ HasPasswordChar.GetHashCode();
                hashCode = (hashCode * 397) ^ HasListWidth.GetHashCode();
                hashCode = (hashCode * 397) ^ HasBoundColumn.GetHashCode();
                hashCode = (hashCode * 397) ^ HasTextColumn.GetHashCode();
                hashCode = (hashCode * 397) ^ HasColumnCount.GetHashCode();
                hashCode = (hashCode * 397) ^ HasListRows.GetHashCode();
                hashCode = (hashCode * 397) ^ HasColumnInfoCount.GetHashCode();
                hashCode = (hashCode * 397) ^ HasMatchEntry.GetHashCode();
                hashCode = (hashCode * 397) ^ HasListStyle.GetHashCode();
                hashCode = (hashCode * 397) ^ HasShowDropDownWhen.GetHashCode();
                hashCode = (hashCode * 397) ^ HasDropButtonStyle.GetHashCode();
                hashCode = (hashCode * 397) ^ HasMultiSelect.GetHashCode();
                hashCode = (hashCode * 397) ^ HasValue.GetHashCode();
                hashCode = (hashCode * 397) ^ HasCaption.GetHashCode();
                hashCode = (hashCode * 397) ^ HasPicturePosition.GetHashCode();
                hashCode = (hashCode * 397) ^ HasBorderColor.GetHashCode();
                hashCode = (hashCode * 397) ^ HasSpecialEffect.GetHashCode();
                hashCode = (hashCode * 397) ^ HasMouseIcon.GetHashCode();
                hashCode = (hashCode * 397) ^ HasPicture.GetHashCode();
                hashCode = (hashCode * 397) ^ HasAccelerator.GetHashCode();
                hashCode = (hashCode * 397) ^ HasGroupName.GetHashCode();
                return hashCode;
            }
        }

        protected bool Equals(MorphDataPropMask other)
        {
            return HasVariousPropertyBits == other.HasVariousPropertyBits && HasBackColor == other.HasBackColor && HasForeColor == other.HasForeColor
                   && HasMaxLength == other.HasMaxLength && HasBorderStyle == other.HasBorderStyle && HasScrollBars == other.HasScrollBars
                   && HasDisplayStyle == other.HasDisplayStyle && HasMousePointer == other.HasMousePointer && HasSize == other.HasSize
                   && HasPasswordChar == other.HasPasswordChar && HasListWidth == other.HasListWidth && HasBoundColumn == other.HasBoundColumn
                   && HasTextColumn == other.HasTextColumn && HasColumnCount == other.HasColumnCount && HasListRows == other.HasListRows
                   && HasColumnInfoCount == other.HasColumnInfoCount && HasMatchEntry == other.HasMatchEntry && HasListStyle == other.HasListStyle
                   && HasShowDropDownWhen == other.HasShowDropDownWhen && HasDropButtonStyle == other.HasDropButtonStyle && HasMultiSelect == other.HasMultiSelect
                   && HasValue == other.HasValue && HasCaption == other.HasCaption && HasPicturePosition == other.HasPicturePosition
                   && HasBorderColor == other.HasBorderColor && HasSpecialEffect == other.HasSpecialEffect && HasMouseIcon == other.HasMouseIcon
                   && HasPicture == other.HasPicture && HasAccelerator == other.HasAccelerator && HasGroupName == other.HasGroupName;
        }
    }
}
