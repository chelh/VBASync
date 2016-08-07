using System;
using System.IO;
using System.Linq;
using static VbaSync.FrxObjects.StreamDataHelpers;

namespace VbaSync.FrxObjects {
    class MorphDataControl : FrxCommon {
        public byte MinorVersion { get; }
        public byte MajorVersion { get; }
        public MorphDataPropMask PropMask { get; }
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
        public uint MaxLength { get; }
        public BorderStyle BorderStyle { get; }
        public FormScrollBarFlags ScrollBars { get; }
        public byte DisplayStyle { get; }
        public string PasswordChar { get; }
        public uint ListWidth { get; }
        public ushort BoundColumn { get; }
        public short TextColumn { get; }
        public short ColumnCount { get; }
        public ushort ListRows { get; }
        public ushort ColumnInfoCount { get; }
        public byte MatchEntry { get; }
        public byte ListStyle { get; }
        public byte ShowDropButtonWhen { get; }
        public byte DropButtonStyle { get; }
        public byte MultiSelect { get; }
        public OleColor BorderColor { get; }
        public SpecialEffect SpecialEffect { get; }
        public string Value { get; }
        public string GroupName { get; }
        public byte[] Remainder { get; } = new byte[0];

        public MorphDataControl(byte[] b) {
            using (var st = new MemoryStream(b))
            using (var r = new BinaryReader(st)) {
                MinorVersion = r.ReadByte();
                MajorVersion = r.ReadByte();

                var cbMorphData = r.ReadUInt16();
                PropMask = new MorphDataPropMask(r.ReadUInt64());

                BeginDataBlock();
                VariousPropertyBits = ReadAlignedUInt32If(PropMask.HasVariousPropertyBits, r);
                BackColor = ReadAlignedOleColorIf(PropMask.HasBackColor, r);
                ForeColor = ReadAlignedOleColorIf(PropMask.HasForeColor, r);
                MaxLength = ReadAlignedUInt32If(PropMask.HasMaxLength, r);
                BorderStyle = (BorderStyle)ReadByteIf(PropMask.HasBorderStyle, r);
                ScrollBars = new FormScrollBarFlags(ReadByteIf(PropMask.HasScrollBars, r));
                DisplayStyle = ReadByteIf(PropMask.HasDisplayStyle, r);
                MousePointer = (MousePointer)ReadByteIf(PropMask.HasMousePointer, r);
                PasswordChar = ReadAlignedWCharIf(PropMask.HasPasswordChar, r);
                ListWidth = ReadAlignedUInt32If(PropMask.HasListWidth, r);
                BoundColumn = ReadAlignedUInt16If(PropMask.HasBoundColumn, r);
                TextColumn = ReadAlignedInt16If(PropMask.HasTextColumn, r);
                ColumnCount = ReadAlignedInt16If(PropMask.HasColumnCount, r);
                ListRows = ReadAlignedUInt16If(PropMask.HasListRows, r);
                ColumnInfoCount = ReadAlignedUInt16If(PropMask.HasColumnInfoCount, r);
                MatchEntry = ReadByteIf(PropMask.HasMatchEntry, r);
                ListStyle = ReadByteIf(PropMask.HasListStyle, r);
                ShowDropButtonWhen = ReadByteIf(PropMask.HasShowDropDownWhen, r);
                DropButtonStyle = ReadByteIf(PropMask.HasDropButtonStyle, r);
                MultiSelect = ReadByteIf(PropMask.HasMultiSelect, r);
                var valueCcb = ReadAlignedCcbIf(PropMask.HasValue, r);
                var captionCcb = ReadAlignedCcbIf(PropMask.HasCaption, r);
                PicturePosition = (PicturePosition)ReadAlignedUInt32If(PropMask.HasPicturePosition, r);
                BorderColor = ReadAlignedOleColorIf(PropMask.HasBorderColor, r);
                SpecialEffect = (SpecialEffect)ReadAlignedUInt32If(PropMask.HasSpecialEffect, r);
                Ignore2AlignedBytesIf(PropMask.HasMouseIcon, r);
                Ignore2AlignedBytesIf(PropMask.HasPicture, r);
                Accelerator = ReadAlignedWCharIf(PropMask.HasAccelerator, r);
                var groupNameCcb = ReadAlignedCcbIf(PropMask.HasGroupName, r);
                EndDataBlock(r);

                BeginExtraDataBlock();
                Size = ReadAlignedCoordsIf(PropMask.HasSize, r);
                Value = ReadStringFromCcb(valueCcb, r);
                Caption = ReadStringFromCcb(captionCcb, r);
                GroupName = ReadStringFromCcb(groupNameCcb, r);
                EndExtraDataBlock(r);

                if (cbMorphData != 8 + DataBlockBytes + ExtraDataBlockBytes)
                    throw new ApplicationException("Error reading 'o' stream in .frx data: expected cbMorphData size "
                                                   + $"{8 + DataBlockBytes + ExtraDataBlockBytes}, but actual size was {cbMorphData}.");

                // StreamData
                if (PropMask.HasMouseIcon) {
                    MouseIcon = ReadGuidAndPicture(r);
                }
                if (PropMask.HasPicture) {
                    Picture = ReadGuidAndPicture(r);
                }

                TextProps = ReadTextProps(r);

                if (st.Position < st.Length) {
                    Remainder = r.ReadBytes((int)(st.Length - st.Position));
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
            return Equals((MorphDataControl)obj);
        }

        protected bool Equals(MorphDataControl other) {
            return MinorVersion == other.MinorVersion && MajorVersion == other.MajorVersion && Equals(PropMask, other.PropMask) &&
                   Equals(ForeColor, other.ForeColor) && Equals(BackColor, other.BackColor) && VariousPropertyBits == other.VariousPropertyBits &&
                   string.Equals(Caption, other.Caption) && PicturePosition == other.PicturePosition && MousePointer == other.MousePointer &&
                   string.Equals(Accelerator, other.Accelerator) && Equals(Size, other.Size) && Picture.SequenceEqual(other.Picture) &&
                   MouseIcon.SequenceEqual(other.MouseIcon) && Equals(TextProps, other.TextProps) && MaxLength == other.MaxLength && BorderStyle == other.BorderStyle &&
                   Equals(ScrollBars, other.ScrollBars) && DisplayStyle == other.DisplayStyle && string.Equals(PasswordChar, other.PasswordChar) &&
                   ListWidth == other.ListWidth && BoundColumn == other.BoundColumn && TextColumn == other.TextColumn && ColumnCount == other.ColumnCount &&
                   ListRows == other.ListRows && ColumnInfoCount == other.ColumnInfoCount && MatchEntry == other.MatchEntry && ListStyle == other.ListStyle &&
                   ShowDropButtonWhen == other.ShowDropButtonWhen && DropButtonStyle == other.DropButtonStyle && MultiSelect == other.MultiSelect &&
                   Equals(BorderColor, other.BorderColor) && SpecialEffect == other.SpecialEffect && string.Equals(Value, other.Value) &&
                   string.Equals(GroupName, other.GroupName) && Remainder.SequenceEqual(other.Remainder);
        }

        public override int GetHashCode() {
            unchecked {
                var hashCode = MinorVersion.GetHashCode();
                hashCode = (hashCode*397) ^ MajorVersion.GetHashCode();
                hashCode = (hashCode*397) ^ (PropMask?.GetHashCode() ?? 0);
                hashCode = (hashCode*397) ^ (ForeColor?.GetHashCode() ?? 0);
                hashCode = (hashCode*397) ^ (BackColor?.GetHashCode() ?? 0);
                hashCode = (hashCode*397) ^ (int)VariousPropertyBits;
                hashCode = (hashCode*397) ^ (Caption?.GetHashCode() ?? 0);
                hashCode = (hashCode*397) ^ (int)PicturePosition;
                hashCode = (hashCode*397) ^ (int)MousePointer;
                hashCode = (hashCode*397) ^ (Accelerator?.GetHashCode() ?? 0);
                hashCode = (hashCode*397) ^ (Size?.GetHashCode() ?? 0);
                hashCode = (hashCode*397) ^ (Picture?.Length.GetHashCode() ?? 0);
                hashCode = (hashCode*397) ^ (MouseIcon?.Length.GetHashCode() ?? 0);
                hashCode = (hashCode*397) ^ (TextProps?.GetHashCode() ?? 0);
                hashCode = (hashCode*397) ^ (int)MaxLength;
                hashCode = (hashCode*397) ^ (int)BorderStyle;
                hashCode = (hashCode*397) ^ (ScrollBars?.GetHashCode() ?? 0);
                hashCode = (hashCode*397) ^ DisplayStyle.GetHashCode();
                hashCode = (hashCode*397) ^ (PasswordChar?.GetHashCode() ?? 0);
                hashCode = (hashCode*397) ^ (int)ListWidth;
                hashCode = (hashCode*397) ^ BoundColumn.GetHashCode();
                hashCode = (hashCode*397) ^ TextColumn.GetHashCode();
                hashCode = (hashCode*397) ^ ColumnCount.GetHashCode();
                hashCode = (hashCode*397) ^ ListRows.GetHashCode();
                hashCode = (hashCode*397) ^ ColumnInfoCount.GetHashCode();
                hashCode = (hashCode*397) ^ MatchEntry.GetHashCode();
                hashCode = (hashCode*397) ^ ListStyle.GetHashCode();
                hashCode = (hashCode*397) ^ ShowDropButtonWhen.GetHashCode();
                hashCode = (hashCode*397) ^ DropButtonStyle.GetHashCode();
                hashCode = (hashCode*397) ^ MultiSelect.GetHashCode();
                hashCode = (hashCode*397) ^ (BorderColor?.GetHashCode() ?? 0);
                hashCode = (hashCode*397) ^ (int)SpecialEffect;
                hashCode = (hashCode*397) ^ (Value?.GetHashCode() ?? 0);
                hashCode = (hashCode*397) ^ (GroupName?.GetHashCode() ?? 0);
                hashCode = (hashCode*397) ^ (Remainder?.Length.GetHashCode() ?? 0);
                return hashCode;
            }
        }
    }

    class MorphDataPropMask {
        public bool HasVariousPropertyBits { get; }
        public bool HasBackColor { get; }
        public bool HasForeColor { get; }
        public bool HasMaxLength { get; }
        public bool HasBorderStyle { get; }
        public bool HasScrollBars { get; }
        public bool HasDisplayStyle { get; }
        public bool HasMousePointer { get; }
        public bool HasSize { get; }
        public bool HasPasswordChar { get; }
        public bool HasListWidth { get; }
        public bool HasBoundColumn { get; }
        public bool HasTextColumn { get; }
        public bool HasColumnCount { get; }
        public bool HasListRows { get; }
        public bool HasColumnInfoCount { get; }
        public bool HasMatchEntry { get; }
        public bool HasListStyle { get; }
        public bool HasShowDropDownWhen { get; }
        public bool HasDropButtonStyle { get; }
        public bool HasMultiSelect { get; }
        public bool HasValue { get; }
        public bool HasCaption { get; }
        public bool HasPicturePosition { get; }
        public bool HasBorderColor { get; }
        public bool HasSpecialEffect { get; }
        public bool HasMouseIcon { get; }
        public bool HasPicture { get; }
        public bool HasAccelerator { get; }
        public bool HasGroupName { get; }

        public MorphDataPropMask(ulong i) {
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
            return Equals((MorphDataPropMask)obj);
        }

        protected bool Equals(MorphDataPropMask other) {
            return HasVariousPropertyBits == other.HasVariousPropertyBits && HasBackColor == other.HasBackColor && HasForeColor == other.HasForeColor &&
                   HasMaxLength == other.HasMaxLength && HasBorderStyle == other.HasBorderStyle && HasScrollBars == other.HasScrollBars &&
                   HasDisplayStyle == other.HasDisplayStyle && HasMousePointer == other.HasMousePointer && HasSize == other.HasSize &&
                   HasPasswordChar == other.HasPasswordChar && HasListWidth == other.HasListWidth && HasBoundColumn == other.HasBoundColumn &&
                   HasTextColumn == other.HasTextColumn && HasColumnCount == other.HasColumnCount && HasListRows == other.HasListRows &&
                   HasColumnInfoCount == other.HasColumnInfoCount && HasMatchEntry == other.HasMatchEntry && HasListStyle == other.HasListStyle &&
                   HasShowDropDownWhen == other.HasShowDropDownWhen && HasDropButtonStyle == other.HasDropButtonStyle && HasMultiSelect == other.HasMultiSelect &&
                   HasValue == other.HasValue && HasCaption == other.HasCaption && HasPicturePosition == other.HasPicturePosition &&
                   HasBorderColor == other.HasBorderColor && HasSpecialEffect == other.HasSpecialEffect && HasMouseIcon == other.HasMouseIcon &&
                   HasPicture == other.HasPicture && HasAccelerator == other.HasAccelerator && HasGroupName == other.HasGroupName;
        }

        public override int GetHashCode() {
            unchecked {
                var hashCode = HasVariousPropertyBits.GetHashCode();
                hashCode = (hashCode*397) ^ HasBackColor.GetHashCode();
                hashCode = (hashCode*397) ^ HasForeColor.GetHashCode();
                hashCode = (hashCode*397) ^ HasMaxLength.GetHashCode();
                hashCode = (hashCode*397) ^ HasBorderStyle.GetHashCode();
                hashCode = (hashCode*397) ^ HasScrollBars.GetHashCode();
                hashCode = (hashCode*397) ^ HasDisplayStyle.GetHashCode();
                hashCode = (hashCode*397) ^ HasMousePointer.GetHashCode();
                hashCode = (hashCode*397) ^ HasSize.GetHashCode();
                hashCode = (hashCode*397) ^ HasPasswordChar.GetHashCode();
                hashCode = (hashCode*397) ^ HasListWidth.GetHashCode();
                hashCode = (hashCode*397) ^ HasBoundColumn.GetHashCode();
                hashCode = (hashCode*397) ^ HasTextColumn.GetHashCode();
                hashCode = (hashCode*397) ^ HasColumnCount.GetHashCode();
                hashCode = (hashCode*397) ^ HasListRows.GetHashCode();
                hashCode = (hashCode*397) ^ HasColumnInfoCount.GetHashCode();
                hashCode = (hashCode*397) ^ HasMatchEntry.GetHashCode();
                hashCode = (hashCode*397) ^ HasListStyle.GetHashCode();
                hashCode = (hashCode*397) ^ HasShowDropDownWhen.GetHashCode();
                hashCode = (hashCode*397) ^ HasDropButtonStyle.GetHashCode();
                hashCode = (hashCode*397) ^ HasMultiSelect.GetHashCode();
                hashCode = (hashCode*397) ^ HasValue.GetHashCode();
                hashCode = (hashCode*397) ^ HasCaption.GetHashCode();
                hashCode = (hashCode*397) ^ HasPicturePosition.GetHashCode();
                hashCode = (hashCode*397) ^ HasBorderColor.GetHashCode();
                hashCode = (hashCode*397) ^ HasSpecialEffect.GetHashCode();
                hashCode = (hashCode*397) ^ HasMouseIcon.GetHashCode();
                hashCode = (hashCode*397) ^ HasPicture.GetHashCode();
                hashCode = (hashCode*397) ^ HasAccelerator.GetHashCode();
                hashCode = (hashCode*397) ^ HasGroupName.GetHashCode();
                return hashCode;
            }
        }
    }
}
