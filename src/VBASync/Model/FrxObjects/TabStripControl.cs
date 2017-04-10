using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using VBASync.Localization;

namespace VBASync.Model.FrxObjects
{
    internal class TabStripControl
    {
        public TabStripControl(byte[] b)
        {
            using (var st = new MemoryStream(b))
            using (var r = new FrxReader(st))
            {
                MinorVersion = r.ReadByte();
                MajorVersion = r.ReadByte();

                var cbTabStrip = r.ReadUInt16();
                PropMask = new TabStripPropMask(r.ReadUInt32());

                // DataBlock
                ListIndex = PropMask.HasListIndex ? r.ReadInt32() : 0;
                BackColor = PropMask.HasBackColor ? r.ReadOleColor() : null;
                ForeColor = PropMask.HasForeColor ? r.ReadOleColor() : null;
                var itemsSize = PropMask.HasItems ? r.ReadUInt32() : 0;
                MousePointer = PropMask.HasMousePointer ? r.ReadMousePointer() : MousePointer.Arrow;
                TabOrientation = PropMask.HasTabOrientation ? r.ReadUInt32() : 0;
                TabStyle = PropMask.HasTabStyle ? r.ReadUInt32() : 0;
                TabFixedWidth = PropMask.HasTabFixedWidth ? r.ReadUInt32() : 0;
                TabFixedHeight = PropMask.HasTabFixedHeight ? r.ReadUInt32() : 0;
                var tipStringsSize = PropMask.HasTipStrings ? r.ReadUInt32() : 0;
                var namesSize = PropMask.HasNames ? r.ReadUInt32() : 0;
                VariousPropertyBits = PropMask.HasVariousPropertyBits ? r.ReadUInt32() : 0;
                TabsAllocated = PropMask.HasTabsAllocated ? r.ReadUInt32() : 0;
                var tagsSize = PropMask.HasTags ? r.ReadUInt32() : 0;
                TabData = PropMask.HasTabData ? r.ReadUInt32() : 0;
                var acceleratorsSize = PropMask.HasAccelerator ? r.ReadUInt32() : 0;
                if (PropMask.HasMouseIcon)
                {
                    r.Skip2Bytes();
                }

                // ExtraDataBlock
                Size = PropMask.HasSize ? r.ReadCoords() : Tuple.Create(0, 0);
                Items = r.ReadArrayStrings(itemsSize);
                TipStrings = r.ReadArrayStrings(tipStringsSize);
                TabNames = r.ReadArrayStrings(namesSize);
                Tags = r.ReadArrayStrings(tagsSize);
                Accelerators = r.ReadArrayStrings(acceleratorsSize);

                r.AlignTo(4);
                if (cbTabStrip != r.BaseStream.Position - 4)
                {
                    throw new ApplicationException(string.Format(VBASyncResources.ErrorFrxStreamSizeMismatch,
                        "o", "cbTabStrip", r.BaseStream.Position - 4, cbTabStrip));
                }

                // StreamData
                MouseIcon = PropMask.HasMouseIcon ? r.ReadGuidAndPicture() : new byte[0];

                TextProps = r.ReadTextProps();

                Remainder = st.Position < st.Length ? r.Unaligned.ReadBytes((int)(st.Length - st.Position)) : new byte[0];
            }
        }

        public string[] Accelerators { get; }
        public uint AcceleratorsSize { get; }
        public OleColor BackColor { get; }
        public OleColor ForeColor { get; }
        public string[] Items { get; }
        public uint ItemsSize { get; }
        public int ListIndex { get; }
        public byte MajorVersion { get; }
        public byte MinorVersion { get; }
        public byte[] MouseIcon { get; }
        public MousePointer MousePointer { get; }
        public uint NamesSize { get; }
        public TabStripPropMask PropMask { get; }
        public byte[] Remainder { get; }
        public Tuple<int, int> Size { get; }
        public uint TabData { get; }
        public uint TabFixedHeight { get; }
        public uint TabFixedWidth { get; }
        public string[] TabNames { get; }
        public uint TabOrientation { get; }
        public uint TabsAllocated { get; }
        public uint TabStyle { get; }
        public string[] Tags { get; }
        public uint TagsSize { get; }
        public TextProps TextProps { get; }
        public string[] TipStrings { get; }
        public uint TipStringsSize { get; }
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
            return Equals((TabStripControl)obj);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = Accelerators.Length.GetHashCode();
                hashCode = (hashCode * 397) ^ (BackColor?.GetHashCode() ?? 0);
                hashCode = (hashCode * 397) ^ (ForeColor?.GetHashCode() ?? 0);
                hashCode = (hashCode * 397) ^ Items.Length.GetHashCode();
                hashCode = (hashCode * 397) ^ ListIndex.GetHashCode();
                hashCode = (hashCode * 397) ^ MajorVersion.GetHashCode();
                hashCode = (hashCode * 397) ^ MinorVersion.GetHashCode();
                hashCode = (hashCode * 397) ^ (MouseIcon?.Length.GetHashCode() ?? 0);
                hashCode = (hashCode * 397) ^ MousePointer.GetHashCode();
                hashCode = (hashCode * 397) ^ PropMask.GetHashCode();
                hashCode = (hashCode * 397) ^ Remainder.Length.GetHashCode();
                hashCode = (hashCode * 397) ^ Size.GetHashCode();
                hashCode = (hashCode * 397) ^ TabData.GetHashCode();
                hashCode = (hashCode * 397) ^ TabFixedHeight.GetHashCode();
                hashCode = (hashCode * 397) ^ TabFixedWidth.GetHashCode();
                hashCode = (hashCode * 397) ^ TabNames.Length.GetHashCode();
                hashCode = (hashCode * 397) ^ TabOrientation.GetHashCode();
                hashCode = (hashCode * 397) ^ TabsAllocated.GetHashCode();
                hashCode = (hashCode * 397) ^ TabStyle.GetHashCode();
                hashCode = (hashCode * 397) ^ Tags.Length.GetHashCode();
                hashCode = (hashCode * 397) ^ TextProps.GetHashCode();
                hashCode = (hashCode * 397) ^ TipStrings.Length.GetHashCode();
                hashCode = (hashCode * 397) ^ VariousPropertyBits.GetHashCode();
                return hashCode;
            }
        }

        protected bool Equals(TabStripControl other)
        {
            return Accelerators.SequenceEqual(other.Accelerators) && Equals(BackColor, other.BackColor) && Equals(ForeColor, other.ForeColor)
                && Items.SequenceEqual(other.Items) && ListIndex == other.ListIndex && MajorVersion == other.MajorVersion && MinorVersion == other.MinorVersion
                && MouseIcon.SequenceEqual(other.MouseIcon) && MousePointer == other.MousePointer && Equals(PropMask, other.PropMask)
                && Remainder.SequenceEqual(other.Remainder) && Equals(Size, other.Size) && TabData == other.TabData
                && TabFixedHeight == other.TabFixedHeight && TabFixedWidth == other.TabFixedWidth && TabNames.SequenceEqual(other.TabNames)
                && TabOrientation == other.TabOrientation && TabsAllocated == other.TabsAllocated && TabStyle == other.TabStyle
                && Tags.SequenceEqual(other.Tags) && Equals(TextProps, other.TextProps) && TipStrings.SequenceEqual(other.TipStrings)
                && VariousPropertyBits == other.VariousPropertyBits;
        }
    }

    internal class TabStripPropMask
    {
        public TabStripPropMask(uint i)
        {
            Func<int, bool> bit = j => (i & ((uint)1 << j)) != 0;
            HasListIndex = bit(0);
            HasBackColor = bit(1);
            HasForeColor = bit(2);
            HasSize = bit(4);
            HasItems = bit(5);
            HasMousePointer = bit(6);
            HasTabOrientation = bit(8);
            HasTabStyle = bit(9);
            HasMultiRow = bit(10);
            HasTabFixedWidth = bit(11);
            HasTabFixedHeight = bit(12);
            HasTooltips = bit(13);
            HasTipStrings = bit(15);
            HasNames = bit(17);
            HasVariousPropertyBits = bit(18);
            HasNewVersion = bit(19);
            HasTabsAllocated = bit(20);
            HasTags = bit(21);
            HasTabData = bit(22);
            HasAccelerator = bit(23);
            HasMouseIcon = bit(24);
        }

        public bool HasAccelerator { get; }
        public bool HasBackColor { get; }
        public bool HasForeColor { get; }
        public bool HasItems { get; }
        public bool HasListIndex { get; }
        public bool HasMouseIcon { get; }
        public bool HasMousePointer { get; }
        public bool HasMultiRow { get; }
        public bool HasNames { get; }
        public bool HasNewVersion { get; }
        public bool HasSize { get; }
        public bool HasTabData { get; }
        public bool HasTabFixedHeight { get; }
        public bool HasTabFixedWidth { get; }
        public bool HasTabOrientation { get; }
        public bool HasTabsAllocated { get; }
        public bool HasTabStyle { get; }
        public bool HasTags { get; }
        public bool HasTipStrings { get; }
        public bool HasTooltips { get; }
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
            return Equals((TabStripPropMask)obj);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = HasAccelerator.GetHashCode();
                hashCode = (hashCode * 397) ^ HasBackColor.GetHashCode();
                hashCode = (hashCode * 397) ^ HasForeColor.GetHashCode();
                hashCode = (hashCode * 397) ^ HasItems.GetHashCode();
                hashCode = (hashCode * 397) ^ HasListIndex.GetHashCode();
                hashCode = (hashCode * 397) ^ HasMouseIcon.GetHashCode();
                hashCode = (hashCode * 397) ^ HasMousePointer.GetHashCode();
                hashCode = (hashCode * 397) ^ HasMultiRow.GetHashCode();
                hashCode = (hashCode * 397) ^ HasNames.GetHashCode();
                hashCode = (hashCode * 397) ^ HasNewVersion.GetHashCode();
                hashCode = (hashCode * 397) ^ HasSize.GetHashCode();
                hashCode = (hashCode * 397) ^ HasTabData.GetHashCode();
                hashCode = (hashCode * 397) ^ HasTabFixedHeight.GetHashCode();
                hashCode = (hashCode * 397) ^ HasTabFixedWidth.GetHashCode();
                hashCode = (hashCode * 397) ^ HasTabOrientation.GetHashCode();
                hashCode = (hashCode * 397) ^ HasTabsAllocated.GetHashCode();
                hashCode = (hashCode * 397) ^ HasTabStyle.GetHashCode();
                hashCode = (hashCode * 397) ^ HasTags.GetHashCode();
                hashCode = (hashCode * 397) ^ HasTipStrings.GetHashCode();
                hashCode = (hashCode * 397) ^ HasTooltips.GetHashCode();
                hashCode = (hashCode * 397) ^ HasVariousPropertyBits.GetHashCode();
                return hashCode;
            }
        }

        protected bool Equals(TabStripPropMask other)
        {
            return HasAccelerator == other.HasAccelerator && HasBackColor == other.HasBackColor && HasForeColor == other.HasForeColor
                && HasItems == other.HasItems && HasListIndex == other.HasListIndex && HasMouseIcon == other.HasMouseIcon
                && HasMousePointer == other.HasMousePointer && HasMultiRow == other.HasMultiRow && HasNames == other.HasNames
                && HasNewVersion == other.HasNewVersion && HasSize == other.HasSize && HasTabData == other.HasTabData
                && HasTabFixedHeight == other.HasTabFixedHeight && HasTabFixedWidth == other.HasTabFixedWidth
                && HasTabOrientation == other.HasTabOrientation && HasTabsAllocated == other.HasTabsAllocated && HasTags == other.HasTags
                && HasTipStrings == other.HasTipStrings && HasTooltips == other.HasTooltips
                && HasVariousPropertyBits == other.HasVariousPropertyBits;
        }
    }
}
