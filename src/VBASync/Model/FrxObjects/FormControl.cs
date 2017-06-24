using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using VBASync.Localization;

namespace VBASync.Model.FrxObjects
{
    internal enum Cycle
    {
        AllForms = 0x00,
        CurrentForm = 0x02
    }

    internal enum FormScrollBars
    {
        None = 0x00,
        Horizontal = 0x01,
        Vertical = 0x02,
        Both = 0x03
    }

    internal class FormControl
    {
        public FormControl(byte[] b)
        {
            using (var st = new MemoryStream(b))
            using (var r = new FrxReader(st))
            {
                MinorVersion = r.ReadByte();
                MajorVersion = r.ReadByte();

                var cbForm = r.ReadUInt16();
                PropMask = new FormPropMask(r.ReadUInt32());

                // DataBlock
                BackColor = PropMask.HasBackColor ? r.ReadOleColor() : null;
                ForeColor = PropMask.HasForeColor ? r.ReadOleColor() : null;
                NextAvailableId = PropMask.HasNextAvailableId ? r.ReadUInt32() : 0;
                BooleanProperties = new FormFlags(PropMask.HasBooleanProperties ? r.ReadUInt32() : 0);
                BorderStyle = PropMask.HasBorderStyle ? r.ReadBorderStyle() : BorderStyle.None;
                MousePointer = PropMask.HasMousePointer ? r.ReadMousePointer() : MousePointer.Default;
                ScrollBars = new FormScrollBarFlags(PropMask.HasScrollBars ? r.ReadByte() : (byte)0);
                GroupCount = PropMask.HasGroupCount ? r.ReadInt32() : 0;
                // captionCcb is possibly here instead of where it's indicated in [MS-OFORMS]?
                if (PropMask.HasMouseIcon)
                {
                    r.Skip2Bytes();
                }

                Cycle = PropMask.HasCycle ? r.ReadCycle() : Cycle.AllForms;
                SpecialEffect = PropMask.HasSpecialEffect ? r.ReadSpecialEffect() : SpecialEffect.Flat;
                BorderColor = PropMask.HasBorderColor ? r.ReadOleColor() : null;
                var captionCcb = PropMask.HasCaption ? r.ReadCcb() : Tuple.Create(0, false);
                if (PropMask.HasFont)
                {
                    r.Skip2Bytes();
                }

                if (PropMask.HasPicture)
                {
                    r.Skip2Bytes();
                }

                Zoom = PropMask.HasZoom ? r.ReadUInt32() : 0;
                PictureAlignment = PropMask.HasPictureAlignment ? r.ReadPictureAlignment() : PictureAlignment.TopLeft;
                PictureSizeMode = PropMask.HasPictureSizeMode ? r.ReadPictureSizeMode() : PictureSizeMode.Clip;
                ShapeCookie = PropMask.HasShapeCookie ? r.ReadUInt32() : 0;
                DrawBuffer = PropMask.HasDrawBuffer ? r.ReadUInt32() : 0;

                // ExtraDataBlock
                DisplayedSize = PropMask.HasDisplayedSize ? r.ReadCoords() : Tuple.Create(0, 0);
                LogicalSize = PropMask.HasLogicalSize ? r.ReadCoords() : Tuple.Create(0, 0);
                ScrollPosition = PropMask.HasScrollPosition ? r.ReadCoords() : Tuple.Create(0, 0);
                Caption = r.ReadStringFromCcb(captionCcb);

                r.AlignTo(4);
                if (cbForm != r.BaseStream.Position - 4)
                {
                    throw new ApplicationException(string.Format(VBASyncResources.ErrorFrxStreamSizeMismatch,
                        "f", "cbForm", r.BaseStream.Position - 4, cbForm));
                }

                // StreamData
                MouseIcon = PropMask.HasMouseIcon ? r.ReadGuidAndPicture() : new byte[0];
                if (PropMask.HasFont)
                {
                    FontIsStdFont = r.GetFontIsStdFont();
                    if (FontIsStdFont)
                    {
                        FontStdFont = r.ReadStdFont();
                    }
                    else
                    {
                        FontTextProps = r.ReadTextProps();
                    }
                }
                Picture = PropMask.HasPicture ? r.ReadGuidAndPicture() : new byte[0];

                // FormSiteData
                SiteClassInfos = new List<byte[]>();
                ushort siteClassInfoCount = 0;
                if (!PropMask.HasBooleanProperties || BooleanProperties.ClassTablePersisted)
                {
                    siteClassInfoCount = r.Unaligned.ReadUInt16();
                }
                for (var i = 0; i < siteClassInfoCount; i++)
                {
                    st.Seek(2, SeekOrigin.Current); // skip Version
                    SiteClassInfos.Add(r.Unaligned.ReadBytes(r.Unaligned.ReadUInt16()));
                }
                var siteCount = r.Unaligned.ReadUInt32();
                var cbSites = r.Unaligned.ReadUInt32();
                var sitesStartPos = r.BaseStream.Position;
                var depths = new byte[siteCount];
                var types = new byte[siteCount];
                var siteDepthsLeft = siteCount;
                while (siteDepthsLeft > 0)
                {
                    var thisDepth = r.Unaligned.ReadByte();
                    var thisType = r.Unaligned.ReadByte();
                    var thisCount = (byte)1;
                    if ((thisType & 0x80) == 0x80)
                    {
                        thisCount = (byte)(thisType ^ 0x80);
                        thisType = r.Unaligned.ReadByte();
                    }
                    for (var i = 0; i < thisCount; i++)
                    {
                        var siteIdx = siteCount - siteDepthsLeft;
                        depths[siteIdx] = thisDepth;
                        types[siteIdx] = thisType;
                        siteDepthsLeft--;
                    }
                }
                var rem = (r.BaseStream.Position - sitesStartPos) % 4;
                if (rem != 0)
                {
                    r.BaseStream.Seek(4 - rem, SeekOrigin.Current); // add ArrayPadding
                }

                Sites = new OleSiteConcreteControl[siteCount];
                for (var i = 0; i < siteCount; i++)
                {
                    r.BaseStream.Seek(2, SeekOrigin.Current); // ignore Version
                    var cbSite = r.Unaligned.ReadUInt16();
                    Sites[i] = new OleSiteConcreteControl(r.Unaligned.ReadBytes(cbSite));
                }
                if (cbSites != r.BaseStream.Position - sitesStartPos)
                {
                    throw new ApplicationException(string.Format(VBASyncResources.ErrorFrxStreamSizeMismatch,
                        "f", "cbSites", r.BaseStream.Position - sitesStartPos, cbSites));
                }

                Remainder = st.Position < st.Length ? r.Unaligned.ReadBytes((int)(st.Length - st.Position)) : new byte[0];
            }
        }

        public OleColor BackColor { get; }
        public FormFlags BooleanProperties { get; }
        public OleColor BorderColor { get; }
        public BorderStyle BorderStyle { get; }
        public string Caption { get; }
        public Cycle Cycle { get; }
        public Tuple<int, int> DisplayedSize { get; }
        public uint DrawBuffer { get; }
        public bool FontIsStdFont { get; }
        public Tuple<short, byte, short, uint, string> FontStdFont { get; }
        public TextProps FontTextProps { get; }
        public OleColor ForeColor { get; }
        public int GroupCount { get; }
        public Tuple<int, int> LogicalSize { get; }
        public byte MajorVersion { get; }
        public byte MinorVersion { get; }
        public byte[] MouseIcon { get; }
        public MousePointer MousePointer { get; }
        public uint NextAvailableId { get; }
        public byte[] Picture { get; }
        public PictureAlignment PictureAlignment { get; }
        public PictureSizeMode PictureSizeMode { get; }
        public FormPropMask PropMask { get; }
        public byte[] Remainder { get; }
        public FormScrollBarFlags ScrollBars { get; }
        public Tuple<int, int> ScrollPosition { get; }
        public uint ShapeCookie { get; }
        public List<byte[]> SiteClassInfos { get; }
        public OleSiteConcreteControl[] Sites { get; }
        public SpecialEffect SpecialEffect { get; }
        public uint Zoom { get; }

        public override bool Equals(object obj)
        {
            var other = obj as FormControl;
            if (other == null || !(MinorVersion == other.MinorVersion && MajorVersion == other.MajorVersion && Equals(PropMask, other.PropMask)
                  && Equals(BackColor, other.BackColor) && Equals(ForeColor, other.ForeColor) && NextAvailableId == other.NextAvailableId
                  && Equals(BooleanProperties, other.BooleanProperties) && BorderStyle == other.BorderStyle && MousePointer == other.MousePointer
                  && Equals(ScrollBars, other.ScrollBars) && GroupCount == other.GroupCount && Cycle == other.Cycle && SpecialEffect == other.SpecialEffect
                  && Equals(BorderColor, other.BorderColor) && Zoom == other.Zoom && PictureAlignment == other.PictureAlignment
                  && PictureSizeMode == other.PictureSizeMode && ShapeCookie == other.ShapeCookie && DrawBuffer == other.DrawBuffer
                  && Equals(DisplayedSize, other.DisplayedSize) && Equals(LogicalSize, other.LogicalSize) && Equals(ScrollPosition, other.ScrollPosition)
                  && string.Equals(Caption, other.Caption) && MouseIcon.SequenceEqual(other.MouseIcon) && FontIsStdFont == other.FontIsStdFont
                  && Picture.SequenceEqual(other.Picture) && Equals(FontTextProps, other.FontTextProps) && Equals(FontStdFont, other.FontStdFont)
                  && Sites.OrderBy(s => s.Id).SequenceEqual(other.Sites.OrderBy(s => s.Id)) && Remainder.SequenceEqual(other.Remainder)))
            {
                return false;
            }

            if (SiteClassInfos.Count != other.SiteClassInfos.Count)
            {
                return false;
            }

            return !SiteClassInfos.Where((t, i) => !t.SequenceEqual(other.SiteClassInfos[i])).Any();
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = MinorVersion.GetHashCode();
                hashCode = (hashCode * 397) ^ MajorVersion.GetHashCode();
                hashCode = (hashCode * 397) ^ (PropMask?.GetHashCode() ?? 0);
                hashCode = (hashCode * 397) ^ (BackColor?.GetHashCode() ?? 0);
                hashCode = (hashCode * 397) ^ (ForeColor?.GetHashCode() ?? 0);
                hashCode = (hashCode * 397) ^ (FontTextProps?.GetHashCode() ?? 0);
                hashCode = (hashCode * 397) ^ (int)NextAvailableId;
                hashCode = (hashCode * 397) ^ (BooleanProperties?.GetHashCode() ?? 0);
                hashCode = (hashCode * 397) ^ (int)BorderStyle;
                hashCode = (hashCode * 397) ^ (int)MousePointer;
                hashCode = (hashCode * 397) ^ (ScrollBars?.GetHashCode() ?? 0);
                hashCode = (hashCode * 397) ^ GroupCount;
                hashCode = (hashCode * 397) ^ (int)Cycle;
                hashCode = (hashCode * 397) ^ (int)SpecialEffect;
                hashCode = (hashCode * 397) ^ (BorderColor?.GetHashCode() ?? 0);
                hashCode = (hashCode * 397) ^ (int)Zoom;
                hashCode = (hashCode * 397) ^ (int)PictureAlignment;
                hashCode = (hashCode * 397) ^ (int)PictureSizeMode;
                hashCode = (hashCode * 397) ^ (int)ShapeCookie;
                hashCode = (hashCode * 397) ^ (int)DrawBuffer;
                hashCode = (hashCode * 397) ^ (DisplayedSize?.GetHashCode() ?? 0);
                hashCode = (hashCode * 397) ^ (LogicalSize?.GetHashCode() ?? 0);
                hashCode = (hashCode * 397) ^ (ScrollPosition?.GetHashCode() ?? 0);
                hashCode = (hashCode * 397) ^ (Caption?.GetHashCode() ?? 0);
                hashCode = (hashCode * 397) ^ FontIsStdFont.GetHashCode();
                hashCode = (hashCode * 397) ^ (FontStdFont?.GetHashCode() ?? 0);
                return hashCode;
            }
        }
    }

    internal class FormFlags
    {
        public FormFlags(uint i)
        {
            Func<int, bool> bit = j => (i & ((uint)1 << j)) != 0;
            Enabled = bit(2);
            DesignExtenderPropertiesPersisted = bit(14);
            ClassTablePersisted = !bit(15);
        }

        public bool ClassTablePersisted { get; }
        public bool DesignExtenderPropertiesPersisted { get; }
        public bool Enabled { get; }

        public override bool Equals(object obj)
        {
            var other = obj as FormFlags;
            return other != null && Enabled == other.Enabled && DesignExtenderPropertiesPersisted == other.DesignExtenderPropertiesPersisted
                && ClassTablePersisted == other.ClassTablePersisted;
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = Enabled.GetHashCode();
                hashCode = (hashCode * 397) ^ DesignExtenderPropertiesPersisted.GetHashCode();
                hashCode = (hashCode * 397) ^ ClassTablePersisted.GetHashCode();
                return hashCode;
            }
        }
    }

    internal class FormPropMask
    {
        public FormPropMask(uint i)
        {
            Func<int, bool> bit = j => (i & ((uint)1 << j)) != 0;
            HasBackColor = bit(1);
            HasForeColor = bit(2);
            HasNextAvailableId = bit(3);
            HasBooleanProperties = bit(6);
            HasBorderStyle = bit(7);
            HasMousePointer = bit(8);
            HasScrollBars = bit(9);
            HasDisplayedSize = bit(10);
            HasLogicalSize = bit(11);
            HasScrollPosition = bit(12);
            HasGroupCount = bit(13);
            HasMouseIcon = bit(15);
            HasCycle = bit(16);
            HasSpecialEffect = bit(17);
            HasBorderColor = bit(18);
            HasCaption = bit(19);
            HasFont = bit(20);
            HasPicture = bit(21);
            HasZoom = bit(22);
            HasPictureAlignment = bit(23);
            PictureTiling = bit(24);
            HasPictureSizeMode = bit(25);
            HasShapeCookie = bit(26);
            HasDrawBuffer = bit(27);
        }

        public bool HasBackColor { get; }
        public bool HasBooleanProperties { get; }
        public bool HasBorderColor { get; }
        public bool HasBorderStyle { get; }
        public bool HasCaption { get; }
        public bool HasCycle { get; }
        public bool HasDisplayedSize { get; }
        public bool HasDrawBuffer { get; }
        public bool HasFont { get; }
        public bool HasForeColor { get; }
        public bool HasGroupCount { get; }
        public bool HasLogicalSize { get; }
        public bool HasMouseIcon { get; }
        public bool HasMousePointer { get; }
        public bool HasNextAvailableId { get; }
        public bool HasPicture { get; }
        public bool HasPictureAlignment { get; }
        public bool HasPictureSizeMode { get; }
        public bool HasScrollBars { get; }
        public bool HasScrollPosition { get; }
        public bool HasShapeCookie { get; }
        public bool HasSpecialEffect { get; }
        public bool HasZoom { get; }
        public bool PictureTiling { get; }

        public override bool Equals(object obj)
        {
            var other = obj as FormPropMask;
            return other != null && HasBackColor == other.HasBackColor && HasForeColor == other.HasForeColor && HasNextAvailableId == other.HasNextAvailableId
                   && HasBooleanProperties == other.HasBooleanProperties && HasBorderStyle == other.HasBorderStyle && HasMousePointer == other.HasMousePointer
                   && HasScrollBars == other.HasScrollBars && HasDisplayedSize == other.HasDisplayedSize && HasLogicalSize == other.HasLogicalSize
                   && HasScrollPosition == other.HasScrollPosition && HasGroupCount == other.HasGroupCount && HasMouseIcon == other.HasMouseIcon
                   && HasCycle == other.HasCycle && HasSpecialEffect == other.HasSpecialEffect && HasBorderColor == other.HasBorderColor
                   && HasCaption == other.HasCaption && HasFont == other.HasFont && HasPicture == other.HasPicture && HasZoom == other.HasZoom
                   && HasPictureAlignment == other.HasPictureAlignment && PictureTiling == other.PictureTiling && HasPictureSizeMode == other.HasPictureSizeMode
                   && HasShapeCookie == other.HasShapeCookie && HasDrawBuffer == other.HasDrawBuffer;
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = HasBackColor.GetHashCode();
                hashCode = (hashCode * 397) ^ HasForeColor.GetHashCode();
                hashCode = (hashCode * 397) ^ HasNextAvailableId.GetHashCode();
                hashCode = (hashCode * 397) ^ HasBooleanProperties.GetHashCode();
                hashCode = (hashCode * 397) ^ HasBorderStyle.GetHashCode();
                hashCode = (hashCode * 397) ^ HasMousePointer.GetHashCode();
                hashCode = (hashCode * 397) ^ HasScrollBars.GetHashCode();
                hashCode = (hashCode * 397) ^ HasDisplayedSize.GetHashCode();
                hashCode = (hashCode * 397) ^ HasLogicalSize.GetHashCode();
                hashCode = (hashCode * 397) ^ HasScrollPosition.GetHashCode();
                hashCode = (hashCode * 397) ^ HasGroupCount.GetHashCode();
                hashCode = (hashCode * 397) ^ HasMouseIcon.GetHashCode();
                hashCode = (hashCode * 397) ^ HasCycle.GetHashCode();
                hashCode = (hashCode * 397) ^ HasSpecialEffect.GetHashCode();
                hashCode = (hashCode * 397) ^ HasBorderColor.GetHashCode();
                hashCode = (hashCode * 397) ^ HasCaption.GetHashCode();
                hashCode = (hashCode * 397) ^ HasFont.GetHashCode();
                hashCode = (hashCode * 397) ^ HasPicture.GetHashCode();
                hashCode = (hashCode * 397) ^ HasZoom.GetHashCode();
                hashCode = (hashCode * 397) ^ HasPictureAlignment.GetHashCode();
                hashCode = (hashCode * 397) ^ PictureTiling.GetHashCode();
                hashCode = (hashCode * 397) ^ HasPictureSizeMode.GetHashCode();
                hashCode = (hashCode * 397) ^ HasShapeCookie.GetHashCode();
                hashCode = (hashCode * 397) ^ HasDrawBuffer.GetHashCode();
                return hashCode;
            }
        }
    }

    internal class FormScrollBarFlags
    {
        public FormScrollBarFlags(byte b)
        {
            Func<int, bool> bit = j => (b & (1 << j)) != 0;
            Horizontal = bit(0);
            Vertical = bit(1);
            KeepHorizontal = bit(2);
            KeepVertical = bit(3);
            KeepLeft = bit(4);
        }

        public bool Horizontal { get; }
        public bool KeepHorizontal { get; }
        public bool KeepLeft { get; }
        public bool KeepVertical { get; }
        public bool Vertical { get; }

        public override bool Equals(object obj)
        {
            var other = obj as FormScrollBarFlags;
            return other != null && Horizontal == other.Horizontal && Vertical == other.Vertical && KeepHorizontal == other.KeepHorizontal
                && KeepVertical == other.KeepVertical && KeepLeft == other.KeepLeft;
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = Horizontal.GetHashCode();
                hashCode = (hashCode * 397) ^ Vertical.GetHashCode();
                hashCode = (hashCode * 397) ^ KeepHorizontal.GetHashCode();
                hashCode = (hashCode * 397) ^ KeepVertical.GetHashCode();
                hashCode = (hashCode * 397) ^ KeepLeft.GetHashCode();
                return hashCode;
            }
        }
    }

    internal class OleSiteConcreteControl
    {
        public OleSiteConcreteControl(byte[] b)
        {
            using (var st = new MemoryStream(b))
            using (var r = new FrxReader(st))
            {
                PropMask = new SitePropMask(r.ReadUInt32());

                // DataBlock
                var nameCcb = PropMask.HasName ? r.ReadCcb() : Tuple.Create(0, false);
                var tagCcb = PropMask.HasTag ? r.ReadCcb() : Tuple.Create(0, false);
                Id = PropMask.HasId ? r.ReadInt32() : 0;
                HelpContextId = PropMask.HasHelpContextId ? r.ReadInt32() : 0;
                BitFlags = PropMask.HasBitFlags ? r.ReadUInt32() : 0;
                ObjectStreamSize = PropMask.HasObjectStreamSize ? r.ReadUInt32() : 0;
                TabIndex = PropMask.HasTabIndex ? r.ReadInt16() : (short)0;
                ClsidCacheIndex = PropMask.HasClsidCacheIndex ? r.ReadInt16() : (short)0;
                GroupId = PropMask.HasGroupId ? r.ReadUInt16() : (ushort)0;
                var controlTipTextCcb = PropMask.HasControlTipText ? r.ReadCcb() : Tuple.Create(0, false);
                var runtimeLicKeyCcb = PropMask.HasRuntimeLicKey ? r.ReadCcb() : Tuple.Create(0, false);
                var controlSourceCcb = PropMask.HasControlSource ? r.ReadCcb() : Tuple.Create(0, false);
                var rowSourceCcb = PropMask.HasRowSource ? r.ReadCcb() : Tuple.Create(0, false);

                // ExtraDataBlock
                Name = r.ReadStringFromCcb(nameCcb);
                Tag = r.ReadStringFromCcb(tagCcb);
                SitePosition = PropMask.HasPosition ? r.ReadCoords() : Tuple.Create(0, 0);
                ControlTipText = r.ReadStringFromCcb(controlTipTextCcb);
                RuntimeLicKey = r.ReadStringFromCcb(runtimeLicKeyCcb);
                ControlSource = r.ReadStringFromCcb(controlSourceCcb);
                RowSource = r.ReadStringFromCcb(rowSourceCcb);

                if (st.Position < st.Length)
                {
                    throw new ApplicationException(VBASyncResources.ErrorFrxExpectedEndOfOleSiteConcreteControl);
                }
            }
        }

        public uint BitFlags { get; }
        public short ClsidCacheIndex { get; }
        public string ControlSource { get; }
        public string ControlTipText { get; }
        public ushort GroupId { get; }
        public int HelpContextId { get; }
        public int Id { get; }
        public string Name { get; }
        public uint ObjectStreamSize { get; }
        public SitePropMask PropMask { get; }
        public string RowSource { get; }
        public string RuntimeLicKey { get; }
        public short TabIndex { get; }
        public string Tag { get; }
        private Tuple<int, int> SitePosition { get; }

        public override bool Equals(object obj)
        {
            var other = obj as OleSiteConcreteControl;
            return other != null && Equals(PropMask, other.PropMask) && string.Equals(Name, other.Name) && string.Equals(Tag, other.Tag) && Id == other.Id
                   && HelpContextId == other.HelpContextId && BitFlags == other.BitFlags && ObjectStreamSize == other.ObjectStreamSize
                   && TabIndex == other.TabIndex && ClsidCacheIndex == other.ClsidCacheIndex && GroupId == other.GroupId
                   && string.Equals(ControlTipText, other.ControlTipText) && string.Equals(RuntimeLicKey, other.RuntimeLicKey)
                   && string.Equals(ControlSource, other.ControlSource) && string.Equals(RowSource, other.RowSource) && Equals(SitePosition, other.SitePosition);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = PropMask?.GetHashCode() ?? 0;
                hashCode = (hashCode * 397) ^ (Name?.GetHashCode() ?? 0);
                hashCode = (hashCode * 397) ^ (Tag?.GetHashCode() ?? 0);
                hashCode = (hashCode * 397) ^ Id;
                hashCode = (hashCode * 397) ^ HelpContextId;
                hashCode = (hashCode * 397) ^ (int)BitFlags;
                hashCode = (hashCode * 397) ^ (int)ObjectStreamSize;
                hashCode = (hashCode * 397) ^ TabIndex.GetHashCode();
                hashCode = (hashCode * 397) ^ ClsidCacheIndex.GetHashCode();
                hashCode = (hashCode * 397) ^ GroupId.GetHashCode();
                hashCode = (hashCode * 397) ^ (ControlTipText?.GetHashCode() ?? 0);
                hashCode = (hashCode * 397) ^ (RuntimeLicKey?.GetHashCode() ?? 0);
                hashCode = (hashCode * 397) ^ (ControlSource?.GetHashCode() ?? 0);
                hashCode = (hashCode * 397) ^ (RowSource?.GetHashCode() ?? 0);
                hashCode = (hashCode * 397) ^ (SitePosition?.GetHashCode() ?? 0);
                return hashCode;
            }
        }
    }

    internal class SitePropMask
    {
        public SitePropMask(uint i)
        {
            Func<int, bool> bit = j => (i & ((uint)1 << j)) != 0;
            HasName = bit(0);
            HasTag = bit(1);
            HasId = bit(2);
            HasHelpContextId = bit(3);
            HasBitFlags = bit(4);
            HasObjectStreamSize = bit(5);
            HasTabIndex = bit(6);
            HasClsidCacheIndex = bit(7);
            HasPosition = bit(8);
            HasGroupId = bit(9);
            HasControlTipText = bit(11);
            HasRuntimeLicKey = bit(12);
            HasControlSource = bit(13);
            HasRowSource = bit(14);
        }

        public bool HasBitFlags { get; }
        public bool HasClsidCacheIndex { get; }
        public bool HasControlSource { get; }
        public bool HasControlTipText { get; }
        public bool HasGroupId { get; }
        public bool HasHelpContextId { get; }
        public bool HasId { get; }
        public bool HasName { get; }
        public bool HasObjectStreamSize { get; }
        public bool HasPosition { get; }
        public bool HasRowSource { get; }
        public bool HasRuntimeLicKey { get; }
        public bool HasTabIndex { get; }
        public bool HasTag { get; }

        public override bool Equals(object obj)
        {
            var other = obj as SitePropMask;
            return other != null && HasName == other.HasName && HasTag == other.HasTag && HasId == other.HasId && HasHelpContextId == other.HasHelpContextId
                   && HasBitFlags == other.HasBitFlags && HasObjectStreamSize == other.HasObjectStreamSize && HasTabIndex == other.HasTabIndex
                   && HasClsidCacheIndex == other.HasClsidCacheIndex && HasPosition == other.HasPosition && HasGroupId == other.HasGroupId
                   && HasControlTipText == other.HasControlTipText && HasRuntimeLicKey == other.HasRuntimeLicKey && HasControlSource == other.HasControlSource
                   && HasRowSource == other.HasRowSource;
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = HasName.GetHashCode();
                hashCode = (hashCode * 397) ^ HasTag.GetHashCode();
                hashCode = (hashCode * 397) ^ HasId.GetHashCode();
                hashCode = (hashCode * 397) ^ HasHelpContextId.GetHashCode();
                hashCode = (hashCode * 397) ^ HasBitFlags.GetHashCode();
                hashCode = (hashCode * 397) ^ HasObjectStreamSize.GetHashCode();
                hashCode = (hashCode * 397) ^ HasTabIndex.GetHashCode();
                hashCode = (hashCode * 397) ^ HasClsidCacheIndex.GetHashCode();
                hashCode = (hashCode * 397) ^ HasPosition.GetHashCode();
                hashCode = (hashCode * 397) ^ HasGroupId.GetHashCode();
                hashCode = (hashCode * 397) ^ HasControlTipText.GetHashCode();
                hashCode = (hashCode * 397) ^ HasRuntimeLicKey.GetHashCode();
                hashCode = (hashCode * 397) ^ HasControlSource.GetHashCode();
                hashCode = (hashCode * 397) ^ HasRowSource.GetHashCode();
                return hashCode;
            }
        }
    }
}
