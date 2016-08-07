using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using static VbaSync.FrxObjects.AlignmentHelpers;
using static VbaSync.FrxObjects.StreamDataHelpers;

namespace VbaSync.FrxObjects {
    class FormControl : FrxCommon {
        public byte MinorVersion { get; }
        public byte MajorVersion { get; }
        public FormPropMask PropMask { get; }
        public OleColor BackColor { get; }
        public OleColor ForeColor { get; }
        public uint NextAvailableId { get; }
        public FormFlags BooleanProperties { get; }
        public BorderStyle BorderStyle { get; }
        public MousePointer MousePointer { get; }
        public FormScrollBarFlags ScrollBars { get; }
        public int GroupCount { get; }
        public Cycle Cycle { get; }
        public SpecialEffect SpecialEffect { get; }
        public OleColor BorderColor { get; }
        public uint Zoom { get; }
        public PictureAlignment PictureAlignment { get; }
        public PictureSizeMode PictureSizeMode { get; }
        public uint ShapeCookie { get; }
        public uint DrawBuffer { get; }
        public Tuple<int, int> DisplayedSize { get; }
        public Tuple<int, int> LogicalSize { get; }
        public Tuple<int, int> ScrollPosition { get; }
        public string Caption { get; }
        public byte[] MouseIcon { get; } = new byte[0];
        public bool FontIsStdFont { get; }
        public byte[] Picture { get; } = new byte[0];
        public TextProps FontTextProps { get; }
        public Tuple<short, byte, short, uint, string> FontStdFont { get; }
        public List<byte[]> SiteClassInfos { get; }
        public OleSiteConcreteControl[] Sites { get; }
        public byte[] Remainder { get; } = new byte[0];

        public FormControl(byte[] b) {
            using (var st = new MemoryStream(b))
            using (var r = new BinaryReader(st)) {
                MinorVersion = r.ReadByte();
                MajorVersion = r.ReadByte();

                var cbForm = r.ReadUInt16();
                PropMask = new FormPropMask(r.ReadUInt32());

                BeginDataBlock();
                BackColor = ReadAlignedOleColorIf(PropMask.HasBackColor, r);
                ForeColor = ReadAlignedOleColorIf(PropMask.HasForeColor, r);
                NextAvailableId = ReadAlignedUInt32If(PropMask.HasNextAvailableId, r);
                BooleanProperties = new FormFlags(ReadAlignedUInt32If(PropMask.HasBooleanProperties, r));
                BorderStyle = (BorderStyle)ReadByteIf(PropMask.HasBorderStyle, r);
                MousePointer = (MousePointer)ReadByteIf(PropMask.HasMousePointer, r);
                ScrollBars = new FormScrollBarFlags(ReadByteIf(PropMask.HasScrollBars, r));
                GroupCount = ReadAlignedInt32If(PropMask.HasGroupCount, r);
                var captionCcb = ReadAlignedCcbIf(PropMask.HasCaption, r); // this seems to be in a different position than indicated in MS's spec?
                Ignore2AlignedBytesIf(PropMask.HasFont, r);
                Ignore2AlignedBytesIf(PropMask.HasMouseIcon, r);
                Cycle = (Cycle)ReadByteIf(PropMask.HasCycle, r);
                SpecialEffect = (SpecialEffect)ReadByteIf(PropMask.HasSpecialEffect, r);
                BorderColor = ReadAlignedOleColorIf(PropMask.HasBorderColor, r);
                Ignore2AlignedBytesIf(PropMask.HasPicture, r);
                Zoom = ReadAlignedUInt32If(PropMask.HasZoom, r);
                PictureAlignment = (PictureAlignment)ReadByteIf(PropMask.HasPictureAlignment, r);
                PictureSizeMode = (PictureSizeMode)ReadByteIf(PropMask.HasPictureSizeMode, r);
                ShapeCookie = ReadAlignedUInt32If(PropMask.HasShapeCookie, r);
                DrawBuffer = ReadAlignedUInt32If(PropMask.HasDrawBuffer, r);
                EndDataBlock(r);

                BeginExtraDataBlock();
                DisplayedSize = ReadAlignedCoordsIf(PropMask.HasDisplayedSize, r);
                LogicalSize = ReadAlignedCoordsIf(PropMask.HasLogicalSize, r);
                ScrollPosition = ReadAlignedCoordsIf(PropMask.HasScrollPosition, r);
                Caption = ReadStringFromCcb(captionCcb, r);
                EndExtraDataBlock(r);

                if (cbForm != 4 + DataBlockBytes + ExtraDataBlockBytes)
                    throw new ApplicationException("Error reading 'f' stream in .frx data: expected cbForm size "
                                                   + $"{4 + DataBlockBytes + ExtraDataBlockBytes}, but actual size was {cbForm}.");

                // StreamData
                if (PropMask.HasMouseIcon) {
                    MouseIcon = ReadGuidAndPicture(r);
                }
                if (PropMask.HasFont) {
                    FontIsStdFont = GetFontIsStdFont(r);
                    if (FontIsStdFont) {
                        FontStdFont = ReadStdFont(r);
                    } else {
                        FontTextProps = ReadTextProps(r);
                    }
                }
                if (PropMask.HasPicture) {
                    Picture = ReadGuidAndPicture(r);
                }

                // FormSiteData
                SiteClassInfos = new List<byte[]>();
                ushort siteClassInfoCount = 0;
                if (!PropMask.HasBooleanProperties || BooleanProperties.ClassTablePersisted) {
                    siteClassInfoCount = r.ReadUInt16();
                }
                for (var i = 0; i < siteClassInfoCount; i++) {
                    st.Seek(2, SeekOrigin.Current); // skip Version
                    SiteClassInfos.Add(r.ReadBytes(r.ReadUInt16()));
                }
                var siteCount = r.ReadUInt32();
                var cbSites = r.ReadUInt32();
                uint sitesBytes = 0;
                var depths = new byte[siteCount];
                var types = new byte[siteCount];
                var siteDepthsLeft = siteCount;
                while (siteDepthsLeft > 0) {
                    var thisDepth = r.ReadByte();
                    var thisType = r.ReadByte();
                    sitesBytes += 2;
                    var thisCount = (byte)1;
                    if ((thisType & 0x80) == 0x80) {
                        thisCount = (byte)(thisType ^ 0x80);
                        thisType = r.ReadByte();
                        sitesBytes += 1;
                    }
                    for (var i = 0; i < thisCount; i++) {
                        var siteIdx = siteCount - siteDepthsLeft;
                        depths[siteIdx] = thisDepth;
                        types[siteIdx] = thisType;
                        siteDepthsLeft--;
                    }
                }
                AlignTo(4, st, ref sitesBytes);
                Sites = new OleSiteConcreteControl[siteCount];
                for (var i = 0; i < siteCount; i++) {
                    IgnoreNext(2, st, ref sitesBytes); // ignore Version
                    var cbSite = r.ReadUInt16();
                    sitesBytes += 2;
                    Sites[i] = new OleSiteConcreteControl(r.ReadBytes(cbSite));
                    sitesBytes += cbSite;
                }
                if (cbSites != sitesBytes) {
                    throw new ApplicationException("Error reading 'f' stream in .frx data: expected cbSites size "
                        + $"{sitesBytes} but actual size was {cbSites}.");
                }

                if (st.Position < st.Length) {
                    Remainder = r.ReadBytes((int)(st.Length - st.Position));
                }
            }
        }

        public override bool Equals(object o) {
            var other = o as FormControl;
            if (other == null || !(MinorVersion == other.MinorVersion && MajorVersion == other.MajorVersion && Equals(PropMask, other.PropMask)
                  && Equals(BackColor, other.BackColor) && Equals(ForeColor, other.ForeColor) && NextAvailableId == other.NextAvailableId
                  && Equals(BooleanProperties, other.BooleanProperties) && BorderStyle == other.BorderStyle && MousePointer == other.MousePointer
                  && Equals(ScrollBars, other.ScrollBars) && GroupCount == other.GroupCount && Cycle == other.Cycle && SpecialEffect == other.SpecialEffect
                  && Equals(BorderColor, other.BorderColor) && Zoom == other.Zoom && PictureAlignment == other.PictureAlignment
                  && PictureSizeMode == other.PictureSizeMode && ShapeCookie == other.ShapeCookie && DrawBuffer == other.DrawBuffer
                  && Equals(DisplayedSize, other.DisplayedSize) && Equals(LogicalSize, other.LogicalSize) && Equals(ScrollPosition, other.ScrollPosition)
                  && string.Equals(Caption, other.Caption) && MouseIcon.SequenceEqual(other.MouseIcon) && FontIsStdFont == other.FontIsStdFont
                  && Picture.SequenceEqual(other.Picture) && Equals(FontTextProps, other.FontTextProps) && Equals(FontStdFont, other.FontStdFont)
                  && Sites.SequenceEqual(other.Sites) && Remainder.SequenceEqual(other.Remainder))) return false;
            if (SiteClassInfos.Count != other.SiteClassInfos.Count) return false;
            return !SiteClassInfos.Where((t, i) => !t.SequenceEqual(other.SiteClassInfos[i])).Any();
        }

        public override int GetHashCode() {
            unchecked {
                var hashCode = MinorVersion.GetHashCode();
                hashCode = (hashCode*397) ^ MajorVersion.GetHashCode();
                hashCode = (hashCode*397) ^ (PropMask?.GetHashCode() ?? 0);
                hashCode = (hashCode*397) ^ (BackColor?.GetHashCode() ?? 0);
                hashCode = (hashCode*397) ^ (ForeColor?.GetHashCode() ?? 0);
                hashCode = (hashCode*397) ^ (FontTextProps?.GetHashCode() ?? 0);
                hashCode = (hashCode*397) ^ (int)NextAvailableId;
                hashCode = (hashCode*397) ^ (BooleanProperties?.GetHashCode() ?? 0);
                hashCode = (hashCode*397) ^ (int)BorderStyle;
                hashCode = (hashCode*397) ^ (int)MousePointer;
                hashCode = (hashCode*397) ^ (ScrollBars?.GetHashCode() ?? 0);
                hashCode = (hashCode*397) ^ GroupCount;
                hashCode = (hashCode*397) ^ (int)Cycle;
                hashCode = (hashCode*397) ^ (int)SpecialEffect;
                hashCode = (hashCode*397) ^ (BorderColor?.GetHashCode() ?? 0);
                hashCode = (hashCode*397) ^ (int)Zoom;
                hashCode = (hashCode*397) ^ (int)PictureAlignment;
                hashCode = (hashCode*397) ^ (int)PictureSizeMode;
                hashCode = (hashCode*397) ^ (int)ShapeCookie;
                hashCode = (hashCode*397) ^ (int)DrawBuffer;
                hashCode = (hashCode*397) ^ (DisplayedSize?.GetHashCode() ?? 0);
                hashCode = (hashCode*397) ^ (LogicalSize?.GetHashCode() ?? 0);
                hashCode = (hashCode*397) ^ (ScrollPosition?.GetHashCode() ?? 0);
                hashCode = (hashCode*397) ^ (Caption?.GetHashCode() ?? 0);
                hashCode = (hashCode*397) ^ FontIsStdFont.GetHashCode();
                hashCode = (hashCode*397) ^ (FontStdFont?.GetHashCode() ?? 0);
                return hashCode;
            }
        }
    }

    class OleSiteConcreteControl {
        public SitePropMask PropMask { get; }
        public string Name { get; }
        public string Tag { get; }
        public int Id { get; }
        public int HelpContextId { get; }
        public uint BitFlags { get; }
        public uint ObjectStreamSize { get; }
        public short TabIndex { get; }
        public short ClsidCacheIndex { get; }
        public ushort GroupId { get; }
        public string ControlTipText { get; }
        public string RuntimeLicKey { get; }
        public string ControlSource { get; }
        public string RowSource { get; }
        Tuple<int, int> SitePosition { get; }

        public OleSiteConcreteControl(byte[] b) {
            using (var st = new MemoryStream(b))
            using (var r = new BinaryReader(st)) {
                PropMask = new SitePropMask(r.ReadUInt32());

                // SiteDataBlock
                ushort dataBlockBytes = 0;
                var nameLength = 0;
                var nameCompressed = false;
                if (PropMask.HasName) {
                    nameLength = CcbToLength(r.ReadInt32(), out nameCompressed);
                    dataBlockBytes += 4;
                }
                var tagLength = 0;
                var tagCompressed = false;
                if (PropMask.HasTag) {
                    tagLength = CcbToLength(r.ReadInt32(), out tagCompressed);
                    dataBlockBytes += 4;
                }
                if (PropMask.HasId) {
                    Id = r.ReadInt32();
                    dataBlockBytes += 4;
                }
                if (PropMask.HasHelpContextId) {
                    HelpContextId = r.ReadInt32();
                    dataBlockBytes += 4;
                }
                if (PropMask.HasBitFlags) {
                    BitFlags = r.ReadUInt32();
                    dataBlockBytes += 4;
                }
                if (PropMask.HasObjectStreamSize) {
                    ObjectStreamSize = r.ReadUInt32();
                    dataBlockBytes += 4;
                }
                if (PropMask.HasTabIndex) {
                    TabIndex = r.ReadInt16();
                    dataBlockBytes += 2;
                }
                if (PropMask.HasClsidCacheIndex) {
                    ClsidCacheIndex = r.ReadInt16();
                    dataBlockBytes += 2;
                }
                if (PropMask.HasGroupId) {
                    GroupId = r.ReadUInt16();
                    dataBlockBytes += 2;
                }
                AlignTo(4, st, ref dataBlockBytes);
                var controlTipTextLength = 0;
                var controlTipTextCompressed = false;
                if (PropMask.HasControlTipText) {
                    controlTipTextLength = CcbToLength(r.ReadInt32(), out controlTipTextCompressed);
                    dataBlockBytes += 4;
                }
                var runtimeLicKeyLength = 0;
                var runtimeLicKeyCompressed = false;
                if (PropMask.HasRuntimeLicKey) {
                    runtimeLicKeyLength = CcbToLength(r.ReadInt32(), out runtimeLicKeyCompressed);
                    dataBlockBytes += 4;
                }
                var controlSourceLength = 0;
                var controlSourceCompressed = false;
                if (PropMask.HasControlSource) {
                    controlSourceLength = CcbToLength(r.ReadInt32(), out controlSourceCompressed);
                    dataBlockBytes += 4;
                }
                var rowSourceLength = 0;
                var rowSourceCompressed = false;
                if (PropMask.HasRowSource) {
                    rowSourceLength = CcbToLength(r.ReadInt32(), out rowSourceCompressed);
                    dataBlockBytes += 4;
                }

                // SiteExtraDataBlock
                ushort extraDataBlockBytes = 0;
                if (nameLength > 0) {
                    Name = (nameCompressed ? Encoding.UTF8 : Encoding.Unicode).GetString(r.ReadBytes(nameLength));
                    extraDataBlockBytes += (ushort)nameLength;
                    AlignTo(4, st, ref extraDataBlockBytes);
                }
                if (tagLength > 0) {
                    Tag = (tagCompressed ? Encoding.UTF8 : Encoding.Unicode).GetString(r.ReadBytes(tagLength));
                    extraDataBlockBytes += (ushort)tagLength;
                    AlignTo(4, st, ref extraDataBlockBytes);
                }
                if (PropMask.HasPosition) {
                    SitePosition = Tuple.Create(r.ReadInt32(), r.ReadInt32());
                }
                if (controlTipTextLength > 0) {
                    ControlTipText = (controlTipTextCompressed ? Encoding.UTF8 : Encoding.Unicode).GetString(r.ReadBytes(controlTipTextLength));
                    extraDataBlockBytes += (ushort)controlTipTextLength;
                    AlignTo(4, st, ref extraDataBlockBytes);
                }
                if (runtimeLicKeyLength > 0) {
                    RuntimeLicKey = (runtimeLicKeyCompressed ? Encoding.UTF8 : Encoding.Unicode).GetString(r.ReadBytes(runtimeLicKeyLength));
                    extraDataBlockBytes += (ushort)runtimeLicKeyLength;
                    AlignTo(4, st, ref extraDataBlockBytes);
                }
                if (controlSourceLength > 0) {
                    ControlSource = (controlSourceCompressed ? Encoding.UTF8 : Encoding.Unicode).GetString(r.ReadBytes(controlSourceLength));
                    extraDataBlockBytes += (ushort)controlSourceLength;
                    AlignTo(4, st, ref extraDataBlockBytes);
                }
                if (rowSourceLength > 0) {
                    RowSource = (rowSourceCompressed ? Encoding.UTF8 : Encoding.Unicode).GetString(r.ReadBytes(rowSourceLength));
                    extraDataBlockBytes += (ushort)rowSourceLength;
                    AlignTo(4, st, ref extraDataBlockBytes);
                }

                if (st.Position < st.Length) throw new ApplicationException("Expected end of OleSiteConcreteControl.");
            }
        }

        public override bool Equals(object o) {
            var other = o as OleSiteConcreteControl;
            return other != null && Equals(PropMask, other.PropMask) && string.Equals(Name, other.Name) && string.Equals(Tag, other.Tag) && Id == other.Id
                   && HelpContextId == other.HelpContextId && BitFlags == other.BitFlags && ObjectStreamSize == other.ObjectStreamSize
                   && TabIndex == other.TabIndex && ClsidCacheIndex == other.ClsidCacheIndex && GroupId == other.GroupId
                   && string.Equals(ControlTipText, other.ControlTipText) && string.Equals(RuntimeLicKey, other.RuntimeLicKey)
                   && string.Equals(ControlSource, other.ControlSource) && string.Equals(RowSource, other.RowSource) && Equals(SitePosition, other.SitePosition);
        }

        public override int GetHashCode() {
            unchecked {
                var hashCode = PropMask?.GetHashCode() ?? 0;
                hashCode = (hashCode*397) ^ (Name?.GetHashCode() ?? 0);
                hashCode = (hashCode*397) ^ (Tag?.GetHashCode() ?? 0);
                hashCode = (hashCode*397) ^ Id;
                hashCode = (hashCode*397) ^ HelpContextId;
                hashCode = (hashCode*397) ^ (int)BitFlags;
                hashCode = (hashCode*397) ^ (int)ObjectStreamSize;
                hashCode = (hashCode*397) ^ TabIndex.GetHashCode();
                hashCode = (hashCode*397) ^ ClsidCacheIndex.GetHashCode();
                hashCode = (hashCode*397) ^ GroupId.GetHashCode();
                hashCode = (hashCode*397) ^ (ControlTipText?.GetHashCode() ?? 0);
                hashCode = (hashCode*397) ^ (RuntimeLicKey?.GetHashCode() ?? 0);
                hashCode = (hashCode*397) ^ (ControlSource?.GetHashCode() ?? 0);
                hashCode = (hashCode*397) ^ (RowSource?.GetHashCode() ?? 0);
                hashCode = (hashCode*397) ^ (SitePosition?.GetHashCode() ?? 0);
                return hashCode;
            }
        }
    }

    enum Cycle {
        AllForms = 0x00,
        CurrentForm = 0x02
    }

    class FormFlags {
        public bool Enabled { get; }
        public bool DesignExtenderPropertiesPersisted { get; }
        public bool ClassTablePersisted { get; }

        public FormFlags(uint i) {
            Func<int, bool> bit = j => (i & ((uint)1 << j)) != 0;
            Enabled = bit(2);
            DesignExtenderPropertiesPersisted = bit(14);
            ClassTablePersisted = !bit(15);
        }

        public override bool Equals(object o) {
            var other = o as FormFlags;
            return other != null && Enabled == other.Enabled && DesignExtenderPropertiesPersisted == other.DesignExtenderPropertiesPersisted
                && ClassTablePersisted == other.ClassTablePersisted;
        }

        public override int GetHashCode() {
            unchecked {
                var hashCode = Enabled.GetHashCode();
                hashCode = (hashCode*397) ^ DesignExtenderPropertiesPersisted.GetHashCode();
                hashCode = (hashCode*397) ^ ClassTablePersisted.GetHashCode();
                return hashCode;
            }
        }
    }

    enum FormScrollBars {
        None = 0x00,
        Horizontal = 0x01,
        Vertical = 0x02,
        Both = 0x03
    }

    class FormScrollBarFlags {
        public bool Horizontal { get; }
        public bool Vertical { get; }
        public bool KeepHorizontal { get; }
        public bool KeepVertical { get; }
        public bool KeepLeft { get; }

        public FormScrollBarFlags(byte b) {
            Func<int, bool> bit = j => (b & (1 << j)) != 0;
            Horizontal = bit(0);
            Vertical = bit(1);
            KeepHorizontal = bit(2);
            KeepVertical = bit(3);
            KeepLeft = bit(4);
        }

        public override bool Equals(object o) {
            var other = o as FormScrollBarFlags;
            return other != null && Horizontal == other.Horizontal && Vertical == other.Vertical && KeepHorizontal == other.KeepHorizontal
                && KeepVertical == other.KeepVertical && KeepLeft == other.KeepLeft;
        }

        public override int GetHashCode() {
            unchecked {
                var hashCode = Horizontal.GetHashCode();
                hashCode = (hashCode*397) ^ Vertical.GetHashCode();
                hashCode = (hashCode*397) ^ KeepHorizontal.GetHashCode();
                hashCode = (hashCode*397) ^ KeepVertical.GetHashCode();
                hashCode = (hashCode*397) ^ KeepLeft.GetHashCode();
                return hashCode;
            }
        }
    }

    class FormPropMask {
        public bool HasBackColor { get; }
        public bool HasForeColor { get; }
        public bool HasNextAvailableId { get; }
        public bool HasBooleanProperties { get; }
        public bool HasBorderStyle { get; }
        public bool HasMousePointer { get; }
        public bool HasScrollBars { get; }
        public bool HasDisplayedSize { get; }
        public bool HasLogicalSize { get; }
        public bool HasScrollPosition { get; }
        public bool HasGroupCount { get; }
        public bool HasMouseIcon { get; }
        public bool HasCycle { get; }
        public bool HasSpecialEffect { get; }
        public bool HasBorderColor { get; }
        public bool HasCaption { get; }
        public bool HasFont { get; }
        public bool HasPicture { get; }
        public bool HasZoom { get; }
        public bool HasPictureAlignment { get; }
        public bool PictureTiling { get; }
        public bool HasPictureSizeMode { get; }
        public bool HasShapeCookie { get; }
        public bool HasDrawBuffer { get; }

        public FormPropMask(uint i) {
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

        public override bool Equals(object o) {
            var other = o as FormPropMask;
            return other != null && HasBackColor == other.HasBackColor && HasForeColor == other.HasForeColor && HasNextAvailableId == other.HasNextAvailableId
                   && HasBooleanProperties == other.HasBooleanProperties && HasBorderStyle == other.HasBorderStyle && HasMousePointer == other.HasMousePointer
                   && HasScrollBars == other.HasScrollBars && HasDisplayedSize == other.HasDisplayedSize && HasLogicalSize == other.HasLogicalSize
                   && HasScrollPosition == other.HasScrollPosition && HasGroupCount == other.HasGroupCount && HasMouseIcon == other.HasMouseIcon
                   && HasCycle == other.HasCycle && HasSpecialEffect == other.HasSpecialEffect && HasBorderColor == other.HasBorderColor
                   && HasCaption == other.HasCaption && HasFont == other.HasFont && HasPicture == other.HasPicture && HasZoom == other.HasZoom
                   && HasPictureAlignment == other.HasPictureAlignment && PictureTiling == other.PictureTiling && HasPictureSizeMode == other.HasPictureSizeMode
                   && HasShapeCookie == other.HasShapeCookie && HasDrawBuffer == other.HasDrawBuffer;
        }

        public override int GetHashCode() {
            unchecked {
                var hashCode = HasBackColor.GetHashCode();
                hashCode = (hashCode*397) ^ HasForeColor.GetHashCode();
                hashCode = (hashCode*397) ^ HasNextAvailableId.GetHashCode();
                hashCode = (hashCode*397) ^ HasBooleanProperties.GetHashCode();
                hashCode = (hashCode*397) ^ HasBorderStyle.GetHashCode();
                hashCode = (hashCode*397) ^ HasMousePointer.GetHashCode();
                hashCode = (hashCode*397) ^ HasScrollBars.GetHashCode();
                hashCode = (hashCode*397) ^ HasDisplayedSize.GetHashCode();
                hashCode = (hashCode*397) ^ HasLogicalSize.GetHashCode();
                hashCode = (hashCode*397) ^ HasScrollPosition.GetHashCode();
                hashCode = (hashCode*397) ^ HasGroupCount.GetHashCode();
                hashCode = (hashCode*397) ^ HasMouseIcon.GetHashCode();
                hashCode = (hashCode*397) ^ HasCycle.GetHashCode();
                hashCode = (hashCode*397) ^ HasSpecialEffect.GetHashCode();
                hashCode = (hashCode*397) ^ HasBorderColor.GetHashCode();
                hashCode = (hashCode*397) ^ HasCaption.GetHashCode();
                hashCode = (hashCode*397) ^ HasFont.GetHashCode();
                hashCode = (hashCode*397) ^ HasPicture.GetHashCode();
                hashCode = (hashCode*397) ^ HasZoom.GetHashCode();
                hashCode = (hashCode*397) ^ HasPictureAlignment.GetHashCode();
                hashCode = (hashCode*397) ^ PictureTiling.GetHashCode();
                hashCode = (hashCode*397) ^ HasPictureSizeMode.GetHashCode();
                hashCode = (hashCode*397) ^ HasShapeCookie.GetHashCode();
                hashCode = (hashCode*397) ^ HasDrawBuffer.GetHashCode();
                return hashCode;
            }
        }
    }

    class SitePropMask {
        public bool HasName { get; }
        public bool HasTag { get; }
        public bool HasId { get; }
        public bool HasHelpContextId { get; }
        public bool HasBitFlags { get; }
        public bool HasObjectStreamSize { get; }
        public bool HasTabIndex { get; }
        public bool HasClsidCacheIndex { get; }
        public bool HasPosition { get; }
        public bool HasGroupId { get; }
        public bool HasControlTipText { get; }
        public bool HasRuntimeLicKey { get; }
        public bool HasControlSource { get; }
        public bool HasRowSource { get; }

        public SitePropMask(uint i) {
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

        public override bool Equals(object o) {
            var other = o as SitePropMask;
            return other != null && HasName == other.HasName && HasTag == other.HasTag && HasId == other.HasId && HasHelpContextId == other.HasHelpContextId
                   && HasBitFlags == other.HasBitFlags && HasObjectStreamSize == other.HasObjectStreamSize && HasTabIndex == other.HasTabIndex
                   && HasClsidCacheIndex == other.HasClsidCacheIndex && HasPosition == other.HasPosition && HasGroupId == other.HasGroupId
                   && HasControlTipText == other.HasControlTipText && HasRuntimeLicKey == other.HasRuntimeLicKey && HasControlSource == other.HasControlSource
                   && HasRowSource == other.HasRowSource;
        }

        public override int GetHashCode() {
            unchecked {
                var hashCode = HasName.GetHashCode();
                hashCode = (hashCode*397) ^ HasTag.GetHashCode();
                hashCode = (hashCode*397) ^ HasId.GetHashCode();
                hashCode = (hashCode*397) ^ HasHelpContextId.GetHashCode();
                hashCode = (hashCode*397) ^ HasBitFlags.GetHashCode();
                hashCode = (hashCode*397) ^ HasObjectStreamSize.GetHashCode();
                hashCode = (hashCode*397) ^ HasTabIndex.GetHashCode();
                hashCode = (hashCode*397) ^ HasClsidCacheIndex.GetHashCode();
                hashCode = (hashCode*397) ^ HasPosition.GetHashCode();
                hashCode = (hashCode*397) ^ HasGroupId.GetHashCode();
                hashCode = (hashCode*397) ^ HasControlTipText.GetHashCode();
                hashCode = (hashCode*397) ^ HasRuntimeLicKey.GetHashCode();
                hashCode = (hashCode*397) ^ HasControlSource.GetHashCode();
                hashCode = (hashCode*397) ^ HasRowSource.GetHashCode();
                return hashCode;
            }
        }
    }
}
