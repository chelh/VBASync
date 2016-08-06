using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace VbaSync.FrxObjects {
    class FormControl {
        public byte MinorVersion { get; }
        public byte MajorVersion { get; }
        public PropMask PropMask { get; }
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
        public byte[] FontTextProps { get; } = new byte[0];
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
                PropMask = new PropMask(r.ReadUInt32());

                // DataBlock
                ushort dataBlockBytes = 0;
                if (PropMask.HasBackColor) {
                    BackColor = new OleColor(r.ReadBytes(4));
                    dataBlockBytes += 4;
                }
                if (PropMask.HasForeColor) {
                    ForeColor = new OleColor(r.ReadBytes(4));
                    dataBlockBytes += 4;
                }
                if (PropMask.HasNextAvailableId) {
                    NextAvailableId = r.ReadUInt32();
                    dataBlockBytes += 4;
                }
                if (PropMask.HasBooleanProperties) {
                    BooleanProperties = new FormFlags(r.ReadUInt32());
                    dataBlockBytes += 4;
                }
                if (PropMask.HasBorderStyle) {
                    BorderStyle = (BorderStyle)r.ReadByte();
                    dataBlockBytes += 1;
                }
                if (PropMask.HasMousePointer) {
                    MousePointer = (MousePointer)r.ReadByte();
                    dataBlockBytes += 1;
                }
                if (PropMask.HasScrollBars) {
                    ScrollBars = new FormScrollBarFlags(r.ReadByte());
                    dataBlockBytes += 1;
                }
                AlignTo(4, st, ref dataBlockBytes);
                if (PropMask.HasGroupCount) {
                    GroupCount = r.ReadInt32();
                    dataBlockBytes += 4;
                }
                if (PropMask.HasMouseIcon) IgnoreNext(2, st, ref dataBlockBytes);
                if (PropMask.HasCycle) {
                    Cycle = (Cycle)r.ReadByte();
                    dataBlockBytes += 1;
                }
                if (PropMask.HasSpecialEffect) {
                    SpecialEffect = (SpecialEffect)r.ReadByte();
                    dataBlockBytes += 1;
                }
                AlignTo(4, st, ref dataBlockBytes);
                if (PropMask.HasBorderColor) {
                    BorderColor = new OleColor(r.ReadBytes(4));
                    dataBlockBytes += 4;
                }
                var captionLength = 0;
                var captionCompressed = false;
                if (PropMask.HasCaption) {
                    captionLength = r.ReadInt32();
                    dataBlockBytes += 4;
                    if (captionLength < 0) {
                        captionLength = unchecked((int)(captionLength ^ 0x80000000));
                        captionCompressed = true;
                    }
                }
                if (PropMask.HasFont) IgnoreNext(2, st, ref dataBlockBytes);
                AlignTo(4, st, ref dataBlockBytes);
                if (PropMask.HasPicture) IgnoreNext(2, st, ref dataBlockBytes);
                AlignTo(4, st, ref dataBlockBytes);
                if (PropMask.HasZoom) {
                    Zoom = r.ReadUInt32();
                    dataBlockBytes += 4;
                }
                if (PropMask.HasPictureAlignment) {
                    PictureAlignment = (PictureAlignment)r.ReadByte();
                    dataBlockBytes += 1;
                }
                if (PropMask.HasPictureSizeMode) {
                    PictureSizeMode = (PictureSizeMode)r.ReadByte();
                    dataBlockBytes += 1;
                }
                AlignTo(4, st, ref dataBlockBytes);
                if (PropMask.HasShapeCookie) {
                    ShapeCookie = r.ReadUInt32();
                    dataBlockBytes += 4;
                }
                if (PropMask.HasDrawBuffer) {
                    DrawBuffer = r.ReadUInt32();
                    dataBlockBytes += 4;
                }

                // ExtraDataBlock
                ushort extraDataBlockBytes = 0;
                if (PropMask.HasDisplayedSize) {
                    DisplayedSize = Tuple.Create(r.ReadInt32(), r.ReadInt32());
                    extraDataBlockBytes += 8;
                }
                if (PropMask.HasLogicalSize) {
                    LogicalSize = Tuple.Create(r.ReadInt32(), r.ReadInt32());
                    extraDataBlockBytes += 8;
                }
                if (PropMask.HasScrollPosition) {
                    ScrollPosition = Tuple.Create(r.ReadInt32(), r.ReadInt32());
                    extraDataBlockBytes += 8;
                }
                if (captionLength > 0) {
                    Caption = (captionCompressed ? Encoding.UTF8 : Encoding.Unicode).GetString(r.ReadBytes(captionLength));
                    extraDataBlockBytes += (ushort)captionLength;
                }
                AlignTo(4, st, ref dataBlockBytes);

                if (cbForm != 4 + dataBlockBytes + extraDataBlockBytes) {
                    throw new ApplicationException("Error reading 'f' stream in .frx data: expected cbForm size "
                                                   + $"{4 + dataBlockBytes + extraDataBlockBytes}, but actual size was {cbForm}.");
                }

                // StreamData
                if (PropMask.HasMouseIcon) {
                    st.Seek(20, SeekOrigin.Current); // skip GUID and Preamble
                    MouseIcon = r.ReadBytes(r.ReadInt32());
                }
                if (PropMask.HasFont) {
                    var guid = new Guid(r.ReadBytes(16));
                    if (guid == new Guid(0x0BE35203, 0x8F91, 0x11CE, 0x9D, 0xE3, 0x00, 0xAA, 0x00, 0x4B, 0xB8, 0x51)) {
                        // StdFont
                        FontIsStdFont = true;
                        st.Seek(1, SeekOrigin.Current); // skip Version
                        FontStdFont = Tuple.Create(r.ReadInt16(), r.ReadByte(), r.ReadInt16(), r.ReadUInt32(), Encoding.ASCII.GetString(r.ReadBytes(r.ReadByte())));
                    } else {
                        // TextProps
                        st.Seek(2, SeekOrigin.Current); // skip MinorVersion and MajorVersion
                        FontTextProps = r.ReadBytes(r.ReadUInt16());
                    }
                }
                if (PropMask.HasPicture) {
                    st.Seek(20, SeekOrigin.Current); // skip GUID and Preamble
                    Picture = r.ReadBytes(r.ReadInt32());
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
                  && Picture.SequenceEqual(other.Picture) && FontTextProps.SequenceEqual(other.FontTextProps) && Equals(FontStdFont, other.FontStdFont)
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

        static void AlignTo(ushort alignment, Stream st, ref ushort bytesAccum) {
            if (bytesAccum % alignment == 0) return;
            st.Seek(alignment - bytesAccum % alignment, SeekOrigin.Current);
            bytesAccum += (ushort)(alignment - bytesAccum % alignment);
        }

        static void IgnoreNext(ushort bytes, Stream st, ref ushort bytesAccum) {
            st.Seek(bytes, SeekOrigin.Current);
            bytesAccum += bytes;
        }

        static void AlignTo(uint alignment, Stream st, ref uint bytesAccum) {
            if (bytesAccum % alignment == 0) return;
            st.Seek(alignment - bytesAccum % alignment, SeekOrigin.Current);
            bytesAccum += alignment - bytesAccum % alignment;
        }

        static void IgnoreNext(uint bytes, Stream st, ref uint bytesAccum) {
            st.Seek(bytes, SeekOrigin.Current);
            bytesAccum += bytes;
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
                    nameLength = r.ReadInt32();
                    dataBlockBytes += 4;
                    if (nameLength < 0) {
                        nameLength = unchecked((int)(nameLength ^ 0x80000000));
                        nameCompressed = true;
                    }
                }
                var tagLength = 0;
                var tagCompressed = false;
                if (PropMask.HasTag) {
                    tagLength = unchecked((int)(tagLength ^ 0x80000000));
                    dataBlockBytes += 4;
                    if (tagLength < 0) {
                        tagLength = -tagLength;
                        tagCompressed = true;
                    }
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
                    controlTipTextLength = r.ReadInt32();
                    dataBlockBytes += 4;
                    if (controlTipTextLength < 0) {
                        controlTipTextLength = unchecked((int)(controlTipTextLength ^ 0x80000000));
                        controlTipTextCompressed = true;
                    }
                }
                var runtimeLicKeyLength = 0;
                var runtimeLicKeyCompressed = false;
                if (PropMask.HasRuntimeLicKey) {
                    runtimeLicKeyLength = r.ReadInt32();
                    dataBlockBytes += 4;
                    if (runtimeLicKeyLength < 0) {
                        runtimeLicKeyLength = unchecked((int)(runtimeLicKeyLength ^ 0x80000000));
                        runtimeLicKeyCompressed = true;
                    }
                }
                var controlSourceLength = 0;
                var controlSourceCompressed = false;
                if (PropMask.HasControlSource) {
                    controlSourceLength = r.ReadInt32();
                    dataBlockBytes += 4;
                    if (controlSourceLength < 0) {
                        controlSourceLength = unchecked((int)(controlSourceLength ^ 0x80000000));
                        controlSourceCompressed = true;
                    }
                }
                var rowSourceLength = 0;
                var rowSourceCompressed = false;
                if (PropMask.HasRowSource) {
                    rowSourceLength = r.ReadInt32();
                    dataBlockBytes += 4;
                    if (rowSourceLength < 0) {
                        rowSourceLength = unchecked((int)(rowSourceLength ^ 0x80000000));
                        rowSourceCompressed = true;
                    }
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

        static void AlignTo(ushort alignment, Stream st, ref ushort bytesAccum) {
            if (bytesAccum % alignment == 0)
                return;
            st.Seek(alignment - bytesAccum % alignment, SeekOrigin.Current);
            bytesAccum += (ushort)(alignment - bytesAccum % alignment);
        }
    }

    enum Cycle {
        AllForms = 0x00,
        CurrentForm = 0x02
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

    class OleColor {
        public OleColorType ColorType { get; }
        public ushort PaletteIndex { get; }
        public byte Red { get; }
        public byte Blue { get; }
        public byte Green { get; }

        public OleColor(byte[] b) {
            if (b.Length != 4) throw new ArgumentException($"Error creating {nameof(OleColor)}. Expected 4 bytes but got {b.Length}.", nameof(b));
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
                hashCode = (hashCode*397) ^ Red.GetHashCode();
                hashCode = (hashCode*397) ^ Blue.GetHashCode();
                hashCode = (hashCode*397) ^ Green.GetHashCode();
                return hashCode;
            }
        }
    }

    class PropMask {
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

        public PropMask(uint i) {
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
            var other = o as PropMask;
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
