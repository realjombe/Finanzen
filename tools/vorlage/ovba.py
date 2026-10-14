# -*- coding: utf-8 -*-
"""
Erzeugt ein vbaProject.bin (MS-OVBA in einem OLE/CFB-Container) aus VBA-Quelltexten,
ohne Excel. Aufbau wie der LibreOffice-VBA-Export (Quelltext ohne P-Code, Excel
kompiliert beim Oeffnen neu).

    from ovba import build_vba_project
    data = build_vba_project(
        [("ThisWorkbook", "workbook", code), ("shBuchungen", "sheet", ""), ("modJson", "std", code)],
    )
"""
import random
import struct
import uuid

ENDOFCHAIN = 0xFFFFFFFE
FREESECT = 0xFFFFFFFF
FATSECT = 0xFFFFFFFD
NOSTREAM = 0xFFFFFFFF

STDOLE_LIBID = "*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\System32\\stdole2.tlb#OLE Automation"
VB_BASE = {
    "workbook": "0{00020819-0000-0000-C000-000000000046}",
    "sheet": "0{00020820-0000-0000-C000-000000000046}",
}


# --------------------------------------------------------------------------- MS-OVBA Kompression
def _compress_chunk(chunk):
    """Komprimiert bis zu 4096 Bytes nach MS-OVBA 2.4.1 (LZ77-Variante)."""
    n = len(chunk)
    out = bytearray()
    i = 0
    heads = {}
    while i < n:
        flag_pos = len(out)
        out.append(0)
        flags = 0
        for bit in range(8):
            if i >= n:
                break
            best_len = 0
            best_off = 0
            if i >= 1:
                bc = max(4, (i - 1).bit_length())      # ceil(log2(i)), mind. 4
                max_len = (0xFFFF >> bc) + 3
                max_off = 1 << bc
                key = chunk[i:i + 3]
                if len(key) == 3:
                    for p in reversed(heads.get(key, ())):
                        off = i - p
                        if off > max_off:
                            break
                        l = 0
                        lim = min(max_len, n - i)
                        while l < lim and chunk[p + l] == chunk[i + l]:
                            l += 1
                        if l > best_len:
                            best_len, best_off = l, off
                            if l == lim:
                                break
            if best_len >= 3:
                token = ((best_off - 1) << (16 - bc)) | (best_len - 3)
                out += struct.pack("<H", token)
                flags |= 1 << bit
                for k in range(i, i + best_len):
                    if k + 3 <= n:
                        heads.setdefault(chunk[k:k + 3], []).append(k)
                i += best_len
            else:
                out.append(chunk[i])
                if i + 3 <= n:
                    heads.setdefault(chunk[i:i + 3], []).append(i)
                i += 1
        out[flag_pos] = flags
    return bytes(out)


def compress(data):
    out = bytearray(b"\x01")
    pos = 0
    while pos < len(data):
        chunk = data[pos:pos + 4096]
        pos += 4096
        comp = _compress_chunk(chunk)
        if len(comp) >= 4096:
            raw = chunk + b"\x00" * (4096 - len(chunk))
            out += struct.pack("<H", 0x3000 | (4098 - 3)) + raw
        else:
            out += struct.pack("<H", 0x8000 | 0x3000 | (len(comp) + 2 - 3)) + comp
    return bytes(out)


def decompress(data):
    """Referenz-Dekompressor (zum Testen)."""
    if data[0] != 1:
        raise ValueError("kein komprimierter Container")
    out = bytearray()
    pos = 1
    while pos < len(data):
        hdr, = struct.unpack_from("<H", data, pos)
        pos += 2
        size = (hdr & 0x0FFF) + 3
        compressed = bool(hdr & 0x8000)
        end = pos + size - 2
        start = len(out)
        if not compressed:
            out += data[pos:pos + 4096]
            pos = end
            continue
        while pos < end:
            flags = data[pos]
            pos += 1
            for bit in range(8):
                if pos >= end:
                    break
                if flags & (1 << bit):
                    token, = struct.unpack_from("<H", data, pos)
                    pos += 2
                    d = len(out) - start
                    bc = max(4, (d - 1).bit_length())
                    length = (token & (0xFFFF >> bc)) + 3
                    off = (token >> (16 - bc)) + 1
                    for _ in range(length):
                        out.append(out[-off])
                else:
                    out.append(data[pos])
                    pos += 1
    return bytes(out)


# --------------------------------------------------------------------------- Verschluesselung (PROJECT: CMG/DPB/GC)
def _encrypt(data, proj_key, rnd):
    seed = rnd.randrange(256)
    version = 2
    out = [seed, seed ^ version, seed ^ proj_key]
    e1 = seed ^ proj_key
    e2 = seed ^ version
    u1 = proj_key
    ignored = (seed & 6) >> 1
    for _ in range(ignored):
        t = rnd.randrange(256)
        b = t ^ ((e2 + u1) & 0xFF)
        out.append(b)
        e2, e1, u1 = e1, b, t
    for b0 in struct.pack("<I", len(data)) + bytes(data):
        b = b0 ^ ((e2 + u1) & 0xFF)
        out.append(b)
        e2, e1, u1 = e1, b, b0
    return "".join("%02X" % b for b in out)


def _proj_key(project_id):
    k = 0
    for ch in project_id:
        k = (k + ord(ch)) & 0xFF
    return k


# --------------------------------------------------------------------------- dir-Stream
def _rec(rid, data):
    return struct.pack("<HI", rid, len(data)) + data


def _build_dir(modules, project_name, codepage, lcid):
    enc = "cp%d" % codepage
    d = bytearray()
    d += _rec(0x0001, struct.pack("<I", 1))            # SYSKIND Win32
    d += _rec(0x0002, struct.pack("<I", lcid))         # LCID
    d += _rec(0x0014, struct.pack("<I", lcid))         # LCIDINVOKE
    d += _rec(0x0003, struct.pack("<H", codepage))     # CODEPAGE
    d += _rec(0x0004, project_name.encode(enc))        # NAME
    d += _rec(0x0005, b"") + _rec(0x0040, b"")         # DOCSTRING
    d += _rec(0x0006, b"") + _rec(0x003D, b"")         # HELPFILEPATH
    d += _rec(0x0007, struct.pack("<I", 0))            # HELPCONTEXT
    d += _rec(0x0008, struct.pack("<I", 0))            # LIBFLAGS
    d += struct.pack("<HIIH", 0x0009, 4, 1, 0)         # VERSION
    d += _rec(0x000C, b"") + _rec(0x003C, b"")         # CONSTANTS
    d += _rec(0x0016, b"stdole") + _rec(0x003E, "stdole".encode("utf-16-le"))
    libid = STDOLE_LIBID.encode(enc)
    d += _rec(0x000D, struct.pack("<I", len(libid)) + libid + b"\x00\x00\x00\x00\x00\x00")
    d += _rec(0x000F, struct.pack("<H", len(modules)))
    d += _rec(0x0013, struct.pack("<H", 0xFFFF))
    for name, kind, _src in modules:
        nm = name.encode(enc)
        nu = name.encode("utf-16-le")
        d += _rec(0x0019, nm) + _rec(0x0047, nu)       # MODULENAME
        d += _rec(0x001A, nm) + _rec(0x0032, nu)       # MODULESTREAMNAME
        d += _rec(0x001C, b"") + _rec(0x0048, b"")     # MODULEDOCSTRING
        d += _rec(0x0031, struct.pack("<I", 0))        # MODULEOFFSET
        d += _rec(0x001E, struct.pack("<I", 0))        # MODULEHELPCONTEXT
        d += _rec(0x002C, struct.pack("<H", 0xFFFF))   # MODULECOOKIE
        d += _rec(0x0022 if kind in ("workbook", "sheet", "class") else 0x0021, b"")
        d += _rec(0x002B, b"")                         # MODULETERMINATOR
    d += _rec(0x0010, b"")                             # PROJECTTERMINATOR
    return bytes(d)


def _module_source(name, kind, src):
    lines = []
    if kind in ("workbook", "sheet"):
        lines.append('Attribute VB_Name = "%s"' % name)
        lines.append('Attribute VB_Base = "%s"' % VB_BASE[kind])
        lines.append("Attribute VB_GlobalNameSpace = False")
        lines.append("Attribute VB_Creatable = False")
        lines.append("Attribute VB_PredeclaredId = True")
        lines.append("Attribute VB_Exposed = True")
        lines.append("Attribute VB_TemplateDerived = False")
        lines.append("Attribute VB_Customizable = True")
    elif not src.lstrip().startswith("Attribute VB_Name"):
        lines.append('Attribute VB_Name = "%s"' % name)
    body = src.replace("\r\n", "\n").replace("\r", "\n")
    text = "\r\n".join(lines)
    if lines and body:
        text += "\r\n"
    text += body.replace("\n", "\r\n")
    if text and not text.endswith("\r\n"):
        text += "\r\n"
    return text


def _build_project_stream(modules, project_name, project_id, rnd):
    key = _proj_key(project_id)
    lines = ['ID="%s"' % project_id]
    for name, kind, _ in modules:
        if kind in ("workbook", "sheet"):
            lines.append("Document=%s/&H00000000" % name)
    for name, kind, _ in modules:
        if kind == "std":
            lines.append("Module=%s" % name)
    for name, kind, _ in modules:
        if kind == "class":
            lines.append("Class=%s" % name)
    lines.append('Name="%s"' % project_name)
    lines.append('HelpContextID="0"')
    lines.append('VersionCompatible32="393222000"')
    lines.append('CMG="%s"' % _encrypt(struct.pack("<I", 0), key, rnd))
    lines.append('DPB="%s"' % _encrypt(b"\x00", key, rnd))
    lines.append('GC="%s"' % _encrypt(b"\xFF", key, rnd))
    lines.append("")
    lines.append("[Host Extender Info]")
    lines.append("&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000")
    lines.append("")
    lines.append("[Workspace]")
    for name, _, _ in modules:
        lines.append("%s=0, 0, 0, 0, C" % name)
    return "\r\n".join(lines) + "\r\n"


def _build_projectwm(modules, enc):
    out = bytearray()
    for name, _, _ in modules:
        out += name.encode(enc) + b"\x00" + name.encode("utf-16-le") + b"\x00\x00"
    out += b"\x00\x00"
    return bytes(out)


# --------------------------------------------------------------------------- CFB (OLE2) Writer
def _cfb_name_key(name):
    return (len(name), name.upper())


class _Entry:
    __slots__ = ("name", "etype", "data", "children", "left", "right", "child", "start", "size", "eid", "color")

    def __init__(self, name, etype, data=b"", children=None):
        self.name = name
        self.etype = etype          # 5 root, 1 storage, 2 stream
        self.data = data
        self.children = children or []
        self.left = self.right = self.child = NOSTREAM
        self.start = ENDOFCHAIN
        self.size = 0
        self.eid = 0
        self.color = 1              # schwarz


def _build_bst(entries):
    """Ausgeglichener Suchbaum ueber die Geschwister; liefert Wurzel-ID."""
    if not entries:
        return NOSTREAM
    entries = sorted(entries, key=lambda e: _cfb_name_key(e.name))
    mid = len(entries) // 2
    root = entries[mid]
    root.left = _build_bst(entries[:mid])
    root.right = _build_bst(entries[mid + 1:])
    return root.eid


def write_cfb(tree):
    """tree: dict name -> bytes (Stream) oder dict (Storage). Liefert die Datei als bytes."""
    root = _Entry("Root Entry", 5)

    def build(node_dict):
        res = []
        for name, val in node_dict.items():
            if isinstance(val, dict):
                e = _Entry(name, 1, children=build(val))
            else:
                e = _Entry(name, 2, data=bytes(val))
            res.append(e)
        return res

    root.children = build(tree)
    # IDs in Dokumentreihenfolge (Root = 0, dann Tiefensuche)
    all_entries = []

    def collect(e):
        e.eid = len(all_entries)
        all_entries.append(e)
        for c in e.children:
            collect(c)

    collect(root)
    for e in all_entries:
        if e.children:
            e.child = _build_bst(e.children)

    SEC = 512
    MINI = 64
    streams = [e for e in all_entries if e.etype == 2]
    mini_streams = [e for e in streams if 0 < len(e.data) < 4096]
    big_streams = [e for e in streams if len(e.data) >= 4096]

    # Ministream zusammenbauen
    mini = bytearray()
    mini_fat = []
    for e in mini_streams:
        e.start = len(mini) // MINI
        e.size = len(e.data)
        padded = e.data + b"\x00" * ((-len(e.data)) % MINI)
        n = len(padded) // MINI
        mini += padded
        for i in range(n):
            mini_fat.append(e.start + i + 1 if i < n - 1 else ENDOFCHAIN)
    for e in streams:
        if len(e.data) == 0:
            e.start = ENDOFCHAIN
            e.size = 0
    for e in big_streams:
        e.size = len(e.data)

    n_dir = len(all_entries)
    dir_sectors = (n_dir * 128 + SEC - 1) // SEC
    minifat_sectors = (len(mini_fat) * 4 + SEC - 1) // SEC
    mini_sectors = (len(mini) + SEC - 1) // SEC
    big_sector_counts = [(len(e.data) + SEC - 1) // SEC for e in big_streams]
    body = dir_sectors + minifat_sectors + mini_sectors + sum(big_sector_counts)
    fat_sectors = 1
    while (body + fat_sectors) > fat_sectors * (SEC // 4):
        fat_sectors += 1
    total = body + fat_sectors
    if fat_sectors > 109:
        raise ValueError("Datei zu gross fuer diesen einfachen CFB-Writer")

    fat = [FREESECT] * (fat_sectors * (SEC // 4))
    sectors = {}
    nxt = 0
    for i in range(fat_sectors):
        fat[nxt] = FATSECT
        nxt += 1

    def chain(blob, count):
        nonlocal nxt
        first = nxt if count else ENDOFCHAIN
        for i in range(count):
            sectors[nxt] = blob[i * SEC:(i + 1) * SEC]
            fat[nxt] = nxt + 1 if i < count - 1 else ENDOFCHAIN
            nxt += 1
        return first

    dir_first = nxt                      # Platzhalter, Inhalt kommt spaeter
    nxt += dir_sectors
    for i in range(dir_sectors):
        fat[dir_first + i] = dir_first + i + 1 if i < dir_sectors - 1 else ENDOFCHAIN

    minifat_blob = b"".join(struct.pack("<I", v) for v in mini_fat)
    minifat_blob += b"\xFF" * ((-len(minifat_blob)) % SEC)
    minifat_first = chain(minifat_blob, minifat_sectors)

    mini_blob = bytes(mini) + b"\x00" * ((-len(mini)) % SEC)
    mini_first = chain(mini_blob, mini_sectors)
    root.start = mini_first
    root.size = len(mini)

    for e, cnt in zip(big_streams, big_sector_counts):
        blob = e.data + b"\x00" * ((-len(e.data)) % SEC)
        e.start = chain(blob, cnt)

    assert nxt == total

    # Verzeichnis
    dir_blob = bytearray()
    for e in all_entries:
        nm = e.name.encode("utf-16-le") + b"\x00\x00"
        if len(nm) > 64:
            raise ValueError("Name zu lang: " + e.name)
        rec = bytearray(128)
        rec[0:len(nm)] = nm
        struct.pack_into("<H", rec, 64, len(nm))
        rec[66] = e.etype
        rec[67] = e.color
        struct.pack_into("<III", rec, 68, e.left, e.right, e.child)
        # CLSID 80..96, StateBits 96..100, Zeiten 100..116 bleiben 0
        start = e.start if e.etype in (2, 5) else 0
        if e.etype == 5 and mini_sectors == 0:
            start = ENDOFCHAIN
        struct.pack_into("<I", rec, 116, start)
        struct.pack_into("<Q", rec, 120, e.size if e.etype in (2, 5) else 0)
        dir_blob += rec
    while len(dir_blob) < dir_sectors * SEC:
        rec = bytearray(128)
        struct.pack_into("<III", rec, 68, NOSTREAM, NOSTREAM, NOSTREAM)
        dir_blob += rec
    for i in range(dir_sectors):
        sectors[dir_first + i] = bytes(dir_blob[i * SEC:(i + 1) * SEC])

    for i in range(fat_sectors):
        part = fat[i * (SEC // 4):(i + 1) * (SEC // 4)]
        sectors[i] = b"".join(struct.pack("<I", v) for v in part)

    hdr = bytearray(SEC)
    hdr[0:8] = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"
    struct.pack_into("<HHHHH", hdr, 24, 0x003E, 0x0003, 0xFFFE, 9, 6)
    struct.pack_into("<I", hdr, 40, 0)                 # Anzahl Verzeichnissektoren (v3: 0)
    struct.pack_into("<I", hdr, 44, fat_sectors)
    struct.pack_into("<I", hdr, 48, dir_first)
    struct.pack_into("<I", hdr, 52, 0)
    struct.pack_into("<I", hdr, 56, 4096)
    struct.pack_into("<I", hdr, 60, minifat_first if minifat_sectors else ENDOFCHAIN)
    struct.pack_into("<I", hdr, 64, minifat_sectors)
    struct.pack_into("<I", hdr, 68, ENDOFCHAIN)
    struct.pack_into("<I", hdr, 72, 0)
    for i in range(109):
        struct.pack_into("<I", hdr, 76 + 4 * i, i if i < fat_sectors else FREESECT)

    out = bytes(hdr)
    for s in range(total):
        sec = sectors.get(s, b"\x00" * SEC)
        assert len(sec) == SEC, (s, len(sec))
        out += sec
    return out


# --------------------------------------------------------------------------- oeffentlich
def build_vba_project(modules, project_name="VBAProject", codepage=1252, lcid=0x0409, project_id=None, seed=None):
    """
    modules: Liste (name, kind, source) mit kind in 'workbook' | 'sheet' | 'std' | 'class'.
             Dokumentmodule (workbook/sheet) zuerst in der Reihenfolge der Blaetter.
    """
    enc = "cp%d" % codepage
    rnd = random.Random(seed)
    if project_id is None:
        project_id = "{%s}" % str(uuid.UUID(int=rnd.getrandbits(128), version=4)).upper()
    vba = {"_VBA_PROJECT": b"\xCC\x61\xFF\xFF\x00\x00\x00", "dir": compress(_build_dir(modules, project_name, codepage, lcid))}
    for name, kind, src in modules:
        text = _module_source(name, kind, src)
        vba[name] = compress(text.encode(enc))
    tree = {
        "PROJECT": _build_project_stream(modules, project_name, project_id, rnd).encode(enc),
        "PROJECTwm": _build_projectwm(modules, enc),
        "VBA": vba,
    }
    return write_cfb(tree)


def project_stream_text(data, codepage=1252):
    import olefile
    o = olefile.OleFileIO(data)
    return o.openstream("PROJECT").read().decode("cp%d" % codepage)
