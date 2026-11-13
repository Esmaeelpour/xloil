"""
    Writes a VBA project (vbaProject.bin) from VBA source text, without Excel.

    Follows Microsoft's published formats: [MS-CFB] for the compound file and
    [MS-OVBA] for the VBA streams. Only source code is stored (no compiled
    p-code): the project version is 0xFFFF, so Office compiles the source when it
    first opens the file. This lets xlOil.xlam be built on any computer, including
    ones whose policy locks "Trust access to the VBA project object model".
"""
import struct
import uuid

# ---------------------------------------------------------------------------
# [MS-OVBA] 2.4.1 compression
# ---------------------------------------------------------------------------

def _copy_token_bits(difference: int) -> int:
    bits = 4
    while (1 << bits) < difference:
        bits += 1
    return bits


def _compress_chunk(data: bytes) -> bytes:
    out = bytearray()
    pos = 0
    while pos < len(data):
        flag_pos = len(out)
        out.append(0)
        flags = 0
        for bit in range(8):
            if pos >= len(data):
                break
            bits = _copy_token_bits(pos) if pos else 4
            max_len = (0xFFFF >> bits) + 3
            best_len, best_off = 0, 0
            for start in range(pos - 1, -1, -1):
                length = 0
                limit = min(max_len, len(data) - pos)
                while length < limit and data[start + length] == data[pos + length]:
                    length += 1
                if length > best_len:
                    best_len, best_off = length, pos - start
                    if length == limit:
                        break
            if best_len >= 3:
                token = ((best_off - 1) << (16 - bits)) | (best_len - 3)
                out += struct.pack("<H", token)
                flags |= 1 << bit
                pos += best_len
            else:
                out.append(data[pos])
                pos += 1
        out[flag_pos] = flags
    return bytes(out)


def compress(data: bytes) -> bytes:
    out = bytearray(b"\x01")
    for i in range(0, len(data), 4096):
        chunk = data[i:i + 4096]
        body = _compress_chunk(chunk)
        if len(body) > 4096:
            if len(chunk) < 4096:
                raise ValueError("Incompressible final chunk")
            out += struct.pack("<H", 0x3000 | 0x0FFF) + chunk  # raw chunk
        else:
            out += struct.pack("<H", 0x8000 | 0x3000 | (len(body) + 2 - 3)) + body
    return bytes(out)


def decompress(data: bytes) -> bytes:
    """ Inverse of compress, used to self-check the output """
    assert data[0] == 1
    out = bytearray()
    pos = 1
    while pos < len(data):
        header, = struct.unpack_from("<H", data, pos)
        size = (header & 0x0FFF) + 3
        end = pos + size
        pos += 2
        start = len(out)
        if not header & 0x8000:
            out += data[pos:pos + 4096]
            pos += 4096
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
                    bits = _copy_token_bits(len(out) - start)
                    length = (token & (0xFFFF >> bits)) + 3
                    offset = (token >> (16 - bits)) + 1
                    for _ in range(length):
                        out.append(out[-offset])
                else:
                    out.append(data[pos])
                    pos += 1
    return bytes(out)


# ---------------------------------------------------------------------------
# [MS-OVBA] 2.4.3 data encryption (for the PROJECT stream's CMG/DPB/GC)
# ---------------------------------------------------------------------------

def _encrypt(project_id: str, data: bytes, seed: int = 0x07) -> str:
    version = 2
    proj_key = sum(project_id.encode("ascii")) & 0xFF
    ignored_length = (seed & 6) // 2
    out = bytearray([seed, seed ^ version, seed ^ proj_key])
    unencrypted1, encrypted1, encrypted2 = proj_key, out[2], out[1]

    def put(byte):
        nonlocal unencrypted1, encrypted1, encrypted2
        enc = byte ^ ((encrypted2 + unencrypted1) & 0xFF)
        out.append(enc)
        encrypted2, encrypted1, unencrypted1 = encrypted1, enc, byte

    for _ in range(ignored_length):
        put(0)
    for b in struct.pack("<I", len(data)) + data:
        put(b)
    return out.hex().upper()


def _decrypt(text: str) -> bytes:
    enc = bytes.fromhex(text)
    seed, version_enc, proj_key_enc = enc[0], enc[1], enc[2]
    assert seed ^ version_enc == 2
    unencrypted1, encrypted1, encrypted2 = seed ^ proj_key_enc, proj_key_enc, version_enc
    plain = bytearray()
    for e in enc[3:]:
        b = e ^ ((encrypted2 + unencrypted1) & 0xFF)
        plain.append(b)
        encrypted2, encrypted1, unencrypted1 = encrypted1, e, b
    ignored = (seed & 6) // 2
    length, = struct.unpack_from("<I", plain, ignored)
    return bytes(plain[ignored + 4:ignored + 4 + length])


# ---------------------------------------------------------------------------
# VBA streams
# ---------------------------------------------------------------------------

DOCUMENT_BASES = {
    "workbook": "0{00020819-0000-0000-C000-000000000046}",
    "worksheet": "0{00020820-0000-0000-C000-000000000046}",
}

CODEPAGE = 1252


def _rec(rec_id: int, payload: bytes) -> bytes:
    return struct.pack("<HI", rec_id, len(payload)) + payload


def _mbcs(s: str) -> bytes:
    return s.encode("cp1252")


def _utf16(s: str) -> bytes:
    return s.encode("utf-16-le")


def _document_source(name: str, kind: str) -> str:
    return "\r\n".join([
        f'Attribute VB_Name = "{name}"',
        f'Attribute VB_Base = "{DOCUMENT_BASES[kind]}"',
        "Attribute VB_GlobalNameSpace = False",
        "Attribute VB_Creatable = False",
        "Attribute VB_PredeclaredId = True",
        "Attribute VB_Exposed = True",
        "Attribute VB_TemplateDerived = False",
        "Attribute VB_Customizable = True",
        "",
    ])


def _dir_stream(modules) -> bytes:
    d = bytearray()
    d += _rec(0x0001, struct.pack("<I", 3))               # PROJECTSYSKIND: 64-bit Windows
    d += _rec(0x0002, struct.pack("<I", 0x409))           # PROJECTLCID
    d += _rec(0x0014, struct.pack("<I", 0x409))           # PROJECTLCIDINVOKE
    d += _rec(0x0003, struct.pack("<H", CODEPAGE))        # PROJECTCODEPAGE
    d += _rec(0x0004, b"VBAProject")                      # PROJECTNAME
    d += _rec(0x0005, b"") + _rec(0x0040, b"")            # PROJECTDOCSTRING
    d += _rec(0x0006, b"") + _rec(0x003D, b"")            # PROJECTHELPFILEPATH
    d += _rec(0x0007, struct.pack("<I", 0))               # PROJECTHELPCONTEXT
    d += _rec(0x0008, struct.pack("<I", 0))               # PROJECTLIBFLAGS
    d += struct.pack("<HIIH", 0x0009, 4, 1, 0)            # PROJECTVERSION
    d += _rec(0x000C, b"") + _rec(0x003C, b"")            # PROJECTCONSTANTS

    # Reference to OLE Automation (stdole), as Excel's own projects have
    libid = rb"*\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\Windows\System32\stdole2.tlb#OLE Automation"
    d += _rec(0x0016, b"stdole") + _rec(0x003E, _utf16("stdole"))
    d += _rec(0x000D, struct.pack("<I", len(libid)) + libid + struct.pack("<IH", 0, 0))

    d += _rec(0x000F, struct.pack("<H", len(modules)))    # PROJECTMODULES
    d += _rec(0x0013, struct.pack("<H", 0xFFFF))          # PROJECTCOOKIE
    for name, kind, _ in modules:
        d += _rec(0x0019, _mbcs(name))                    # MODULENAME
        d += _rec(0x0047, _utf16(name))                   # MODULENAMEUNICODE
        d += _rec(0x001A, _mbcs(name)) + _rec(0x0032, _utf16(name))  # MODULESTREAMNAME
        d += _rec(0x001C, b"") + _rec(0x0048, b"")        # MODULEDOCSTRING
        d += _rec(0x0031, struct.pack("<I", 0))           # MODULEOFFSET: source at the start
        d += _rec(0x001E, struct.pack("<I", 0))           # MODULEHELPCONTEXT
        d += _rec(0x002C, struct.pack("<H", 0xFFFF))      # MODULECOOKIE
        d += _rec(0x0021 if kind == "module" else 0x0022, b"")  # MODULETYPE
        d += _rec(0x002B, b"")                            # module terminator
    d += _rec(0x0010, b"")                                # dir terminator
    return bytes(d)


def _project_stream(project_id: str, modules) -> bytes:
    lines = [f'ID="{project_id}"']
    for name, kind, _ in modules:
        lines.append(f"Document={name}/&H00000000" if kind != "module" else f"Module={name}")
    lines += [
        'Name="VBAProject"',
        'HelpContextID="0"',
        'VersionCompatible32="393222000"',
        'CMG="' + _encrypt(project_id, struct.pack("<I", 0), 0x07) + '"',
        'DPB="' + _encrypt(project_id, bytes([0x00]), 0x0B) + '"',
        'GC="' + _encrypt(project_id, bytes([0xFF]), 0x0D) + '"',
        "",
        "[Host Extender Info]",
        "&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000",
        "",
    ]
    return _mbcs("\r\n".join(lines))


def _projectwm_stream(modules) -> bytes:
    out = bytearray()
    for name, _, _ in modules:
        out += _mbcs(name) + b"\x00" + _utf16(name) + b"\x00\x00"
    return bytes(out + b"\x00\x00")


def build_vba_project(module_source: dict, documents=(("ThisWorkbook", "workbook"), ("Sheet1", "worksheet")),
                      project_id: str = None) -> bytes:
    """
        module_source: {module name: VBA source, without the Attribute VB_Name line}.
        documents: the workbook's document modules: (code name, "workbook"|"worksheet").
        Returns the bytes of vbaProject.bin.
    """
    project_id = project_id or "{" + str(uuid.uuid5(uuid.NAMESPACE_URL, "xloil.xlam")).upper() + "}"
    modules = [(name, kind, _document_source(name, kind)) for name, kind in documents]
    for name, src in module_source.items():
        body = src.replace("\r\n", "\n").replace("\n", "\r\n")
        modules.append((name, "module", f'Attribute VB_Name = "{name}"\r\n' + body))

    vba = {
        "_VBA_PROJECT": struct.pack("<HHBH", 0x61CC, 0xFFFF, 0x00, 0x0000),
        "dir": compress(_dir_stream(modules)),
    }
    for name, _, src in modules:
        vba[name] = compress(src.encode("cp1252"))
    tree = {
        "VBA": vba,
        "PROJECT": _project_stream(project_id, modules),
        "PROJECTwm": _projectwm_stream(modules),
    }
    return write_cfb(tree)


# ---------------------------------------------------------------------------
# [MS-CFB] compound file writer (version 3, 512-byte sectors)
# ---------------------------------------------------------------------------

FREESECT, ENDOFCHAIN, FATSECT, NOSTREAM = 0xFFFFFFFF, 0xFFFFFFFE, 0xFFFFFFFD, 0xFFFFFFFF
SECTOR, MINI, CUTOFF = 512, 64, 4096


class _Entry:
    def __init__(self, name, kind, data=None, children=None):
        self.name, self.kind, self.data = name, kind, data   # kind: 1 storage, 2 stream, 5 root
        self.children = children or []
        self.left = self.right = self.child = NOSTREAM
        self.color = 1
        self.start, self.size = ENDOFCHAIN, 0
        self.sid = None


def _sort_key(e):
    return (len(e.name), e.name.upper())


def write_cfb(tree: dict) -> bytes:
    def make(name, value):
        if isinstance(value, dict):
            return _Entry(name, 1, children=[make(k, v) for k, v in value.items()])
        return _Entry(name, 2, data=value)

    root = _Entry("Root Entry", 5, children=[make(k, v) for k, v in tree.items()])
    entries = []

    def number(e):
        e.sid = len(entries)
        entries.append(e)
        for c in e.children:
            number(c)
    number(root)

    def colour_balance(nodes):
        """ red-black colouring valid for a tree built by _build_tree """
        n = len(nodes)
        if n == 0:
            return NOSTREAM
        height = n.bit_length()          # levels needed
        complete = (1 << height) - 1 == n

        def build(sub, depth):
            if not sub:
                return NOSTREAM
            mid = len(sub) // 2
            node = sub[mid]
            node.color = 0 if (not complete and depth == height - 1) else 1
            node.left = build(sub[:mid], depth + 1)
            node.right = build(sub[mid + 1:], depth + 1)
            return node.sid
        return build(nodes, 0)

    for e in entries:
        if e.children:
            e.child = colour_balance(sorted(e.children, key=_sort_key))

    # Mini stream for small streams
    mini = bytearray()
    minifat = []
    big = []
    for e in entries:
        if e.kind != 2:
            continue
        e.size = len(e.data)
        if e.size == 0:
            e.start = ENDOFCHAIN
        elif e.size < CUTOFF:
            e.start = len(mini) // MINI
            n = -(-e.size // MINI)
            minifat += [e.start + i + 1 for i in range(n - 1)] + [ENDOFCHAIN]
            mini += e.data + b"\x00" * (n * MINI - e.size)
        else:
            big.append(e)

    def sectors(nbytes):
        return -(-nbytes // SECTOR)

    dir_bytes = len(entries) * 128
    n_dir = sectors(dir_bytes)
    n_minifat = sectors(len(minifat) * 4)
    n_mini = sectors(len(mini))
    n_big = sum(sectors(e.size) for e in big)
    n_fat = 1
    while n_fat * (SECTOR // 4) < n_fat + n_dir + n_minifat + n_mini + n_big:
        n_fat += 1
    if n_fat > 109:
        raise ValueError("File too big for this simple writer")

    fat = []
    body = bytearray()

    def alloc(data: bytes, count: int):
        first = len(fat)
        for i in range(count):
            fat.append(first + i + 1 if i < count - 1 else ENDOFCHAIN)
        body.extend(data + b"\x00" * (count * SECTOR - len(data)))
        return first if count else ENDOFCHAIN

    fat += [FATSECT] * n_fat
    body.extend(b"\x00" * (n_fat * SECTOR))          # filled in below
    first_dir = len(fat)
    fat += [first_dir + i + 1 for i in range(n_dir - 1)] + [ENDOFCHAIN]
    body.extend(b"\x00" * (n_dir * SECTOR))          # filled in below
    first_minifat = alloc(b"".join(struct.pack("<I", x) for x in minifat), n_minifat)
    root.start = alloc(bytes(mini), n_mini)
    root.size = len(mini)
    for e in big:
        e.start = alloc(e.data, sectors(e.size))

    fat_bytes = b"".join(struct.pack("<I", x) for x in fat)
    fat_bytes += struct.pack("<I", FREESECT) * (n_fat * (SECTOR // 4) - len(fat))
    body[0:n_fat * SECTOR] = fat_bytes

    d = bytearray()
    for e in entries:
        name = _utf16(e.name) + b"\x00\x00"
        d += name.ljust(64, b"\x00")
        d += struct.pack("<HBB", len(name), e.kind, e.color)
        d += struct.pack("<III", e.left, e.right, e.child)
        d += b"\x00" * 16 + struct.pack("<I", 0) + b"\x00" * 16
        d += struct.pack("<IQ", e.start, e.size)
    while len(d) % SECTOR:
        d += (b"\x00" * 64 + struct.pack("<HBB", 0, 0, 0) + struct.pack("<III", NOSTREAM, NOSTREAM, NOSTREAM)
              + b"\x00" * 36 + struct.pack("<IQ", 0, 0))
    body[n_fat * SECTOR:(n_fat + n_dir) * SECTOR] = d

    header = bytearray()
    header += bytes.fromhex("D0CF11E0A1B11AE1") + b"\x00" * 16
    header += struct.pack("<HHHHH", 0x003E, 0x0003, 0xFFFE, 9, 6) + b"\x00" * 6
    header += struct.pack("<IIII", 0, n_fat, first_dir, 0)
    header += struct.pack("<IIIII", CUTOFF, first_minifat if n_minifat else ENDOFCHAIN, n_minifat, ENDOFCHAIN, 0)
    difat = list(range(n_fat)) + [FREESECT] * (109 - n_fat)
    header += b"".join(struct.pack("<I", x) for x in difat)
    assert len(header) == 512
    return bytes(header + body)
