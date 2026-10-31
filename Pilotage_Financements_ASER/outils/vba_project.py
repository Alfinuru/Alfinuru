"""Construction d'un fichier vbaProject.bin (MS-OVBA) dans un conteneur OLE (MS-CFB).

Le projet ne contient que le code source compressé (pas de cache p-code) :
la version 0xFFFF du flux _VBA_PROJECT force Office à recompiler le code
source à l'ouverture, quelle que soit la version / l'architecture d'Excel.
"""
import random
import struct
import uuid

CODEPAGE = 1252
ENDOFCHAIN, FREESECT, FATSECT, NOSTREAM = 0xFFFFFFFE, 0xFFFFFFFF, 0xFFFFFFFD, 0xFFFFFFFF


# ----------------------------------------------------------------------------
# MS-OVBA 2.4.1 : compression
# ----------------------------------------------------------------------------
def _copy_token_help(pos):
    bit_count = 4
    while (1 << bit_count) < pos:
        bit_count += 1
    length_mask = 0xFFFF >> bit_count
    return bit_count, length_mask, length_mask + 3


def _compress_chunk(data):
    out = bytearray()
    pos = 0
    n = len(data)
    index = {}  # préfixe de 3 octets -> positions déjà vues (fenêtre = début du bloc)

    def remember(p):
        if p + 3 <= n:
            index.setdefault(data[p:p + 3], []).append(p)

    while pos < n:
        flag_index = len(out)
        out.append(0)
        flags = 0
        for bit in range(8):
            if pos >= n:
                break
            best_len, best_off = 0, 0
            if pos > 0 and pos + 3 <= n:
                _, _, max_len = _copy_token_help(pos)
                limit = min(max_len, n - pos)
                for cand in reversed(index.get(data[pos:pos + 3], ())):
                    l = 3
                    while l < limit and data[cand + l] == data[pos + l]:
                        l += 1
                    if l > best_len:
                        best_len, best_off = l, pos - cand
                        if l == limit:
                            break
            if best_len >= 3:
                bit_count, _, _ = _copy_token_help(pos)
                token = ((best_off - 1) << (16 - bit_count)) | (best_len - 3)
                out += struct.pack('<H', token)
                flags |= 1 << bit
                for p in range(pos, pos + best_len):
                    remember(p)
                pos += best_len
            else:
                out.append(data[pos])
                remember(pos)
                pos += 1
        out[flag_index] = flags
    return bytes(out)


def compress(data: bytes) -> bytes:
    out = bytearray(b'\x01')
    for i in range(0, max(len(data), 1), 4096):
        chunk = data[i:i + 4096]
        if not chunk:
            break
        comp = _compress_chunk(chunk)
        if len(comp) > 4096 and len(chunk) == 4096:
            out += struct.pack('<H', (4096 + 2 - 3) | (0b011 << 12))  # bloc brut
            out += chunk
        else:
            assert len(comp) <= 4096, "bloc final non compressible"
            out += struct.pack('<H', (len(comp) + 2 - 3) | (0b011 << 12) | 0x8000)
            out += comp
    return bytes(out)


# ----------------------------------------------------------------------------
# MS-OVBA 2.4.3 : chiffrement des propriétés CMG / DPB / GC du flux PROJECT
# ----------------------------------------------------------------------------
def _encrypt(data: bytes, project_id: str, rnd: random.Random) -> str:
    seed = rnd.randrange(256)
    version = 2
    proj_key = sum(project_id.encode('ascii')) & 0xFF
    version_enc = seed ^ version
    proj_key_enc = seed ^ proj_key
    out = bytearray([seed, version_enc, proj_key_enc])
    unenc1, enc1, enc2 = proj_key, proj_key_enc, version_enc
    ignored = (seed & 6) // 2
    payload = bytes(rnd.randrange(256) for _ in range(ignored)) + struct.pack('<I', len(data)) + data
    for b in payload:
        e = b ^ ((enc2 + unenc1) & 0xFF)
        out.append(e)
        enc2, enc1, unenc1 = enc1, e, b
    return out.hex().upper()


# ----------------------------------------------------------------------------
# MS-OVBA 2.3.4.2 : flux dir
# ----------------------------------------------------------------------------
def _rec(rid, payload):
    return struct.pack('<HI', rid, len(payload)) + payload


def _mbcs(s):
    return s.encode('cp1252')


def _u16(s):
    return s.encode('utf-16-le')


def build_dir(modules):
    d = bytearray()
    d += _rec(0x0001, struct.pack('<I', 1))          # SYSKIND Win32 (recompilé)
    d += _rec(0x0002, struct.pack('<I', 0x409))      # LCID
    d += _rec(0x0014, struct.pack('<I', 0x409))      # LCIDINVOKE
    d += _rec(0x0003, struct.pack('<H', CODEPAGE))   # CODEPAGE
    d += _rec(0x0004, b'VBAProject')                 # NAME
    d += _rec(0x0005, b'') + _rec(0x0040, b'')       # DOCSTRING
    d += _rec(0x0006, b'') + _rec(0x003D, b'')       # HELPFILEPATH
    d += _rec(0x0007, struct.pack('<I', 0))          # HELPCONTEXT
    d += _rec(0x0008, struct.pack('<I', 0))          # LIBFLAGS
    d += struct.pack('<HIIH', 0x0009, 4, 1, 0)       # VERSION
    d += _rec(0x000C, b'') + _rec(0x003C, b'')       # CONSTANTS
    # Référence stdole (OLE Automation)
    libid = b'*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\System32\\stdole2.tlb#OLE Automation'
    d += _rec(0x0016, b'stdole') + _rec(0x003E, _u16('stdole'))
    d += _rec(0x000D, struct.pack('<I', len(libid)) + libid + struct.pack('<IH', 0, 0))
    # Modules
    d += _rec(0x000F, struct.pack('<H', len(modules)))
    d += _rec(0x0013, struct.pack('<H', 0xFFFF))
    for m in modules:
        d += _rec(0x0019, _mbcs(m['name'])) + _rec(0x0047, _u16(m['name']))
        d += _rec(0x001A, _mbcs(m['name'])) + _rec(0x0032, _u16(m['name']))
        d += _rec(0x001C, b'') + _rec(0x0048, b'')
        d += _rec(0x0031, struct.pack('<I', 0))      # MODULEOFFSET : source au début du flux
        d += _rec(0x001E, struct.pack('<I', 0))
        d += _rec(0x002C, struct.pack('<H', 0xFFFF))
        d += struct.pack('<HI', 0x0022 if m['document'] else 0x0021, 0)
        d += struct.pack('<HI', 0x002B, 0)
    d += struct.pack('<HI', 0x0010, 0)
    return bytes(d)


def build_project_stream(modules, project_id, rnd):
    lines = [f'ID="{project_id}"']
    for m in modules:
        if m['document']:
            lines.append(f"Document={m['name']}/&H00000000")
    for m in modules:
        if not m['document']:
            lines.append(f"Module={m['name']}")
    lines += [
        'Name="VBAProject"',
        'HelpContextID="0"',
        'VersionCompatible32="393222000"',
        'CMG="' + _encrypt(struct.pack('<I', 0), project_id, rnd) + '"',
        'DPB="' + _encrypt(bytes([0x00]), project_id, rnd) + '"',
        'GC="' + _encrypt(bytes([0xFF]), project_id, rnd) + '"',
        '',
        '[Host Extender Info]',
        '&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000',
        '',
        '[Workspace]',
    ]
    for m in modules:
        lines.append(f"{m['name']}=0, 0, 0, 0, C")
    return ('\r\n'.join(lines) + '\r\n').encode('cp1252')


def build_projectwm(modules):
    out = bytearray()
    for m in modules:
        out += _mbcs(m['name']) + b'\x00' + _u16(m['name']) + b'\x00\x00'
    return bytes(out + b'\x00\x00')


# ----------------------------------------------------------------------------
# MS-CFB : écriture d'un fichier composite version 3 (secteurs de 512 octets)
# ----------------------------------------------------------------------------
class _Entry:
    def __init__(self, name, kind, data=b''):
        self.name, self.kind, self.data = name, kind, data
        self.children = []
        self.left = self.right = self.child = NOSTREAM
        self.color = 1
        self.start, self.size = ENDOFCHAIN, 0
        self.sid = None


def _cfb_key(e):
    return (len(e.name), e.name.upper())


def _build_tree(entries, depth, max_depth_holder):
    if not entries:
        return None
    entries = sorted(entries, key=_cfb_key)
    mid = len(entries) // 2
    node = entries[mid]
    node._depth = depth
    max_depth_holder[0] = max(max_depth_holder[0], depth)
    l = _build_tree(entries[:mid], depth + 1, max_depth_holder)
    r = _build_tree(entries[mid + 1:], depth + 1, max_depth_holder)
    node._l, node._r = l, r
    return node


def _link_tree(children):
    holder = [0]
    root = _build_tree(children, 0, holder)
    nodes = []

    def walk(n):
        if n is None:
            return
        nodes.append(n)
        walk(n._l)
        walk(n._r)
    walk(root)
    count = len(nodes)
    perfect = (count & (count + 1)) == 0
    for n in nodes:
        n.left = n._l.sid if n._l else NOSTREAM
        n.right = n._r.sid if n._r else NOSTREAM
        # Arbre équilibré : les feuilles du niveau le plus profond sont rouges
        # si l'arbre n'est pas parfait (propriétés rouge-noir respectées).
        n.color = 0 if (not perfect and n._depth == holder[0] and n._depth > 0) else 1
    return root.sid if root else NOSTREAM


def write_cfb(tree):
    """tree : dict {nom: bytes | dict} sous la racine."""
    root = _Entry('Root Entry', 5)
    all_entries = [root]

    def add(parent, d):
        for name, val in d.items():
            if isinstance(val, dict):
                e = _Entry(name, 1)
                all_entries.append(e)
                parent.children.append(e)
                add(e, val)
            else:
                e = _Entry(name, 2, val)
                all_entries.append(e)
                parent.children.append(e)
    add(root, tree)
    for i, e in enumerate(all_entries):
        e.sid = i
    for e in all_entries:
        if e.kind in (1, 5):
            e.child = _link_tree(e.children)

    SS, MSS, CUTOFF = 512, 64, 4096
    mini_stream = bytearray()
    minifat = []
    big = []
    for e in all_entries:
        if e.kind != 2:
            continue
        e.size = len(e.data)
        if e.size == 0:
            e.start = ENDOFCHAIN
        elif e.size < CUTOFF:
            e.start = len(minifat)
            n = (e.size + MSS - 1) // MSS
            for k in range(n):
                minifat.append(e.start + k + 1 if k < n - 1 else ENDOFCHAIN)
            mini_stream += e.data + b'\x00' * (n * MSS - e.size)
        else:
            big.append(e)

    def nsec(nbytes):
        return (nbytes + SS - 1) // SS

    n_dir = nsec(len(all_entries) * 128)
    n_minifat = nsec(len(minifat) * 4)
    n_mini = nsec(len(mini_stream))
    n_big = sum(nsec(e.size) for e in big)
    n_fat = 1
    while True:
        total = n_fat + n_dir + n_minifat + n_mini + n_big
        if (total + 127) // 128 <= n_fat:
            break
        n_fat += 1
    assert n_fat <= 109

    fat = [FREESECT] * (n_fat * 128)
    cur = 0

    def alloc(n):
        nonlocal cur
        start = cur
        for k in range(n):
            fat[cur] = cur + 1 if k < n - 1 else ENDOFCHAIN
            cur += 1
        return start if n else ENDOFCHAIN

    for k in range(n_fat):
        fat[cur] = FATSECT
        cur += 1
    dir_start = alloc(n_dir)
    minifat_start = alloc(n_minifat)
    mini_start = alloc(n_mini)
    for e in big:
        e.start = alloc(nsec(e.size))
    root.start = mini_start if n_mini else ENDOFCHAIN
    root.size = len(mini_stream)

    # Directory
    dir_bytes = bytearray()
    for e in all_entries:
        name16 = e.name.encode('utf-16-le') + b'\x00\x00'
        assert len(name16) <= 64
        dir_bytes += name16.ljust(64, b'\x00')
        dir_bytes += struct.pack('<HBB', len(name16), e.kind, e.color)
        dir_bytes += struct.pack('<III', e.left, e.right, e.child)
        dir_bytes += b'\x00' * 16 + struct.pack('<I', 0) + b'\x00' * 16
        dir_bytes += struct.pack('<IQ', e.start, e.size)
    while len(dir_bytes) < n_dir * SS:
        dir_bytes += b'\x00' * 64 + struct.pack('<HBB', 0, 0, 0)
        dir_bytes += struct.pack('<III', NOSTREAM, NOSTREAM, NOSTREAM)
        dir_bytes += b'\x00' * 36 + struct.pack('<IQ', 0, 0)

    minifat_bytes = b''.join(struct.pack('<I', x) for x in minifat)
    minifat_bytes += struct.pack('<I', FREESECT) * ((n_minifat * SS - len(minifat_bytes)) // 4)

    header = bytearray()
    header += bytes.fromhex('D0CF11E0A1B11AE1') + b'\x00' * 16
    header += struct.pack('<HHHHH', 0x003E, 0x0003, 0xFFFE, 9, 6)
    header += b'\x00' * 6
    header += struct.pack('<IIIIIIIII', 0, n_fat, dir_start, 0, CUTOFF,
                          minifat_start if n_minifat else ENDOFCHAIN, n_minifat, ENDOFCHAIN, 0)
    difat = list(range(n_fat)) + [FREESECT] * (109 - n_fat)
    header += b''.join(struct.pack('<I', x) for x in difat)
    assert len(header) == 512

    body = bytearray()
    body += b''.join(struct.pack('<I', x) for x in fat)
    body += dir_bytes
    body += minifat_bytes
    body += bytes(mini_stream) + b'\x00' * (n_mini * SS - len(mini_stream))
    for e in big:
        body += e.data + b'\x00' * (nsec(e.size) * SS - e.size)
    return bytes(header + body)


# ----------------------------------------------------------------------------
def build_vba_project(modules, seed=1):
    """modules : liste de dict {name, code (str), document (bool)} ;
    le code est encodé en Windows-1252 avec fins de ligne CRLF."""
    rnd = random.Random(seed)
    project_id = '{' + str(uuid.UUID(int=rnd.getrandbits(128))).upper() + '}'
    vba = {
        '_VBA_PROJECT': b'\xCC\x61\xFF\xFF\x00\x00\x00',
        'dir': compress(build_dir(modules)),
    }
    for m in modules:
        code = m['code'].replace('\r\n', '\n').replace('\n', '\r\n')
        vba[m['name']] = compress(code.encode('cp1252'))
    tree = {
        'VBA': vba,
        'PROJECT': build_project_stream(modules, project_id, rnd),
        'PROJECTwm': build_projectwm(modules),
    }
    return write_cfb(tree)
