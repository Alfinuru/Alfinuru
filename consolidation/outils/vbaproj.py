# -*- coding: utf-8 -*-
"""
Générateur de vbaProject.bin « source seule » (MS-OVBA / MS-CFB).

Le projet produit ne contient que le code source compressé des modules (pas de
p-code) : Excel le recompile à l'ouverture. C'est le même format que le projet
VBA du classeur DPER d'origine, déjà accepté par Excel.

Usage :
    from vbaproj import construire_vba_project
    data = construire_vba_project(modules)   # modules = [(nom, type, source), ...]
    type : "doc" (ThisWorkbook / feuilles) ou "std" (module standard)
"""
import struct

# ----------------------------------------------------------------------------
# Compression MS-OVBA (2.4.1)
# ----------------------------------------------------------------------------

def _compresser_bloc(data):
    """Compresse un bloc de 4096 octets max ; renvoie (octets, compresse?)."""
    out = bytearray()
    pos = 0
    n = len(data)
    while pos < n:
        flag_pos = len(out)
        out.append(0)
        flag = 0
        for bit in range(8):
            if pos >= n:
                break
            # nombre de bits pour l'offset selon la position dans le bloc
            diff = pos
            bit_count = 4
            while (1 << bit_count) < diff:
                bit_count += 1
            length_mask = 0xFFFF >> bit_count
            max_len = length_mask + 3
            best_len, best_off = 0, 0
            if pos > 0:
                lim = min(max_len, n - pos)
                start = pos - 1
                while start >= 0:
                    l = 0
                    while l < lim and data[start + l] == data[pos + l]:
                        l += 1
                    if l > best_len:
                        best_len, best_off = l, pos - start
                        if l == lim:
                            break
                    start -= 1
            if best_len >= 3:
                token = ((best_off - 1) << (16 - bit_count)) | (best_len - 3)
                out += struct.pack('<H', token)
                flag |= (1 << bit)
                pos += best_len
            else:
                out.append(data[pos])
                pos += 1
        out[flag_pos] = flag
    return bytes(out)


def compresser(data):
    """Conteneur compressé MS-OVBA complet."""
    res = bytearray(b'\x01')
    for i in range(0, max(len(data), 1), 4096):
        bloc = data[i:i + 4096]
        if not bloc:
            break
        comp = _compresser_bloc(bloc)
        if len(comp) <= 4096 or len(bloc) < 4096:
            header = 0xB000 | (len(comp) + 2 - 3)
            res += struct.pack('<H', header) + comp
        else:  # bloc incompressible de 4096 octets : stocké brut
            res += struct.pack('<H', 0x3000 | 0x0FFF) + bloc
    return bytes(res)


# ----------------------------------------------------------------------------
# Écriture d'un fichier composé (MS-CFB v3, secteurs de 512 octets)
# ----------------------------------------------------------------------------
FREESECT, ENDOFCHAIN, FATSECT = 0xFFFFFFFF, 0xFFFFFFFE, 0xFFFFFFFD
NOSTREAM = 0xFFFFFFFF


class _Entree:
    def __init__(self, nom, typ, data=b''):
        self.nom, self.typ, self.data = nom, typ, data
        self.enfants = []
        self.id = None
        self.gauche = self.droite = self.enfant = NOSTREAM
        self.debut = ENDOFCHAIN if typ != 2 or not data else 0
        self.taille = len(data) if typ == 2 else 0


def _cle_tri(nom):
    return (len(nom), nom.upper())


def _arbre(enfants):
    """Arbre binaire équilibré (tous les nœuds noirs) ; renvoie la racine."""
    if not enfants:
        return None
    m = len(enfants) // 2
    racine = enfants[m]
    g = _arbre(enfants[:m])
    d = _arbre(enfants[m + 1:])
    racine.gauche = g.id if g else NOSTREAM
    racine.droite = d.id if d else NOSTREAM
    return racine


def ecrire_cfb(arbo):
    """arbo : dict {nom_flux: bytes | dict (stockage)} -> bytes du fichier composé."""
    racine = _Entree('Root Entry', 5)
    entrees = [racine]

    def construire(parent, contenu):
        for nom, val in contenu.items():
            if isinstance(val, dict):
                e = _Entree(nom, 1)
                entrees.append(e)
                parent.enfants.append(e)
                construire(e, val)
            else:
                e = _Entree(nom, 2, bytes(val))
                entrees.append(e)
                parent.enfants.append(e)

    construire(racine, arbo)
    for i, e in enumerate(entrees):
        e.id = i
    for e in entrees:
        if e.enfants:
            e.enfants.sort(key=lambda x: _cle_tri(x.nom))
            r = _arbre(e.enfants)
            e.enfant = r.id

    # --- mini-flux (flux < 4096 octets) ---
    mini = bytearray()
    minifat = []
    for e in entrees:
        if e.typ == 2 and 0 < e.taille < 4096:
            nb = (e.taille + 63) // 64
            e.debut = len(minifat)
            for k in range(nb):
                minifat.append(e.debut + k + 1 if k < nb - 1 else ENDOFCHAIN)
            mini += e.data + b'\x00' * (nb * 64 - e.taille)

    # --- secteurs ordinaires ---
    secteurs = []      # liste de blocs de 512 octets
    fat = []

    def allouer(data):
        if not data:
            return ENDOFCHAIN
        nb = (len(data) + 511) // 512
        debut = len(secteurs)
        for k in range(nb):
            secteurs.append(data[k * 512:(k + 1) * 512].ljust(512, b'\x00'))
            fat.append(debut + k + 1 if k < nb - 1 else ENDOFCHAIN)
        return debut

    for e in entrees:
        if e.typ == 2 and e.taille >= 4096:
            e.debut = allouer(e.data)
    racine.debut = allouer(bytes(mini)) if mini else ENDOFCHAIN
    racine.taille = len(mini)
    mf = b''.join(struct.pack('<I', x) for x in minifat)
    premier_minifat = allouer(mf) if minifat else ENDOFCHAIN
    nb_minifat = (len(mf) + 511) // 512 if minifat else 0

    # répertoire
    rep = bytearray()
    for e in entrees:
        nom = e.nom.encode('utf-16-le')
        b = nom + b'\x00\x00'
        b = b.ljust(64, b'\x00')
        b += struct.pack('<H', len(nom) + 2)
        b += bytes([e.typ, 1])                      # couleur : noir
        b += struct.pack('<III', e.gauche, e.droite, e.enfant)
        b += b'\x00' * 16 + struct.pack('<I', 0) + b'\x00' * 16
        debut = e.debut if (e.typ == 2 and e.taille > 0) or e.typ == 5 else 0
        if e.typ == 1:
            debut = 0
        b += struct.pack('<I', debut) + struct.pack('<Q', e.taille)
        assert len(b) == 128
        rep += b
    while len(rep) % 512:
        rep += (b'\x00' * 64 + b'\x00\x00' + bytes([0, 0]) +
                struct.pack('<III', NOSTREAM, NOSTREAM, NOSTREAM) + b'\x00' * 36 +
                struct.pack('<I', 0) + struct.pack('<Q', 0))
    premier_rep = allouer(bytes(rep))

    # secteurs FAT (ils se décrivent eux-mêmes)
    nb_fat = 1
    while True:
        total = len(secteurs) + nb_fat
        if (total + 127) // 128 <= nb_fat:
            break
        nb_fat += 1
    assert nb_fat <= 109, 'projet trop volumineux pour un DIFAT sans extension'
    debut_fat = len(secteurs)
    for _ in range(nb_fat):
        fat.append(FATSECT)
    fat += [FREESECT] * (nb_fat * 128 - len(fat))
    fat_bytes = b''.join(struct.pack('<I', x) for x in fat)
    for k in range(nb_fat):
        secteurs.append(fat_bytes[k * 512:(k + 1) * 512])

    entete = bytearray()
    entete += bytes.fromhex('D0CF11E0A1B11AE1') + b'\x00' * 16
    entete += struct.pack('<HHHHH', 0x003E, 0x0003, 0xFFFE, 9, 6)
    entete += b'\x00' * 6
    entete += struct.pack('<I', 0)                  # secteurs de répertoire (v3 : 0)
    entete += struct.pack('<I', nb_fat)
    entete += struct.pack('<I', premier_rep)
    entete += struct.pack('<I', 0)                  # signature de transaction
    entete += struct.pack('<I', 4096)
    entete += struct.pack('<I', premier_minifat)
    entete += struct.pack('<I', nb_minifat)
    entete += struct.pack('<I', ENDOFCHAIN)          # pas de DIFAT étendu
    entete += struct.pack('<I', 0)
    difat = [debut_fat + k for k in range(nb_fat)] + [FREESECT] * (109 - nb_fat)
    entete += b''.join(struct.pack('<I', x) for x in difat)
    assert len(entete) == 512
    return bytes(entete) + b''.join(secteurs)


# ----------------------------------------------------------------------------
# Projet VBA
# ----------------------------------------------------------------------------
# Identifiant et état de protection repris du projet DPER d'origine (projet non
# protégé, sans mot de passe). CMG/DPB/GC sont chiffrés avec une clé dérivée de
# l'ID : ils doivent rester cohérents avec lui.
PROJET_ID = '{805AE138-A344-4B79-96D5-BE28DDF8B370}'
PROJET_CMG = 'B8BA50A654A654A654A654'
PROJET_DPB = '7072985B995B995B'
PROJET_GC = '282AC013C113C1EC'

ENTETE_DOC = (
    'Attribute VB_Name = "{nom}"\r\n'
    'Attribute VB_Base = "0{{{base}}}"\r\n'
    'Attribute VB_GlobalNameSpace = False\r\n'
    'Attribute VB_Creatable = False\r\n'
    'Attribute VB_PredeclaredId = True\r\n'
    'Attribute VB_Exposed = True\r\n'
    'Attribute VB_TemplateDerived = False\r\n'
    'Attribute VB_Customizable = True\r\n')
BASE_CLASSEUR = '00020819-0000-0000-C000-000000000046'
BASE_FEUILLE = '00020820-0000-0000-C000-000000000046'


def _rec(rid, data=b''):
    return struct.pack('<HI', rid, len(data)) + data


def _dir_stream(modules):
    d = bytearray()
    d += _rec(0x0001, struct.pack('<I', 3))          # SYSKIND : Win64
    d += _rec(0x004A, struct.pack('<I', 6))          # COMPATVERSION
    d += _rec(0x0002, struct.pack('<I', 0x409))      # LCID
    d += _rec(0x0014, struct.pack('<I', 0x409))      # LCIDINVOKE
    d += _rec(0x0003, struct.pack('<H', 1252))       # CODEPAGE
    d += _rec(0x0004, b'VBAProject')
    d += _rec(0x0005) + _rec(0x0040)                 # DOCSTRING
    d += _rec(0x0006) + _rec(0x003D)                 # HELPFILEPATH
    d += _rec(0x0007, struct.pack('<I', 0))          # HELPCONTEXT
    d += _rec(0x0008, struct.pack('<I', 0))          # LIBFLAGS
    d += struct.pack('<HIIH', 0x0009, 4, 1, 0)        # VERSION (taille 4 + mineur)
    d += _rec(0x000C) + _rec(0x003C)                 # CONSTANTS
    # Référence stdole (identique au projet DPER d'origine)
    libid = b'*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\System32\\stdole2.tlb#OLE Automation'
    d += _rec(0x0016, b'stdole') + _rec(0x003E, 'stdole'.encode('utf-16-le'))
    d += _rec(0x000D, struct.pack('<I', len(libid)) + libid + struct.pack('<IH', 0, 0))
    d += _rec(0x000F, struct.pack('<H', len(modules)))
    d += _rec(0x0013, struct.pack('<H', 0xFFFF))
    for nom, typ, _src in modules:
        a = nom.encode('cp1252')
        u = nom.encode('utf-16-le')
        d += _rec(0x0019, a) + _rec(0x0047, u)
        d += _rec(0x001A, a) + _rec(0x0032, u)
        d += _rec(0x001C) + _rec(0x0048)
        d += _rec(0x0031, struct.pack('<I', 0))      # MODULEOFFSET : source au début
        d += _rec(0x001E, struct.pack('<I', 0))
        d += _rec(0x002C, struct.pack('<H', 0xFFFF))
        d += _rec(0x0022 if typ == 'doc' else 0x0021)
        d += _rec(0x002B)
    d += _rec(0x0010)
    return bytes(d)


def _project_stream(modules):
    lignes = ['ID="%s"' % PROJET_ID]
    for nom, typ, _ in modules:
        if typ == 'doc':
            lignes.append('Document=%s/&H00000000' % nom)
    for nom, typ, _ in modules:
        if typ == 'std':
            lignes.append('Module=%s' % nom)
    lignes += ['Name="VBAProject"', 'HelpContextID="0"', 'VersionCompatible32="393222000"',
               'CMG="%s"' % PROJET_CMG, 'DPB="%s"' % PROJET_DPB, 'GC="%s"' % PROJET_GC, '',
               '[Host Extender Info]',
               '&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000', '',
               '[Workspace]']
    for nom, _, _ in modules:
        lignes.append('%s=0, 0, 0, 0, C' % nom)
    return ('\r\n'.join(lignes) + '\r\n').encode('cp1252')


def _projectwm(modules):
    b = bytearray()
    for nom, _, _ in modules:
        b += nom.encode('cp1252') + b'\x00' + nom.encode('utf-16-le') + b'\x00\x00'
    b += b'\x00\x00'
    return bytes(b)


def source_module(nom, typ, corps, base=None):
    """Assemble l'en-tête d'attributs et le corps (fins de ligne CRLF, cp1252)."""
    corps = corps.replace('\r\n', '\n').replace('\n', '\r\n')
    if typ == 'doc':
        tete = ENTETE_DOC.format(nom=nom, base=base or BASE_FEUILLE)
    else:
        tete = 'Attribute VB_Name = "%s"\r\n' % nom
    txt = tete + corps
    if not txt.endswith('\r\n'):
        txt += '\r\n'
    return txt.encode('cp1252')   # lève une erreur si un caractère n'est pas en cp1252


def construire_vba_project(modules):
    """modules : liste (nom, 'doc'|'std', source_bytes). Renvoie vbaProject.bin."""
    vba = {
        '_VBA_PROJECT': bytes.fromhex('CC61FFFF000300'),
        'dir': compresser(_dir_stream(modules)),
    }
    for nom, _typ, src in modules:
        vba[nom] = compresser(src)
    arbo = {'VBA': vba, 'PROJECT': _project_stream(modules), 'PROJECTwm': _projectwm(modules)}
    return ecrire_cfb(arbo)
