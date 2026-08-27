# -*- coding: utf-8 -*-
"""
build_geo.py — KMZ/KML дан navoi_geo.js ни ТОПОЛОГИК ТЎҒРИ қуради.

НИМА УЧУН БУ КЕРАК
  Эски parse_kml.ps1 ҳар 3-нуқтадан биттасини қолдирар (simplifyFactor=3)
  ва ҳар полигонни АЛОҲИДА сийраклаштирар эди. Икки қўшни маҳалла KML'да
  бир хил чегарани улашса ҳам, улардан ҳар бири ўша чегаранинг ҲАР ХИЛ
  нуқталарини сақлаб қоларди — натижада орада тирқиш ёки устма-устлик
  пайдо бўларди. 152 870 нуқтадан 68 362 таси қолган: 55% йўқотилган.
  Кейин fix_geo_gaps.py ўша тирқишларни ямарди — оқибатни даволаш эди.

  Бу скрипт бошқача ишлайди: тирқишни ямамайди, унинг ПАЙДО БЎЛИШИГА
  ЙЎЛ ҚЎЙМАЙДИ.

ҚАНДАЙ ИШЛАЙДИ — ЮЗАЛАР УСУЛИ (planar face assignment)
  1. KML тўлиқ аниқликда ўқилади: ҳамма нуқта, ҳамма ички тешик
     (innerBoundaryIs), MultiGeometry ичидаги ҳамма полигон.
  2. Координаталар ЯГОНА тўрга ўтказилади (set_precision). Тўр умумий
     бўлгани учун икки полигоннинг бир хил чегара нуқтаси АЙНАН бир хил
     қийматга тушади — улашилган чегара улашилганлигича қолади.
  3. Ҳамма полигоннинг чегара чизиқлари БИРГА тугунланади (noding):
     unary_union(boundaries). Шунда икки маҳалла орасидаги умумий чегара
     БИТТА ёй бўлади — иккита эмас.
  4. polygonize() — ёйлар тўридан ЮЗАЛАР (faces) қурилади. Юзалар
     текисликни қоплайди: улар орасида тирқиш ҳам, устма-устлик ҳам
     БЎЛИШИ МУМКИН ЭМАС — бу геометриянинг ўз хоссаси.
  5. Ҳар юза эгасига берилади:
       - юза ичидаги нуқтани фақат битта полигон ўз ичига олса → ўшаники;
       - бир нечтаси олса (устма-уст жой) → МАЙДОНИ КИЧИГИНИКИ. Кичик
         полигон — аниқ чизилган маҳалла, каттаси — қўпол чизилган чўл
         ОФЙси; шу боис аниқроғи устун туради;
       - ҳеч ким олмаса (манбадаги тирқиш) → энг яқин қўшниники, агар у
         MAX_GAP_M дан яқин бўлса. Шу билан манбадаги тирқиш ҳам ёпилади.
  6. Ҳар объектнинг юзалари бирлаштирилади.

  Натижа: устма-устлик 0 — таъминланган. Ичкаридаги тирқиш 0 —
  таъминланган. Ҳеч бир полигон қирқилмайди ва парчаланмайди: у фақат
  ўзига тегишли юзаларни олади.

  7. Четга очиладиган тирқишлар ёпиқ ҳалқа ҳосил қилмагани учун юзага
     айланмайди — улар морфологик ёпиш билан алоҳида текширилади.

ХАВФСИЗЛИК
  --apply берилмаса ҲЕЧ НАРСА ёзилмайди. Ёзишдан олдин navoi_geo.js
  захираланади ва туман номлари эски файл билан солиштирилади: янги
  KMZ'да туман номи бошқача ёзилган бўлса, платформадаги Excel маълумоти
  харитага боғланмай қолади — шу боис фарқ топилса ЁЗИЛМАЙДИ
  (--allow-new-districts билан мажбурлаш мумкин).

ИШЛАТИШ
    python build_geo.py "C:\\path\\Navoiy.kmz"            # синов, ёзилмайди
    python build_geo.py "C:\\path\\Navoiy.kmz" --apply    # ёзади
    python build_geo.py "...kmz" --simplify 2 --apply     # ~2 м топологик
                                                          # соддалаштириш
Талаб: pip install shapely
"""
import io, json, math, os, re, shutil, sys, zipfile, collections
import xml.etree.ElementTree as ET
import shapely
from shapely.geometry import shape, mapping, Polygon
from shapely.ops import unary_union, polygonize
from shapely import make_valid, STRtree

try: sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception: pass

OUT       = 'navoi_geo.js'
PREFIX    = 'window.NAVOI_GEOJSON = '
GRID      = 1e-7     # ~0.011 м — умумий координата тўри
ND        = 6        # ёзишдаги ўнлик хона (~0.11 м)
MAX_GAP_M = 400.0    # эгасиз юза шунчадан яқин қўшнига берилади
SEAM_M    = 60.0     # четга очиладиган тирқиш кенглиги чегараси

M_LAT = 111132.0
M_LON = 111320.0 * math.cos(math.radians(41))
def km2(a): return a * M_LAT * M_LON / 1e6
def m2(a):  return a * M_LAT * M_LON


# ------------------------------------------------------------ геометрия
def polys_of(g):
    if g is None or g.is_empty: return []
    if g.geom_type == 'Polygon': return [g] if g.area > 0 else []
    if g.geom_type in ('MultiPolygon', 'GeometryCollection'):
        out = []
        for p in g.geoms: out += polys_of(p)
        return out
    return []

def snap(g):
    """Умумий тўрга туширади ва яроқли ҳолга келтиради."""
    if g is None or g.is_empty: return None
    if not g.is_valid: g = make_valid(g)
    g = shapely.set_precision(g, GRID)
    if not g.is_valid: g = make_valid(g)
    parts = polys_of(g)
    return unary_union(parts) if parts else None

def to_geojson(g):
    def rr(ring): return [[round(x, ND), round(y, ND)] for x, y in ring]
    m = mapping(g)
    if m['type'] == 'Polygon':
        return {'type': 'Polygon', 'coordinates': [rr(r) for r in m['coordinates']]}
    return {'type': 'MultiPolygon',
            'coordinates': [[rr(r) for r in poly] for poly in m['coordinates']]}

def npoints(g):
    if g is None: return 0
    m = mapping(g)
    if m['type'] == 'Polygon': return sum(len(r) for r in m['coordinates'])
    return sum(len(r) for poly in m['coordinates'] for r in poly)

def clusters(g):
    """Полигон нечта АЛОҲИДА бўлакдан иборат (тегиб турганлари битта)."""
    if g is None: return 0
    return len(polys_of(unary_union(polys_of(g))))


# ---------------------------------------------------------------- KML ўқиш
def local(tag):
    return tag.rsplit('}', 1)[-1] if '}' in tag else tag

def kml_bytes(path):
    """KMZ (zip) бўлса ичидаги .kml ни, KML бўлса файлнинг ўзини қайтаради."""
    if zipfile.is_zipfile(path):
        with zipfile.ZipFile(path) as z:
            names = [n for n in z.namelist() if n.lower().endswith('.kml')]
            if not names: sys.exit('ХАТО: KMZ ичида .kml топилмади')
            name = next((n for n in names if n.lower().endswith('doc.kml')),
                        max(names, key=lambda n: z.getinfo(n).file_size))
            print('KMZ ичидан ўқилди   : %s' % name)
            return z.read(name)
    return io.open(path, 'rb').read()

def sanitize_xml(data):
    """Эълон қилинмаган namespace префиксларини эълон қилиб қўяди.

    Google Earth экспорти кўпинча `xsi:schemaLocation` ни ёзади, аммо
    `xmlns:xsi` ни эълон қилмайди. Стандарт XML ўқувчи бунга «unbound
    prefix» деб тўхтайди, ҳолбуки геометрияга ҳеч қандай алоқаси йўқ.
    Шу боис етишмаган префиксларни илдиз тегига ўзимиз қўшамиз.
    """
    txt = data.decode('utf-8-sig', 'replace')
    used = set(re.findall(r'[<\s]([A-Za-z][\w.-]*):[A-Za-z_]', txt))
    declared = set(re.findall(r'xmlns:([\w.-]+)\s*=', txt))
    missing = used - declared - {'xmlns'}
    if not missing: return txt
    m = re.search(r'<([A-Za-z_][\w.:-]*)([^>]*)>', txt)
    if not m: return txt
    add = ''.join(' xmlns:%s="urn:x-auto:%s"' % (p, p) for p in sorted(missing))
    print('эълонсиз префикс    : %s — ўзимиз эълон қилдик'
          % ', '.join(sorted(missing)))
    return txt[:m.end() - 1] + add + txt[m.end() - 1:]

def props_of(pm):
    """Placemark'дан ҳамма калит=қиймат жуфтини йиғади.

    Уч манбадан: ExtendedData/SimpleData, ExtendedData/Data ва description
    ичидаги HTML жадвал — янги KMZ бошқа тузилишда бўлса ҳам ишлаши учун.
    """
    out = {}
    for el in pm.iter():
        t = local(el.tag)
        if t == 'SimpleData' and el.get('name'):
            out.setdefault(el.get('name').strip(), (el.text or '').strip())
        elif t == 'Data' and el.get('name'):
            v = next((c.text for c in el if local(c.tag) == 'value'), None)
            out.setdefault(el.get('name').strip(), (v or '').strip())
        elif t == 'description' and el.text:
            for k, v in re.findall(
                    r'<t[dh][^>]*>(.*?)</t[dh]>\s*<t[dh][^>]*>(.*?)</t[dh]>',
                    el.text, re.S | re.I):
                k = re.sub(r'<[^>]+>', '', k).strip()
                v = re.sub(r'<[^>]+>', '', v).strip()
                if k: out.setdefault(k, v)
    return out

def pick(props, *pats):
    """Калит номи қандай ёзилганидан қатъи назар қийматни топади."""
    for pat in pats:
        rx = re.compile(pat, re.I)
        for k, v in props.items():
            if rx.search(k) and str(v).strip(): return str(v).strip()
    return ''

def rings_of(poly_el):
    """Битта <Polygon> дан ташқи ва ички ҳалқаларни олади."""
    outer, inners = None, []
    for b in poly_el:
        t = local(b.tag)
        if t not in ('outerBoundaryIs', 'innerBoundaryIs'): continue
        co = next((e for e in b.iter() if local(e.tag) == 'coordinates'), None)
        if co is None or not co.text: continue
        ring = []
        for p in co.text.split():
            xy = p.split(',')
            if len(xy) >= 2:
                try: ring.append((float(xy[0]), float(xy[1])))
                except ValueError: pass
        if len(ring) < 4: continue
        if ring[0] != ring[-1]: ring.append(ring[0])
        if t == 'outerBoundaryIs': outer = ring
        else: inners.append(ring)
    return outer, inners

def read_kml(path):
    root = ET.fromstring(sanitize_xml(kml_bytes(path)))
    feats, skipped, raw_pts = [], 0, 0
    for pm in (e for e in root.iter() if local(e.tag) == 'Placemark'):
        pr = props_of(pm)
        district = pick(pr, r'tuman\s*nomi', r'^tuman$', r'туман', r'district', r'^rayon')
        mfy      = pick(pr, r'mahalla', r'махалла', r'мфй', r'^mfy$', r'mfy[_\s]*nom')
        mfy_id   = pick(pr, r'mfy[_\s]*id', r'^id$', r'^kod')
        if not district:
            skipped += 1; continue
        parts = []
        for pe in (e for e in pm.iter() if local(e.tag) == 'Polygon'):
            outer, inners = rings_of(pe)
            if not outer: continue
            raw_pts += len(outer) + sum(len(r) for r in inners)
            try: p = Polygon(outer, inners)
            except Exception: continue
            if not p.is_valid: p = make_valid(p)
            parts += polys_of(p)
        if not parts:
            skipped += 1; continue
        feats.append({'district': district, 'mfy': mfy, 'mfy_id': mfy_id,
                      'geom': unary_union(parts)})
    if skipped:
        print('маълумотсиз Placemark  : %d — ўтказиб юборилди' % skipped)
    return feats, raw_pts


# ------------------------------------------------- юзаларни эгасига бериш
def assign_faces(geoms, simplify_m=0.0):
    """Планар юзаларни қуради ва ҳар бирини эгасига беради.

    Қайтаради: (янги геометриялар, ҳисобот dict).
    """
    live = [(i, g) for i, g in enumerate(geoms) if g is not None]
    # 1) чегара чизиқларини БИРГА тугунлаш — умумий қирра битта ёй бўлади
    bounds = unary_union([g.boundary for _, g in live])
    rep = {'arcs_before': len(getattr(bounds, 'geoms', [bounds]))}

    # 2) ихтиёрий: ҳар ёйни БИР МАРТА соддалаштириш. Ёй икки қўшнининг
    #    умумий қиррасигина бўлгани учун иккаласи БИР ХИЛ ўзгаради —
    #    шу боис тирқиш ҳосил бўлмайди. Учлари (тугунлар) қимирламайди.
    if simplify_m > 0:
        tol = simplify_m / M_LON
        arcs = list(getattr(bounds, 'geoms', [bounds]))
        bounds = unary_union([a.simplify(tol, preserve_topology=True) for a in arcs])
    rep['arcs'] = len(getattr(bounds, 'geoms', [bounds]))

    # 3) юзалар
    faces = [f for f in polygonize(bounds) if f.area > 0]
    rep['faces'] = len(faces)

    # 4) эгасини аниқлаш
    idxs  = [i for i, _ in live]
    polys = [g for _, g in live]
    tree  = STRtree(polys)
    order = {i: k for k, i in enumerate(sorted(idxs, key=lambda i: geoms[i].area))}

    own = collections.defaultdict(list)
    n_single = n_over = n_orphan = n_drop = 0
    orphan_area = 0.0
    maxd = MAX_GAP_M / M_LON
    for f in faces:
        pt = f.representative_point()
        cand = [idxs[j] for j in tree.query(pt) if polys[j].contains(pt)]
        if len(cand) == 1:
            own[cand[0]].append(f); n_single += 1
        elif len(cand) > 1:
            # устма-уст жой — МАЙДОНИ КИЧИК полигонники (аниқроқ чизилган)
            own[min(cand, key=lambda i: order[i])].append(f); n_over += 1
        else:
            # эгасиз юза = манбадаги тирқиш. Энг яқин қўшнига берамиз.
            best, bd = None, maxd
            for j in tree.query(f.buffer(maxd)):
                d = polys[j].distance(f)
                if d < bd: bd, best = d, idxs[j]
            if best is None:
                n_drop += 1
            else:
                own[best].append(f); n_orphan += 1; orphan_area += f.area

    rep.update(single=n_single, overlap=n_over, orphan=n_orphan,
               orphan_km2=km2(orphan_area), dropped=n_drop)

    out = list(geoms)
    for i in idxs:
        fs = own.get(i)
        out[i] = unary_union(fs) if fs else None
    return out, rep


# ------------------------------------------- четга очиладиган тирқишлар
def close_seams(geoms, max_gap_m=SEAM_M):
    """Ташқарига очилиб турган тор тирқишларни ёпади.

    Юзалар усули фақат ЁПИҚ ҳалқа ҳосил қиладиган тирқишни тутади. Икки
    полигон орасидаги тирқиш бир учи билан ташқарига чиқиб турса, у
    ҳалқа эмас — юзага айланмайди. Уни морфологик ёпиш билан топамиз.

    МУҲИМ: топилган бўлаклардан фақат ИККИ ва ундан ортиқ полигонга
    тегадиганлари олинади — улар ҳақиқий тирқиш. Битта полигонга
    тегадигани ўша полигоннинг ЎЗ ШАКЛИ (табиий ботиғи); уни тўлдириш
    туман чегарасини нотўғри катталаштирган бўларди.
    """
    d = max_gap_m / M_LON
    live = [g for g in geoms if g is not None]
    u = unary_union(live)
    gap = u.buffer(d, join_style=2).buffer(-d, join_style=2).difference(u)
    cand = [p for p in polys_of(gap) if m2(p.area) > 200]
    if not cand:
        print('   чет тирқиш: йўқ'); return 0

    parts_idx, parts_geom = [], []
    for i, g in enumerate(geoms):
        for pp in polys_of(g):
            parts_idx.append(i); parts_geom.append(pp)
    tree = STRtree(parts_geom)

    assign, seams, shape_only = collections.defaultdict(list), 0, 0
    for pc in cand:
        pb = pc.boundary.buffer(d / 3)
        touch = {}
        for j in tree.query(pc.buffer(d)):
            L = parts_geom[j].boundary.intersection(pb).length
            if L > 0: touch[parts_idx[j]] = touch.get(parts_idx[j], 0) + L
        if len(touch) < 2:
            shape_only += 1; continue
        seams += 1
        assign[max(touch.items(), key=lambda kv: kv[1])[0]].append(pc)

    for i, ps in assign.items():
        geoms[i] = snap(unary_union([geoms[i]] + ps))
    print('   чет тирқиш: %d ёпилди, %d бўлак полигоннинг ўз шакли — тегилмади'
          % (seams, shape_only))
    return seams


# ------------------------------------------------------------- текшириш
def audit(geoms, label):
    live = [g for g in geoms if g is not None]
    u = unary_union(live)
    holes = [Polygon(r) for p in polys_of(u) for r in p.interiors]
    holes = [h for h in holes if h.area > 0]
    ov = sum(g.area for g in live) - u.area
    bad = sum(1 for g in live if not g.is_valid)
    print('%-13s ичкари тешик %4d (%8.0f м²) | устма-устлик %9.0f м² | '
          'яроқсиз %d | майдон %.1f км²'
          % (label, len(holes), m2(sum(h.area for h in holes)),
             m2(max(ov, 0)), bad, km2(u.area)))
    return {'holes': len(holes), 'overlap_m2': m2(max(ov, 0)),
            'invalid': bad, 'area_km2': km2(u.area)}


# ------------------------------------------------------------------ асосий
def main():
    args = sys.argv[1:]
    if not args or args[0].startswith('--'):
        sys.exit(__doc__)
    src = args[0]
    apply_ = '--apply' in args
    allow_new = '--allow-new-districts' in args
    simplify_m = 0.0
    if '--simplify' in args:
        try: simplify_m = float(args[args.index('--simplify') + 1])
        except Exception: sys.exit('ХАТО: --simplify дан кейин метр сони керак')
    if not os.path.exists(src): sys.exit('ХАТО: файл топилмади: ' + src)

    print('манба               : %s (%.1f MB)' % (src, os.path.getsize(src) / 1e6))
    feats, raw_pts = read_kml(src)
    if not feats: sys.exit('ХАТО: KML дан бирорта полигон ўқилмади')
    print('объект (Placemark)  : %d' % len(feats))
    print('манбадаги нуқта     : %d' % raw_pts)
    if simplify_m: print('соддалаштириш       : ~%.1f м (топологик, умумий қирра бўйича)' % simplify_m)
    print()

    src_geoms = [snap(f['geom']) for f in feats]
    geoms = list(src_geoms)
    a0 = audit(geoms, 'манба ҳолати:')

    print()
    print('ЮЗАЛАРНИ ҚУРИШ ВА ЭГАСИГА БЕРИШ')
    geoms, rep = assign_faces(geoms, simplify_m)
    print('   ёй (умумий қирра): %d' % rep['arcs'])
    print('   юза              : %d' % rep['faces'])
    print('   эгаси аниқ       : %d' % rep['single'])
    print('   устма-уст жой    : %d — кичик полигонга берилди' % rep['overlap'])
    print('   манбадаги тирқиш : %d (%.3f км²) — энг яқин қўшнига берилди'
          % (rep['orphan'], rep['orphan_km2']))
    if rep['dropped']:
        print('   эгасиз қолди     : %d — %d м дан узоқ, ташланди'
              % (rep['dropped'], int(MAX_GAP_M)))
    print()
    print('ЧЕТГА ОЧИЛАДИГАН ТИРҚИШЛАР')
    close_seams(geoms)
    print()
    a1 = audit(geoms, 'натижа:      ')

    # ---- парчаланиш ва йўқолиш назорати ----
    print()
    print('БУТУНЛИК НАЗОРАТИ (объект даражасида)')
    frag, lost, shrunk = [], [], []
    for i, f in enumerate(feats):
        c0 = clusters(src_geoms[i]); c1 = clusters(geoms[i])
        a_0 = src_geoms[i].area if src_geoms[i] is not None else 0
        a_1 = geoms[i].area if geoms[i] is not None else 0
        if geoms[i] is None or a_1 <= 0: lost.append(f)
        elif c1 > c0: frag.append((f, c0, c1))
        elif a_0 > 0 and a_1 < a_0 * 0.80: shrunk.append((f, a_0, a_1))
    print('   парчаланган объект: %d' % len(frag))
    for f, c0, c1 in frag[:8]:
        print('      %s / %s : %d → %d бўлак' % (f['district'], f['mfy'], c0, c1))
    print('   йўқолган объект   : %d' % len(lost))
    for f in lost[:8]: print('      %s / %s' % (f['district'], f['mfy']))
    print('   20%%+ кичрайган    : %d' % len(shrunk))
    for f, a_0, a_1 in shrunk[:8]:
        print('      %s / %s : %.2f → %.2f км²' % (f['district'], f['mfy'], km2(a_0), km2(a_1)))

    # ---- туман даражасида ----
    print()
    print('ТУМАНЛАР (бўлаклар сони — 1 бўлса туман бутун)')
    byd = collections.defaultdict(list)
    for i, f in enumerate(feats):
        if geoms[i] is not None: byd[f['district']].append(geoms[i])
    print('   %-24s %5s %5s %12s' % ('туман', 'МФЙ', 'бўлак', 'майдон км²'))
    for d in sorted(byd):
        u = unary_union(byd[d])
        print('   %-24s %5d %5d %12.1f' % (d[:24], len(byd[d]), len(polys_of(u)), km2(u.area)))

    # ---- туман номлари эски файл билан мос келадими ----
    new_names = set(byd)
    old_names = set()
    if os.path.exists(OUT):
        t = io.open(OUT, encoding='utf-8-sig').read()
        old = json.loads(t[t.index('{'):t.rindex('}') + 1])
        old_names = {str(f['properties'].get('district')) for f in old['features']}
    missing = old_names - new_names
    added   = new_names - old_names
    print()
    if old_names and (missing or added):
        print('!!! ДИҚҚАТ: туман номлари эски файлдан ФАРҚ ҚИЛАДИ')
        if missing: print('    йўқолган : %s' % ', '.join(sorted(missing)))
        if added:   print('    янги     : %s' % ', '.join(sorted(added)))
        print('    Платформа Excel маълумотини харитага АНА ШУ номлар')
        print('    орқали боғлайди — мос келмаса пинлар йўқолади.')
    elif old_names:
        print('туман номлари эски файл билан тўлиқ мос — боғланиш сақланади')

    pts = sum(npoints(g) for g in geoms if g is not None)
    print()
    print('=== ХУЛОСА ===')
    print('ичкари тешик    : %d → %d' % (a0['holes'], a1['holes']))
    print('устма-устлик    : %.0f → %.0f м²' % (a0['overlap_m2'], a1['overlap_m2']))
    print('яроқсиз шакл    : %d → %d' % (a0['invalid'], a1['invalid']))
    print('майдон          : %.1f → %.1f км²' % (a0['area_km2'], a1['area_km2']))
    print('нуқта           : манбада %d → файлда %d' % (raw_pts, pts))

    if not apply_:
        print()
        print('>>> СИНОВ РЕЖИМИ — ҳеч нарса ёзилмади. Ёзиш учун: --apply')
        return
    if (missing or added) and not allow_new:
        print()
        sys.exit('>>> ЁЗИЛМАДИ: туман номлари мос эмас (юқорига қаранг).\n'
                 '    Атайлаб бўлса --allow-new-districts қўшинг.')

    if os.path.exists(OUT):
        shutil.copy2(OUT, OUT + '.bak')
        print()
        print('захира          : %s.bak' % OUT)
    out = []
    for f, g in zip(feats, geoms):
        if g is None or g.is_empty: continue
        out.append({'type': 'Feature',
                    'properties': {'district': f['district'], 'mfy': f['mfy'],
                                   'mfy_id': f['mfy_id']},
                    'geometry': to_geojson(g)})
    txt = PREFIX + json.dumps({'features': out, 'type': 'FeatureCollection'},
                              ensure_ascii=False, separators=(',', ':')) + ';'
    io.open(OUT, 'w', encoding='utf-8-sig', newline='').write(txt)
    print('ёзилди          : %s (%d объект, %.2f MB)'
          % (OUT, len(out), len(txt.encode('utf-8')) / 1e6))
    print()
    print('КЕЙИНГИ ҚАДАМ: navoi_platform.html даги navoi_geo.js?v=... ни')
    print('янгиланг — Vercel .js файлларини 1 йил кэшлайди.')


if __name__ == '__main__':
    main()
