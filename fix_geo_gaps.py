# -*- coding: utf-8 -*-
"""
navoi_geo.js — маҳалла/туман полигонларининг топологиясини тозалайди.

Муаммо: полигонлар бир-биридан мустақил рақамланган, шу боис
  (1) чегаралар мос тушмайди → орада ОЧИҚ ЖОЙЛАР (gaps) қолади;
  (2) айрим полигонлар бир-бирини ҚОПЛАЙДИ (overlaps) → чегарадаги
      участка иккита маҳаллага тушиб, қайси бири олдин топилса ўша танланади.

Босқичлар:
  A. Тешикларни ёпиш — unary_union ичидаги ҳар тешик ЎЗИ БИЛАН ЭНГ УЗУН
     чегарани бўлишадиган полигонга қўшилади (шу боис ўзи ўралиб турган
     туманга тушади). Тешик қолмагунча такрорланади.
  B. Устма-устликни ечиш — полигонлар КИЧИГИДАН БОШЛАБ навбатга қўйилади,
     ҳар бири ўзидан аввалгилар эгаллаган жойни бўшатади. Кичик полигонлар
     (аниқ чизилган маҳаллалар) тўлиқ сақланади, катталари (қўпол чизилган
     чўл ОФЙлари) қирқилади. Union ўзгармайди — янги тешик пайдо бўлмайди.
  C. Аниқлик артефактлари учун А босқичи қайта юритилади ва иккала
     кўрсаткич ҳам 0 экани текширилади.

set_precision(1e-6) — GEOS'нинг snap-rounding'и: ҳамма координата ~0.11 м
тўрига тушади ва геометрия ЯРОҚЛИ бўлиб қолади.

IDEMPOTENT: иккинчи марта ишга туширилса геометрия ўзгармайди (тешик 0,
устма-устлик артефакт даражасида, майдон ўша-ўша). Файл байтлари бир оз
фарқ қилиши мумкин —
бу фақат полигон тепа нуқталарининг тартиби, шакл эмас.

`parse_kml.ps1` navoi_geo.js ни қайта яратгандан СЎНГ шуни ишлатинг:
    python fix_geo_gaps.py
Талаб: pip install shapely
"""
import io, json, math, collections, sys
import shapely
from shapely.geometry import shape, mapping, Polygon
from shapely.ops import unary_union
from shapely import make_valid, STRtree

# Windows консоли (cp1251) ўзбекча ҳарфларни чиқара олмайди — UTF-8 га ўтказамиз
try: sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception: pass

SRC    = 'navoi_geo.js'
PREFIX = 'window.NAVOI_GEOJSON = '
GRID   = 1e-6      # ~0.11 м — координата тўри
ND     = 6         # GRID га мос ўнлик хона сони
GROW   = 0.0006    # ~50 м — тешик атрофида қўшни излаш оралиғи
MIN_KEEP = 0.02    # полигон майдонининг камида шунча улуши сақланиши шарт

M_LAT = 111132.0
M_LON = 111320.0 * math.cos(math.radians(41))
def km2(a): return a * M_LAT * M_LON / 1e6

def unwrap(v):
    if isinstance(v, dict) and 'value' in v: return unwrap(v['value'])
    if isinstance(v, list): return [unwrap(x) for x in v]
    return v

def polys_of(g):
    if g is None or g.is_empty: return []
    if g.geom_type == 'Polygon': return [g] if g.area > 0 else []
    if g.geom_type in ('MultiPolygon', 'GeometryCollection'):
        out = []
        for p in g.geoms: out += polys_of(p)
        return out
    return []

def clean(g):
    """Яроқли + айнан ЁЗИЛАДИГАН аниқликдаги полигон(лар) қайтаради.

    Муҳим: фақат set_precision етарли эмас — файлга ёзишдаги ND хонали
    яхлитлаш ўзи янги микро-тирқиш ҳосил қилади. Шу боис геометрияни шу
    ерда round-trip қиламиз (ёзиш → қайта ўқиш), токи ўлчовларимиз
    браузер кўрадиган ҳақиқий геометрияга тегишли бўлсин.
    """
    if g is None or g.is_empty: return None
    if not g.is_valid: g = make_valid(g)
    g = shapely.set_precision(g, GRID)
    if not g.is_valid: g = make_valid(g)
    parts = polys_of(g)
    if not parts: return None
    g = unary_union(parts)
    g = shape(to_geojson(g))          # ёзилиш аниқлигига келтирамиз
    if not g.is_valid: g = make_valid(g)
    parts = polys_of(g)
    return unary_union(parts) if parts else None

def to_geojson(g):
    def rr(ring): return [[round(x, ND), round(y, ND)] for x, y in ring]
    m = mapping(g)
    if m['type'] == 'Polygon':
        return {'type':'Polygon','coordinates':[rr(r) for r in m['coordinates']]}
    return {'type':'MultiPolygon',
            'coordinates':[[rr(r) for r in poly] for poly in m['coordinates']]}

def find_holes(geoms):
    u = unary_union([g for g in geoms if g is not None])
    hs = [Polygon(r) for p in polys_of(u) for r in p.interiors]
    return [h for h in hs if h.area > 0], u

def close_gaps(geoms, label):
    """Тешик қолмагунча ёпади. Нечта ёпилганини қайтаради."""
    closed, prev = 0, None
    for it in range(1, 9):
        holes, _ = find_holes(geoms)
        if not holes:
            print('  %s pass %d: тешик йўқ' % (label, it)); break
        print('  %s pass %d: %d gap (%.3f km2)'
              % (label, it, len(holes), sum(km2(h.area) for h in holes)))
        # Snap-rounding айрим микро-тирқишларни ёпгач қайта ҳосил қилиши мумкин.
        # Илгарилаш тўхтаса — тебранмасдан чиқамиз (қолдиқ 1 м² дан кичик).
        if prev is not None and len(holes) >= prev:
            print('  %s: илгарилаш йўқ — %d микро-тирқиш қолди' % (label, len(holes)))
            break
        prev = len(holes)
        parts_idx, parts_geom = [], []
        for i, g in enumerate(geoms):
            for p in polys_of(g):
                parts_idx.append(i); parts_geom.append(p)
        tree = STRtree(parts_geom)
        assign = collections.defaultdict(list)
        for h in holes:
            hb = h.boundary.buffer(GROW / 3)
            best, best_len = None, 0.0
            for j in tree.query(h.buffer(GROW)):
                inter = parts_geom[j].boundary.intersection(hb)
                L = inter.length if not inter.is_empty else 0.0
                if L > best_len: best_len, best = L, parts_idx[j]
            if best is not None: assign[best].append(h)
        if not assign:
            print('  %s: қўшни топилмади — тўхтадик' % label); break
        for i, hs in assign.items():
            geoms[i] = clean(unary_union([geoms[i]] + hs))
        closed += sum(len(v) for v in assign.values())
    return closed

def drop_debris(cut, orig_area):
    """Қирқишдан кейин қолган МАЙДА ПАРЧАЛАРНИ олиб ташлайди.

    Устма-устликни ечишда катта полигон кичиклари билан кесилади ва
    чеккаларида қоғоз қириндисидек майда бўлаклар қолади. Мисол: Томди
    туманининг «Sharq OFY» полигони Зарафшон маҳаллалари билан кесилгач,
    39 та 0.0-0.1 км² лик парчага айланиб қолди — харитада туман
    майдаланиб кетгандек кўринарди.

    Парча деб ҳисобланади: (а) ўз полигонининг 1% идан кичик ВА
    (б) 5 гектардан кичик. Иккала шарт бирга — шунда ҳақиқий кичик
    маҳалла тасодифан ўчиб кетмайди.
    Ташлангани бўш жой қолдирмайди: у ер аллақачон қўшни полигон
    остида, чунки бу устма-уст тушган жойнинг қолдиғи.
    """
    if cut is None or cut.is_empty: return cut
    parts = polys_of(cut)
    if len(parts) <= 1: return cut
    # 15 гектар чегараси. Ҳақиқий кичик маҳалла тасодифан ўчиб кетмайди,
    # чунки ИККИНЧИ шарт ҳимоя қилади: кичик маҳалланинг ўз бўлаги унинг
    # майдонининг 100% и бўлади, яъни 1% дан катта — сақланади.
    keep = [p for p in parts
            if p.area >= orig_area * 0.01 or p.area * M_LAT * M_LON >= 150000]
    if not keep: return cut          # ҳаммаси майда бўлса — тегмаймиз
    return unary_union(keep)

def close_seams(geoms, max_gap_m=60.0):
    """ЧЕТГА ОЧИЛАДИГАН тирқишларни ёпади.

    close_gaps() фақат unary_union ичидаги ЁПИҚ тешикларни топади. Аммо икки
    полигон орасидаги тирқиш ташқарига очилиб турса, у «тешик» эмас —
    union'нинг ташқи чегарасидаги ботиқ бўлиб қолади ва сезилмай кетади.
    Харитада эса у худди тешикдек кўринади.

    Бу ерда морфологик ёпиш (buffer(+d).buffer(-d)) билан тор бўйинли
    жойлар топилади. МУҲИМ: улардан фақат ИККИ ВА УНДАН ОРТИҚ полигонга
    тегадиганлари олинади — улар ҳақиқий тирқиш. Битта полигонга
    тегадигани эса ўша полигоннинг ЎЗ ШАКЛИ (табиий ботиғи); уни
    тўлдириш туман чегарасини нотўғри катталаштирган бўларди.
    """
    d = max_gap_m / M_LON
    live = [g for g in geoms if g is not None]
    u = unary_union(live)
    gap = u.buffer(d, join_style=2).buffer(-d, join_style=2).difference(u)
    cand = [p for p in polys_of(gap) if p.area * M_LAT * M_LON > 200]
    if not cand:
        print('  чет тирқиш: йўқ'); return 0

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
            shape_only += 1        # полигоннинг ўз шакли — тегмаймиз
            continue
        seams += 1
        assign[max(touch.items(), key=lambda kv: kv[1])[0]].append(pc)

    for i, ps in assign.items():
        geoms[i] = clean(unary_union([geoms[i]] + ps))
    print('  чет тирқиш: %d ёпилди, %d бўлак полигоннинг ўз шакли — тегилмади'
          % (seams, shape_only))
    return seams

def overlap_area(geoms):
    live = [g for g in geoms if g is not None]
    return sum(g.area for g in live) - unary_union(live).area

# ---------- юклаш ----------
src = io.open(SRC, encoding='utf-8-sig').read()
gj = json.loads(src[src.index('{'):src.rindex('}')+1])
feats = gj['features']

geoms, invalid = [], 0
for f in feats:
    g = dict(f['geometry']); g['coordinates'] = unwrap(g['coordinates'])
    s = shape(g)
    if not s.is_valid: invalid += 1
    geoms.append(clean(s))

_, u0 = find_holes(geoms)
area_before = u0.area
ov_before = overlap_area(geoms)
print('features            :', len(feats))
print('invalid -> repaired :', invalid)
print('overlap before      : %.2f km2' % km2(ov_before))
print()

# ---------- A: тешикларни ёпиш ----------
print('A. тешикларни ёпиш')
print('   ёпилди: %d' % close_gaps(geoms, 'A'))
print()

# ---------- B: устма-устликни ечиш ----------
# Қолдиқ snap-rounding артефактидан иборат бўлса — тегмаймиз, акс ҳолда
# ҳар ишга туширишда полигонлар қайта қирқилиб, файл беҳуда шишаверади.
OV_SKIP = 0.01   # km2
print('B. устма-устликни ечиш (кичик полигон устувор)')
if km2(ov_before) < OV_SKIP:
    print('   қолдиқ %.4f km2 — артефакт даражасида, ўтказиб юборилди' % km2(ov_before))
    order = []
else:
    order = sorted((i for i, g in enumerate(geoms) if g is not None),
                   key=lambda i: geoms[i].area)
done_geom = []
trimmed, protected = 0, []
for i in order:
    g = geoms[i]
    if done_geom:
        hits = [done_geom[j] for j in STRtree(done_geom).query(g)]
        if hits:
            cut = clean(g.difference(unary_union(hits)))
            cut = drop_debris(cut, g.area)
            if cut is None or cut.is_empty or cut.area < g.area * MIN_KEEP:
                # Полигон бутунлай йўқолиб кетмасин — қирқмаймиз
                protected.append(i)
            elif cut.area < g.area - 1e-15:
                g = cut; trimmed += 1
    geoms[i] = g
    done_geom.append(g)
print('   қирқилди: %d полигон' % trimmed)
if protected:
    print('   тегилмади (йўқолиб кетарди): %d — %s'
          % (len(protected), ', '.join(feats[i]['properties'].get('mfy','?') for i in protected[:6])))
print()

# ---------- C: артефактларни тозалаш ----------
print('C. қолдиқ тешикларни ёпиш')
print('   ёпилди: %d' % close_gaps(geoms, 'C'))
print()

# ---------- D: четга очиладиган тирқишлар ----------
print('D. чет тирқишларини ёпиш (икки полигон орасидагилар)')
close_seams(geoms)
close_gaps(geoms, 'D')      # ёпишдан кейин майда артефакт қолиши мумкин
print()

# ---------- текшириш ----------
holes2, u2 = find_holes(geoms)
bad = sum(1 for g in geoms if g is not None and not g.is_valid)
ov_after = overlap_area(geoms)
print('=== НАТИЖА ===')
print('gaps remaining      : %d (%.4f km2)' % (len(holes2), sum(km2(h.area) for h in holes2)))
print('overlap remaining   : %.4f km2  (аввал %.2f)' % (km2(ov_after), km2(ov_before)))
print('invalid geometries  : %d' % bad)
print('area before/after   : %.1f / %.1f km2' % (km2(area_before), km2(u2.area)))
print('empty features      : %d' % sum(1 for g in geoms if g is None or g.is_empty))

# ---------- ёзиш ----------
out = []
for f, g in zip(feats, geoms):
    nf = {'properties': f['properties'], 'type': 'Feature'}
    nf['geometry'] = to_geojson(g) if (g is not None and not g.is_empty) else f['geometry']
    out.append(nf)
txt = PREFIX + json.dumps({'features': out, 'type': 'FeatureCollection'},
                          ensure_ascii=False, separators=(',', ':')) + ';'
io.open(SRC, 'w', encoding='utf-8-sig', newline='').write(txt)
print('written             : %s (%.2f MB)' % (SRC, len(txt.encode('utf-8'))/1e6))
