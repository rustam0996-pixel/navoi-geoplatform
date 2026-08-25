# -*- coding: utf-8 -*-
"""
navoi_geo.js — қўшни полигонлар орасидаги тирқишларни (slivers) ёпади.

Муаммо: маҳалла полигонлари бир-биридан мустақил рақамланган, шу боис
чегаралар мос тушмайди — харитада туманлар орасида очиқ жойлар кўринади.

Мантиқ:
  1. Ҳар feature шапели геометрияга ўтказилади (PowerShell'нинг {value,Count}
     ўрами ечилади, нотўғри полигонлар make_valid билан тузатилади).
  2. set_precision билан ҳамма координата 1e-6 (~0.11 м) тўрига ўтказилади —
     бу GEOS'нинг snap-rounding'и, геометрия ЯРОҚЛИ бўлиб қолади.
  3. unary_union олинади → ичидаги "тешик"лар айнан ўша очиқ жойлар.
  4. Ҳар тешик ўзи билан энг узун чегарани бўлишадиган featureга қўшилади,
     шунинг учун у ўралиб турган туманга тушади — харита бузилмайди.
  5. Натижа яна тўрга ўтказилиб ёзилади.

IDEMPOTENT: иккинчи марта ишга туширилса геометрия ўзгармайди (тешик 0,
майдон ўша-ўша). Файл байтлари бир оз фарқ қилиши мумкин — бу фақат
полигон тепа нуқталарининг тартиби, шакл эмас.

`parse_kml.ps1` navoi_geo.js ни қайта яратгандан СЎНГ шуни ишлатинг:
    python fix_geo_gaps.py
Талаб: pip install shapely
"""
import io, json, math, collections, sys
import shapely

# Windows консоли (cp1251) ўзбекча ҳарфларни чиқара олмайди — UTF-8 га ўтказамиз
try: sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception: pass
from shapely.geometry import shape, mapping, Polygon
from shapely.ops import unary_union
from shapely import make_valid, STRtree

SRC    = 'navoi_geo.js'
PREFIX = 'window.NAVOI_GEOJSON = '
GRID   = 1e-6      # ~0.11 м — координата тўри
ND     = 6         # GRID га мос ўнлик хона сони
GROW   = 0.0006    # ~50 м — тешик атрофида қўшни излаш оралиғи

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
    """Яроқли + тўрга ўтказилган полигон(лар) қайтаради."""
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
        return {'type':'Polygon','coordinates':[rr(r) for r in m['coordinates']]}
    return {'type':'MultiPolygon',
            'coordinates':[[rr(r) for r in poly] for poly in m['coordinates']]}

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

print('features            :', len(feats))
print('invalid -> repaired :', invalid)

# ---------- тешикларни топиш ва ёпиш ----------
# Snap-rounding ҳар ўтишда жуда кичик (бир неча м2) артефактлар қолдириши
# мумкин, шу боис ҳеч тешик қолмагунча такрорлаймиз. Амалда 3-4 ўтиш кифоя.
u = unary_union([g for g in geoms if g is not None])
area_before = u.area
total_closed, total_area = 0, 0.0

for it in range(1, 9):
    u_now = unary_union([g for g in geoms if g is not None])
    holes = [Polygon(r) for p in polys_of(u_now) for r in p.interiors]
    holes = [h for h in holes if h.area > 0]
    if not holes:
        print('pass %d              : тешик йўқ — тугади' % it)
        break
    print('pass %d              : %d gap (%.3f km2)' % (it, len(holes), sum(km2(h.area) for h in holes)))

    parts_idx, parts_geom = [], []
    for i, g in enumerate(geoms):
        for p in polys_of(g):
            parts_idx.append(i); parts_geom.append(p)
    tree = STRtree(parts_geom)

    assign = collections.defaultdict(list)
    unassigned = 0
    for h in holes:
        hb = h.boundary.buffer(GROW / 3)
        best, best_len = None, 0.0
        for j in tree.query(h.buffer(GROW)):
            inter = parts_geom[j].boundary.intersection(hb)
            L = inter.length if not inter.is_empty else 0.0
            if L > best_len: best_len, best = L, parts_idx[j]
        if best is None: unassigned += 1
        else: assign[best].append(h)

    if not assign:
        print('  қўшни топилмади — тўхтадик (%d gap қолди)' % unassigned)
        break
    for i, hs in assign.items():
        geoms[i] = clean(unary_union([geoms[i]] + hs))
    total_closed += sum(len(v) for v in assign.values())
    total_area   += sum(km2(h.area) for h in holes)

print('gaps closed (total) :', total_closed)

# ---------- текшириш ----------
u2 = unary_union([g for g in geoms if g is not None])
holes2 = [Polygon(r) for p in polys_of(u2) for r in p.interiors if Polygon(r).area > 0]
bad = sum(1 for g in geoms if g is not None and not g.is_valid)
print('gaps remaining      :', len(holes2), '(%.4f km2)' % sum(km2(h.area) for h in holes2))
print('invalid geometries  :', bad)
print('area before/after   : %.1f / %.1f km2' % (km2(area_before), km2(u2.area)))

# ---------- ёзиш ----------
out = []
for f, g in zip(feats, geoms):
    nf = {'properties': f['properties'], 'type': 'Feature'}
    nf['geometry'] = to_geojson(g) if g is not None else f['geometry']
    out.append(nf)
txt = PREFIX + json.dumps({'features': out, 'type': 'FeatureCollection'},
                          ensure_ascii=False, separators=(',', ':')) + ';'
io.open(SRC, 'w', encoding='utf-8-sig', newline='').write(txt)
print('written             : %s (%.2f MB)' % (SRC, len(txt.encode('utf-8'))/1e6))
