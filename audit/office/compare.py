# 정답(MS Office PDF)과 우리 화면을 나란히 놓고 점수를 낸다.
# 사용: python audit/office/compare.py <출력 폴더(docs/quality-audit/out)>
#  - sbs/<이름>-view.png : 왼쪽 Office, 오른쪽 우리 보기 화면이 그리는 PDF(LibreOffice 변환본) — 쪽마다 한 줄
#  - sbs/<이름>-edit.png : 왼쪽 Office 쪽들, 오른쪽 편집 화면 그림
#  - scores.json / scores.md : 쪽 수 일치, 글자 일치율(본문 글자 순서 비교), 모양 일치율(쪽 그림 차이) — 같은 명령으로 다시 재면 고친 효과가 숫자로 나온다
import difflib, json, os, re, sys
import pymupdf
from PIL import Image, ImageChops, ImageDraw, ImageOps, ImageStat

OUT = sys.argv[1] if len(sys.argv) > 1 else 'docs/quality-audit/out'
SBS = os.path.join(OUT, 'sbs'); os.makedirs(SBS, exist_ok=True)
manifest = []
for f in sorted(os.listdir(OUT)):
    if f.startswith('manifest') and f.endswith('.json'):
        manifest += json.load(open(os.path.join(OUT, f), encoding='utf-8'))


def pages(pdf, zoom=1.0):
    if not os.path.exists(pdf): return []
    doc = pymupdf.open(pdf)
    out = []
    for p in doc:
        pix = p.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), alpha=False)
        out.append(Image.frombytes('RGB', (pix.width, pix.height), pix.samples))
    return out


def text_of(pdf):
    if not os.path.exists(pdf): return ''
    return ''.join(p.get_text() for p in pymupdf.open(pdf))


norm = lambda s: re.sub(r'\s+', '', s or '')


def text_ratio(a, b):
    a, b = norm(a), norm(b)
    if not a and not b: return 1.0
    return difflib.SequenceMatcher(None, a, b, autojunk=False).ratio()


def look_ratio(a, b):
    # 잉크(글자·선) 겹침 점수 1(같음)~0. 두 그림을 같은 폭으로 줄이고 잉크를 몇 픽셀 번지게 해서 서로 덮는 비율(F1)을 낸다 —
    # 한두 픽셀 밀림은 같다고 보고, 빠진 그림·줄 바뀜·쪽 밀림은 깎인다
    from PIL import ImageFilter
    size = (400, max(1, int(400 * a.height / a.width)))
    ink = lambda im: ImageOps.grayscale(im.resize(size)).point(lambda v: 255 if v < 200 else 0)
    ia, ib = ink(a), ink(b)
    da, db = ia.filter(ImageFilter.MaxFilter(5)), ib.filter(ImageFilter.MaxFilter(5))
    cnt = lambda im: sum(im.histogram()[255:])
    na, nb = cnt(ia), cnt(ib)
    if na == 0 and nb == 0: return 1.0
    if na == 0 or nb == 0: return 0.0
    recall = cnt(ImageChops.multiply(ia, db)) / na; precision = cnt(ImageChops.multiply(ib, da)) / nb
    return 0.0 if recall + precision == 0 else 2 * recall * precision / (recall + precision)


def label(im, text, color=(200, 0, 0)):
    d = ImageDraw.Draw(im); d.rectangle([0, 0, 9 * len(text) + 12, 18], fill=(255, 255, 255)); d.text((4, 3), text, fill=color)
    return im


def hstack(ims, gap=12, bg=(150, 150, 150)):
    ims = [i for i in ims if i is not None]
    h = max(i.height for i in ims); w = sum(i.width for i in ims) + gap * (len(ims) - 1)
    out = Image.new('RGB', (w, h), bg); x = 0
    for i in ims: out.paste(i, (x, 0)); x += i.width + gap
    return out


def vstack(ims, gap=12, bg=(150, 150, 150)):
    ims = [i for i in ims if i is not None]
    if not ims: return None
    w = max(i.width for i in ims); h = sum(i.height for i in ims) + gap * (len(ims) - 1)
    out = Image.new('RGB', (w, h), bg); y = 0
    for i in ims: out.paste(i, (0, y)); y += i.height + gap
    return out


def fit_w(im, w):
    return im if im.width == w else im.resize((w, max(1, int(im.height * w / im.width))))


def save(im, path, max_h=12000):
    if im.height > max_h: im = im.resize((int(im.width * max_h / im.height), max_h))
    im.save(path, optimize=True)


rows = []
for e in manifest:
    name, ext = e['name'], e['ext']
    ref_pdf = os.path.join(OUT, 'ref', name + '.pdf'); lo_pdf = os.path.join(OUT, 'lo', name + '.pdf')
    ref = pages(ref_pdf); row = {'name': name, 'ext': ext, 'errors': e.get('errors', []), 'refPages': len(ref), 'convertMs': e.get('convertMs')}
    if not ref: row['note'] = 'Office 정답 없음'
    if ext in ('docx', 'pptx'):
        lo = pages(lo_pdf); row['viewPages'] = len(lo)
        n = max(len(ref), len(lo))
        if ref and lo:
            row['viewText'] = round(text_ratio(text_of(ref_pdf), text_of(lo_pdf)), 3)
            looks = [look_ratio(ref[i], lo[i].resize(ref[i].size)) for i in range(min(len(ref), len(lo)))]
            row['viewLook'] = round(sum(looks) / n, 3)  # 짝 없는 쪽은 0점
            pairs = []
            for i in range(n):
                a = label(ref[i].copy(), f'Office p{i + 1}') if i < len(ref) else Image.new('RGB', ref[0].size, (255, 220, 220))
                b = label(lo[i].copy(), f'LibreOffice(view) p{i + 1}', (0, 0, 200)) if i < len(lo) else Image.new('RGB', a.size, (255, 220, 220))
                pairs.append(hstack([a, fit_w(b, a.width) if abs(b.width - a.width) > 2 else b]))
            save(vstack(pairs), os.path.join(SBS, f'{name}-view.png'))
        # 편집 화면
        edit = e.get('edit') or {}
        if ext == 'docx':
            shot = os.path.join(OUT, 'edit', name + '.png')
            if ref and os.path.exists(shot):
                row['editText'] = round(text_ratio(text_of(ref_pdf), edit.get('text', '')), 3)
                left = vstack([label(p.copy(), f'Office p{i + 1}') for i, p in enumerate(ref)])
                right = label(fit_w(Image.open(shot).convert('RGB'), left.width), 'editor (ours)', (0, 120, 0))
                save(hstack([left, right]), os.path.join(SBS, f'{name}-edit.png'))
        else:
            slides = [os.path.join(OUT, 'edit', f'{name}-s{i + 1}.png') for i in range(edit.get('slides', 0))]
            row['editSlides'] = len(slides)
            if ref and slides:
                ref_txt = [p.get_text() for p in pymupdf.open(ref_pdf)]
                row['editText'] = round(sum(text_ratio(ref_txt[i], edit['texts'][i]) if i < len(edit['texts']) else 0 for i in range(len(ref_txt))) / max(len(ref_txt), len(slides)), 3)
                looks, pairs = [], []
                for i in range(max(len(ref), len(slides))):
                    a = label(ref[i].copy(), f'Office s{i + 1}') if i < len(ref) else None
                    b = Image.open(slides[i]).convert('RGB') if i < len(slides) and os.path.exists(slides[i]) else None
                    if a and b:
                        b = b.resize(a.size); looks.append(look_ratio(a, b))
                    pairs.append(hstack([a, label(b, 'editor (ours)', (0, 120, 0)) if b else None]))
                row['editLook'] = round(sum(looks) / max(len(ref), len(slides)), 3)
                save(vstack(pairs), os.path.join(SBS, f'{name}-edit.png'))
    else:  # xlsx: 정답 쪽들 | 보기 격자 | 편집 격자 (시트마다)
        view = e.get('view') or []; edit = e.get('edit') or []
        row['sheetsView'] = len(view); row['sheetsEdit'] = len(edit)
        left = vstack([label(p.copy(), f'Excel p{i + 1}') for i, p in enumerate(ref)]) if ref else None
        cols = []
        for i in range(max(len(view), len(edit))):
            v = label(Image.open(view[i]['file']).convert('RGB'), f'viewer {view[i]["tab"]}', (0, 0, 200)) if i < len(view) else None
            ed = label(Image.open(edit[i]['file']).convert('RGB'), f'editor {edit[i]["tab"]}', (0, 120, 0)) if i < len(edit) else None
            cols.append(vstack([v, ed]))
        right = vstack(cols)
        if left is not None and right is not None:
            save(hstack([left, right]), os.path.join(SBS, f'{name}-grid.png'))
    row['_entry'] = e
    rows.append(row)

# 점검표(checklist.json): Office에 보이는 글자·그림이 우리 화면에도 있는가
check = json.load(open(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'checklist.json'), encoding='utf-8'))['items']
by_name = {r['name']: r for r in rows}
results = []
for it in check:
    r = by_name.get(it['doc'])
    if not r: continue
    e = r['_entry']
    for where in it['where']:
        if where == 'editImages':
            n = (e.get('edit') or {}).get('images', 0)
            results.append({'id': it['id'], 'doc': it['doc'], 'where': where, 'item': f">={it['min']} 그림", 'ok': n >= it['min'], 'got': n}); continue
        if r['ext'] == 'xlsx':
            txt = '\n'.join(s.get('text', '') for s in (e.get(where) or []))
        elif where == 'view':
            txt = text_of(os.path.join(OUT, 'lo', r['name'] + '.pdf'))
        else:
            ed = e.get('edit') or {}; txt = ed.get('text') or '\n'.join(ed.get('texts', []))
        sq = norm(txt)
        for h in it.get('has', []):
            results.append({'id': it['id'], 'doc': it['doc'], 'where': where, 'item': h, 'ok': norm(h) in sq})
        for h in it.get('not', []):
            results.append({'id': it['id'], 'doc': it['doc'], 'where': where, 'item': 'not ' + h, 'ok': norm(h) not in sq})
# PPT 편집 화면 DOM 점검(pptx_checks.json, capture.mjs → pptx_dom.mjs가 잰 결과)
for r in rows:
    ed = r['_entry'].get('edit')
    for c in ((ed.get('dom') if isinstance(ed, dict) else None) or []):
        results.append({'id': c['id'], 'doc': c['doc'], 'where': 'editDom', 'item': f"s{c['slide']} {c['item']}", 'ok': c['ok'], 'got': {k: v for k, v in c['got'].items() if k != 'found'} if c['got'].get('found') else '없음'})
for r in rows: r.pop('_entry', None)

def mean(xs):
    xs = [x for x in xs if x is not None]
    return round(sum(xs) / len(xs), 3) if xs else None
summary = {}
for ext in ('docx', 'pptx', 'xlsx'):
    rs = [r for r in rows if r['ext'] == ext]
    cs = [c for c in results if by_name[c['doc']]['ext'] == ext]
    summary[ext] = {
        'viewLook': mean([r.get('viewLook') for r in rs]), 'editLook': mean([r.get('editLook') for r in rs]),
        'viewText': mean([r.get('viewText') for r in rs]), 'editText': mean([r.get('editText') for r in rs]),
        'checksView': f"{sum(c['ok'] for c in cs if c['where'] == 'view')}/{sum(1 for c in cs if c['where'] == 'view')}",
        'checksEdit': f"{sum(c['ok'] for c in cs if c['where'] != 'view')}/{sum(1 for c in cs if c['where'] != 'view')}",
    }
json.dump({'summary': summary, 'documents': rows, 'checks': results}, open(os.path.join(OUT, 'scores.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=2)
lines = ['## 형식별 요약', '', '| 형식 | 보기 모양 | 편집 모양 | 보기 글자 | 편집 글자 | 점검표(보기) | 점검표(편집) |', '|---|---|---|---|---|---|---|']
for ext, v in summary.items():
    lines.append(f"| {ext} | {v['viewLook']} | {v['editLook']} | {v['viewText']} | {v['editText']} | {v['checksView']} | {v['checksEdit']} |")
lines += ['', '## 문서별', '', '| 문서 | 정답 쪽 | 보기 쪽/시트 | 보기 글자 | 보기 모양 | 편집 글자 | 편집 모양 | 오류 |', '|---|---|---|---|---|---|---|---|']
for r in rows:
    g = lambda k: '' if r.get(k) is None else str(r[k])
    lines.append(f"| {r['name']}.{r['ext']} | {g('refPages')} | {g('viewPages') or g('sheetsView')} | {g('viewText')} | {g('viewLook')} | {g('editText')} | {g('editLook')} | {len(r['errors'])} |")
lines += ['', '## 점검표 실패 항목', '', '| 결함 | 문서 | 화면 | 항목 |', '|---|---|---|---|']
for c in results:
    if not c['ok']: lines.append(f"| {c['id']} | {c['doc']} | {c['where']} | {c['item']}{' (' + str(c['got']) + ')' if 'got' in c else ''} |")
open(os.path.join(OUT, 'scores.md'), 'w', encoding='utf-8').write('\n'.join(lines) + '\n')
print('\n'.join(lines[:7]))
