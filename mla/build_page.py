#!/usr/bin/env python3
"""Render the study note into a single self-contained HTML page.

Slide images are inlined as data URIs so the page works offline and as a
published artifact. Run after editing the markdown:  python3 build_page.py
"""
import base64, html, os, re, sys

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, 'MLA-1-윤활-및-오일분석-정리.md')
IMGDIR = os.path.join(HERE, 'images') + os.sep
TPL = os.path.join(HERE, 'page_template.html')
OUT = os.path.join(HERE, 'index.html')

FORMULAS = {
    'H_{min}': '<var>H</var><sub>min</sub> <span class="op">=</span> '
               '<span class="frac"><span><var>G</var><sup>0.54</sup> <var>U</var><sup>0.7</sup></span>'
               '<span><var>W</var><sup>0.13</sup></span></span>',
    '\\kappa': '<var>&kappa;</var> <span class="op">=</span> '
               '<span class="frac"><span>&nu;</span><span>&nu;<sub>1</sub></span></span>'
               '<span class="op">=</span>'
               '<span class="frac"><span>실제 운전온도에서의 점도</span><span>요구(정격) 점도</span></span>',
    '\\beta_x': '<var>&beta;</var><sub><var>x</var></sub> <span class="op">=</span> '
                '<span class="frac"><span>상류 입자수 (&gt;<var>x</var> &micro;m)</span>'
                '<span>하류 입자수 (&gt;<var>x</var> &micro;m)</span></span>'
                '<span class="op">,</span>&nbsp;&nbsp; <var>&eta;</var> <span class="op">=</span> '
                '(1 &minus; 1/<var>&beta;</var><sub><var>x</var></sub>) &times; 100&thinsp;%',
    'G\\,[': '<var>G</var> [g] <span class="op">=</span> 0.005 &times; <var>D</var> &times; <var>B</var>',
}

CAPTIONS = {
    '01': ('증점제 계통', '비누계(단순/복합)와 비비누계로 갈리는 증점제 분류, 그리고 증점제별 8개 성능 항목 매트릭스 (VG/G/F/P).'),
    '02': ('유막을 정하는 두 축', '왼쪽은 기계가 정하는 변수, 오른쪽은 윤활유가 정하는 변수. 정비 담당자가 손댈 수 있는 쪽은 대체로 오른쪽이다.'),
    '03': ('ISO VG 선정표', '베어링 속도(rpm)와 운전 온도(°C)로 읽는 점도 등급 — 속도가 오르면 낮추고, 온도가 오르면 올린다.'),
    '04': ('점도 노모그램과 κ', '평균 지름·회전수로 요구 점도 ν₁을 읽고, 점도-온도 선도에서 실제 점도 ν를 읽어 κ = ν/ν₁ 을 구한다.'),
    '05': ('API 기어오일 GL 분류', 'GL-1부터 GL-6, MT-1까지의 용도와 조성. GL-4와 GL-5의 EP 함량 차이가 핵심.'),
    '06': ('같은 등급, 다른 내용물', '산업용과 자동차용 기어오일의 조성 비교 — 첨가제 블록의 두께가 다르고, 그 안에서 AW/EP/기타로 다시 갈린다.'),
    '07': ('비용이 커지는 방향', '점도지수와 무아연 축을 따라 오른쪽 위로 갈수록 가격이 오른다.'),
    '08': ('유압유 선정 사분면', '표준/고 VI × 아연계/무아연. 위쪽은 EAL(환경 민감), 오른쪽 밖은 합성유(극한 온도).'),
    '09': ('OEM 펌프 시험 지도', '각 시험이 커버하는 압력×속도 영역. 내 장비가 어느 영역인지 보고 그 시험을 통과한 유체를 고른다.'),
    '10': ('대표 시료, 그리고 두 전략', 'Predictive는 risk-based("징후가 잡히면"), Proactive는 condition-based("원인을 미리"). 초록 구간이 압도적으로 싸다.'),
    '11': ('산화 · 질화 · 열분해', '왼쪽은 각 열화의 결과와 촉진 인자, 오른쪽 위는 산화 라디칼 사이클과 Arrhenius 규칙(10°C = 수명 절반).'),
    '12': ('가수분해와 D2619', '제조 반응이 물 때문에 역방향으로 돌아 산을 만든다. 오른쪽은 구리·물·48시간·93°C의 Coke Bottle Test.'),
    '13': ('첨가제가 사라지는 세 경로', '분해 · 분리 · 흡착. 여과와 원심분리가 "분리" 칸에 들어 있다는 점이 실무의 함정이다.'),
    '14': ('ICP 원소와 발생원', '마모 금속 / 오염물 / 첨가제 — 같은 원소가 두 칸에 걸치는 경우(Cr, Cu, Si, Na)를 구분해 읽어야 한다.'),
    '15': ('공기의 네 얼굴', '용해 · 혼입 · 거품 · 자유 공기. 오른쪽 텍스트가 혼입 공기의 피해를, 왼쪽이 ASTM D892 기포성 시험 표기법을 보여준다.'),
    '16': ('ISO 4406 코드 읽기', '>4/>6/>14 µm 입자수를 코드로 변환 — 8050 → 20, 95 → 14, 11 → 11, 즉 20/14/11.'),
    '17': ('입자 계수 4가지', 'ISO 4407(현미경), ISO 11500(광차폐 자동), 패치 테스트(종류 판별), 차압/레이저 방식.'),
    '18': ('필터와 바이패스', '표면 · 심층 · 자석. 아래 단면도의 요점은 ΔP가 한계를 넘으면 바이패스가 열려 여과가 0이 된다는 것.'),
    '19': ('치면이 말해주는 마모', '접착(scuffing) · 연삭(scoring) · 부식 · 표면 피로(micropitting) — 형태가 원인을 지목한다.'),
    '20': ('유막 두께와 스커핑', '지수를 보라: 속도 0.7, 점도 0.54, 하중 0.13. 유막은 하중보다 속도·점도에 훨씬 민감하다.'),
    '21': ('수동 재윤활의 톱니', '주입 직후 과다(붉은 위쪽), 주기 말기 과소(붉은 아래쪽). 정량 구간에 머무는 시간이 짧다.'),
    '22': ('자동 재윤활의 밴드', '잦고 작은 양으로 항상 최적 구간 안에 머문다 — 자동 윤활기를 쓰는 유일한 이유.'),
    '23': ('단일점 자동 윤활기', '가스 발생 또는 스프링으로 밀어내는 방식. 접근이 어려운 곳에 주로 붙는다.'),
    '25': ('레벨과 스플래시 윤활', '왼쪽은 일정 레벨 급유기, 오른쪽은 회전 방향(파란 화살표)에 따라 실제 유면이 달라지는 상황.'),
}


def inline(t: str) -> str:
    t = html.escape(t, quote=False)
    t = re.sub(r'`([^`]+)`', lambda m: '<code>' + m.group(1) + '</code>', t)
    t = re.sub(r'\[([^\]]+)\]\(([^)\s]+)\)',
               lambda m: '<a href="%s" target="_blank" rel="noopener">%s</a>' % (m.group(2), m.group(1)), t)
    t = re.sub(r'\*\*([^*]+)\*\*', lambda m: '<strong>' + m.group(1) + '</strong>', t)
    t = re.sub(r'(?<![\w*])\*([^*\n]+)\*(?![\w*])', lambda m: '<em>' + m.group(1) + '</em>', t)
    t = t.replace('$$', '')
    t = re.sub(r'\$([^$]+)\$', lambda m: '<code>' + m.group(1) + '</code>', t)
    return t


def img_tag(md: str) -> str:
    src = re.search(r'src="images/([^"]+)"', md).group(1)
    alt = re.search(r'alt="([^"]*)"', md)
    alt = alt.group(1) if alt else ''
    with open(IMGDIR + src, 'rb') as fh:
        b64 = base64.b64encode(fh.read()).decode()
    label, body = CAPTIONS.get(src[:2], ('슬라이드', alt))
    return ('<figure><img src="data:image/jpeg;base64,%s" alt="%s">'
            '<figcaption><b>%s</b><span>%s</span></figcaption></figure>'
            % (b64, html.escape(alt, quote=True), html.escape(label), html.escape(body)))


TAGS = [
    (re.compile(r'^\*\*원 메모(?:\(([^)]*)\))?\*\*\s*(?:—|-|:)?\s*(.*)$'), 'memo', '원 메모'),
    (re.compile(r'^\*\*보강\*\*\s*(?:—|-|:)?\s*(.*)$'), 'add', '보강'),
    (re.compile(r'^\*\*보강\s*[—-]\s*(.*)$'), 'add', '보강'),
    (re.compile(r'^⚠️\s*\*\*(?:정정|보강/정정|정정/구체화|보강/정정 필요|정정 필요)[^*]*\*\*\s*(?:—|-|:)?\s*(.*)$'), 'fix', '정정'),
    (re.compile(r'^⚠️\s*(.*)$'), 'fix', '주의'),
    (re.compile(r'^✅\s*(.*)$'), 'ok', '확인'),
]


def classify(line: str):
    """Return (kind, tag_label, remaining_text) for a provenance-marked line."""
    if line.startswith('**원 메모'):
        m = re.match(r'^\*\*원 메모\s*(\([^)]*\))?\*\*\s*(?:—|-|:)?\s*(.*)$', line)
        if m:
            label = '원 메모' + (' ' + m.group(1) if m.group(1) else '')
            return 'memo', label, m.group(2)
    if line.startswith('**보강'):
        m = re.match(r'^\*\*보강([^*]*)\*\*\s*(?:—|-|:)?\s*(.*)$', line)
        if m:
            extra = m.group(1).strip(' —-')
            label = '보강' + (' · ' + extra if extra else '')
            return 'add', label, m.group(2)
    if line.startswith('⚠️'):
        rest = line[2:].strip()
        m = re.match(r'^\*\*([^*]+)\*\*\s*(?:—|-|:)?\s*(.*)$', rest)
        if m:
            return 'fix', m.group(1).strip(), m.group(2)
        return 'fix', '주의', rest
    if line.startswith('✅'):
        return 'ok', '확인', line[1:].strip()
    return None, None, None


def is_list(s: str) -> bool:
    return bool(re.match(r'^(?:[-*] |\d+\. )', s))


def take_list(md_lines, i):
    """Consume one list block starting at line i; return (html, next_index)."""
    n = len(md_lines)
    ordered = bool(re.match(r'^\d+\. ', md_lines[i].strip()))
    items = []
    while i < n:
        cur = md_lines[i].rstrip('\n')
        if re.match(r'^\s*(?:[-*]|\d+\.)\s', cur):
            items.append((len(cur) - len(cur.lstrip()), re.sub(r'^\s*(?:[-*]|\d+\.)\s+', '', cur)))
        elif cur.strip() and cur.startswith(('  ', '\t')) and items:
            items[-1] = (items[-1][0], items[-1][1] + ' ' + cur.strip())
        else:
            break
        i += 1
    tag = 'ol' if ordered else 'ul'
    buf, depth, base = ['<%s>' % tag], 0, items[0][0]
    for ind, txt in items:
        lvl = 1 if ind > base else 0
        if lvl > depth:
            buf.append('<ul>')
        elif lvl < depth:
            buf.append('</ul>')
        depth = lvl
        buf.append('<li>%s</li>' % inline(txt))
    buf.append('</ul>' * depth)
    buf.append('</%s>' % tag)
    return ''.join(buf), i


def render(md_lines):
    out, nav = [], []
    i, n = 0, len(md_lines)
    sec = 0
    # state for open callout
    open_note = None

    def close_note():
        nonlocal open_note
        if open_note:
            out.append('</div>')
            open_note = None

    while i < n:
        line = md_lines[i].rstrip('\n')
        s = line.strip()

        if s.startswith('## 목차'):                      # nav is generated instead
            i += 1
            while i < n and not md_lines[i].startswith('---'):
                i += 1
            i += 1
            continue

        if not s or s == '---':
            i += 1
            continue

        if s.startswith('# '):                            # title lives in the masthead
            i += 1
            while i < n and (md_lines[i].startswith('>') or not md_lines[i].strip()):
                i += 1
            continue

        if s.startswith('## '):
            close_note()
            sec += 1
            body = s[3:]
            m = re.match(r'^(\d+)\.\s*(.*)$', body)
            num, txt = (m.group(1), m.group(2)) if m else ('', body)
            sid = 's%d' % sec
            out.append('<h2 id="%s"><span class="num">%s</span>%s</h2>'
                       % (sid, ('SECTION ' + num.zfill(2)) if num else '', inline(txt)))
            nav.append('<li><a href="#%s"><span>%s</span>%s</a></li>'
                       % (sid, num.zfill(2) if num else '·', inline(txt)))
            i += 1
            continue

        if s.startswith('### '):
            close_note()
            body = s[4:]
            m = re.match(r'^([\d.]+)\s+(.*)$', body)
            hid = 'h' + re.sub(r'\W+', '', m.group(1)) if m else None
            if m:
                out.append('<h3%s><span class="num">%s</span>%s</h3>'
                           % (' id="%s"' % hid, m.group(1), inline(m.group(2))))
            else:
                out.append('<h3>%s</h3>' % inline(body))
            i += 1
            continue

        if s.startswith('<img '):
            block = s
            while '>' not in block:
                i += 1
                block += md_lines[i].strip()
            close_note()
            out.append(img_tag(block))
            i += 1
            continue

        if s.startswith('$$'):
            block = []
            while i < n:
                block.append(md_lines[i].strip())
                if md_lines[i].strip().endswith('$$') and (len(block) > 1 or md_lines[i].strip() != '$$'):
                    break
                i += 1
            tex = ' '.join(block)
            rendered = next((v for k, v in FORMULAS.items() if k in tex), None)
            out.append('<div class="formula">%s</div>' % (rendered or html.escape(tex.strip('$ '))))
            i += 1
            continue

        if s.startswith('```'):
            i += 1
            buf = []
            while i < n and not md_lines[i].startswith('```'):
                buf.append(md_lines[i].rstrip('\n'))
                i += 1
            i += 1
            out.append('<pre><code>%s</code></pre>' % html.escape('\n'.join(buf)))
            continue

        if s.startswith('|'):                            # table
            close_note()
            rows = []
            while i < n and md_lines[i].strip().startswith('|'):
                rows.append([c.strip() for c in md_lines[i].strip().strip('|').split('|')])
                i += 1
            head, body = rows[0], rows[2:] if len(rows) > 1 and set(rows[1][0]) <= set('-: ') else rows[1:]
            t = ['<div class="tablewrap"><table><thead><tr>']
            t += ['<th>%s</th>' % inline(c) for c in head]
            t.append('</tr></thead><tbody>')
            for r in body:
                t.append('<tr>' + ''.join('<td>%s</td>' % inline(c) for c in r) + '</tr>')
            t.append('</tbody></table></div>')
            out.append(''.join(t))
            continue

        if s.startswith('> '):
            close_note()
            buf = []
            while i < n and md_lines[i].strip().startswith('>'):
                buf.append(md_lines[i].strip().lstrip('>').strip())
                i += 1
            paras = ' '.join(buf).split('  ')
            out.append('<blockquote>%s</blockquote>'
                       % ''.join('<p>%s</p>' % inline(p) for p in paras if p.strip()))
            continue

        if is_list(s):
            close_note()
            frag, i = take_list(md_lines, i)
            out.append(frag)
            continue

        kind, label, rest = classify(s)
        if kind:
            close_note()
            out.append('<div class="note note--%s"><span class="note__tag">%s</span>' % (kind, html.escape(label)))
            open_note = kind
            para = [rest.strip()] if rest.strip() else []
            i += 1
            # absorb the immediately following lines that belong to this callout
            while i < n:
                st = md_lines[i].strip()
                if not st or st.startswith(('#', '|', '>', '<img', '$$', '```', '---')) or classify(st)[0]:
                    break
                if is_list(st):
                    # a list directly under the marker is part of the callout
                    if para:
                        out.append('<p>%s</p>' % inline(' '.join(para)))
                        para = []
                    frag, i = take_list(md_lines, i)
                    out.append(frag)
                    break
                para.append(st)
                i += 1
            if para:
                out.append('<p>%s</p>' % inline(' '.join(para)))
            close_note()
            continue

        close_note()
        # markdown wraps paragraphs across lines; a blank line ends the paragraph
        para = [s]
        i += 1
        while i < n:
            st = md_lines[i].strip()
            if not st or st.startswith(('#', '|', '>', '<img', '$$', '```', '---')) \
               or is_list(st) or classify(st)[0]:
                break
            para.append(st)
            i += 1
        out.append('<p>%s</p>' % inline(' '.join(para)))

    close_note()
    return '\n'.join(out), '\n'.join(nav)


def main():
    with open(SRC, encoding='utf-8') as fh:
        lines = fh.readlines()
    body, nav = render(lines)
    body += ('\n<footer>원본: Lube_Win_MLA_1.docx (YouTube MLA 1 강의 메모) · '
             '정리 및 보강 2026-08 · 슬라이드 캡처 24장 · 표준 인용은 본문 22장 참고 자료 참조</footer>')
    with open(TPL, encoding='utf-8') as fh:
        tpl = fh.read()
    out = tpl.replace('<!--NAV-->', nav).replace('<!--CONTENT-->', body)
    with open(OUT, 'w', encoding='utf-8') as fh:
        fh.write(out)
    print('wrote %s  (%.1f MB)' % (OUT, os.path.getsize(OUT) / 1e6))


if __name__ == '__main__':
    main()
