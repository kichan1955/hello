from docx import Document
import re, html as esc

doc = Document("/root/.claude/uploads/36a16a87-ae6a-5826-a0db-a084126bbb76/d8073541-____________________200p.docx")
all_paras = [para.text.strip() for para in doc.paragraphs]

def get_range(start, end):
    return [t for t in all_paras[start:end] if t.strip()]

def paras_html(texts):
    out = []
    for t in texts:
        e = esc.escape(t)
        # subsection headings like "1-1. ..."
        if re.match(r'^\d+[-–]\d+\.\s', t) or re.match(r'^[①②③④⑤]', t):
            out.append(f'<h3 class="sub-heading">{e}</h3>')
        elif re.match(r'^◆|^【|^\[0[0-9]\]|^\[[0-9]+\]', t):
            out.append(f'<h4 class="box-heading">{e}</h4>')
        elif re.match(r'^▶', t):
            out.append(f'<p class="practice">{e}</p>')
        elif re.match(r'^—\s', t):
            out.append(f'<p class="quote-attr">{e}</p>')
        elif len(t) < 60 and t.endswith('다') == False and not t.endswith('?') and re.match(r'^[가-힣A-Za-z]', t) and '.' not in t[:40]:
            out.append(f'<p class="bold-line">{e}</p>')
        else:
            out.append(f'<p>{e}</p>')
    return '\n'.join(out)

sections = {
    'recommend': get_range(19, 30),
    'preface':   get_range(30, 45),
    'prologue':  get_range(79, 101),
    'ch1':       get_range(101, 131),
    'ch2':       get_range(131, 157),
    'ch3':       get_range(157, 177),
    'ch4':       get_range(177, 201),
    'ch5':       get_range(201, 218),
    'ch6':       get_range(218, 231),
    'ch7':       get_range(231, 241),
    'ch8':       get_range(241, 251),
    'ch9':       get_range(251, 267),
    'ch10':      get_range(267, 289),
    'ch13_14':   get_range(289, 319),
    'ch15':      get_range(319, 336),
    'ch16':      get_range(336, 353),
    'epilogue':  get_range(353, 370),
    'q100':      get_range(370, 410),
}

# ── SVG illustrations (one per major section) ──────────────────────────────

SVG = {}

# 추천의 글 – 별이 빛나는 펼쳐진 책
SVG['recommend'] = '''<svg viewBox="0 0 800 380" xmlns="http://www.w3.org/2000/svg">
  <rect width="800" height="380" fill="#0d1b2a"/>
  <rect x="160" y="80" width="220" height="220" rx="6" fill="#1a2d42" stroke="#c9a84c" stroke-width="2"/>
  <rect x="420" y="80" width="220" height="220" rx="6" fill="#1a2d42" stroke="#c9a84c" stroke-width="2"/>
  <rect x="375" y="75" width="50" height="230" fill="#c9a84c" opacity="0.3"/>
  <line x1="400" y1="80" x2="400" y2="300" stroke="#c9a84c" stroke-width="3"/>
  <line x1="180" y1="120" x2="360" y2="120" stroke="#c9a84c" stroke-width="1" opacity="0.5"/>
  <line x1="180" y1="145" x2="360" y2="145" stroke="#e0c57a" stroke-width="1" opacity="0.3"/>
  <line x1="180" y1="170" x2="360" y2="170" stroke="#e0c57a" stroke-width="1" opacity="0.3"/>
  <line x1="180" y1="195" x2="360" y2="195" stroke="#e0c57a" stroke-width="1" opacity="0.3"/>
  <line x1="180" y1="220" x2="360" y2="220" stroke="#e0c57a" stroke-width="1" opacity="0.3"/>
  <line x1="440" y1="120" x2="620" y2="120" stroke="#c9a84c" stroke-width="1" opacity="0.5"/>
  <line x1="440" y1="145" x2="620" y2="145" stroke="#e0c57a" stroke-width="1" opacity="0.3"/>
  <line x1="440" y1="170" x2="620" y2="170" stroke="#e0c57a" stroke-width="1" opacity="0.3"/>
  <line x1="440" y1="195" x2="620" y2="195" stroke="#e0c57a" stroke-width="1" opacity="0.3"/>
  <circle cx="120" cy="60" r="4" fill="#c9a84c" opacity="0.8"/>
  <circle cx="680" cy="40" r="3" fill="#e0c57a" opacity="0.7"/>
  <circle cx="700" cy="100" r="2" fill="#c9a84c" opacity="0.6"/>
  <circle cx="80" cy="140" r="3" fill="#e0c57a" opacity="0.5"/>
  <circle cx="740" cy="200" r="4" fill="#c9a84c" opacity="0.7"/>
  <circle cx="60" cy="260" r="2" fill="#e0c57a" opacity="0.6"/>
  <circle cx="720" cy="320" r="3" fill="#c9a84c" opacity="0.5"/>
  <text x="400" y="345" text-anchor="middle" fill="#c9a84c" font-size="16" font-family="serif" opacity="0.9">추천의 글</text>
</svg>'''

# 저자서문 – 강단의 교수
SVG['preface'] = '''<svg viewBox="0 0 800 380" xmlns="http://www.w3.org/2000/svg">
  <defs><linearGradient id="pg" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#0d1b2a"/><stop offset="1" stop-color="#1a3a5c"/></linearGradient></defs>
  <rect width="800" height="380" fill="url(#pg)"/>
  <!-- lectern -->
  <polygon points="340,280 460,280 480,310 320,310" fill="#2e4a6a" stroke="#c9a84c" stroke-width="1.5"/>
  <rect x="340" y="200" width="120" height="82" fill="#1a2d42" stroke="#c9a84c" stroke-width="1.5"/>
  <!-- person -->
  <circle cx="400" cy="155" r="30" fill="#c9a84c" opacity="0.85"/>
  <rect x="375" y="185" width="50" height="15" fill="#c9a84c" opacity="0.8"/>
  <!-- paper on lectern -->
  <rect x="355" y="215" width="90" height="55" fill="#faf8f3" rx="2" opacity="0.9"/>
  <line x1="362" y1="228" x2="438" y2="228" stroke="#555" stroke-width="1"/>
  <line x1="362" y1="240" x2="438" y2="240" stroke="#555" stroke-width="1"/>
  <line x1="362" y1="252" x2="415" y2="252" stroke="#555" stroke-width="1"/>
  <!-- audience dots -->
  <circle cx="200" cy="330" r="12" fill="#2e4a6a" stroke="#c9a84c" stroke-width="1"/>
  <circle cx="240" cy="340" r="12" fill="#2e4a6a" stroke="#c9a84c" stroke-width="1"/>
  <circle cx="280" cy="330" r="12" fill="#2e4a6a" stroke="#c9a84c" stroke-width="1"/>
  <circle cx="510" cy="330" r="12" fill="#2e4a6a" stroke="#c9a84c" stroke-width="1"/>
  <circle cx="550" cy="340" r="12" fill="#2e4a6a" stroke="#c9a84c" stroke-width="1"/>
  <circle cx="590" cy="330" r="12" fill="#2e4a6a" stroke="#c9a84c" stroke-width="1"/>
  <circle cx="320" cy="345" r="12" fill="#2e4a6a" stroke="#c9a84c" stroke-width="1"/>
  <circle cx="360" cy="355" r="12" fill="#2e4a6a" stroke="#c9a84c" stroke-width="1"/>
  <circle cx="440" cy="355" r="12" fill="#2e4a6a" stroke="#c9a84c" stroke-width="1"/>
  <circle cx="480" cy="345" r="12" fill="#2e4a6a" stroke="#c9a84c" stroke-width="1"/>
  <text x="400" y="50" text-anchor="middle" fill="#c9a84c" font-size="20" font-family="serif">저자서문</text>
  <text x="400" y="80" text-anchor="middle" fill="#e0c57a" font-size="13" font-family="serif" opacity="0.8">히든 챔피언의 히든 스토리</text>
</svg>'''

# 프롤로그 – 공중전화 부스
SVG['prologue'] = '''<svg viewBox="0 0 800 380" xmlns="http://www.w3.org/2000/svg">
  <defs><linearGradient id="sky" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#1a3a5c"/><stop offset="1" stop-color="#0d1b2a"/></linearGradient></defs>
  <rect width="800" height="380" fill="url(#sky)"/>
  <!-- ground -->
  <rect x="0" y="310" width="800" height="70" fill="#0d1b2a"/>
  <!-- phone booth -->
  <rect x="320" y="90" width="100" height="220" fill="#1a2d42" stroke="#c9a84c" stroke-width="2"/>
  <rect x="330" y="100" width="80" height="160" fill="#2e4a6a" opacity="0.5" rx="2"/>
  <rect x="340" y="255" width="60" height="55" fill="#1a2d42"/>
  <!-- phone on wall -->
  <rect x="348" y="170" width="44" height="60" fill="#0d1b2a" rx="3" stroke="#c9a84c" stroke-width="1"/>
  <circle cx="370" cy="185" r="12" fill="none" stroke="#c9a84c" stroke-width="1.5"/>
  <rect x="356" y="205" width="28" height="18" fill="#c9a84c" opacity="0.3" rx="2"/>
  <!-- coin slot -->
  <rect x="362" y="162" width="16" height="4" fill="#c9a84c" rx="1"/>
  <!-- 5-won coin -->
  <circle cx="500" cy="200" r="22" fill="none" stroke="#c9a84c" stroke-width="3"/>
  <text x="500" y="205" text-anchor="middle" fill="#c9a84c" font-size="14" font-weight="bold">5원</text>
  <!-- person silhouette -->
  <circle cx="280" cy="175" r="22" fill="#c9a84c" opacity="0.7"/>
  <rect x="265" y="197" width="30" height="55" fill="#c9a84c" opacity="0.6" rx="3"/>
  <rect x="255" y="210" width="12" height="38" fill="#c9a84c" opacity="0.6" rx="3"/>
  <rect x="297" y="207" width="12" height="40" fill="#c9a84c" opacity="0.6" rx="3"/>
  <!-- thought bubble -->
  <circle cx="470" cy="130" r="5" fill="#c9a84c" opacity="0.6"/>
  <circle cx="490" cy="115" r="8" fill="#c9a84c" opacity="0.6"/>
  <circle cx="515" cy="100" r="13" fill="#c9a84c" opacity="0.6"/>
  <circle cx="545" cy="88" r="20" fill="#1a2d42" stroke="#c9a84c" stroke-width="1.5"/>
  <text x="545" y="83" text-anchor="middle" fill="#c9a84c" font-size="8">자판기</text>
  <text x="545" y="95" text-anchor="middle" fill="#c9a84c" font-size="8">사업!</text>
  <!-- year label -->
  <text x="400" y="365" text-anchor="middle" fill="#c9a84c" font-size="14" opacity="0.8">1974년 여름, 영등포역</text>
</svg>'''

# 제1장 – 소년 이한구 (아이스케키 장수)
SVG['ch1'] = '''<svg viewBox="0 0 800 380" xmlns="http://www.w3.org/2000/svg">
  <defs><linearGradient id="bg1" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#1a3a5c"/><stop offset="1" stop-color="#0d1b2a"/></linearGradient></defs>
  <rect width="800" height="380" fill="url(#bg1)"/>
  <rect x="0" y="310" width="800" height="70" fill="#142236"/>
  <!-- sun -->
  <circle cx="650" cy="70" r="45" fill="#c9a84c" opacity="0.25"/>
  <circle cx="650" cy="70" r="32" fill="#c9a84c" opacity="0.4"/>
  <!-- boy with ice cream cart -->
  <circle cx="260" cy="200" r="22" fill="#e0c57a" opacity="0.8"/>
  <rect x="247" y="222" width="26" height="50" fill="#e0c57a" opacity="0.7" rx="3"/>
  <!-- cart -->
  <rect x="290" y="250" width="100" height="55" fill="#2e4a6a" stroke="#c9a84c" stroke-width="2" rx="4"/>
  <text x="340" y="283" text-anchor="middle" fill="#c9a84c" font-size="14" font-weight="bold">아이스케키</text>
  <circle cx="305" cy="312" r="10" fill="#1a2d42" stroke="#c9a84c" stroke-width="2"/>
  <circle cx="375" cy="312" r="10" fill="#1a2d42" stroke="#c9a84c" stroke-width="2"/>
  <line x1="275" y1="268" x2="292" y2="268" stroke="#c9a84c" stroke-width="2"/>
  <!-- poverty home outline -->
  <polygon points="480,170 560,120 640,170" fill="none" stroke="#c9a84c" stroke-width="1.5" opacity="0.5"/>
  <rect x="490" y="170" width="140" height="100" fill="#1a2d42" stroke="#c9a84c" stroke-width="1.5" opacity="0.6"/>
  <rect x="545" y="220" width="30" height="50" fill="#0d1b2a" opacity="0.8"/>
  <!-- text -->
  <text x="400" y="50" text-anchor="middle" fill="#c9a84c" font-size="18" font-family="serif">제 1 장</text>
  <text x="400" y="76" text-anchor="middle" fill="#e0c57a" font-size="13" font-family="serif">가난이 단련시킨 창업정신</text>
  <text x="260" y="370" text-anchor="middle" fill="#c9a84c" font-size="12" opacity="0.7">1949 인천 창영동 → 소년가장</text>
</svg>'''

# 제2장 – 껌 자판기
SVG['ch2'] = '''<svg viewBox="0 0 800 380" xmlns="http://www.w3.org/2000/svg">
  <rect width="800" height="380" fill="#0d1b2a"/>
  <!-- vending machine body -->
  <rect x="290" y="60" width="220" height="290" rx="10" fill="#1a2d42" stroke="#c9a84c" stroke-width="3"/>
  <!-- display window -->
  <rect x="310" y="80" width="180" height="120" rx="4" fill="#0d1b2a" stroke="#c9a84c" stroke-width="1.5"/>
  <!-- gum packs in window -->
  <rect x="320" y="90" width="38" height="25" fill="#c9a84c" opacity="0.8" rx="2"/>
  <rect x="362" y="90" width="38" height="25" fill="#e0c57a" opacity="0.8" rx="2"/>
  <rect x="404" y="90" width="38" height="25" fill="#c9a84c" opacity="0.7" rx="2"/>
  <rect x="446" y="90" width="38" height="25" fill="#e0c57a" opacity="0.7" rx="2"/>
  <rect x="320" y="120" width="38" height="25" fill="#e0c57a" opacity="0.6" rx="2"/>
  <rect x="362" y="120" width="38" height="25" fill="#c9a84c" opacity="0.6" rx="2"/>
  <rect x="404" y="120" width="38" height="25" fill="#e0c57a" opacity="0.6" rx="2"/>
  <rect x="446" y="120" width="38" height="25" fill="#c9a84c" opacity="0.6" rx="2"/>
  <text x="380" y="168" text-anchor="middle" fill="#c9a84c" font-size="11">껌 자동판매기</text>
  <!-- coin slot -->
  <rect x="340" y="215" width="120" height="30" rx="4" fill="#0d1b2a" stroke="#c9a84c" stroke-width="1.5"/>
  <rect x="388" y="218" width="24" height="8" rx="2" fill="#c9a84c" opacity="0.8"/>
  <text x="400" y="240" text-anchor="middle" fill="#e0c57a" font-size="10">10원</text>
  <!-- output tray -->
  <rect x="330" y="265" width="140" height="25" rx="3" fill="#0d1b2a" stroke="#c9a84c" stroke-width="1"/>
  <!-- coin arrow -->
  <circle cx="190" cy="190" r="28" fill="none" stroke="#c9a84c" stroke-width="2.5"/>
  <text x="190" y="186" text-anchor="middle" fill="#c9a84c" font-size="15" font-weight="bold">10</text>
  <text x="190" y="202" text-anchor="middle" fill="#c9a84c" font-size="12">원</text>
  <line x1="220" y1="190" x2="285" y2="220" stroke="#c9a84c" stroke-width="2" marker-end="url(#arr)"/>
  <defs><marker id="arr" markerWidth="8" markerHeight="8" refX="6" refY="3" orient="auto"><path d="M0,0 L0,6 L8,3 z" fill="#c9a84c"/></marker></defs>
  <text x="400" y="42" text-anchor="middle" fill="#c9a84c" font-size="18" font-family="serif">제 2 장</text>
  <text x="400" y="64" text-anchor="middle" fill="#e0c57a" font-size="12" font-family="serif">동전 하나의 혁명 — 껌 자동판매기</text>
  <text x="400" y="368" text-anchor="middle" fill="#c9a84c" font-size="11" opacity="0.7">1974~1978 · 국내 최초 껌 자판기 개발</text>
</svg>'''

# 제3장 – 부도 후 재기
SVG['ch3'] = '''<svg viewBox="0 0 800 380" xmlns="http://www.w3.org/2000/svg">
  <defs>
    <linearGradient id="bg3" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#1a0a0a"/><stop offset="1" stop-color="#0d1b2a"/></linearGradient>
    <radialGradient id="fire" cx="50%" cy="100%" r="60%"><stop offset="0" stop-color="#c9a84c" stop-opacity="0.8"/><stop offset="1" stop-color="#0d1b2a" stop-opacity="0"/></radialGradient>
  </defs>
  <rect width="800" height="380" fill="url(#bg3)"/>
  <!-- fire glow at bottom -->
  <ellipse cx="400" cy="360" rx="200" ry="80" fill="url(#fire)"/>
  <!-- rising figure -->
  <circle cx="400" cy="160" r="28" fill="#c9a84c" opacity="0.85"/>
  <rect x="383" y="188" width="34" height="60" fill="#c9a84c" opacity="0.75" rx="4"/>
  <!-- arms raised -->
  <line x1="383" y1="210" x2="340" y2="175" stroke="#c9a84c" stroke-width="10" stroke-linecap="round" opacity="0.75"/>
  <line x1="417" y1="210" x2="460" y2="175" stroke="#c9a84c" stroke-width="10" stroke-linecap="round" opacity="0.75"/>
  <!-- broken chain -->
  <path d="M320,290 Q340,260 360,290 Q380,320 400,290" fill="none" stroke="#888" stroke-width="2.5" stroke-dasharray="8,4"/>
  <!-- 부도 text crossed out -->
  <text x="220" y="240" fill="#cc3333" font-size="36" font-family="serif" opacity="0.5" font-weight="bold">부도</text>
  <line x1="210" y1="215" x2="295" y2="255" stroke="#cc3333" stroke-width="3" opacity="0.6"/>
  <!-- arrow up -->
  <line x1="400" y1="130" x2="400" y2="80" stroke="#c9a84c" stroke-width="3" marker-end="url(#ua)"/>
  <defs><marker id="ua" markerWidth="10" markerHeight="10" refX="5" refY="5" orient="auto"><path d="M0,10 L5,0 L10,10 z" fill="#c9a84c"/></marker></defs>
  <text x="400" y="42" text-anchor="middle" fill="#c9a84c" font-size="18" font-family="serif">제 3 장</text>
  <text x="400" y="64" text-anchor="middle" fill="#e0c57a" font-size="12" font-family="serif">부도의 아픔 속에서 얻은 평생의 철학</text>
  <text x="400" y="370" text-anchor="middle" fill="#c9a84c" font-size="12" opacity="0.7">결핍을 자원으로 — 무차입·무부도 원칙의 탄생</text>
</svg>'''

# 제4장 – 브라운관에서 산업용 모니터로
SVG['ch4'] = '''<svg viewBox="0 0 800 380" xmlns="http://www.w3.org/2000/svg">
  <rect width="800" height="380" fill="#0d1b2a"/>
  <!-- CRT monitor old -->
  <rect x="80" y="120" width="180" height="150" rx="8" fill="#1a2d42" stroke="#888" stroke-width="3"/>
  <rect x="96" y="135" width="148" height="110" rx="4" fill="#0a1520" stroke="#666" stroke-width="1"/>
  <ellipse cx="170" cy="190" rx="60" ry="45" fill="#1a3a5c" opacity="0.6"/>
  <rect x="145" y="270" width="50" height="20" fill="#1a2d42"/>
  <rect x="120" y="288" width="100" height="10" fill="#888" rx="2"/>
  <text x="170" y="340" text-anchor="middle" fill="#888" font-size="11">브라운관 TV</text>
  <!-- arrow -->
  <line x1="285" y1="195" x2="375" y2="195" stroke="#c9a84c" stroke-width="3" marker-end="url(#arr4)"/>
  <defs><marker id="arr4" markerWidth="10" markerHeight="10" refX="8" refY="3" orient="auto"><path d="M0,0 L0,6 L10,3 z" fill="#c9a84c"/></marker></defs>
  <text x="330" y="185" text-anchor="middle" fill="#c9a84c" font-size="11">혁신</text>
  <!-- Industrial monitor new -->
  <rect x="390" y="100" width="230" height="165" rx="6" fill="#1a2d42" stroke="#c9a84c" stroke-width="3"/>
  <rect x="405" y="112" width="200" height="138" rx="3" fill="#0d1b2a"/>
  <!-- screen content - data lines -->
  <line x1="415" y1="135" x2="595" y2="135" stroke="#c9a84c" stroke-width="1.5" opacity="0.8"/>
  <line x1="415" y1="155" x2="560" y2="155" stroke="#e0c57a" stroke-width="1" opacity="0.6"/>
  <line x1="415" y1="175" x2="580" y2="175" stroke="#e0c57a" stroke-width="1" opacity="0.6"/>
  <!-- waveform -->
  <polyline points="415,215 435,200 455,230 475,195 495,225 515,205 535,220 555,200 575,215 595,208" fill="none" stroke="#c9a84c" stroke-width="2"/>
  <rect x="480" y="268" width="70" height="15" fill="#1a2d42"/>
  <rect x="440" y="282" width="150" height="10" fill="#888" rx="2"/>
  <text x="505" y="340" text-anchor="middle" fill="#c9a84c" font-size="11">산업용 모니터 (코텍)</text>
  <text x="400" y="42" text-anchor="middle" fill="#c9a84c" font-size="18" font-family="serif">제 4 장</text>
  <text x="400" y="64" text-anchor="middle" fill="#e0c57a" font-size="12" font-family="serif">브라운관에서 산업용 모니터로</text>
  <text x="400" y="370" text-anchor="middle" fill="#c9a84c" font-size="11" opacity="0.7">1987 세주전자(코텍) 창업 · 삼성전관 5,000만원 담보 계약</text>
</svg>'''

# 제5장 – 배반과 질병, 두 개의 전선
SVG['ch5'] = '''<svg viewBox="0 0 800 380" xmlns="http://www.w3.org/2000/svg">
  <rect width="800" height="380" fill="#0d1b2a"/>
  <!-- center figure -->
  <circle cx="400" cy="190" r="32" fill="#c9a84c" opacity="0.75"/>
  <rect x="382" y="222" width="36" height="65" fill="#c9a84c" opacity="0.65" rx="4"/>
  <!-- left threat: betrayal -->
  <g transform="translate(0,0)">
    <circle cx="180" cy="180" r="50" fill="#cc3333" opacity="0.15" stroke="#cc3333" stroke-width="2"/>
    <text x="180" y="170" text-anchor="middle" fill="#cc3333" font-size="14" font-weight="bold">배반</text>
    <text x="180" y="190" text-anchor="middle" fill="#cc3333" font-size="11">핵심인력 이탈</text>
    <text x="180" y="207" text-anchor="middle" fill="#cc3333" font-size="10">세계전자 설립</text>
  </g>
  <!-- right threat: illness -->
  <g>
    <circle cx="620" cy="180" r="50" fill="#cc6633" opacity="0.15" stroke="#cc6633" stroke-width="2"/>
    <text x="620" y="170" text-anchor="middle" fill="#cc6633" font-size="14" font-weight="bold">질병</text>
    <text x="620" y="190" text-anchor="middle" fill="#cc6633" font-size="11">B형 간염</text>
    <text x="620" y="207" text-anchor="middle" fill="#cc6633" font-size="10">해외출장 강행</text>
  </g>
  <!-- arrows attacking -->
  <line x1="232" y1="180" x2="368" y2="190" stroke="#cc3333" stroke-width="2.5" stroke-dasharray="6,3" marker-end="url(#r1)"/>
  <line x1="568" y1="180" x2="432" y2="190" stroke="#cc6633" stroke-width="2.5" stroke-dasharray="6,3" marker-end="url(#r2)"/>
  <defs>
    <marker id="r1" markerWidth="8" markerHeight="8" refX="6" refY="3" orient="auto"><path d="M0,0 L0,6 L8,3 z" fill="#cc3333"/></marker>
    <marker id="r2" markerWidth="8" markerHeight="8" refX="6" refY="3" orient="auto"><path d="M0,0 L0,6 L8,3 z" fill="#cc6633"/></marker>
  </defs>
  <!-- shield -->
  <path d="M400,158 L432,175 L432,210 L400,225 L368,210 L368,175 Z" fill="none" stroke="#c9a84c" stroke-width="2.5" opacity="0.8"/>
  <text x="400" y="198" text-anchor="middle" fill="#c9a84c" font-size="11">신뢰</text>
  <text x="400" y="42" text-anchor="middle" fill="#c9a84c" font-size="18" font-family="serif">제 5 장</text>
  <text x="400" y="64" text-anchor="middle" fill="#e0c57a" font-size="12" font-family="serif">배반과 질병, 두 개의 전선</text>
  <text x="400" y="370" text-anchor="middle" fill="#c9a84c" font-size="11" opacity="0.7">1989년 이중 위기 — 신뢰 4원칙의 확립</text>
</svg>'''

# 제6장 – Free Voltage · 글로벌 돌파
SVG['ch6'] = '''<svg viewBox="0 0 800 380" xmlns="http://www.w3.org/2000/svg">
  <rect width="800" height="380" fill="#0d1b2a"/>
  <!-- globe -->
  <circle cx="400" cy="200" r="130" fill="none" stroke="#c9a84c" stroke-width="2"/>
  <ellipse cx="400" cy="200" rx="130" ry="50" fill="none" stroke="#c9a84c" stroke-width="1" opacity="0.4"/>
  <ellipse cx="400" cy="200" rx="80" ry="130" fill="none" stroke="#c9a84c" stroke-width="1" opacity="0.4"/>
  <line x1="270" y1="200" x2="530" y2="200" stroke="#c9a84c" stroke-width="1" opacity="0.3"/>
  <line x1="400" y1="70" x2="400" y2="330" stroke="#c9a84c" stroke-width="1" opacity="0.3"/>
  <!-- continents simplified -->
  <ellipse cx="350" cy="175" rx="30" ry="20" fill="#2e4a6a" opacity="0.7"/>
  <ellipse cx="440" cy="195" rx="25" ry="18" fill="#2e4a6a" opacity="0.7"/>
  <ellipse cx="370" cy="230" rx="20" ry="15" fill="#2e4a6a" opacity="0.6"/>
  <!-- lightning bolt Free Voltage -->
  <polygon points="415,120 400,170 415,165 395,220 425,155 410,160" fill="#c9a84c" opacity="0.9"/>
  <!-- voltage labels -->
  <rect x="100" y="175" width="70" height="25" rx="4" fill="#1a2d42" stroke="#888" stroke-width="1.5"/>
  <text x="135" y="192" text-anchor="middle" fill="#888" font-size="12">110V</text>
  <rect x="630" y="175" width="70" height="25" rx="4" fill="#1a2d42" stroke="#888" stroke-width="1.5"/>
  <text x="665" y="192" text-anchor="middle" fill="#888" font-size="12">220V</text>
  <!-- center FREE VOLTAGE label -->
  <rect x="490" y="105" width="120" height="30" rx="4" fill="#c9a84c" opacity="0.9"/>
  <text x="550" y="125" text-anchor="middle" fill="#0d1b2a" font-size="12" font-weight="bold">FREE VOLTAGE</text>
  <line x1="492" y1="120" x2="425" y2="148" stroke="#c9a84c" stroke-width="1.5" opacity="0.7"/>
  <text x="400" y="42" text-anchor="middle" fill="#c9a84c" font-size="18" font-family="serif">제 6 장</text>
  <text x="400" y="64" text-anchor="middle" fill="#e0c57a" font-size="12" font-family="serif">Free Voltage와 글로벌 돌파</text>
  <text x="400" y="370" text-anchor="middle" fill="#c9a84c" font-size="11" opacity="0.7">세계 어느 전압에서도 작동 — 코텍 최초의 글로벌 혁신</text>
</svg>'''

# 제7장 – IGT와의 첫 접촉
SVG['ch7'] = '''<svg viewBox="0 0 800 380" xmlns="http://www.w3.org/2000/svg">
  <rect width="800" height="380" fill="#0d1b2a"/>
  <!-- slot machine -->
  <rect x="280" y="80" width="240" height="260" rx="12" fill="#1a2d42" stroke="#c9a84c" stroke-width="3"/>
  <rect x="300" y="100" width="200" height="130" rx="6" fill="#0a1520" stroke="#c9a84c" stroke-width="1.5"/>
  <!-- reels -->
  <rect x="308" y="108" width="56" height="114" rx="4" fill="#1a3050"/>
  <rect x="372" y="108" width="56" height="114" rx="4" fill="#1a3050"/>
  <rect x="436" y="108" width="56" height="114" rx="4" fill="#1a3050"/>
  <!-- symbols -->
  <text x="336" y="170" text-anchor="middle" fill="#c9a84c" font-size="36">7</text>
  <text x="400" y="170" text-anchor="middle" fill="#e0c57a" font-size="36">7</text>
  <text x="464" y="170" text-anchor="middle" fill="#c9a84c" font-size="36">7</text>
  <!-- jackpot line -->
  <line x1="302" y1="165" x2="498" y2="165" stroke="#c9a84c" stroke-width="2" stroke-dasharray="4,3" opacity="0.8"/>
  <!-- coin slot + lever -->
  <rect x="360" y="248" width="80" height="20" rx="4" fill="#0d1b2a" stroke="#c9a84c" stroke-width="1.5"/>
  <rect x="393" y="250" width="14" height="6" rx="2" fill="#c9a84c"/>
  <circle cx="530" cy="195" r="12" fill="#c9a84c"/>
  <rect x="527" y="140" width="6" height="58" fill="#c9a84c" rx="3"/>
  <!-- IGT label -->
  <text x="160" y="180" text-anchor="middle" fill="#c9a84c" font-size="28" font-weight="bold" font-family="serif">IGT</text>
  <text x="160" y="205" text-anchor="middle" fill="#e0c57a" font-size="11">International</text>
  <text x="160" y="220" text-anchor="middle" fill="#e0c57a" font-size="11">Gaming Technology</text>
  <line x1="220" y1="200" x2="278" y2="200" stroke="#c9a84c" stroke-width="2" stroke-dasharray="5,3"/>
  <!-- monitor on slot -->
  <rect x="308" y="108" width="182" height="114" rx="4" fill="none" stroke="#c9a84c" stroke-width="2" opacity="0.5"/>
  <text x="400" y="42" text-anchor="middle" fill="#c9a84c" font-size="18" font-family="serif">제 7 장</text>
  <text x="400" y="64" text-anchor="middle" fill="#e0c57a" font-size="12" font-family="serif">IGT와의 첫 접촉</text>
  <text x="400" y="370" text-anchor="middle" fill="#c9a84c" font-size="11" opacity="0.7">1989년 · 카지노 모니터 세계 시장 개척</text>
</svg>'''

# 제8장 – 100% 리콜 · 신뢰
SVG['ch8'] = '''<svg viewBox="0 0 800 380" xmlns="http://www.w3.org/2000/svg">
  <defs><radialGradient id="glow8" cx="50%" cy="50%" r="50%"><stop offset="0" stop-color="#c9a84c" stop-opacity="0.2"/><stop offset="1" stop-color="#0d1b2a" stop-opacity="0"/></radialGradient></defs>
  <rect width="800" height="380" fill="#0d1b2a"/>
  <circle cx="400" cy="195" r="150" fill="url(#glow8)"/>
  <!-- shield -->
  <path d="M400,70 L510,115 L510,210 L400,270 L290,210 L290,115 Z" fill="#1a2d42" stroke="#c9a84c" stroke-width="3"/>
  <!-- 100% inside shield -->
  <text x="400" y="165" text-anchor="middle" fill="#c9a84c" font-size="42" font-weight="bold" font-family="serif">100%</text>
  <text x="400" y="200" text-anchor="middle" fill="#e0c57a" font-size="18" font-family="serif">리콜</text>
  <text x="400" y="230" text-anchor="middle" fill="#c9a84c" font-size="13">신 뢰</text>
  <!-- checkmark -->
  <polyline points="358,170 385,195 440,145" fill="none" stroke="#c9a84c" stroke-width="4" opacity="0.3"/>
  <!-- side labels -->
  <text x="150" y="175" text-anchor="middle" fill="#cc3333" font-size="14" opacity="0.8">30% 불량</text>
  <text x="150" y="198" text-anchor="middle" fill="#cc3333" font-size="11" opacity="0.7">납품 후 발생</text>
  <line x1="210" y1="185" x2="288" y2="185" stroke="#cc3333" stroke-width="1.5" stroke-dasharray="4,3" opacity="0.6"/>
  <text x="650" y="175" text-anchor="middle" fill="#4caf50" font-size="14" opacity="0.9">세계 1위</text>
  <text x="650" y="198" text-anchor="middle" fill="#4caf50" font-size="11" opacity="0.7">신뢰 확보</text>
  <line x1="512" y1="185" x2="590" y2="185" stroke="#4caf50" stroke-width="1.5" stroke-dasharray="4,3" opacity="0.6"/>
  <text x="400" y="42" text-anchor="middle" fill="#c9a84c" font-size="18" font-family="serif">제 8 장</text>
  <text x="400" y="64" text-anchor="middle" fill="#e0c57a" font-size="12" font-family="serif">100% 리콜 — 신뢰를 파는 방법</text>
  <text x="400" y="370" text-anchor="middle" fill="#c9a84c" font-size="11" opacity="0.7">"신뢰를 쌓는 데는 20년, 무너지는 데는 5분" — 워런 버핏</text>
</svg>'''

# 제9장 – 코스닥 상장
SVG['ch9'] = '''<svg viewBox="0 0 800 380" xmlns="http://www.w3.org/2000/svg">
  <rect width="800" height="380" fill="#0d1b2a"/>
  <!-- chart background -->
  <rect x="80" y="80" width="640" height="220" rx="4" fill="#0a1520" stroke="#2e4a6a" stroke-width="1"/>
  <!-- grid lines -->
  <line x1="80" y1="135" x2="720" y2="135" stroke="#2e4a6a" stroke-width="1" opacity="0.4"/>
  <line x1="80" y1="190" x2="720" y2="190" stroke="#2e4a6a" stroke-width="1" opacity="0.4"/>
  <line x1="80" y1="245" x2="720" y2="245" stroke="#2e4a6a" stroke-width="1" opacity="0.4"/>
  <line x1="200" y1="80" x2="200" y2="300" stroke="#2e4a6a" stroke-width="1" opacity="0.3"/>
  <line x1="340" y1="80" x2="340" y2="300" stroke="#2e4a6a" stroke-width="1" opacity="0.3"/>
  <line x1="480" y1="80" x2="480" y2="300" stroke="#2e4a6a" stroke-width="1" opacity="0.3"/>
  <line x1="620" y1="80" x2="620" y2="300" stroke="#2e4a6a" stroke-width="1" opacity="0.3"/>
  <!-- stock line -->
  <polyline points="100,265 160,258 220,252 260,245 310,240 360,230 400,215 430,200 460,188 490,170 530,150 570,125 610,105 660,92 710,85"
    fill="none" stroke="#c9a84c" stroke-width="3"/>
  <!-- area under line -->
  <polygon points="100,265 160,258 220,252 260,245 310,240 360,230 400,215 430,200 460,188 490,170 530,150 570,125 610,105 660,92 710,85 710,300 100,300"
    fill="#c9a84c" opacity="0.08"/>
  <!-- listing point -->
  <circle cx="480" cy="170" r="8" fill="#c9a84c"/>
  <line x1="480" y1="80" x2="480" y2="162" stroke="#c9a84c" stroke-width="1.5" stroke-dasharray="4,3"/>
  <text x="510" y="95" fill="#c9a84c" font-size="11">코스닥 상장</text>
  <text x="510" y="110" fill="#e0c57a" font-size="10">2000년</text>
  <!-- axis labels -->
  <text x="90" y="275" fill="#888" font-size="10">1987</text>
  <text x="460" y="275" fill="#888" font-size="10">2000</text>
  <text x="690" y="275" fill="#888" font-size="10">2014</text>
  <!-- y label -->
  <text x="68" y="95" fill="#c9a84c" font-size="10" text-anchor="middle">주가</text>
  <text x="68" y="110" fill="#c9a84c" font-size="10" text-anchor="middle">성장</text>
  <text x="400" y="42" text-anchor="middle" fill="#c9a84c" font-size="18" font-family="serif">제 9 장</text>
  <text x="400" y="64" text-anchor="middle" fill="#e0c57a" font-size="12" font-family="serif">코스닥 상장 — 세상에 얼굴을 내밀다</text>
  <text x="400" y="345" text-anchor="middle" fill="#c9a84c" font-size="11" opacity="0.7">27년간 한 우물 · 히든챔피언 전략 · Pro-Kortek</text>
</svg>'''

# 제10장 – 最知信 (1% Plus)
SVG['ch10'] = '''<svg viewBox="0 0 800 380" xmlns="http://www.w3.org/2000/svg">
  <rect width="800" height="380" fill="#0d1b2a"/>
  <!-- triangle -->
  <polygon points="400,70 570,300 230,300" fill="none" stroke="#c9a84c" stroke-width="3"/>
  <!-- inner glow triangle -->
  <polygon points="400,100 548,285 252,285" fill="#1a2d42" opacity="0.5"/>
  <!-- three nodes -->
  <circle cx="400" cy="70" r="38" fill="#1a2d42" stroke="#c9a84c" stroke-width="2.5"/>
  <circle cx="570" cy="300" r="38" fill="#1a2d42" stroke="#c9a84c" stroke-width="2.5"/>
  <circle cx="230" cy="300" r="38" fill="#1a2d42" stroke="#c9a84c" stroke-width="2.5"/>
  <!-- characters -->
  <text x="400" y="62" text-anchor="middle" fill="#c9a84c" font-size="28" font-weight="bold">最</text>
  <text x="400" y="84" text-anchor="middle" fill="#e0c57a" font-size="10">최고</text>
  <text x="570" y="292" text-anchor="middle" fill="#c9a84c" font-size="28" font-weight="bold">知</text>
  <text x="570" y="314" text-anchor="middle" fill="#e0c57a" font-size="10">기술력</text>
  <text x="230" y="292" text-anchor="middle" fill="#c9a84c" font-size="28" font-weight="bold">信</text>
  <text x="230" y="314" text-anchor="middle" fill="#e0c57a" font-size="10">신뢰</text>
  <!-- center text -->
  <text x="400" y="192" text-anchor="middle" fill="#c9a84c" font-size="22" font-weight="bold" font-family="serif">1% Plus</text>
  <text x="400" y="215" text-anchor="middle" fill="#e0c57a" font-size="13" font-family="serif">남들보다 1%만 더</text>
  <!-- subtitle decorations -->
  <line x1="290" y1="192" x2="340" y2="192" stroke="#c9a84c" stroke-width="1" opacity="0.6"/>
  <line x1="460" y1="192" x2="510" y2="192" stroke="#c9a84c" stroke-width="1" opacity="0.6"/>
  <text x="400" y="42" text-anchor="middle" fill="#c9a84c" font-size="18" font-family="serif">제 10 장</text>
  <text x="400" y="64" text-anchor="middle" fill="#e0c57a" font-size="12" font-family="serif">最·知·信 — 코스닥 상장과 경영철학</text>
  <text x="400" y="370" text-anchor="middle" fill="#c9a84c" font-size="11" opacity="0.7">Auto Color Bias 특허 · 1등과 2등의 차이는 1%</text>
</svg>'''

# 제13·14장 – 인사철학 + 조직문화
SVG['ch13_14'] = '''<svg viewBox="0 0 800 380" xmlns="http://www.w3.org/2000/svg">
  <rect width="800" height="380" fill="#0d1b2a"/>
  <!-- org chart -->
  <!-- CEO -->
  <circle cx="400" cy="100" r="28" fill="#c9a84c" opacity="0.9"/>
  <text x="400" y="106" text-anchor="middle" fill="#0d1b2a" font-size="12" font-weight="bold">CEO</text>
  <!-- lines to second level -->
  <line x1="372" y1="128" x2="250" y2="175" stroke="#c9a84c" stroke-width="2"/>
  <line x1="400" y1="128" x2="400" y2="175" stroke="#c9a84c" stroke-width="2"/>
  <line x1="428" y1="128" x2="550" y2="175" stroke="#c9a84c" stroke-width="2"/>
  <!-- second level -->
  <circle cx="250" cy="195" r="22" fill="#1a2d42" stroke="#c9a84c" stroke-width="2"/>
  <text x="250" y="199" text-anchor="middle" fill="#c9a84c" font-size="9">생산</text>
  <circle cx="400" cy="195" r="22" fill="#1a2d42" stroke="#c9a84c" stroke-width="2"/>
  <text x="400" y="199" text-anchor="middle" fill="#c9a84c" font-size="9">영업</text>
  <circle cx="550" cy="195" r="22" fill="#1a2d42" stroke="#c9a84c" stroke-width="2"/>
  <text x="550" y="199" text-anchor="middle" fill="#c9a84c" font-size="9">개발</text>
  <!-- third level people dots -->
  <circle cx="200" cy="268" r="14" fill="#2e4a6a" stroke="#c9a84c" stroke-width="1.5"/>
  <circle cx="235" cy="268" r="14" fill="#2e4a6a" stroke="#c9a84c" stroke-width="1.5"/>
  <circle cx="265" cy="268" r="14" fill="#2e4a6a" stroke="#c9a84c" stroke-width="1.5"/>
  <circle cx="300" cy="268" r="14" fill="#2e4a6a" stroke="#c9a84c" stroke-width="1.5"/>
  <circle cx="365" cy="268" r="14" fill="#2e4a6a" stroke="#c9a84c" stroke-width="1.5"/>
  <circle cx="400" cy="268" r="14" fill="#2e4a6a" stroke="#c9a84c" stroke-width="1.5"/>
  <circle cx="435" cy="268" r="14" fill="#2e4a6a" stroke="#c9a84c" stroke-width="1.5"/>
  <circle cx="510" cy="268" r="14" fill="#2e4a6a" stroke="#c9a84c" stroke-width="1.5"/>
  <circle cx="550" cy="268" r="14" fill="#2e4a6a" stroke="#c9a84c" stroke-width="1.5"/>
  <circle cx="590" cy="268" r="14" fill="#2e4a6a" stroke="#c9a84c" stroke-width="1.5"/>
  <!-- profit distribution box -->
  <rect x="120" y="310" width="560" height="38" rx="6" fill="#1a2d42" stroke="#c9a84c" stroke-width="1.5"/>
  <text x="235" y="333" text-anchor="middle" fill="#c9a84c" font-size="11">유보이익 50%</text>
  <line x1="350" y1="312" x2="350" y2="348" stroke="#c9a84c" stroke-width="1" opacity="0.5"/>
  <text x="400" y="333" text-anchor="middle" fill="#c9a84c" font-size="11">주주 25%</text>
  <line x1="450" y1="312" x2="450" y2="348" stroke="#c9a84c" stroke-width="1" opacity="0.5"/>
  <text x="565" y="333" text-anchor="middle" fill="#c9a84c" font-size="11">직원 성과급 25%</text>
  <text x="400" y="42" text-anchor="middle" fill="#c9a84c" font-size="18" font-family="serif">제 13 · 14 장</text>
  <text x="400" y="64" text-anchor="middle" fill="#e0c57a" font-size="12" font-family="serif">인사철학 · 조직문화 — 사람이 전략이다</text>
</svg>'''

# 제15장 – 승계의 철학
SVG['ch15'] = '''<svg viewBox="0 0 800 380" xmlns="http://www.w3.org/2000/svg">
  <rect width="800" height="380" fill="#0d1b2a"/>
  <!-- torch left person -->
  <circle cx="220" cy="170" r="30" fill="#c9a84c" opacity="0.7"/>
  <rect x="207" y="200" width="26" height="55" fill="#c9a84c" opacity="0.6" rx="3"/>
  <!-- torch -->
  <rect x="245" y="125" width="10" height="50" fill="#888" rx="2"/>
  <!-- flame -->
  <ellipse cx="250" cy="118" rx="14" ry="18" fill="#c9a84c" opacity="0.9"/>
  <ellipse cx="250" cy="112" rx="8" ry="12" fill="#e0c57a"/>
  <ellipse cx="250" cy="108" rx="4" ry="7" fill="#fff7e0" opacity="0.8"/>
  <!-- passing motion line -->
  <path d="M255,140 Q400,100 545,140" fill="none" stroke="#c9a84c" stroke-width="2" stroke-dasharray="6,4" opacity="0.7" marker-end="url(#arr15)"/>
  <defs><marker id="arr15" markerWidth="8" markerHeight="8" refX="6" refY="3" orient="auto"><path d="M0,0 L0,6 L8,3 z" fill="#c9a84c"/></marker></defs>
  <!-- right person (successor) -->
  <circle cx="570" cy="170" r="30" fill="#2e4a6a" stroke="#c9a84c" stroke-width="2.5"/>
  <rect x="557" y="200" width="26" height="55" fill="#2e4a6a" stroke="#c9a84c" stroke-width="1.5" rx="3"/>
  <!-- labels -->
  <text x="220" y="280" text-anchor="middle" fill="#c9a84c" font-size="13">이한구 회장</text>
  <text x="220" y="298" text-anchor="middle" fill="#888" font-size="11">창업자</text>
  <text x="570" y="280" text-anchor="middle" fill="#c9a84c" font-size="13">후계자</text>
  <text x="570" y="298" text-anchor="middle" fill="#888" font-size="11">능력 · 열정 · 기업가정신</text>
  <text x="400" y="160" text-anchor="middle" fill="#c9a84c" font-size="12" opacity="0.7">기업은 기업의 것이다</text>
  <!-- cross out family -->
  <text x="400" y="320" text-anchor="middle" fill="#888" font-size="12" opacity="0.6">혈연 승계 거부 → 능력 승계 선택</text>
  <text x="400" y="42" text-anchor="middle" fill="#c9a84c" font-size="18" font-family="serif">제 15 장</text>
  <text x="400" y="64" text-anchor="middle" fill="#e0c57a" font-size="12" font-family="serif">승계의 철학 — 기업은 기업의 것이다</text>
  <text x="400" y="370" text-anchor="middle" fill="#c9a84c" font-size="11" opacity="0.7">2년간 50명 인터뷰 · 가장 아름다운 승계</text>
</svg>'''

# 제16장 – M&A
SVG['ch16'] = '''<svg viewBox="0 0 800 380" xmlns="http://www.w3.org/2000/svg">
  <rect width="800" height="380" fill="#0d1b2a"/>
  <!-- two hands shaking -->
  <!-- left hand -->
  <path d="M180,215 L250,180 L270,175 L290,178 L310,185 L320,195 L305,205 L285,198 L275,200 L370,225 L365,250 L260,228 L260,240 L355,262 L350,286 L245,260 L245,272 L335,293 L330,316 L225,290 L200,295 L175,280 L170,250 Z" fill="#c9a84c" opacity="0.65"/>
  <!-- right hand -->
  <path d="M620,215 L550,180 L530,175 L510,178 L490,185 L480,195 L495,205 L515,198 L525,200 L430,225 L435,250 L540,228 L540,240 L445,262 L450,286 L555,260 L555,272 L465,293 L470,316 L575,290 L600,295 L625,280 L630,250 Z" fill="#c9a84c" opacity="0.65"/>
  <!-- handshake area -->
  <ellipse cx="400" cy="215" rx="45" ry="30" fill="#c9a84c" opacity="0.5" stroke="#c9a84c" stroke-width="2"/>
  <!-- labels -->
  <text x="210" y="340" text-anchor="middle" fill="#e0c57a" font-size="13">코텍 · 이한구</text>
  <text x="590" y="340" text-anchor="middle" fill="#e0c57a" font-size="13">아이앤아이 · 김영달</text>
  <!-- deal label -->
  <rect x="330" y="88" width="140" height="45" rx="6" fill="#1a2d42" stroke="#c9a84c" stroke-width="2"/>
  <text x="400" y="110" text-anchor="middle" fill="#c9a84c" font-size="15" font-weight="bold">700억</text>
  <text x="400" y="127" text-anchor="middle" fill="#e0c57a" font-size="11">단 3분 만에 결정</text>
  <text x="400" y="42" text-anchor="middle" fill="#c9a84c" font-size="18" font-family="serif">제 16 장</text>
  <text x="400" y="64" text-anchor="middle" fill="#e0c57a" font-size="12" font-family="serif">김영달과의 M&amp;A — 완벽한 승계</text>
  <text x="400" y="370" text-anchor="middle" fill="#c9a84c" font-size="11" opacity="0.7">주가 8,000원 → 15,000원 · Win-Win M&amp;A</text>
</svg>'''

# 에필로그 – SER-M 다이어그램
SVG['epilogue'] = '''<svg viewBox="0 0 800 380" xmlns="http://www.w3.org/2000/svg">
  <rect width="800" height="380" fill="#0d1b2a"/>
  <!-- SER-M framework -->
  <!-- S Subject -->
  <circle cx="400" cy="140" r="50" fill="#1a2d42" stroke="#c9a84c" stroke-width="2.5"/>
  <text x="400" y="133" text-anchor="middle" fill="#c9a84c" font-size="24" font-weight="bold">S</text>
  <text x="400" y="155" text-anchor="middle" fill="#e0c57a" font-size="11">주체(이한구)</text>
  <!-- E Environment -->
  <circle cx="220" cy="270" r="44" fill="#1a2d42" stroke="#c9a84c" stroke-width="2"/>
  <text x="220" y="263" text-anchor="middle" fill="#c9a84c" font-size="22" font-weight="bold">E</text>
  <text x="220" y="282" text-anchor="middle" fill="#e0c57a" font-size="10">환경</text>
  <!-- R Resource -->
  <circle cx="400" cy="300" r="44" fill="#1a2d42" stroke="#c9a84c" stroke-width="2"/>
  <text x="400" y="293" text-anchor="middle" fill="#c9a84c" font-size="22" font-weight="bold">R</text>
  <text x="400" y="312" text-anchor="middle" fill="#e0c57a" font-size="10">자원</text>
  <!-- M Mechanism -->
  <circle cx="580" cy="270" r="44" fill="#1a2d42" stroke="#c9a84c" stroke-width="2"/>
  <text x="580" y="263" text-anchor="middle" fill="#c9a84c" font-size="22" font-weight="bold">M</text>
  <text x="580" y="282" text-anchor="middle" fill="#e0c57a" font-size="10">메커니즘</text>
  <!-- connecting lines -->
  <line x1="358" y1="175" x2="252" y2="240" stroke="#c9a84c" stroke-width="1.5" opacity="0.6"/>
  <line x1="400" y1="190" x2="400" y2="256" stroke="#c9a84c" stroke-width="1.5" opacity="0.6"/>
  <line x1="442" y1="175" x2="548" y2="240" stroke="#c9a84c" stroke-width="1.5" opacity="0.6"/>
  <line x1="260" y1="278" x2="356" y2="294" stroke="#c9a84c" stroke-width="1" opacity="0.4"/>
  <line x1="444" y1="294" x2="536" y2="278" stroke="#c9a84c" stroke-width="1" opacity="0.4"/>
  <!-- outcome -->
  <text x="400" y="50" text-anchor="middle" fill="#c9a84c" font-size="20" font-family="serif">에 필 로 그</text>
  <text x="400" y="74" text-anchor="middle" fill="#e0c57a" font-size="13" font-family="serif">SER-M으로 보는 이한구 경영학</text>
  <text x="400" y="370" text-anchor="middle" fill="#c9a84c" font-size="11" opacity="0.7">발로 걷는 삶 · 비워야 다시 채울 수 있다</text>
</svg>'''

print("SVG count:", len(SVG))
print("Done preparing SVGs")
