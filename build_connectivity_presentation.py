from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.dml.color import RGBColor

OUT = Path('A3-PS015-TC142_Explainable_Connectivity_Diagnosis.pptx')
prs = Presentation()
prs.slide_width = Inches(13.333333)
prs.slide_height = Inches(7.5)
blank = prs.slide_layouts[6]

BLUE = RGBColor(0, 0, 210)
NAVY = RGBColor(0, 0, 135)
BLACK = RGBColor(20, 20, 20)
GRAY = RGBColor(82, 82, 82)
LIGHT = RGBColor(239, 244, 255)
MID = RGBColor(196, 210, 255)
WHITE = RGBColor(255, 255, 255)

def tx(slide, text, x, y, w, h, size=18, color=BLACK, bold=False, align=PP_ALIGN.LEFT,
       valign=MSO_ANCHOR.TOP, font='Aptos'):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear(); tf.word_wrap = True; tf.vertical_anchor = valign
    p = tf.paragraphs[0]; p.alignment = align
    run = p.add_run(); run.text = text
    run.font.name = font; run.font.size = Pt(size); run.font.bold = bold; run.font.color.rgb = color
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    return box

def rect(slide, x, y, w, h, fill=WHITE, line=None, radius=False):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE,
                                   Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid(); shape.fill.fore_color.rgb = fill
    shape.line.color.rgb = line if line else fill
    return shape

def frame(title, section, n):
    s = prs.slides.add_slide(blank)
    rect(s, 0, 0, 13.333, 7.5)
    tx(s, 'AIORI 3', 9.65, .10, .8, .18, 7.2, BLUE, True, PP_ALIGN.RIGHT)
    tx(s, section, 10.6, .10, 2.35, .18, 7.2, BLUE, True, PP_ALIGN.LEFT)
    tx(s, title, .59, .58, 11.9, .55, 23, BLUE, True)
    tx(s, 'Varshini SD & Aditya  •  A3-PS015-TC142  •  October 14, 2026', .59, 7.25, 7.7, .13, 7, BLUE, True)
    tx(s, f'{n} / 11', 12.08, 7.25, .65, .13, 7, BLUE, False, PP_ALIGN.RIGHT)
    return s

def bullet_text(slide, items, x=.60, y=2.0, w=12.0, size=15.5, gap=.46):
    for i, item in enumerate(items):
        tx(slide, '•', x, y + i*gap, .2, .25, size, BLUE, True)
        tx(slide, item, x+.3, y + i*gap, w-.3, .34, size, BLACK)

def node(slide, label, x, y, w, h, fill=LIGHT, size=12):
    sh = rect(slide, x, y, w, h, fill, BLUE, True)
    tx(slide, label, x+.07, y+.08, w-.14, h-.16, size, NAVY, True, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    return sh

def arrow(slide, x1, y1, x2, y2):
    ln = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    ln.line.color.rgb = BLUE; ln.line.width = Pt(1.5); ln.line.end_arrowhead = True
    return ln

# 1 — title
s = prs.slides.add_slide(blank)
rect(s, 0, 0, 13.333, 7.5)
tx(s, 'AIORI 3', 9.65, .10, .8, .18, 7.2, BLUE, True, PP_ALIGN.RIGHT)
tx(s, 'Hackathon Submission', 10.6, .10, 2.35, .18, 7.2, BLUE, True)
tx(s, 'Explainable Home &\nLast-Mile Connectivity Diagnosis', .61, 1.40, 11.9, 1.35, 29, NAVY, True)
tx(s, 'An AI-powered system that identifies likely connectivity causes\nand gives users clear, actionable next steps.', .62, 3.05, 9.9, .68, 18, BLACK)
tx(s, 'A3-PS015-TC142', .62, 4.78, 4.3, .28, 16, BLACK, True)
tx(s, 'Varshini SD  •  Aditya', .62, 5.16, 4.3, .25, 15, BLACK)
tx(s, 'Faculty Mentor: Dr. Monika Puttaramaiah', .62, 5.51, 5.7, .25, 14, BLACK)
tx(s, 'AIORI Hackathon 3  •  October 14, 2026', .62, 5.86, 5.7, .25, 14, BLACK)

# 2 — problem
s = frame('Problem Description', 'Abstract', 2)
tx(s, 'Users can see that their connection is poor — but rarely why.', .60, 1.60, 12.0, .30, 16, BLACK, True)
bullet_text(s, [
    'Slow speeds, buffering, high latency, packet loss, and dropouts disrupt work, learning, calls, and gaming.',
    'The root cause can be Wi‑Fi signal weakness, interference, congestion, router faults, or an ISP/last-mile issue.',
    'Conventional speed tests report metrics but do not explain the probable cause or what the user should do next.',
    'This leads to blind restarts, unnecessary setting changes, and avoidable ISP support calls.'
], y=2.14, size=15.2, gap=.61)

# 3 — background
s = frame('Background', 'Background', 3)
tx(s, 'Home connectivity is a multi-layer experience.', .60, 1.62, 12, .28, 16, BLACK, True)
for label, x in [('Device & Wi‑Fi', .75), ('Home Router', 3.45), ('ISP Last Mile', 6.15), ('Internet Service', 8.85)]:
    node(s, label, x, 2.45, 1.85, .68)
for x in [2.60, 5.30, 8.00]: arrow(s, x, 2.79, x+.80, 2.79)
bullet_text(s, [
    'The same symptom can originate at different layers, so a single speed value is not enough for diagnosis.',
    'Useful observations include latency, jitter, packet loss, download/upload throughput, Wi‑Fi signal, and time patterns.',
    'A helpful tool must translate measurements into confidence-aware, human-readable guidance.'
], y=3.70, size=15, gap=.58)

# 4 — review
s = frame('Literature & Standards Review', 'Literature & RFC Review', 4)
tx(s, 'What existing measurement practice gives us — and what it misses.', .60, 1.55, 12, .3, 16, BLACK, True)
rows = [
    ('Speed tests', 'Download/upload throughput; ping', 'Usually no root-cause explanation'),
    ('RFC 6349', 'TCP throughput testing framework', 'Metrics need interpretation for home users'),
    ('RFC 2681', 'One-way delay measurement concepts', 'Delay alone does not isolate the failing layer'),
    ('Our approach', 'Multi-metric diagnosis + explanations', 'Actionable, personalized troubleshooting')]
xs = [.60, 3.2, 7.0]; ws = [2.45, 3.65, 5.7]
for x,w,t in zip(xs,ws,['Source / Practice','Contribution','Gap / Opportunity']):
    rect(s,x,2.15,w,.43,BLUE,BLUE); tx(s,t,x+.08,2.25,w-.16,.18,11,WHITE,True)
for r,(a,b,c) in enumerate(rows):
    y=2.59+r*.66
    for x,w,textv in zip(xs,ws,[a,b,c]):
        rect(s,x,y,w,.65,LIGHT if r%2==0 else WHITE,MID)
        tx(s,textv,x+.08,y+.10,w-.16,.44,11.3,BLACK, r==3)

# 5 — solution
s = frame('Proposed Solution', 'Proposed Solution', 5)
tx(s, 'An explainable diagnosis engine turns raw network tests into practical advice.', .60, 1.52, 12, .30, 16, BLACK, True)
steps = [('1. Measure','Latency • jitter • loss\nthroughput • Wi‑Fi signal'),('2. Analyze','Rules + ML classify\nlikely bottleneck'),('3. Explain','Show evidence and\nconfidence'),('4. Act','Give safe, prioritized\nrecommendations')]
for i,(head,sub) in enumerate(steps):
    x=.67+i*3.12
    node(s,head,x,2.30,2.2,.58,RGBColor(225,232,255),13)
    tx(s,sub,x+.08,3.02,2.04,.55,12,BLACK,False,PP_ALIGN.CENTER)
    if i<3: arrow(s,x+2.25,2.59,x+2.93,2.59)
bullet_text(s, [
    'Example: weak signal plus normal wired performance → “Move closer to the router or use 5 GHz.”',
    'Example: healthy local link but persistent external loss/latency → “Collect evidence and investigate ISP-side issues.”'
], y=4.45, size=14.5, gap=.70)

# 6 — optimization
s = frame('Optimization Proposed', 'Solution Optimization', 6)
tx(s, 'From raw metrics to a diagnosis users can trust.', .60, 1.55, 12, .3, 16, BLACK, True)
rows = [
    ('Single snapshot', 'Repeated tests over time', 'Separates transient congestion from persistent faults'),
    ('One aggregate score', 'Layer-aware feature set', 'Distinguishes Wi‑Fi, router, and ISP signals'),
    ('Opaque classification', 'Evidence + confidence display', 'Makes the diagnosis explainable'),
    ('Generic advice', 'Ranked, safe recommendations', 'Reduces unnecessary resets and support calls')]
xs=[.60,3.43,6.85,10.02]; ws=[2.75,3.32,3.10,2.70]
for x,w,t in zip(xs,ws,['Before','Enhancement','After','User value']):
    rect(s,x,2.12,w,.43,BLUE,BLUE); tx(s,t,x+.08,2.23,w-.16,.17,11,WHITE,True)
for i,row in enumerate(rows):
    y=2.56+i*.70
    for x,w,textv in zip(xs,ws,row):
        rect(s,x,y,w,.69,LIGHT if i%2==0 else WHITE,MID)
        tx(s,textv,x+.08,y+.11,w-.16,.45,10.7,BLACK, i==3)

# 7 — architecture
s = frame('Solution Architecture & Design', 'Architecture Design (UML Diagram)', 7)
tx(s, 'Component view', .60, 1.43, 12, .26, 16, BLACK, True)
node(s,'User App\nWeb / Mobile',.65,2.32,1.8,.82)
node(s,'Measurement\nCollector',3.25,2.32,1.8,.82)
node(s,'Feature\nExtractor',5.86,2.32,1.8,.82)
node(s,'Diagnosis\nEngine',8.46,2.32,1.8,.82)
node(s,'Explanation &\nRecommendation',11.05,2.32,1.7,.82)
for x in [2.48,5.08,7.69,10.28]: arrow(s,x,2.73,x+.70,2.73)
node(s,'Metrics Store\n(time-series history)',5.20,4.28,2.4,.72,RGBColor(242,242,255),11)
arrow(s,6.06,4.28,6.55,3.18); arrow(s,7.00,4.28,9.15,3.18)
tx(s,'Inputs: latency, jitter, packet loss, download/upload, signal strength, test history',.95,5.58,11.4,.25,12.5,GRAY,False,PP_ALIGN.CENTER)
tx(s,'Figure: Explainable Connectivity Diagnosis Architecture',.60,6.55,5.8,.18,10,BLACK)

# 8 — DFD
s = frame('Data Flow Diagram', 'Data Flow Diagram', 8)
tx(s, 'Data flow from a connectivity test to a user decision.', .60, 1.43, 12, .26, 16, BLACK, True)
node(s,'User\nstarts test',.70,2.45,1.65,.75)
node(s,'Collect local +\nnetwork metrics',3.25,2.45,2.15,.75)
node(s,'Normalize &\nvalidate readings',6.30,2.45,2.15,.75)
node(s,'Infer likely\nroot cause',9.35,2.45,1.8,.75)
node(s,'Explain +\nrecommend',11.85,2.45,1.1,.75,RGBColor(225,232,255),10)
for x1,x2 in [(2.38,3.20),(5.43,6.25),(8.48,9.30),(11.18,11.80)]: arrow(s,x1,2.82,x2,2.82)
node(s,'Diagnostic history /\nbaseline profile',5.40,4.35,2.4,.72,RGBColor(242,242,255),11)
arrow(s,6.60,4.35,7.15,3.25)
tx(s,'Output includes: likely cause • supporting evidence • confidence • next-best action',.82,5.65,11.7,.25,13,GRAY,False,PP_ALIGN.CENTER)
tx(s,'Figure: Data Flow Diagram (DFD)',.60,6.55,4.0,.18,10,BLACK)

# 9 — timeline
s = frame('Deployment Timeline', 'Deployment Timeline', 9)
tx(s, 'Phased delivery keeps the system measurable, safe, and useful.', .60, 1.55, 12, .3, 16, BLACK, True)
phases=[('Phase 1','Weeks 1–2','Define metrics, test cases, and baseline rules'),('Phase 2','Weeks 3–4','Build collector, test runner, and data pipeline'),('Phase 3','Weeks 5–6','Train/tune diagnosis logic and explanation layer'),('Phase 4','Weeks 7–8','Pilot, validate recommendations, and iterate')]
for i,(a,b,c) in enumerate(phases):
    x=.70+i*3.10
    rect(s,x,2.28,2.55,2.32,LIGHT,BLUE,True)
    tx(s,a,x+.15,2.55,2.25,.24,15,BLUE,True,PP_ALIGN.CENTER)
    tx(s,b,x+.15,2.97,2.25,.20,11,GRAY,True,PP_ALIGN.CENTER)
    tx(s,c,x+.18,3.42,2.18,.75,11.5,BLACK,False,PP_ALIGN.CENTER)
    if i<3: arrow(s,x+2.57,3.44,x+3.03,3.44)
tx(s,'Success measures: diagnosis accuracy • explanation usefulness • recommendation completion • reduced unnecessary support contacts',.80,5.53,11.7,.4,13,GRAY,False,PP_ALIGN.CENTER)

# 10 — bibliography
s = frame('References', 'Bibliography', 10)
refs=[
    '[1] IETF RFC 6349. Framework for TCP Throughput Testing. August 2011.',
    '[2] IETF RFC 2681. A Round-trip Delay Metric for IPPM. September 1999.',
    '[3] IETF RFC 3393. IP Packet Delay Variation Metric for IPPM. November 2002.',
    '[4] ITU-T Recommendation Y.1540. Internet protocol data communication service — IP packet transfer and availability performance parameters.']
tx(s,'Selected references supporting network measurement concepts.',.60,1.55,12,.30,16,BLACK,True)
for i,r in enumerate(refs):
    tx(s,r,.72,2.30+i*.72,11.65,.45,14,BLACK)

# 11 — closing
s = prs.slides.add_slide(blank)
rect(s,0,0,13.333,7.5)
tx(s,'AIORI 3',9.65,.10,.8,.18,7.2,BLUE,True,PP_ALIGN.RIGHT)
tx(s,'Closing',10.6,.10,2.35,.18,7.2,BLUE,True)
tx(s,'Thank You',.62,2.25,12.0,.6,30,BLUE,True,PP_ALIGN.CENTER)
tx(s,'Making home connectivity problems understandable,\nexplainable, and easier to resolve.',.80,3.20,11.75,.68,18,BLACK,False,PP_ALIGN.CENTER)
tx(s,'Questions?',.62,5.15,12.0,.35,18,NAVY,True,PP_ALIGN.CENTER)

prs.save(OUT)
print(OUT.resolve())
