#!/usr/bin/env python3
"""Hassan Raza — CV, 'AI-native builder' edition (2 pages, ATS-friendly).
Angle: 12 years of shipping mobile/web + a one-person studio (BigInt Games) run end-to-end with AI agents."""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import nsdecls, qn
from docx.oxml import parse_xml
import os, io

BLUE=RGBColor(0x1D,0x4E,0xD8); DARK=RGBColor(0x0B,0x12,0x2B); GRAY=RGBColor(0x64,0x74,0x8B); TEXT=RGBColor(0x2B,0x36,0x4B); ORANGE=RGBColor(0xEA,0x58,0x0C)
doc=Document()
for s in doc.sections: s.top_margin=Cm(1.1); s.bottom_margin=Cm(1.0); s.left_margin=Cm(1.5); s.right_margin=Cm(1.5)
st=doc.styles['Normal']; st.font.name='Calibri'; st.font.size=Pt(9.5); st.font.color.rgb=TEXT
st.paragraph_format.space_after=Pt(1); st.paragraph_format.space_before=Pt(0); st.paragraph_format.line_spacing=1.06

def run(p,text,size=9.5,bold=False,color=TEXT,italic=False):
    r=p.add_run(text); r.font.size=Pt(size); r.font.bold=bold; r.font.italic=italic; r.font.color.rgb=color; r.font.name='Calibri'; return r
def rule(color=BLUE,sz=6,before=1,after=3):
    p=doc.add_paragraph(); p.paragraph_format.space_before=Pt(before); p.paragraph_format.space_after=Pt(after)
    p._p.get_or_add_pPr().append(parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="{sz}" w:space="1" w:color="{color}"/></w:pBdr>'))
def heading(text):
    p=doc.add_paragraph(); p.paragraph_format.space_before=Pt(8); p.paragraph_format.space_after=Pt(0)
    run(p,text.upper(),10.5,True,BLUE); rule(BLUE,4,0,1)
def bullet(text,size=9.2):
    p=doc.add_paragraph(style='List Bullet'); p.paragraph_format.space_after=Pt(0); p.paragraph_format.left_indent=Cm(0.55)
    # lead-in before an em dash goes bold
    if " — " in text:
        a,b=text.split(" — ",1); run(p,a+" — ",size,True,DARK); run(p,b,size)
    else: run(p,text,size)
def shade(cell,hexcolor):
    cell._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="{hexcolor}"/>'))
def noborders(table):
    for row in table.rows:
        for cell in row.cells:
            cell._tc.get_or_add_tcPr().append(parse_xml(f'<w:tcBorders {nsdecls("w")}>'+"".join(f'<w:{e} w:val="none" w:sz="0" w:space="0" w:color="auto"/>' for e in ("top","left","bottom","right"))+'</w:tcBorders>'))
def job(role,company,date,loc,bullets):
    p=doc.add_paragraph(); p.paragraph_format.space_before=Pt(6); p.paragraph_format.space_after=Pt(0); run(p,role,10.5,True,DARK)
    p=doc.add_paragraph(); p.paragraph_format.space_after=Pt(1); run(p,company,9.5,True,BLUE); run(p,f"   {date}   ·   {loc}",8.8,False,GRAY)
    for b in bullets: bullet(b)

# ---------- header ----------
t=doc.add_table(rows=1,cols=2); t.alignment=WD_TABLE_ALIGNMENT.CENTER; noborders(t)
c=t.rows[0].cells[0]; c.width=Cm(14.5); p=c.paragraphs[0]; run(p,"HASSAN RAZA",25,True,DARK)
p=c.add_paragraph(); run(p,"Senior Software Engineer  ·  AI-Native Product Builder  ·  Founder, BigInt Games",11.5,True,BLUE)
p=c.add_paragraph(); p.paragraph_format.space_before=Pt(3); run(p,"+971 58 557 8686   |   mhassanreza@gmail.com   |   Dubai, UAE",9,False,GRAY)
p=c.add_paragraph(); run(p,"linkedin.com/in/muhammadhassanraza   |   bigintgames.com   |   mhassanreza.github.io/mhassanreza-cv",9,True,BLUE)
pc=t.rows[0].cells[1]; pc.width=Cm(3.4); pp=pc.paragraphs[0]; pp.alignment=WD_ALIGN_PARAGRAPH.RIGHT
photo=os.path.join(os.path.dirname(os.path.abspath(__file__)),"profile.jpg")
if os.path.exists(photo):
    from PIL import Image, ImageDraw
    im=Image.open(photo).convert("RGBA"); s=min(im.size); im=im.crop(((im.width-s)//2,(im.height-s)//2,(im.width-s)//2+s,(im.height-s)//2+s)).resize((500,500),Image.LANCZOS)
    m=Image.new("L",(500,500),0); ImageDraw.Draw(m).ellipse((0,0,500,500),fill=255)
    out=Image.new("RGBA",(500,500),(0,0,0,0)); out.paste(im,(0,0),m); ImageDraw.Draw(out).ellipse((0,0,499,499),outline=(29,78,216,255),width=12)
    b=io.BytesIO(); out.save(b,"PNG"); b.seek(0); pp.add_run().add_picture(b,width=Cm(3.0))
rule(BLUE,10,2,4)

# ---------- the pitch (shaded box) ----------
box=doc.add_table(rows=1,cols=1); box.alignment=WD_TABLE_ALIGNMENT.CENTER; noborders(box); cell=box.rows[0].cells[0]; shade(cell,"EEF3FF")
p=cell.paragraphs[0]; p.paragraph_format.space_before=Pt(4); p.paragraph_format.space_after=Pt(2)
run(p,"I build and ship software with AI as a first-class engineer on the team. ",10,True,DARK)
run(p,"Twelve years of Android, iOS, cross-platform and full-stack delivery for enterprise and government clients — and since 2026 a one-person games studio, "
      "BigInt Games, that has put five new titles on Google Play and the App Store in eight months by running the entire pipeline through AI agents: "
      "specification, code, art generation, QA harnesses, localization into 10 languages, store optimization and the marketing site. "
      "Result: 5–6× the throughput of a conventional team, with production apps serving 750,000+ downloads.",9.5)
p=cell.add_paragraph(); p.paragraph_format.space_after=Pt(4)
run(p,"Toolchain: ",9,True,ORANGE); run(p,"Claude Code (agentic coding), GitHub Copilot, Cursor, Google Gemini API (in-app AI features), AI image generation with programmatic validation.",9)

# ---------- impact strip (single line — renders the same in Word, Pages, Quick Look, ATS parsers) ----------
p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_before=Pt(6); p.paragraph_format.space_after=Pt(2)
for i,(big,small) in enumerate([("750K+"," lifetime downloads"),("8"," titles live on both stores"),("10"," languages shipped"),("5–6×"," delivery speed with AI")]):
    if i: run(p,"     ·     ",11,False,GRAY)
    run(p,big,14,True,BLUE); run(p,small,9,False,GRAY)

# ---------- how I work with AI ----------
heading("How I build with AI")
for b in [
 "Agentic development — I direct AI coding agents (Claude Code) through whole features: spec → implementation → headless build → device test; I own product decisions, architecture and verification on real devices.",
 "Automated 'render-and-look' QA — every build captures every screen headlessly; an agent reviews the screenshots against the mockups before anything reaches a phone. Caught dozens of layout bugs no asset check would.",
 "AI art pipelines with hard gates — generated backgrounds, sprites and animation sheets pass programmatic checks (dedupe, scale-lock, spoiler and text scans) before import.",
 "Localization at scale — 60 authored mystery cases and the full UI in 10 languages via agent fleets with pilot-first machine validation.",
 "Growth engineering — generated a 130-page SEO site (structured data, hreflang, sitemaps), ASO sheets for both stores, GA4/UTM attribution, Search Console — all agent-assisted.",
 "AI inside the product — Gemini-powered food scanning (calories & macros from a photo) shipped to a 750K-download app with a Premium subscription.",
]: bullet(b)

# ---------- experience ----------
heading("Experience")
job("Founder & Lead Engineer","BigInt Games — independent mobile studio & publisher","2020 – Present (games studio since 2026)","Dubai, UAE",[
 "Ship complete products solo — design, Unity/C# and native code, Firebase back end, monetization, release and marketing — using AI agents as the engineering team.",
 "2026 line-up — Agent X (daily murder-mystery, 10 languages, live content pipeline), Puck Clash, Cannon Valley (tower-defense remake), Brick Broke, and Anatomy: AI Diet & Food Scan (iOS), all live on Google Play and/or the App Store.",
 "Gym Workout: AI Diet & Planner — 750,000+ downloads; AI Food Scan (Gemini), 4-step plan generator, 7-day diet, Premium subscription + lifetime IAP, AdMob.",
 "Platform stack — Firebase Auth/Firestore/Hosting/FCM/Crashlytics, Play Billing & StoreKit 2, AdMob with UMP consent and ATT, Apple/Google account linking, cloud save, push campaigns.",
 "bigintgames.com — generated studio hub with per-title pages, localized content pages, free health tools, legal pages and press kit; verified in Search Console with full structured data.",
])
job("Senior Software Engineer — Mobile & Full Stack","Inlogic IT Solutions LLC","Jan 2020 – Present","Dubai, UAE",[
 "Deliver web and mobile solutions with React.js/Next.js, React Native, Node.js and native Android (Kotlin, Jetpack Compose) / iOS (SwiftUI), integrating REST APIs and SQL/NoSQL data layers.",
 "Bilingual (Arabic/English, RTL) e-services portals and admin dashboards; secure auth flows (OAuth2, JWT, SSO) and role-based access for enterprise data governance.",
 "Integrated Microsoft Dynamics 365 / Dataverse and MS SQL Server back ends; CI/CD on GitHub Actions and Azure DevOps.",
 "AI-assisted workflow — introduced Copilot, Claude and Cursor to the team: 5–6× faster delivery on Himenus, OOKAAZ and Olamop without quality regressions.",
])
job("Senior Software Engineer (Outsourced — Government Sector)","AAAID — Arab Authority for Agricultural Investment and Development","Sep 2022 – Jan 2025","Dubai, UAE",[
 "Delivered the public-facing E-Services portal and native apps (Android & iOS) for a pan-Arab governmental authority serving 500+ internal and external users.",
 "Tarasul correspondence system (Xamarin Forms): SSO, PDF annotation, audio recording, dynamic theming, audit-compliant data flows.",
 "Bilingual RTL interfaces, accessibility-compliant design, interactive dashboards; aligned delivery with the authority's digital-transformation roadmap.",
])
job("Principal Analyst Developer — Android","Cheetay","Sep 2018 – Dec 2019","Lahore, Pakistan",[
 "Led Android and iOS teams for a high-traffic e-commerce/delivery platform; owned sprint deliverables, technical roadmap, code review and CI/CD discipline; shipped 4 Android apps end to end.",
])
job("Android Developer","Insight","Mar 2015 – Sep 2018","Lahore, Pakistan",[
 "Offline-first applications with synchronized SQL data layers and REST integration; live production support and root-cause analysis for business-critical systems.",
])
job("Trainee Android Developer","Techverx","Aug 2013 – Jan 2015","Lahore, Pakistan",[
 "Built and maintained production Android features under senior guidance; established engineering fundamentals.",
])

# ---------- skills ----------
heading("Skills")
sk=doc.add_table(rows=4,cols=2); sk.alignment=WD_TABLE_ALIGNMENT.CENTER; noborders(sk)
rows=[("AI-assisted engineering","Claude Code, GitHub Copilot, Cursor, Gemini API, prompt & agent-workflow design, automated visual QA"),
      ("Mobile","Kotlin, Jetpack Compose, SwiftUI, Unity/C#, React Native, Xamarin Forms/.NET MAUI, Samsung Tizen, Android TV"),
      ("Web & back end","React.js, Next.js, TypeScript, Node.js/Express, REST, MS SQL/T-SQL, Dynamics 365/Dataverse, Firebase (Auth, Firestore, Hosting, FCM)"),
      ("Product & growth","AdMob/UMP/ATT, IAP & subscriptions (Play Billing, StoreKit 2), ASO, SEO & structured data, GA4/UTM, CI/CD (GitHub Actions, Azure DevOps), Figma")]
for i,(k,v) in enumerate(rows):
    a=sk.rows[i].cells[0]; a.width=Cm(4.2); b=sk.rows[i].cells[1]; b.width=Cm(13.7)
    p=a.paragraphs[0]; p.paragraph_format.space_after=Pt(1); run(p,k,9,True,DARK)
    p=b.paragraphs[0]; p.paragraph_format.space_after=Pt(1); run(p,v,9)

# ---------- selected products ----------
heading("Selected products")
def prod(name,plat,desc):
    p=doc.add_paragraph(); p.paragraph_format.space_before=Pt(2); p.paragraph_format.space_after=Pt(0)
    run(p,name,9.5,True,DARK); run(p,f"  [{plat}]",8.3,False,GRAY); run(p,f" — {desc}",9)
prod("Agent X: Detective","Android · iOS","daily murder-mystery game — 60 authored cases, 10 languages, cloud content pipeline, informant chat, seasons & medals.")
prod("Gym Workout: AI Diet & Planner / Anatomy: AI Diet & Food Scan","Android · iOS","750K+ downloads; Gemini food scan, interactive muscle atlas, diet planner, Premium.")
prod("Cannon Valley · Puck Clash · Brick Broke","Android · iOS","modern remake of a classic tower defense; sling-puck arcade; neon brick-breaker with 120 levels — Unity, Firebase, AdMob, IAP.")
prod("CAST4k","Android · iOS · Android TV · Tizen","smart IPTV streaming platform with multi-device sync and EPG — 10,000+ downloads.")
prod("AAAID E-Services & Tarasul","Web · Android · iOS","government e-services portal and correspondence system — SSO, PDF annotation, bilingual RTL, 500+ users.")
prod("Olamop · Himenus · OOKAAZ · iB2B","Web · Android · iOS","on-demand services with Stripe/Apple Pay; restaurant ordering; e-commerce; real-time team messaging with calls.")

# ---------- education ----------
heading("Education & languages")
p=doc.add_paragraph(); run(p,"Bachelor of Computer Science",9.5,True,DARK); run(p,"  ·  University of Management and Technology (UMT), Lahore  ·  2009 – 2013",9,False,GRAY)
p=doc.add_paragraph(); run(p,"English (proficient)  ·  Urdu (native)  ·  Arabic (working, via bilingual product delivery)",9)

out=os.path.join(os.path.dirname(os.path.abspath(__file__)),"Hassan Raza - CV - AI Builder.docx"); doc.save(out); print("saved",out)
