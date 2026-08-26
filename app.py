import io, os, re, unicodedata
from pathlib import Path

import fitz
import openpyxl
import streamlit as st
from PIL import Image, ImageDraw, ImageFont
import arabic_reshaper
from bidi.algorithm import get_display

SUBJECTS = ['القراءة','الرياضيات','العلوم']
NAVY='#123E5A'; TEAL='#006E73'; PALE='#F4F9FA'; BORDER='#B9D5D8'
RED='#C62828'; ORANGE='#C86B00'; GREEN='#166534'; GRAY='#6B7280'; WHITE='#FFFFFF'


def ar(s):
    return get_display(arabic_reshaper.reshape(str(s)))


def norm(s):
    s=str(s or '')
    s=unicodedata.normalize('NFKC',s)
    s=s.replace('أ','ا').replace('إ','ا').replace('آ','ا').replace('ى','ي').replace('ة','ه')
    s=re.sub(r'[\u200e\u200f\u2066-\u2069]','',s)
    s=re.sub(r'\s+',' ',s).strip()
    return s


def words_by_block(page):
    groups={}
    W=page.rect.width
    for x0,y0,x1,y1,text,b,l,wno in page.get_text('words'):
        groups.setdefault(b,[]).append({'x0':x0,'y0':y0,'x1':x1,'y1':y1,'text':text,'line':l})
    out=[]
    for b,items in groups.items():
        items=sorted(items,key=lambda z:(z['line'],z['x0']))
        out.append({'b':b,'items':items,'texts':[i['text'] for i in items],'W':W})
    return out


def pct_values(block):
    vals=[]
    for i in block['items']:
        m=re.fullmatch(r'(\d+(?:\.\d+)?)%',i['text'])
        if m: vals.append((float(m.group(1)),i['x0']/block['W']))
    return vals


def num_values(block):
    vals=[]
    for i in block['items']:
        m=re.fullmatch(r'(\d+(?:\.\d+)?)',i['text'])
        if m: vals.append((float(m.group(1)),i['x0']/block['W']))
    return vals


def signed_value(block):
    for i in block['items']:
        t=i['text'].replace('−','-')
        m=re.fullmatch(r'(\d+(?:\.\d+)?)([+-])',t)
        if m:
            v=float(m.group(1)); return v if m.group(2)=='+' else -v
        m=re.fullmatch(r'([+-])(\d+(?:\.\d+)?)',t)
        if m:
            v=float(m.group(2)); return v if m.group(1)=='+' else -v
    return None


def priority(texts):
    s=' '.join(texts)
    if 'عاجلة' in s and 'أولوية' in s: return 'أولوية عاجلة'
    if 'مرتفعة' in s and 'أولوية' in s: return 'أولوية مرتفعة'
    if 'تحسين' in s and 'نوعي' in s: return 'تحسين نوعي'
    if 'استدامة' in s: return 'استدامة'
    return ''


def extract_school_name(pdf_name, doc):
    name=re.sub(r'^[\u200e\u200f\u2066-\u2069\s]*\d+\s*-\s*','',pdf_name)
    name=re.sub(r'\s*-\s*تقرير تحليل فجوات.*$','',name)
    name=re.sub(r'\.pdf$','',name,flags=re.I).strip()
    if re.search(r'(ابتدائية|متوسطة|ثانوية)',name): return name
    bls=words_by_block(doc[0])
    for b in bls:
        if any(k in b['texts'] for k in ['ابتدائية','متوسطة','ثانوية']):
            texts=sorted(b['items'],key=lambda z:z['x0'],reverse=True)
            return ' '.join(x['text'] for x in texts)
    return 'اسم المدرسة'



GRADE_WORDS = [
    'الأول','الثاني','الثالث','الرابع','الخامس','السادس',
    'السابع','الثامن','التاسع','العاشر','الحادي عشر','الثاني عشر'
]

def detect_grade(texts):
    joined=' '.join(texts)
    # Match longer grade names first.
    for g in sorted(GRADE_WORDS,key=len,reverse=True):
        if f'الصف {g}' in joined or ('الصف' in texts and g in joined):
            return f'الصف {g}'
    return 'صف غير محدد'


def unique_by_grade(records):
    out={}
    for r in records:
        g=r['grade']
        # Prefer the record that contains a signed gap/change.
        if g not in out:
            out[g]=r
    return out


def grade_options(raw, subject):
    p2=set(raw.get('p2_candidates',{}).get(subject,{}).keys())
    p3=set(raw.get('p3_candidates',{}).get(subject,{}).keys())
    common=sorted(p2 & p3)
    if common:
        return common
    return sorted(p2 | p3)


def resolve_data(raw, selections):
    data={
        'school_name':raw['school_name'],
        'p2':{},'p3':{},
        'p7':raw['p7'],'p10':raw['p10']
    }
    for s in SUBJECTS:
        grade=selections.get(s)
        if grade:
            if grade in raw.get('p2_candidates',{}).get(s,{}):
                data['p2'][s]=raw['p2_candidates'][s][grade]['data']
            if grade in raw.get('p3_candidates',{}).get(s,{}):
                data['p3'][s]=raw['p3_candidates'][s][grade]['data']
    data['selected_grades']=selections.copy()
    return data


def parse_pdf(pdf_bytes, pdf_name):
    doc=fitz.open(stream=pdf_bytes,filetype='pdf')
    raw={'school_name':extract_school_name(pdf_name,doc)}

    # صفحة 2: نجمع جميع الصفوف المتاحة لكل مادة ولا نختار من تلقاء أنفسنا.
    p2_candidates={s:{} for s in SUBJECTS}
    bls=words_by_block(doc[1])
    for s in SUBJECTS:
        records=[]
        for b in bls:
            if s in b['texts'] and 'الصف' in b['texts'] and len(pct_values(b))>=2 and signed_value(b) is not None:
                ps=pct_values(b)
                school=min(ps,key=lambda x:abs(x[1]-0.68))[0]
                admin=min(ps,key=lambda x:abs(x[1]-0.33))[0]
                records.append({
                    'grade':detect_grade(b['texts']),
                    'data':{'school':school,'admin':admin,'gap':signed_value(b)}
                })
        p2_candidates[s]=unique_by_grade(records)
    raw['p2_candidates']=p2_candidates

    # صفحة 3: نجمع جميع الصفوف المتاحة لكل مادة من جدول التغير بين العامين.
    p3_candidates={s:{} for s in SUBJECTS}
    bls=words_by_block(doc[2])
    for s in SUBJECTS:
        records=[]
        for b in bls:
            if s in b['texts'] and 'الصف' in b['texts'] and len(pct_values(b))>=2 and signed_value(b) is not None:
                ps=pct_values(b)
                y1446=min(ps,key=lambda x:abs(x[1]-0.40))[0]
                y1445=min(ps,key=lambda x:abs(x[1]-0.53))[0]
                records.append({
                    'grade':detect_grade(b['texts']),
                    'data':{'y1445':y1445,'y1446':y1446,'change':signed_value(b)}
                })
        p3_candidates[s]=unique_by_grade(records)
    raw['p3_candidates']=p3_candidates

    # صفحة 7
    p7={}; bls=words_by_block(doc[6])
    keys={
        'التحصيل العلمي':['التحصيل','العلمي'],
        'نواتج التعلم':['نواتج','التعلم'],
        'التعليم والتعلم':['التعليم','والتعلم']
    }
    for k,toks in keys.items():
        cands=[b for b in bls if all(t in b['texts'] for t in toks) and len(pct_values(b))>=2]
        if cands:
            b=cands[0]; ps=pct_values(b)
            degree=min(ps,key=lambda x:abs(x[1]-0.56))[0]
            p7[k]={'value':degree,'priority':priority(b['texts'])}
    raw['p7']=p7

    # صفحة 10
    p10={}; bls=words_by_block(doc[9])
    codes={'القراءة':'3-1-1-1','الرياضيات':'3-1-1-2','العلوم':'3-1-1-3'}
    for s,code in codes.items():
        cands=[b for b in bls if code in b['texts']]
        if cands:
            b=cands[0]; nums=num_values(b)
            degree=min(nums,key=lambda x:abs(x[1]-0.35))[0]
            p10[s]={'value':degree,'priority':priority(b['texts'])}
    raw['p10']=p10
    return raw


def extract_proposal(xlsx_bytes, school_name):
    wb=openpyxl.load_workbook(io.BytesIO(xlsx_bytes),data_only=True)
    target=norm(school_name)
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=True):
            vals=['' if v is None else str(v) for v in row]
            joined=' | '.join(vals)
            if target and target in norm(joined):
                candidates=[v.strip() for v in vals if re.search(
                    r'لا حاجة|اقتراح زيادة|زيادة عدد حصص|زيادة عدد الحصص',v)]
                if candidates:
                    return max(candidates,key=len)
    return 'لم يتم العثور على المقترح في ملف Excel.'


def validate(data):
    labels={'p2':'الفقرة 1','p3':'الفقرة 2','p7':'الفقرة 3','p10':'الفقرة 4'}
    return [labels[k] for k in labels if len(data.get(k,{}))!=3]


def font_path(bold=False):
    """اختر خطًا يدعم العربية فعليًا، ولا تستخدم أي خط عشوائي."""
    import subprocess

    families = (
        ['Noto Kufi Arabic', 'Noto Sans Arabic', 'DejaVu Sans', 'Liberation Sans']
        if bold else
        ['Noto Sans Arabic', 'Noto Kufi Arabic', 'DejaVu Sans', 'Liberation Sans']
    )

    # fc-match هو الأكثر موثوقية على Linux/Streamlit لأنه يعيد المسار الحقيقي للخط.
    for family in families:
        try:
            style = ':style=Bold' if bold else ''
            result = subprocess.run(
                ['fc-match', '-f', '%{file}', family + style],
                capture_output=True, text=True, timeout=3
            )
            path=result.stdout.strip()
            if path and os.path.exists(path):
                return path
        except Exception:
            pass

    # مسارات معروفة كحل احتياطي، لكن فقط لخطوط تدعم العربية.
    candidates = [
        '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf' if bold
        else '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',
        '/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf' if bold
        else '/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf',
    ]
    for p in candidates:
        if os.path.exists(p):
            return p

    return None


def proposal_lines(text, max_chars=70):
    text=str(text).strip()
    if len(text)<=max_chars:
        return [text]
    words=text.split()
    lines=[]; cur=[]
    for w in words:
        test=' '.join(cur+[w])
        if len(test)>max_chars and cur:
            lines.append(' '.join(cur)); cur=[w]
        else:
            cur.append(w)
    if cur: lines.append(' '.join(cur))
    return lines[:3]


def make_report_image(data, proposal):
    W,H=1600,2263
    im=Image.new('RGB',(W,H),WHITE); d=ImageDraw.Draw(im)
    FB=font_path(True); FR=font_path(False)

    def F(sz,b=True):
        path = FB if b else FR
        if not path:
            raise RuntimeError('لم يتم العثور على خط يدعم العربية على خادم Streamlit.')
        return ImageFont.truetype(path, sz)
    def center(txt,x,y,sz,color=NAVY,b=True):
        t=ar(txt); f=F(sz,b); box=d.textbbox((0,0),t,font=f)
        d.text((x-(box[2]-box[0])/2,y),t,font=f,fill=color)
    def rr(x1,y1,x2,y2,fill=WHITE,outline=BORDER,r=22,w=3):
        d.rounded_rectangle((x1,y1,x2,y2),radius=r,fill=fill,outline=outline,width=w)
    def sec(y,title):
        d.rounded_rectangle((30,y,1570,y+70),radius=18,fill=NAVY)
        center(title,800,y+12,27,WHITE)

    # رأس التقرير
    center('المملكة العربية السعودية',230,32,24)
    center('وزارة التعليم',230,70,24)
    center('الإدارة العامة للتعليم بمنطقة الباحة',230,108,21)
    center('مساعد مدير عام التعليم',230,144,20)
    center('جودة خدمات المركز الوطني للمناهج',230,180,20)

    center('مقترحات الخطط الدراسية',820,28,48)
    rr(480,100,1140,180,fill=TEAL,outline=TEAL,r=28)
    center(data['school_name'],810,118,30,WHITE)

    logo_path=Path(__file__).with_name('moe_logo.png')
    if logo_path.exists():
        logo=Image.open(logo_path).convert('RGB')
        logo.thumbnail((280,190))
        im.paste(logo,(1270,30))

    # 1
    sec(225,'1. الفرق بين المدرسة والإدارة')
    rr(30,300,1570,820)
    xs=[250,700,1150]
    for s,x in zip(SUBJECTS,xs):
        v=data['p2'][s]
        center(s,x,330,23)
        center(data.get('selected_grades',{}).get(s,''),x,365,15,GRAY)
        rr(x-170,395,x+170,720,fill=PALE)
        center('نتيجة المدرسة',x,420,18,TEAL)
        center(f"{v['school']:.1f}%",x,463,40,TEAL)
        center('نتيجة الإدارة',x,535,18,GRAY)
        center(f"{v['admin']:.1f}%",x,578,40,GRAY)
        center('الفارق',x,650,17,RED)
        center(f"{v['gap']:.1f} نقطة",x,682,25,RED)

    # 2
    sec(850,'2. التغيير بين عامي 1445هـ و1446هـ')
    rr(30,925,1570,1325)
    for s,x in zip(SUBJECTS,xs):
        v=data['p3'][s]
        rr(x-190,970,x+190,1265,fill=PALE)
        center(s,x,980,23)
        center(data.get('selected_grades',{}).get(s,''),x,1015,15,GRAY)
        center('1445هـ',x-80,1050,17,GRAY)
        center(f"{v['y1445']:.1f}%",x-80,1090,31,GREEN)
        center('1446هـ',x+80,1050,17,GRAY)
        center(f"{v['y1446']:.1f}%",x+80,1090,31,RED if v['change']<0 else GREEN)
        col=RED if v['change']<0 else GREEN
        center(f"التغير: {v['change']:+.1f} نقطة",x,1180,22,col)

    # 3 و4
    d.rounded_rectangle((30,1355,775,1425),radius=18,fill=NAVY)
    d.rounded_rectangle((825,1355,1570,1425),radius=18,fill=NAVY)
    center('3. نتائج المدرسة في المجالات',402,1367,25,WHITE)
    center('4. مستوى المدرسة في المجالات (مؤشرات)',1195,1367,22,WHITE)
    rr(30,1435,775,1900); rr(825,1435,1570,1900)

    keys=['التحصيل العلمي','نواتج التعلم','التعليم والتعلم']
    for k,x in zip(keys,[160,402,644]):
        v=data['p7'][k]
        rr(x-100,1490,x+100,1765,fill=PALE)
        center(k,x,1520,17)
        center(f"{v['value']:.2f}%",x,1600,30)
        color=RED if 'عاجلة' in v['priority'] else ORANGE if 'مرتفعة' in v['priority'] else GREEN
        center(v['priority'],x,1685,15,color)

    for i,s in enumerate(SUBJECTS):
        y=1495+i*105; v=data['p10'][s]
        rr(880,y,1515,y+82,fill=PALE)
        center(s,1420,y+17,19)
        center(f"{v['value']:.2f}%",1050,y+10,31)
        color=RED if 'عاجلة' in v['priority'] else ORANGE if 'مرتفعة' in v['priority'] else GREEN
        center(v['priority'],1230,y+20,16,color)

    # الاقتراح
    rr(30,1940,1570,2180,fill='#F8FCFC',outline=TEAL,w=4)
    d.rounded_rectangle((620,1910,980,1980),radius=18,fill=TEAL)
    center('الاقتراح',800,1923,28,WHITE)

    lines=proposal_lines(proposal)
    start_y=2025 if len(lines)==1 else 2005
    for i,line in enumerate(lines):
        center(line,800,start_y+i*48,25,GREEN)

    bio=io.BytesIO()
    im.save(bio,format='PNG',quality=95)
    return bio.getvalue()


def png_to_pdf(png_bytes):
    image=Image.open(io.BytesIO(png_bytes)).convert('RGB')
    bio=io.BytesIO()
    image.save(bio,format='PDF',resolution=150.0)
    return bio.getvalue()


def priority_color(p):
    if 'عاجلة' in p: return RED
    if 'مرتفعة' in p: return ORANGE
    if 'استدامة' in p: return GREEN
    return TEAL


st.set_page_config(page_title='مولد تقارير المدارس',layout='wide')

st.markdown("""
<style>
.block-container {max-width: 1450px; padding-top: 1.3rem;}
h1,h2,h3,p,div,span,label {direction: rtl; text-align: right;}
[data-testid="stDataFrame"] {direction: rtl;}
.report-card {
    border:1px solid #d8e6e8; border-radius:14px; padding:14px;
    background:#f8fbfb; margin-bottom:8px;
}
.metric-title {font-weight:700; color:#123E5A; margin-bottom:6px;}
.metric-value {font-size:28px; font-weight:800; color:#123E5A;}
.proposal-box {
    border:2px solid #006E73; border-radius:16px; padding:18px;
    background:#f7fcfc; color:#166534; font-size:22px; font-weight:700;
}
</style>
""", unsafe_allow_html=True)

st.title('مولد تقارير مقترحات الخطط الدراسية')

st.caption('يرفع المستخدم تقرير المدرسة وملف Excel فقط. الأرقام من الصفحات 2 و3 و7 و10، والمقترح من Excel. يدعم الصفوف المختلفة دون افتراض صف ثابت.')

c1,c2=st.columns(2)
with c1:
    pdf=st.file_uploader('تقرير المدرسة PDF',type=['pdf'])
with c2:
    xlsx=st.file_uploader('ملف Excel المعتمد',type=['xlsx','xlsm'])

if 'raw' not in st.session_state:
    st.session_state.raw=None
    st.session_state.data=None
    st.session_state.proposal=None
    st.session_state.png=None
    st.session_state.pdf=None

if pdf and xlsx:
    if st.button('استخراج البيانات',type='primary'):
        try:
            raw=parse_pdf(pdf.getvalue(),pdf.name)
            proposal=extract_proposal(xlsx.getvalue(),raw['school_name'])
            st.session_state.raw=raw
            st.session_state.proposal=proposal
            st.session_state.data=None
            st.session_state.png=None
            st.session_state.pdf=None
            st.success('تم استخراج البيانات الأولية. حدّد الصف المطلوب إذا ظهر أكثر من صف.')
        except Exception as e:
            st.exception(e)
else:
    st.info('ارفع ملف PDF وملف Excel أولًا.')

raw=st.session_state.raw
if raw:
    st.divider()
    st.subheader('اختيار الصفوف لنتائج نافس')
    st.caption('إذا وجد البرنامج صفًا واحدًا للمادة فسيعتمده تلقائيًا. إذا وجد أكثر من صف فاختر الصف الذي تريد أن يبنى عليه التقرير.')

    selections={}
    all_ok=True
    cols=st.columns(3)
    for col,s in zip(cols,SUBJECTS):
        opts=grade_options(raw,s)
        with col:
            if not opts:
                st.error(f'{s}: لم يتم العثور على صف صالح في الصفحتين 2 و3.')
                all_ok=False
            elif len(opts)==1:
                selections[s]=opts[0]
                st.success(f'{s}: {opts[0]}')
            else:
                selections[s]=st.selectbox(
                    f'اختر الصف لمادة {s}',
                    opts,
                    key=f'grade_{s}'
                )

    if all_ok and st.button('اعتماد الصفوف وإظهار المراجعة',type='primary'):
        data=resolve_data(raw,selections)
        problems=validate(data)
        if problems:
            st.error('لم يكتمل الاستخراج بعد اختيار الصف في: '+ '، '.join(problems))
            st.session_state.data=None
        else:
            st.session_state.data=data
            st.session_state.png=None
            st.session_state.pdf=None
            st.success('تم اعتماد الصفوف واستخراج جميع الفقرات بنجاح.')

data=st.session_state.data
proposal=st.session_state.proposal

if data:
    st.divider()
    st.subheader(data['school_name'])

    st.markdown('### مراجعة القيم قبل إنشاء التقرير النهائي')

    # الفقرة 1
    st.markdown('#### 1. الفرق بين المدرسة والإدارة')
    cols=st.columns(3)
    for col,s in zip(cols,SUBJECTS):
        a=data['p2'][s]
        with col:
            st.markdown(
                f"""<div class="report-card">
                <div class="metric-title">{s}</div>
                <div style="color:#6B7280;font-size:13px">{data.get('selected_grades',{}).get(s,'')}</div>
                <div>المدرسة: <b>{a['school']:.1f}%</b></div>
                <div>الإدارة: <b>{a['admin']:.1f}%</b></div>
                <div style="color:{RED};font-weight:700">الفارق: {a['gap']:.1f} نقطة مئوية</div>
                </div>""", unsafe_allow_html=True)

    # الفقرة 2
    st.markdown('#### 2. التغيير بين عامي 1445هـ و1446هـ')
    cols=st.columns(3)
    for col,s in zip(cols,SUBJECTS):
        b=data['p3'][s]
        colr=RED if b['change']<0 else GREEN
        with col:
            st.markdown(
                f"""<div class="report-card">
                <div class="metric-title">{s}</div>
                <div style="color:#6B7280;font-size:13px">{data.get('selected_grades',{}).get(s,'')}</div>
                <div>1445هـ: <b>{b['y1445']:.1f}%</b></div>
                <div>1446هـ: <b>{b['y1446']:.1f}%</b></div>
                <div style="color:{colr};font-weight:700">التغير: {b['change']:+.1f} نقطة مئوية</div>
                </div>""", unsafe_allow_html=True)

    # الفقرة 3
    st.markdown('#### 3. نتائج المدرسة في المجالات')
    cols=st.columns(3)
    for col,k in zip(cols,['التحصيل العلمي','نواتج التعلم','التعليم والتعلم']):
        v=data['p7'][k]
        with col:
            st.markdown(
                f"""<div class="report-card">
                <div class="metric-title">{k}</div>
                <div class="metric-value">{v['value']:.2f}%</div>
                <div style="color:{priority_color(v['priority'])};font-weight:700">{v['priority']}</div>
                </div>""", unsafe_allow_html=True)

    # الفقرة 4
    st.markdown('#### 4. مستوى المدرسة في المجالات (مؤشرات)')
    cols=st.columns(3)
    for col,s in zip(cols,SUBJECTS):
        v=data['p10'][s]
        with col:
            st.markdown(
                f"""<div class="report-card">
                <div class="metric-title">{s}</div>
                <div class="metric-value">{v['value']:.2f}%</div>
                <div style="color:{priority_color(v['priority'])};font-weight:700">{v['priority']}</div>
                </div>""", unsafe_allow_html=True)

    st.markdown('#### الاقتراح')
    st.markdown(f'<div class="proposal-box">{proposal}</div>',unsafe_allow_html=True)

    st.warning('راجع القيم أعلاه. إذا كانت صحيحة اضغط زر إنشاء التقرير النهائي.')

    if st.button('إنشاء التقرير النهائي',type='primary'):
        png=make_report_image(data,proposal)
        pdf_bytes=png_to_pdf(png)
        st.session_state.png=png
        st.session_state.pdf=pdf_bytes

if st.session_state.get('png'):
    st.divider()
    st.markdown('## التقرير النهائي')
    st.image(st.session_state.png,use_container_width=True)
    c1,c2=st.columns(2)
    with c1:
        st.download_button(
            'تحميل التقرير PNG',
            st.session_state.png,
            file_name=f"تقرير_{data['school_name']}.png",
            mime='image/png',
            use_container_width=True
        )
    with c2:
        st.download_button(
            'تحميل التقرير PDF',
            st.session_state.pdf,
            file_name=f"تقرير_{data['school_name']}.pdf",
            mime='application/pdf',
            use_container_width=True
        )
