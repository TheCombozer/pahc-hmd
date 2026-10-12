from pathlib import Path
from io import BytesIO
import base64, json, re
import pandas as pd
import requests
import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(page_title='Annual Broiler Health Monitor', layout='wide')
BASE=Path(__file__).resolve().parent
TEMPLATE=(BASE/'dashboard_template.html').read_text('utf-8')
BUNDLED=(BASE/'data.xlsx').read_bytes()
ADMIN_PASSWORD='0000'  # Fixed admin password requested by owner.

# Fill these once before deployment. The Excel workbook is the source of truth.
GITHUB_OWNER='TheCombozer'
GITHUB_REPO='pahc-hmd'
GITHUB_BRANCH='main'
GITHUB_DATA_PATH='data.xlsx'
API_VERSION='2022-11-28'

NAME_MAP={'ARS':'Airsacculitis','ASC':'Ascites','BD3':'Bursa Damage','BSM':'Bursameter Score','BoW':'Body Weight','BuW':'Bursa Weight','CS':'Cellular Sloughing','DBJT':'Jejunum Tone','DbBAL':'Ballooning','DbFP':'Feed Passage','DbIAC':'Ileum Abnormal Content','DbIH':'Ileum Hyperemia','DbIT':'Ileum Tone','DbITH':'Ileum Thickness','DbJAC':'Jejunum Abnormal Content','DbJH':'Jejunum Hyperemia','DbJTH':'Jejunum Thickness','EYED':'Eyelid Irritation','FHN1':'Femoral Head Necrosis','GIZ':'Gizzard Erosion','HCK':'Hock Burn','HRT':'Heart Lesion','IP':'Inflammatory Process','LE':'Litter Eater','LNG':'Pneumonia','ML':'Mouth Lesion','PDM2':'Pododermatitis','PER':'Peritonitis','RY':'Retained Yolk Sac','SCR':'Skin Scratch','Sex':'Sex','TD':'Tibial Dyschondroplasia','THY':'Thymus Atrophy','TRA':'Tracheitis','WB':'Woody Breast','WS':'White Striping','gAC':'Gross E. acervulina','gMX':'Gross E. maxima','gTN':'Gross E. tenella','mMX':'Microscopic E. maxima','TMLS':'Total Mean Lesion Score'}
REQUIRED=['Farm','House','Bird','Age (days)','Variable','Variable Score','Metric','Description','System','Complex','Customer','State','Year','Month','Season','Session date','Reference','Prevalence']
SYSTEM_ORDER=['All','Respiratory','Digestive','Coccidiosis','Immune','Musculoskeletal','General Observation & Welfare','Metrics','Other']

def final_system(code,source):
    if code in {'gAC','gMX','gTN','mMX','TMLS'}: return 'Coccidiosis'
    if code in {'HRT','EYED','ARS','TRA','LNG','ASC'}: return 'Respiratory'
    if code in {'IP','HCK','PDM2','SCR'}: return 'General Observation & Welfare'
    if code in {'FHN1','TD','WB','WS'}: return 'Musculoskeletal'
    if code=='PER': return 'Other'
    if source=='Skeletal': return 'Musculoskeletal'
    if source=='Integument': return 'General Observation & Welfare'
    if source=='Metabolic': return 'Respiratory'
    return source

def read_raw(blob):
    xls=pd.ExcelFile(BytesIO(blob),engine='openpyxl')
    if 'Raw Data' not in xls.sheet_names: raise ValueError("Sheet 'Raw Data' was not found.")
    df=pd.read_excel(xls,sheet_name='Raw Data')
    missing=[x for x in REQUIRED if x not in df.columns]
    if missing: raise ValueError('Missing columns: '+', '.join(missing))
    return df

def compact_payload(blob):
    df=read_raw(blob).copy()
    for c in ['House','Bird','Age (days)','Variable Score','Metric','Year','Month','Prevalence']:
        df[c]=pd.to_numeric(df[c],errors='coerce')
    for c in ['Farm','Variable','Customer']:
        if df[c].fillna('').astype(str).str.strip().eq('').any(): raise ValueError(f'{c} contains blank values.')
    if df['Year'].isna().any() or df['Month'].isna().any(): raise ValueError('Year or Month contains invalid values.')
    idx=['Farm','House','Bird','Age (days)','Variable','Variable Score','Metric','Description','System','Complex','Customer','State','Year','Month','Season','Session date','Reference','Prevalence']
    data=[]
    for row in df[idx].itertuples(index=False,name=None):
        r=list(row);r[4]=str(r[4]).strip();r[8]=final_system(r[4],str(r[8]) if pd.notna(r[8]) else '')
        r=[None if pd.isna(x) else x.item() if hasattr(x,'item') else x for x in r];data.append(r)
    # Add derived TMLS per bird/reference from gross scores only.
    parts={};base={}
    for r in data:
        if r[4] in {'gAC','gMX','gTN'}:
            k=(r[0],r[1],r[2],r[15],r[16]);parts.setdefault(k,{})[r[4]]=r[5];base[k]=r
    for k,q in parts.items():
        if all(isinstance(q.get(x),(int,float)) for x in ('gAC','gMX','gTN')):
            r=base[k].copy();r[4]='TMLS';r[5]=q['gAC']+q['gMX']+q['gTN'];r[6]=None;r[7]='Total Mean Lesion Score';r[8]='Coccidiosis';r[17]=None;data.append(r)
    meta={}
    for code in sorted({r[4] for r in data}):
        rr=next(r for r in data if r[4]==code);meta[code]={'system':rr[8],'desc':NAME_MAP.get(code,code)}
    systems=[x for x in SYSTEM_ORDER if x=='All' or any(r[8]==x for r in data)]
    return data,meta,systems,len(df)

def github_ready(): return GITHUB_OWNER!='CHANGE_ME' and GITHUB_REPO!='CHANGE_ME'
def token():
    try:return st.secrets['GITHUB_TOKEN']
    except Exception:return None
def api_url(): return f'https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/contents/{GITHUB_DATA_PATH}'
def headers():
    h={'Accept':'application/vnd.github+json','X-GitHub-Api-Version':API_VERSION}
    if token():h['Authorization']=f'Bearer {token()}'
    return h

def github_get():
    if not github_ready(): return BUNDLED,'Bundled Excel'
    r=requests.get(api_url(),headers=headers(),params={'ref':GITHUB_BRANCH},timeout=30)
    r.raise_for_status();obj=r.json();return base64.b64decode(obj['content']),'GitHub Excel'

def github_put(blob):
    if not github_ready(): raise ValueError('Set GITHUB_OWNER and GITHUB_REPO in app.py first.')
    if not token(): raise ValueError('GITHUB_TOKEN is missing from Streamlit Secrets.')
    g=requests.get(api_url(),headers=headers(),params={'ref':GITHUB_BRANCH},timeout=30)
    sha=g.json().get('sha') if g.ok else None
    payload={'message':'Update dashboard Raw Data Excel','content':base64.b64encode(blob).decode(),'branch':GITHUB_BRANCH}
    if sha:payload['sha']=sha
    r=requests.put(api_url(),headers=headers(),json=payload,timeout=60)
    if not r.ok:raise ValueError(f'GitHub update failed: {r.status_code} {r.text[:300]}')
    return r.json()['commit']['html_url']

st.sidebar.header('Admin Excel update')
pw=st.sidebar.text_input('Password',type='password')
if pw and pw!=ADMIN_PASSWORD:st.sidebar.error('Incorrect password')
if pw==ADMIN_PASSWORD:
    upload=st.sidebar.file_uploader("Excel workbook with sheet 'Raw Data'",type=['xlsx'])
    if upload:
        try:
            data,meta,systems,n=compact_payload(upload.getvalue());st.sidebar.success(f'Validated {n:,} Raw Data rows')
            if st.sidebar.button('Commit Excel to GitHub',type='primary'):
                commit=github_put(upload.getvalue());st.cache_data.clear();st.sidebar.success('Committed to GitHub');st.sidebar.markdown(f'[View commit]({commit})');st.rerun()
        except Exception as e:st.sidebar.error(str(e))

@st.cache_data(ttl=60,show_spinner=False)
def build_dashboard():
    blob,source=github_get();data,meta,systems,n=compact_payload(blob)
    html=TEMPLATE.replace('__DATA__',json.dumps(data,ensure_ascii=False,separators=(',',':'))).replace('__META__',json.dumps(meta,ensure_ascii=False,separators=(',',':'))).replace('__SYSTEMS__',json.dumps(systems,ensure_ascii=False))
    return html,source,n
try:html,source,n=build_dashboard()
except Exception as e:st.error(f'Dashboard loading failed: {e}');st.stop()
st.caption(f'Data source: {source} • {n:,} Raw Data rows')
components.html(html,height=1350,scrolling=True)
st.download_button('Download current dashboard as HTML',html,file_name='Annual_Report_Dashboard.html',mime='text/html')
