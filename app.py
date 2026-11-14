import streamlit as st
import pandas as pd
import re, os
from io import BytesIO
from zipfile import ZipFile, ZIP_DEFLATED
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font
from difflib import SequenceMatcher

st.set_page_config(page_title='MEP BOQ Automation', page_icon='ðŸ“Š', layout='wide')
st.title('MEP BOQ Automation System')
st.caption('System-wise separation | Multi-OEM comparison | Rate mapping')

SYSTEM_KEYWORDS = {
 'Fire Alarm System':['fire alarm','smoke detector','heat detector','mcp','manual call point','hooter','sounder','fire panel','beam detector','monitor module','control module','telephone jack','strobe'],
 'Public Address System':['public address','pa system','speaker','amplifier','microphone','paviro','voice evacuation','ceiling speaker'],
 'CCTV System':['cctv','camera','nvr','dvr','video surveillance','ip camera'],
 'Access Control System':['access control','card reader','door controller','magnetic lock','em lock','access reader'],
 'BMS System':['bms','ddc','bacnet','modbus','field controller','bms panel'],
 'DALI Automation':['dali','dali gateway','dali driver','dimming','dali controller'],
 'KNX Automation':['knx','knx gateway','knx switch','knx sensor','knx actuator'],
 'Structured Cabling':['structured cabling','patch panel','cat6','cat 6','fiber optic','network rack','data point'],
 'Fire Fighting System':['sprinkler','fire hydrant','fire pump','landing valve','hose reel','fire extinguisher'],
 'Plumbing System':['plumbing','water supply','drainage','soil pipe','waste pipe','sanitary'],
 'HVAC System':['hvac','ahu','fcu','chiller','ducting','diffuser','grille']}
SYSTEMS=list(SYSTEM_KEYWORDS)+['Other ELV Systems','Unassigned']


def clean(v):
    if v is None or (isinstance(v,float) and pd.isna(v)): return ''
    return re.sub(r'\s+',' ',str(v)).strip()
def norm(v): return re.sub(r'[^a-z0-9]+',' ',clean(v).lower()).strip()
def safe(v): return re.sub(r'[^A-Za-z0-9_-]+','_',str(v))
def identify_system(d,c=''):
    t=(clean(d)+' '+clean(c)).lower(); found=[s for s,ks in SYSTEM_KEYWORDS.items() if any(k in t for k in ks)]
    return found[0] if len(found)==1 else 'Unassigned'
def header_map(ws,hr):
    return {f'{get_column_letter(c)} - {clean(ws.cell(hr,c).value)}':c for c in range(1,ws.max_column+1) if clean(ws.cell(hr,c).value)}
def auto_col(headers, patterns, exclude=()):
    # Prefer explicit phrase matches, then token matches.
    scored=[]
    for col,h in headers.items():
        x=norm(h)
        if any(norm(e) in x for e in exclude): continue
        score=max([len(norm(p)) for p in patterns if norm(p) and norm(p) in x] or [0])
        if score: scored.append((score,col))
    return max(scored)[1] if scored else None

def guess_header(rows):
    best=(0,0)
    for i,row in enumerate(rows[:25]):
        text=' '.join(norm(x) for x in row if clean(x))
        score=sum(k in text for k in ['description','item','supply','installation','rate','make','part code','qty','quantity','amount'])
        if score>best[1]: best=(i,score)
    return best[0] if best[1]>=1 else 0

def excel_oem(file_bytes, filename):
    wb=load_workbook(BytesIO(file_bytes),data_only=True)
    out=[]
    for ws in wb.worksheets:
        rows=list(ws.iter_rows(values_only=True))
        if not rows: continue
        hi=guess_header(rows); headers=[clean(x) or f'Column {j+1}' for j,x in enumerate(rows[hi])]
        # make duplicate headings unique
        seen={}; unique=[]
        for h in headers:
            seen[h]=seen.get(h,0)+1; unique.append(h if seen[h]==1 else f'{h}_{seen[h]}')
        df=pd.DataFrame(rows[hi+1:],columns=unique).dropna(how='all')
        if df.empty: continue
        hm={h:h for h in unique}
        desc=auto_col(hm,['description of item','item description','description','particulars','item details','material description','product description','item'])
        code=auto_col(hm,['part code','item code','product code','model no','model number','sku','catalogue no','catalog no'])
        sr=auto_col(hm,['supply rate','unit supply rate','basic rate','unit rate','rate','supply amount','supply'])
        ir=auto_col(hm,['installation rate','install rate','labour rate','installation amount','installation'])
        make=auto_col(hm,['make','brand','manufacturer','offered make'])
        qty=auto_col(hm,['quantity','qty'])
        if not desc: continue
        for ridx,row in df.iterrows():
            d=clean(row.get(desc))
            if not d or norm(d) in ['description','item description','total','grand total','sub total','subtotal']: continue
            supply=row.get(sr) if sr else None; install=row.get(ir) if ir else None
            # Exclude lines with no rate values; do not mistake qty for a rate.
            if not is_number(supply) and not is_number(install): continue
            out.append({'OEM File':filename,'OEM Sheet':ws.title,'OEM Source Row':int(ridx)+hi+2,'Description':d,'Part Code':clean(row.get(code)) if code else '', 'Supply Rate':num(supply),'Installation Rate':num(install),'Make':clean(row.get(make)) if make else '', 'OEM Qty':num(row.get(qty)) if qty else None})
    return out

def is_number(x):
    try: return x is not None and not pd.isna(x) and float(str(x).replace(',','').replace('â‚¹','').strip())>=0
    except: return False
def num(x):
    try: return float(str(x).replace(',','').replace('â‚¹','').strip()) if is_number(x) else None
    except: return None

def ocr_extract(data,filename):
    import pytesseract
    from PIL import Image
    ext=os.path.splitext(filename.lower())[1]; texts=[]
    if ext=='.pdf':
        import pdfplumber
        with pdfplumber.open(BytesIO(data)) as pdf:
            for page in pdf.pages:
                t=page.extract_text() or ''
                if t.strip(): texts.append(t)
                else:
                    from pdf2image import convert_from_bytes
                    imgs=convert_from_bytes(data,first_page=page.page_number,last_page=page.page_number,dpi=220)
                    texts.append(pytesseract.image_to_string(imgs[0]) if imgs else '')
    else:
        texts.append(pytesseract.image_to_string(Image.open(BytesIO(data)).convert('RGB')))
    return '\n'.join(texts)

def ocr_rows(text,filename):
    rows=[]
    for line_no,line in enumerate(text.splitlines(),1):
        line=clean(line)
        if len(line)<5: continue
        nums=re.findall(r'(?<![A-Za-z])(?:â‚¹\s*)?\d[\d,]*(?:\.\d{1,2})?',line)
        vals=[]
        for x in nums:
            try: vals.append(float(x.replace(',','').replace('â‚¹','').strip()))
            except: pass
        if not vals: continue
        desc=line
        for x in nums: desc=desc.replace(x,' ',1)
        desc=clean(desc.strip(' |:-'))
        if len(desc)<3: continue
        # OCR cannot reliably infer whether a single number is supply or install.
        rows.append({'OEM File':filename,'OEM Sheet':'OCR','OEM Source Row':line_no,'Description':desc,'Part Code':'','Supply Rate':vals[-2] if len(vals)>=2 else vals[-1],'Installation Rate':vals[-1] if len(vals)>=2 else None,'Make':'','OEM Qty':None,'OCR Source Text':line})
    return rows

def similarity(a,b):
    a,b=norm(a),norm(b)
    if not a or not b:return 0
    if a==b:return 1
    return SequenceMatcher(None,a,b).ratio()

def client_items(ws,hr,headers):
    dc=auto_col(headers,['description of item','item description','description','particulars','item details'])
    sc=auto_col(headers,['technical specifications','specifications','specification'])
    qc=auto_col(headers,['quantity','qty'])
    uc=auto_col(headers,['unit','uom'])
    cc=auto_col(headers,['part code','item code','product code','model no'])
    if not dc:return [],{}
    items=[]
    for r in range(hr+1,ws.max_row+1):
        d=clean(ws.cell(r,headers[dc]).value)
        if not d or re.search(r'\b(total|summary|grand total)\b',d,re.I):continue
        q=ws.cell(r,headers[qc]).value if qc else None
        code=clean(ws.cell(r,headers[cc]).value) if cc else ''
        serial=clean(ws.cell(r,1).value)
        if q is None and not code and not re.match(r'^\d+(\.\d+)?$',serial):continue
        items.append({'Client Row':r,'Client Description':d,'Client Specification':clean(ws.cell(r,headers[sc]).value) if sc else '', 'Client Code':code,'Client Qty':q,'Client Unit':clean(ws.cell(r,headers[uc]).value) if uc else ''})
    return items,{'desc':headers[dc],'spec':headers.get(sc) if sc else None,'qty':headers.get(qc) if qc else None,'unit':headers.get(uc) if uc else None,'code':headers.get(cc) if cc else None}

def compare(client,oem):
    odf=pd.DataFrame(oem); used=set(); results=[]
    for ci in client:
        best=None; score=0
        for i,o in odf.iterrows():
            if i in used:continue
            ccode=norm(ci['Client Code']); ocode=norm(o.get('Part Code',''))
            s=1.0 if ccode and ocode and ccode==ocode else similarity(ci['Client Description']+' '+ci['Client Specification'],o.get('Description',''))
            if s>score:best=i;score=s
        if best is not None and score>=.50:
            used.add(best); o=odf.iloc[best]
            status='Matched' if score>=.84 else 'Review'
            results.append({**ci,'OEM File':o.get('OEM File'),'OEM Description':o.get('Description'),'Match Score %':round(score*100,1),'Status':status,'Supply Rate':o.get('Supply Rate'),'Installation Rate':o.get('Installation Rate'),'Make':o.get('Make'),'Approved':score>=.84})
        else:
            results.append({**ci,'OEM File':'','OEM Description':'','Match Score %':0,'Status':'Missing in OEM','Supply Rate':None,'Installation Rate':None,'Make':'','Approved':False})
    extras=[]
    for i,o in odf.iterrows():
        if i not in used: extras.append({**o.to_dict(),'Status':'Extra in OEM'})
    return pd.DataFrame(results),pd.DataFrame(extras)

def style_sheet(ws):
    for c in ws[1]: c.font=Font(bold=True,color='FFFFFF'); c.fill=PatternFill('solid',fgColor='1F4E78')
    for row in ws.iter_rows(min_row=2):
        status=clean(row[7].value) if len(row)>7 else ''
        color='E2F0D9' if status=='Matched' else ('FFF2CC' if status=='Review' else 'FCE4D6')
        for c in row:c.fill=PatternFill('solid',fgColor=color)


tab_split,tab_rates=st.tabs(['1. System-wise BOQ Splitter','2. Multi-OEM Rate Mapping'])
with tab_split:
    st.subheader('Separate client BOQ into system-wise Excel files')
    up=st.file_uploader('Upload client Excel BOQ',type=['xlsx'],key='split_up')
    if up:
        raw=up.getvalue(); wb=load_workbook(BytesIO(raw),data_only=False); sn=st.selectbox('Select worksheet',wb.sheetnames,key='sp_sheet'); ws=wb[sn]
        hr=int(st.number_input('Header row',1,max(1,ws.max_row),2,key='sp_hr')); hm=header_map(ws,hr)
        if hm:
            dl=st.selectbox('Description column',list(hm),index=next((i for i,x in enumerate(hm) if 'description' in x.lower()),0),key='sp_desc')
            cl=st.selectbox('Part code column (optional)',['None']+list(hm),key='sp_code')
            if st.button('Identify systems',key='sp_ident'):
                rec=[]
                for r in range(hr+1,ws.max_row+1):
                    d=ws.cell(r,hm[dl]).value; c=ws.cell(r,hm[cl]).value if cl!='None' else ''
                    if clean(d): rec.append({'Excel Row':r,'Description':clean(d),'Suggested System':identify_system(d,c),'Final System':identify_system(d,c)})
                st.session_state['split_rec']=pd.DataFrame(rec);st.session_state['split_raw']=raw;st.session_state['split_sheet']=sn;st.session_state['split_hr']=hr
            if 'split_rec' in st.session_state:
                ed=st.data_editor(st.session_state['split_rec'],use_container_width=True,hide_index=True,disabled=['Excel Row','Description','Suggested System'],column_config={'Final System':st.column_config.SelectboxColumn(options=SYSTEMS,required=True)},key='sp_editor');st.session_state['split_rec']=ed
                if st.button('Generate system-wise Excel ZIP',key='sp_gen'):
                    z=BytesIO()
                    with ZipFile(z,'w',ZIP_DEFLATED) as zf:
                        for sys in ed['Final System'].dropna().unique():
                            if sys=='Unassigned':continue
                            selected=set(ed.loc[ed['Final System']==sys,'Excel Row'].astype(int)); ow=load_workbook(BytesIO(st.session_state['split_raw'])); sw=ow[st.session_state['split_sheet']]
                            for rr in range(sw.max_row,st.session_state['split_hr'],-1):
                                if rr not in selected:sw.delete_rows(rr,1)
                            b=BytesIO();ow.save(b);zf.writestr(safe(sys)+'_BOQ.xlsx',b.getvalue())
                    st.download_button('Download system-wise BOQs ZIP',z.getvalue(),'MEP_System_Wise_BOQs.zip','application/zip',type='primary')

with tab_rates:
    st.subheader('Automatic multi-OEM comparison and client rate mapping')
    st.info('Upload one client blank BOQ and multiple OEM quotations. The app auto-detects likely columns; review extracted values and approve matches before filling rates.')
    cf=st.file_uploader('A. Client blank BOQ (.xlsx)',type=['xlsx'],key='multi_client')
    ofs=st.file_uploader('B. Upload multiple OEM files (Excel / PDF / JPG / PNG)',type=['xlsx','xls','pdf','jpg','jpeg','png'],accept_multiple_files=True,key='multi_oems')
    if cf:
        cbytes=cf.getvalue(); cwb=load_workbook(BytesIO(cbytes),data_only=False); csh=st.selectbox('Client worksheet',cwb.sheetnames,key='multi_csh'); cws=cwb[csh]
        # auto-detect client header among first 20 rows
        candidates=[]
        for rr in range(1,min(cws.max_row,20)+1):
            hm0=header_map(cws,rr); score=sum(any(k in norm(x) for k in ['description','qty','quantity','unit','rate']) for x in hm0)
            candidates.append((score,rr,hm0))
        _,chr0,chm=max(candidates,key=lambda x:x[0]) if candidates else (0,1,{})
        chr1=int(st.number_input('Client header row (auto-detected; change only if needed)',1,max(1,cws.max_row),int(chr0 or 1),key='multi_chr'))
        chm=header_map(cws,chr1)
        items,ccols=client_items(cws,chr1,chm)
        st.caption(f'Auto-detected client item rows: {len(items)} | Description / Qty / Unit / Specification columns detected automatically where available.')
        if not items: st.warning('Client items not detected. Please check the selected worksheet and header row.')
        if ofs and items:
            all_oem=[]; errors=[]; ocr_previews=[]
            for f in ofs:
                data=f.getvalue(); ext=os.path.splitext(f.name.lower())[1]
                try:
                    if ext in ['.xlsx','.xls']:
                        rows=excel_oem(data,f.name)
                    else:
                        txt=ocr_extract(data,f.name); ocr_previews.append((f.name,txt)); rows=ocr_rows(txt,f.name)
                    all_oem.extend(rows)
                    if not rows: errors.append(f'{f.name}: no rate-bearing item rows auto-detected')
                except Exception as e: errors.append(f'{f.name}: {e}')
            st.caption(f'OEM files uploaded: {len(ofs)} | Extracted item rows: {len(all_oem)}')
            for e in errors: st.warning(e)
            for fn,txt in ocr_previews:
                with st.expander(f'OCR text preview â€” {fn}'):
                    st.text_area('Extracted text',txt,height=160,key='ocr_'+safe(fn))
            if all_oem:
                oemdf=pd.DataFrame(all_oem)
                st.markdown('**Review and correct automatically extracted OEM rows**')
                oemdf=st.data_editor(oemdf,num_rows='dynamic',use_container_width=True,key='multi_oem_editor')
                if st.button('Compare all OEMs with Client BOQ',type='primary',key='multi_compare'):
                    # Compare per OEM separately so duplicate products across suppliers are not treated as duplicates.
                    comps=[]; extras=[]
                    for fname in oemdf['OEM File'].dropna().unique():
                        part=oemdf[oemdf['OEM File']==fname].reset_index(drop=True)
                        comp,extra=compare(items,part.to_dict('records'))
                        comp.insert(0,'OEM Supplier File',fname); comps.append(comp)
                        if not extra.empty: extra.insert(0,'OEM Supplier File',fname); extras.append(extra)
                    combined=pd.concat(comps,ignore_index=True) if comps else pd.DataFrame()
                    extra_df=pd.concat(extras,ignore_index=True) if extras else pd.DataFrame()
                    st.session_state['multi_comp']=combined;st.session_state['multi_extra']=extra_df;st.session_state['multi_client_bytes']=cbytes;st.session_state['multi_client_sheet']=csh;st.session_state['multi_client_header']=chr1;st.session_state['multi_client_cols']=ccols
            if 'multi_comp' in st.session_state:
                st.markdown('### OEM-wise matching results')
                comp=st.data_editor(st.session_state['multi_comp'],use_container_width=True,hide_index=True,disabled=['OEM Supplier File','Client Row','Client Description','Client Qty','Client Unit','OEM Description','Match Score %','Status'],key='multi_match_editor',column_config={'Approved':st.column_config.CheckboxColumn('Approve this OEM rate'),'Supply Rate':st.column_config.NumberColumn(format='%.2f'),'Installation Rate':st.column_config.NumberColumn(format='%.2f')})
                st.session_state['multi_comp']=comp
                extra=st.session_state['multi_extra']
                st.markdown('### Extra items in OEM files')
                if extra.empty:st.success('No extra OEM-only items detected.')
                else:st.dataframe(extra,use_container_width=True,hide_index=True)
                missing=comp[comp.Status=='Missing in OEM']; reviews=comp[comp.Status=='Review']
                a,b,c,d=st.columns(4);a.metric('OEM comparisons',comp['OEM Supplier File'].nunique());b.metric('Missing in OEM',len(missing));c.metric('Extra in OEM',len(extra));d.metric('Needs review',len(reviews))
                st.download_button('Download comparison report (CSV)',comp.to_csv(index=False).encode('utf-8-sig'),'Multi_OEM_Comparison.csv','text/csv')
                # choose one OEM for final rate fill, preserving the comparison of all suppliers
                supplier=st.selectbox('Select OEM whose approved rates should be applied',comp['OEM Supplier File'].dropna().unique(),key='final_supplier')
                if st.button('Apply approved rates and create final client BOQ',type='primary',key='multi_apply'):
                    out=load_workbook(BytesIO(st.session_state['multi_client_bytes'])); ws=out[st.session_state['multi_client_sheet']]; hdr=st.session_state['multi_client_header']; hm=header_map(ws,hdr); cols=st.session_state['multi_client_cols']
                    supply_col=auto_col(hm,['supply rate','unit supply rate','basic rate','supply amount'])
                    install_col=auto_col(hm,['installation rate','install rate','installation amount'])
                    make_col=auto_col(hm,['make','brand','manufacturer'])
                    qty_col=cols.get('qty')
                    total_supply=auto_col(hm,['total supply amount','total supply amt','supply total'])
                    total_install=auto_col(hm,['total installation amount','total installation amt','installation total'])
                    chosen=comp[(comp['OEM Supplier File']==supplier)&(comp['Approved'].fillna(False))]
                    if not supply_col and not install_col:st.error('Client BOQ supply/installation rate columns could not be detected. Add those column headings to the client template and retry.')
                    else:
                        for _,r in chosen.iterrows():
                            rr=int(r['Client Row'])
                            if supply_col and pd.notna(r.get('Supply Rate')):ws.cell(rr,supply_col).value=float(r['Supply Rate'])
                            if install_col and pd.notna(r.get('Installatio