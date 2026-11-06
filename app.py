import io, os, re, html, uuid
from datetime import datetime
import pandas as pd
import streamlit as st
import database as db
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.units import inch
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
try:
    from openai import OpenAI
except ImportError:
    OpenAI=None

st.set_page_config(page_title='Interactive Q&A and Reporting',layout='wide')
db.init_db()
st.title('Interactive Q&A with Uploaded Files and Reporting')
st.caption('Create projects, save every answer version, choose the version for reporting, and resume later.')

def ai_client():
    key=os.getenv('OPENAI_API_KEY','').strip()
    try:
        if st.secrets.get('OPENAI_API_KEY'): key=st.secrets['OPENAI_API_KEY'].strip()
    except Exception: pass
    return OpenAI(api_key=key) if key and OpenAI else None

def ask_ai(system,prompt,fallback):
    c=ai_client()
    if not c:return fallback
    try:
        r=c.chat.completions.create(model='gpt-4o-mini',messages=[{'role':'system','content':system},{'role':'user','content':prompt}],temperature=0.4)
        return r.choices[0].message.content.strip()
    except Exception as e:
        st.error(f'AI request failed: {e}'); return None

def generate_answer(question,info,note=''):
    prompt=f'''Company information:\n{info}\n\nQuestion:\n{question}\n\nAdditional instructions:\n{note or 'None'}\n\nGive clear, evidence-grounded business analysis. Do not invent facts.'''
    fallback=f'[PREVIEW MODE — SAMPLE RESPONSE]\n\nQuestion: {question}\n\nCompany information supplied: {str(info)[:1000]}\n\nAdditional focus: {note or "None"}\n\nThis is simulated preview content, not a real AI-generated answer.'
    return ask_ai('You are a careful market and business analyst.',prompt,fallback)

def read_csv(upload):
    upload.seek(0); return pd.read_csv(upload).fillna('')

def col(df,names):
    lookup={str(x).strip().lower():x for x in df.columns}
    return next((lookup[x.lower()] for x in names if x.lower() in lookup),None)

def load_project(pid):
    p=db.get_project(pid)
    if not p: st.error('Saved project not found.'); return
    st.session_state.active_project_id=pid; st.session_state.project_name=p['name']
    st.session_state.company_df=p['company']; st.session_state.questions_df=p['questions']; st.session_state.template_df=p['template']
    st.session_state.report_rows=[]; st.session_state.report_generated=False; st.session_state.current_report_id=None
    # Restore all saved Q&A rows and their latest/selected version metadata.
    refresh_answers()

def refresh_answers():
    pid=st.session_state.active_project_id
    history=[]
    if pid:
        for a in db.get_answers(pid):
            vs=db.versions(a['answer_id']); chosen=next((v for v in vs if v['version_id']==a['selected_version_id']),None)
            history.append({'ID':a['answer_id'],'CompanyURL':a['company_url'],'CompanyInfo':a['company_info'],'Question':a['question'],'Answer':a['latest_text'],'LatestVersion':a['latest_number'],'SelectedVersionID':a['selected_version_id'],'SelectedVersionNumber':chosen['version_number'] if chosen else None})
    st.session_state.qna_history=history

# Session defaults
for k,v in {'active_project_id':None,'project_name':'','company_df':None,'questions_df':None,'template_df':None,'qna_history':[],'report_rows':[],'report_generated':False,'current_report_id':None}.items():
    if k not in st.session_state: st.session_state[k]=v

# 1. Project selection and creation
st.subheader('1. Resume an existing project or create a new one')
projects=db.list_projects()
if projects:
    opts={f"{p['name']} — updated {p['updated_at'][:16]}":p['project_id'] for p in projects}
    selected=st.selectbox('Saved projects',['Select a project...']+list(opts),key='project_to_resume')
    if st.button('Resume selected project',disabled=selected=='Select a project...'):
        load_project(opts[selected]); st.rerun()
with st.expander('Create a new project',expanded=not bool(st.session_state.active_project_id)):
    pname=st.text_input('Project name',key='new_project_name')
    c1,c2,c3=st.columns(3)
    with c1: company_file=st.file_uploader('Company information CSV',type=['csv'],key='new_company_csv')
    with c2: questions_file=st.file_uploader('Questions CSV',type=['csv'],key='new_questions_csv')
    with c3: template_file=st.file_uploader('Report template CSV (optional)',type=['csv'],key='new_template_csv')
    if st.button('Create and save project'):
        if not pname.strip() or not company_file or not questions_file: st.warning('Enter a project name and upload company and questions CSV files.')
        else:
            try:
                cdf=read_csv(company_file); qdf=read_csv(questions_file); tdf=read_csv(template_file) if template_file else None
                pid=db.create_project(pname.strip(),cdf,qdf,tdf)
                st.session_state.active_project_id=pid; st.session_state.project_name=pname.strip()
                st.session_state.company_df=cdf; st.session_state.questions_df=qdf; st.session_state.template_df=tdf
                st.session_state.qna_history=[]; st.session_state.report_rows=[]; st.session_state.report_generated=False
                st.success('Project created and saved.'); st.rerun()
            except Exception as e: st.error(f'Could not create project: {e}')

if not st.session_state.active_project_id:
    st.info('Create or resume a project to continue.'); st.stop()
st.markdown(f"**Active project:** {st.session_state.project_name}")
if st.button('Close current project'):
    for k,v in {'active_project_id':None,'project_name':'','company_df':None,'questions_df':None,'template_df':None,'qna_history':[],'report_rows':[],'report_generated':False,'current_report_id':None}.items():st.session_state[k]=v
    st.rerun()
with st.expander('View saved input files'):
    st.write('Company information'); st.dataframe(st.session_state.company_df,use_container_width=True)
    st.write('Questions'); st.dataframe(st.session_state.questions_df,use_container_width=True)
    if st.session_state.template_df is not None: st.write('Report template'); st.dataframe(st.session_state.template_df,use_container_width=True)

# 2. Q&A generation
st.subheader('2. Generate Q&A responses')
existing_answers = db.get_answers(st.session_state.active_project_id)
if existing_answers:
    st.info(f"This project already has {len(existing_answers)} saved Q&A responses. You can continue editing, regenerating individual answers, reviewing versions, approving selections, and generating reports below. Bulk Q&A generation is disabled here to prevent accidental duplicate answers.")
else:
    generate_qna_clicked = st.button('Generate Q&A Responses',type='primary')

if not existing_answers and generate_qna_clicked:
    companies=st.session_state.company_df; questions=st.session_state.questions_df
    urlc=col(companies,['company URL','company URL.','CompanyURL','URL'])
    infoc=col(companies,['Information about the company.','Information about the company','Company Information'])
    qc=col(questions,['Questions.','Questions','Question'])
    if not all([urlc,infoc,qc]): st.error('Required columns not found. Check company URL, company information, and Questions columns.')
    else:
        total=max(len(companies)*len(questions),1); progress=st.progress(0); n=0
        for _,qr in questions.iterrows():
            question=str(qr[qc]).strip()
            if not question: continue
            for _,cr in companies.iterrows():
                info=str(cr[infoc]); answer=generate_answer(question,info)
                if answer: db.create_answer(st.session_state.active_project_id,str(uuid.uuid4()),str(cr[urlc]),info,question,answer)
                n+=1; progress.progress(min(n/total,1.0))
        refresh_answers(); st.session_state.report_rows=[]; st.session_state.report_generated=False
        st.success('Generated Q&A responses and saved original answers to SQLite.'); st.rerun()

# 3. Version history / selection / regeneration / manual edits
st.subheader('3. Review answers and compare versions')
refresh_answers()
if st.session_state.qna_history:
    st.info('Saved Q&A responses have been restored for this project. Continue editing or regenerating individual answers; every change is saved as a new version.')
    for item in st.session_state.qna_history:
        aid=item['ID']; allv=db.versions(aid)
        with st.container(border=True):
            st.markdown(f"**Company:** {item['CompanyURL']}  \n**Question:** {item['Question']}")
            st.caption(f"Latest: V{item['LatestVersion']} | Selected for report: V{item['SelectedVersionNumber'] or 'Not selected'}")
            with st.expander(f'Preview all {len(allv)} versions'):
                for v in allv:
                    st.markdown(f"**Version {v['version_number']}** · {v['created_at']} · {v['status']}")
                    st.write(v['answer_text'])
                    if v['note']: st.caption('Change: '+v['note'])
                    if v['version_id']==item['SelectedVersionID']: st.success('Currently selected for report')
                    st.divider()
                labels={f"Version {v['version_number']} — {v['status']} — {v['created_at']}":v['version_id'] for v in allv}
                current=next((x for x,y in labels.items() if y==item['SelectedVersionID']),list(labels)[-1])
                choice=st.selectbox('Choose version for report',list(labels),index=list(labels).index(current),key=f'version_choice_{aid}')
                if st.button('Use this version for report',key=f'use_version_{aid}'):
                    db.choose(aid,labels[choice]); st.success('Report version selection saved.'); st.rerun()
            edited=st.text_area('Latest answer (editable)',item['Answer'],height=150,key=f'edit_{aid}')
            if st.button('Save manual edit as new version',key=f'save_edit_{aid}'):
                if edited.strip() and edited!=item['Answer']:
                    db.add_version(aid,edited,'Manual edit'); st.success('Manual edit saved as a new version.'); st.rerun()
                else: st.info('No changes to save.')
            note=st.text_input('Instructions for regeneration',key=f'note_{aid}',placeholder='e.g. Be more concise; focus on Canadian competitors.')
            if st.button('Regenerate this answer',key=f'regenerate_{aid}'):
                with st.spinner('Regenerating selected answer...'): revised=generate_answer(item['Question'],item['CompanyInfo'],note)
                if revised:
                    db.add_version(aid,revised,'Regenerated: '+(note or 'No additional instructions'))
                    st.session_state.report_rows=[]; st.session_state.report_generated=False
                    st.success('New version saved.'); st.rerun()
    export=[]
    for item in st.session_state.qna_history:
        for v in db.versions(item['ID']):
            export.append({'Answer ID':item['ID'],'Company URL':item['CompanyURL'],'Question':item['Question'],'Version':v['version_number'],'Answer':v['answer_text'],'Change note':v['note'],'Status':v['status'],'Created at':v['created_at'],'Selected for report':v['version_id']==item['SelectedVersionID']})
    st.download_button('Download complete Q&A version history CSV',pd.DataFrame(export).to_csv(index=False).encode(), 'QnA_Version_History.csv','text/csv')
else: st.caption('No Q&A responses saved for this project yet.')

# Approval controls for the selected version
if st.session_state.qna_history:
    with st.expander('Approve selected versions before report generation'):
        for item in st.session_state.qna_history:
            chosen=next((v for v in db.versions(item['ID']) if v['version_id']==item['SelectedVersionID']),None)
            if chosen:
                st.write(f"{item['CompanyURL']} — {item['Question']} — V{chosen['version_number']}")
                if chosen['status']=='Approved': st.success('Approved')
                elif st.button('Approve selected version',key=f"approve_{chosen['version_id']}"):
                    db.approve(chosen['version_id']); st.success('Approved.'); st.rerun()

# 4. Report generation from selected approved versions
st.subheader('4. Generate report from selected versions')
def make_report():
    template=st.session_state.template_df
    if template is None:return None,'No report template is saved in this project.'
    qc=col(template,['Report Questions','Report Question','Question']); nc=col(template,['Report Question number','Report Question Number','Question number','Number'])
    if qc is None:return None,"Template needs a 'Report Questions' column."
    selected=db.selected_answers(st.session_state.active_project_id)
    if not selected:return None,'Generate Q&A responses first.'
    pending=[x for x in selected if x['status']!='Approved']
    if pending:return None,f'{len(pending)} selected version(s) still need approval. Approve each selected version above.'
    context='\n\n'.join(f"Company: {x['company_url']}\nQuestion: {x['question']}\nSelected approved answer (V{x['version_number']}): {x['answer_text']}" for x in selected)
    rows=[]
    for _,r in template.iterrows():
        q=str(r[qc]).strip()
        if not q:continue
        prompt=f'Report section/question:\n{q}\n\nUse only these selected, approved Q&A responses:\n{context}\n\nWrite a structured answer. Do not add unsupported facts.'
        fallback=f'[PREVIEW MODE — SAMPLE REPORT SECTION]\n\nReport question: {q}\n\nThis section would be drafted from the selected, approved Q&A.'
        answer=ask_ai('You are a business research analyst writing a report from supplied evidence.',prompt,fallback)
        if answer is None:return None,'Report generation stopped due to an AI error.'
        row={'Report Question':q,'Report Answer':answer}
        if nc:row['Report Question number']=r[nc]
        rows.append(row)
    sources=[{'answer_id':x['answer_id'],'version_id':x['version_id'],'version_number':x['version_number']} for x in selected]
    return {'rows':rows,'sources':sources},None
if st.button('Generate Report',type='primary'):
    with st.spinner('Generating report sections from selected versions...'):result,error=make_report()
    if error:st.error(error)
    else:
        st.session_state.report_rows=result['rows']; st.session_state.report_generated=True
        st.session_state.current_report_id=db.save_report(st.session_state.active_project_id,result['rows'],result['sources'])
        st.success('Report generated and saved.'); st.rerun()

# Load an earlier report without re-running Q&A or report generation
saved_reports=db.reports(st.session_state.active_project_id)
if saved_reports:
    with st.expander('Resume a previously generated report'):
        ropts={f"{r['created_at']} — {r['report_id'][:8]}":r['report_id'] for r in saved_reports}
        rchoice=st.selectbox('Saved reports', ['Choose a report...']+list(ropts),key='saved_report_choice')
        if st.button('Load saved report',disabled=rchoice=='Choose a report...'):
            old=db.get_report(ropts[rchoice]); st.session_state.report_rows=old['rows']; st.session_state.report_generated=True; st.session_state.current_report_id=ropts[rchoice]; st.rerun()

# 5. Original export formats preserved

def pdf_bytes(rows):
    out=io.BytesIO(); doc=SimpleDocTemplate(out,pagesize=letter,rightMargin=.65*inch,leftMargin=.65*inch,topMargin=.65*inch,bottomMargin=.65*inch)
    styles=getSampleStyleSheet(); heading=ParagraphStyle('Section',parent=styles['Heading2'],textColor=colors.HexColor('#1F4E79'),spaceBefore=10,spaceAfter=7); body=ParagraphStyle('Body',parent=styles['BodyText'],fontSize=9.5,leading=14,spaceAfter=7)
    story=[Paragraph('Report Analysis',styles['Title']),Paragraph('Generated '+datetime.now().strftime('%B %d, %Y'),styles['Italic']),Spacer(1,14)]
    for row in rows:
        story.append(Paragraph(html.escape(str(row['Report Question'])),heading))
        for line in str(row['Report Answer']).splitlines():
            line=line.strip()
            if not line:story.append(Spacer(1,5))
            elif line.startswith('#'):story.append(Paragraph('<b>'+html.escape(line.lstrip('# ').strip())+'</b>',heading))
            elif line.startswith(('- ','* ')):story.append(Paragraph('&#8226; '+html.escape(line[2:]),body))
            else:story.append(Paragraph(html.escape(line).replace('**',''),body))
        story.append(Spacer(1,9))
    doc.build(story); return out.getvalue()

def pptx_bytes(rows):
    prs=Presentation(); first=prs.slides.add_slide(prs.slide_layouts[0]); first.shapes.title.text='Report Analysis'; first.placeholders[1].text='Generated '+datetime.now().strftime('%B %d, %Y')
    for row in rows:
        paras=[p.strip() for p in re.split(r'\n+',str(row['Report Answer'])) if p.strip()]; chunks=[]; current=[]; count=0
        for para in paras:
            words=para.split()
            while len(words)>85:
                if current:chunks.append(current); current=[]; count=0
                chunks.append([' '.join(words[:85])]); words=words[85:]
            if count+len(words)>85 and current:chunks.append(current); current=[]; count=0
            if words:current.append(' '.join(words)); count+=len(words)
        if current:chunks.append(current)
        if not chunks:chunks=[['']]
        for j,chunk in enumerate(chunks,1):
            slide=prs.slides.add_slide(prs.slide_layouts[5]); tb=slide.shapes.add_textbox(Inches(.55),Inches(.35),Inches(12.2),Inches(.8)); tf=tb.text_frame; tf.text=str(row['Report Question'])[:170]+(f' (continued {j})' if j>1 else '')
            for run in tf.paragraphs[0].runs:run.font.size=Pt(22); run.font.bold=True; run.font.color.rgb=RGBColor(31,78,121)
            box=slide.shapes.add_textbox(Inches(.7),Inches(1.35),Inches(11.9),Inches(5.7)); frame=box.text_frame; frame.word_wrap=True
            for k,txt in enumerate(chunk):
                p=frame.paragraphs[0] if k==0 else frame.add_paragraph(); p.text=re.sub(r'[*#]','',txt); p.font.size=Pt(16); p.space_after=Pt(9)
    out=io.BytesIO(); prs.save(out); return out.getvalue()

if st.session_state.report_generated and st.session_state.report_rows:
    st.subheader('5. Review and edit report sections')
    edited=[]
    for i,row in enumerate(st.session_state.report_rows):
        with st.expander(f"Section {i+1}: {row['Report Question']}",expanded=(i==0)):
            txt=st.text_area('Report answer (editable)',row['Report Answer'],height=170,key=f"report_{st.session_state.current_report_id}_{i}")
            edited.append({**row,'Report Answer':txt})
    st.session_state.report_rows=edited
    st.subheader('6. Download report')
    a,b,c=st.columns(3)
    with a:st.download_button('Download Report CSV',pd.DataFrame(edited).to_csv(index=False).encode(),'ReportGenerated.csv','text/csv')
    with b:st.download_button('Download Report PDF',pdf_bytes(edited),'ReportGenerated.pdf','application/pdf')
    with c:st.download_button('Download Report PowerPoint',pptx_bytes(edited),'ReportGenerated.pptx','application/vnd.openxmlformats-officedocument.presentationml.presentation')

if ai_client() is None:st.info('Preview mode: no OPENAI_API_KEY is configured. AI outputs are simulated; database and export workflows remain testable.')
else:st.success('OpenAI API key detected. AI generation is enabled.')
