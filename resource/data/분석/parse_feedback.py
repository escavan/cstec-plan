"""
설문후기 87개 엑셀 통합 파서.
실행: python parse_feedback.py (출력: consolidated_objective.csv, consolidated_freetext.csv)
"""
import pandas as pd
import os, glob, warnings
warnings.filterwarnings('ignore')

FOLDER = r'c:\Users\kayang\claudecode\cstec-plan\reference\2025-2026 설문후기'
OUT = os.path.dirname(os.path.abspath(__file__))

SLOT_RULES = {
    'diff_free': lambda q: '난이도' in q and ('조절' in q or '초점' in q),
    'difficulty': lambda q: '난이도' in q,
    'sat_free': lambda q: '만족도' in q and ('높거나' in q or '낮은' in q or '이유' in q),
    'satisfaction': lambda q: '만족도' in q and '강의' not in q and '과목' not in q,
    'util_free': lambda q: '활용도' in q and ('낮다면' in q or '이유' in q),
    'utility': lambda q: '활용도' in q,
    'wishlist': lambda q: ('받고 싶' in q) or ('교육받고 싶' in q) or ('수강하고 싶' in q),
    'suggestion': lambda q: '건의' in q,
    'needs': lambda q: '꼭 필요한' in q,
    'duration': lambda q: '시간' in q and '적절' in q,
    'diff_inst': lambda q: '차별성' in q,
    'instructor': lambda q: ('과목' in q or '강사' in q),
    'goal': lambda q: '목표' in q and '달성' in q,
}

def slot_of(qt):
    if not qt: return 'other'
    for slot, rule in SLOT_RULES.items():
        if rule(qt): return slot
    return 'other'

def parse_file(path):
    df = pd.read_excel(path, sheet_name=0, header=None)
    meta = {}
    for i in range(min(4, len(df))):
        if pd.notna(df.iloc[i,0]) and pd.notna(df.iloc[i,1]):
            meta[str(df.iloc[i,0])] = str(df.iloc[i,1])
    course = meta.get('과정명', os.path.basename(path).replace('설문 결과(','').replace('.xlsx',''))
    header_row = None
    for i in range(len(df)):
        if str(df.iloc[i,0]).strip() == '번호':
            header_row = i; break
    if header_row is None:
        return None
    questions, current_q = [], None
    for i in range(header_row+1, len(df)):
        v0,v1,v2,v3,v4 = df.iloc[i,0], df.iloc[i,1], df.iloc[i,2], df.iloc[i,3], df.iloc[i,4]
        if pd.notna(v0) and isinstance(v0,(int,float)) and v0==int(v0):
            current_q = {'q_num':int(v0),'q_type':str(v1) if pd.notna(v1) else '',
                         'q_text':str(v2) if pd.notna(v2) else '',
                         'respondents':int(v3) if pd.notna(v3) and isinstance(v3,(int,float)) else None,
                         'options':[], 'responses':[]}
            questions.append(current_q)
        elif current_q is not None and pd.isna(v0) and pd.notna(v1):
            if current_q['q_type']=='객관식':
                current_q['options'].append({'text':str(v2) if pd.notna(v2) else '',
                                              'count':int(v3) if pd.notna(v3) and isinstance(v3,(int,float)) else 0,
                                              'pct':float(v4) if pd.notna(v4) and isinstance(v4,(int,float)) else 0.0})
            elif current_q['q_type']=='주관식':
                t=str(v2) if pd.notna(v2) else ''
                if t and t.lower()!='nan': current_q['responses'].append(t)
    return {'course':course,'questions':questions,'file':os.path.basename(path)}

if __name__ == '__main__':
    files = sorted(glob.glob(os.path.join(FOLDER,'*.xlsx')))
    data = [parse_file(fn) for fn in files]
    data = [d for d in data if d]
    print(f'Parsed: {len(data)}/{len(files)}')
    obj_rows, ft_rows = [], []
    for d in data:
        c = d['course']
        for q in d['questions']:
            s = slot_of(q['q_text'])
            if q['q_type']=='객관식':
                for o in q['options']:
                    obj_rows.append({'course':c,'q_num':q['q_num'],'slot':s,'q_text':q['q_text'],
                                     'option':o['text'],'count':o['count'],'pct':o['pct'],'n':q['respondents']})
            elif q['q_type']=='주관식':
                for r in q['responses']:
                    ft_rows.append({'course':c,'q_num':q['q_num'],'slot':s,'q_text':q['q_text'],'response':r})
    pd.DataFrame(obj_rows).to_csv(os.path.join(OUT,'consolidated_objective.csv'), index=False, encoding='utf-8-sig')
    pd.DataFrame(ft_rows).to_csv(os.path.join(OUT,'consolidated_freetext.csv'), index=False, encoding='utf-8-sig')
    print(f'Saved {len(obj_rows)} obj / {len(ft_rows)} freetext')
