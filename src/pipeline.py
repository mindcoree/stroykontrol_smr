"""Download, audit, aggregate, explore and fit two transparent baselines.

Run from any directory: python -m src.pipeline (from the repository root).
No synthetic data, row sampling or silently ignored CSV parse errors.
"""
from __future__ import annotations

import argparse
import hashlib
import importlib.metadata
import json
import platform
import re
import time
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import duckdb
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import requests
from sklearn.compose import ColumnTransformer
from sklearn.dummy import DummyClassifier
from sklearn.impute import SimpleImputer
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import (accuracy_score, average_precision_score, balanced_accuracy_score,
                            brier_score_loss, confusion_matrix, f1_score, precision_score,
                            recall_score, roc_auc_score)
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import OneHotEncoder, StandardScaler

ROOT = Path(__file__).resolve().parents[1]
RAW = ROOT / 'data/raw'
OUT = ROOT / 'data/processed'
REP = ROOT / 'reports'
FIG = REP / 'figures'
URL = 'https://data.cityofchicago.org/api/views/22u3-xenr/rows.csv?accessType=DOWNLOAD'
META_URL = 'https://data.cityofchicago.org/api/views/22u3-xenr.json'
CUTOFF = '2026-10-09'
CAT = ['department_bureau', 'inspection_category']
NUM = ['year', 'month_sin', 'month_cos', 'dayofweek', 'latitude', 'longitude']
FEATURES = CAT + NUM
ORANGE, DARK, GRAY, PAPER = '#E8591A', '#1E2227', '#5B636E', '#F2EFE8'


def save_json(path, value):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2, default=str), encoding='utf-8')


def sha256(path):
    h = hashlib.sha256()
    with path.open('rb') as f:
        for block in iter(lambda: f.read(4 * 1024 * 1024), b''):
            h.update(block)
    return h.hexdigest()


def download():
    RAW.mkdir(parents=True, exist_ok=True)
    path = RAW / 'building_violations.csv'
    if not path.exists():
        tmp = path.with_suffix('.csv.part')
        with requests.get(URL, stream=True, timeout=(30, 600)) as r:
            r.raise_for_status()
            with tmp.open('wb') as f:
                for block in r.iter_content(4 * 1024 * 1024):
                    f.write(block)
        tmp.replace(path)
    manifest = RAW / 'manifest.json'
    if manifest.exists():
        old = json.loads(manifest.read_text())
        if old['sha256'] != sha256(path):
            raise ValueError('CSV differs from its manifest; use a new snapshot directory.')
    else:
        r = requests.get(META_URL, timeout=60)
        r.raise_for_status()
        meta = r.json()
        save_json(RAW / 'source_metadata.json', meta)
        save_json(manifest, {
            'url': URL, 'dataset_id': '22u3-xenr',
            'retrieved_at': datetime.now(ZoneInfo('Asia/Almaty')).isoformat(),
            'rows_updated_at_unix': meta.get('rowsUpdatedAt'),
            'sha256': sha256(path), 'csv_bytes': path.stat().st_size,
            'analysis_cutoff': CUTOFF,
            'note': 'Live archive; local checksum identifies the analyzed snapshot.'})
    return path


def audit_and_prepare(csv):
    OUT.mkdir(parents=True, exist_ok=True)
    REP.mkdir(parents=True, exist_ok=True)
    con = duckdb.connect()
    con.execute('SET threads=4')
    source = f"read_csv('{csv.as_posix()}', header=true, all_varchar=true)"
    cols = con.sql(f'DESCRIBE SELECT * FROM {source}').df()['column_name'].tolist()
    dates = {'violation_date', 'violation_status_date', 'violation_last_modified_date'}
    expr = []
    for c in cols:
        s = re.sub('[^a-z0-9]+', '_', c.lower()).strip('_')
        q = '"' + c.replace('"', '""') + '"'
        base = f"nullif(trim({q}), '')"
        if s in dates:
            base = f"coalesce(try_strptime({base}, '%m/%d/%Y'), try_strptime({base}, '%m/%d/%Y %I:%M:%S %p'), try_cast({base} as timestamp))"
        elif s in {'latitude', 'longitude'}:
            base = f'try_cast({base} as double)'
        if s in {'inspection_status', 'department_bureau', 'inspection_category', 'inspection_waived'}:
            base = f'upper({base})'
        expr.append(f'{base} AS "{s}"')
    pq = RAW / 'building_violations.parquet'
    t = time.perf_counter()
    con.execute(f"COPY (SELECT {', '.join(expr)} FROM {source}) TO '{pq.as_posix()}' (FORMAT PARQUET, COMPRESSION ZSTD)")
    convert_s = time.perf_counter() - t
    con.execute(f"CREATE VIEW raw AS SELECT * FROM read_parquet('{pq.as_posix()}')")
    n = con.sql('SELECT count(*) FROM raw').fetchone()[0]
    names = [re.sub('[^a-z0-9]+', '_', c.lower()).strip('_') for c in cols]
    nulls = con.sql('SELECT ' + ', '.join(f'count(*)-count("{c}") AS "{c}"' for c in names) + ' FROM raw').df().iloc[0]
    missing = pd.DataFrame({'column': names, 'missing_n': nulls.values.astype(int)})
    missing['missing_pct'] = missing['missing_n'] / n * 100
    missing.sort_values('missing_pct', ascending=False).to_csv(REP / 'missingness.csv', index=False)
    checks = con.sql(f"""SELECT count(*)-count(DISTINCT id) AS duplicate_ids,
      count(*)-(SELECT count(*) FROM (SELECT DISTINCT * FROM raw)) AS duplicate_rows,
      count(*) FILTER (WHERE violation_date IS NULL OR violation_date<'2006-01-01' OR violation_date>'{CUTOFF}') AS bad_dates,
      count(*) FILTER (WHERE violation_status_date<violation_date) AS reversed_status_dates,
      count(*) FILTER (WHERE latitude IS NULL OR longitude IS NULL) AS missing_coordinates,
      count(*) FILTER (WHERE latitude IS NOT NULL AND longitude IS NOT NULL AND NOT (latitude BETWEEN 41.6 AND 42.1 AND longitude BETWEEN -88 AND -87.5)) AS outside_coordinates,
      count(*) FILTER (WHERE inspection_number IS NULL OR property_group IS NULL) AS missing_keys,
      min(violation_date) AS date_min, max(violation_date) AS date_max,
      count(DISTINCT property_group) AS objects, count(DISTINCT inspection_number) AS inspection_ids,
      count(DISTINCT violation_code) AS violation_types FROM raw""").df().iloc[0].to_dict()
    # Check failed casts against raw strings rather than assuming typed NULLs were absent.
    date_parse = {}
    for c in cols:
        s = re.sub('[^a-z0-9]+', '_', c.lower()).strip('_')
        if s in dates:
            date_parse[s] = int(con.sql(f'''SELECT count(*) FROM {source} a JOIN raw b ON a."ID"=b.id
              WHERE nullif(trim(a."{c}"),'') IS NOT NULL AND b."{s}" IS NULL''').fetchone()[0])
    con.execute("""CREATE VIEW dedup AS SELECT * EXCLUDE (rn) FROM
      (SELECT *, row_number() OVER (PARTITION BY id ORDER BY violation_last_modified_date DESC NULLS LAST) rn FROM raw) WHERE rn=1""")
    # Reject ambiguous inspection labels/metadata, instead of arbitrarily choosing first/last.
    con.execute(f"""CREATE TABLE inspections AS SELECT inspection_number,
       min(violation_date) AS inspection_date, min(property_group) AS property_group,
       min(department_bureau) AS department_bureau, min(inspection_category) AS inspection_category,
       min(inspection_status) AS inspection_status,
       avg(CASE WHEN latitude BETWEEN 41.6 AND 42.1 AND longitude BETWEEN -88 AND -87.5 THEN latitude END) AS latitude,
       avg(CASE WHEN latitude BETWEEN 41.6 AND 42.1 AND longitude BETWEEN -88 AND -87.5 THEN longitude END) AS longitude,
       count(*) AS n_violations, count(DISTINCT violation_code) AS n_codes,
       count(DISTINCT inspection_status) AS status_variants,
       count(DISTINCT property_group) AS object_variants,
       count(DISTINCT department_bureau) AS bureau_variants,
       count(DISTINCT inspection_category) AS category_variants,
       count(DISTINCT violation_date) AS date_variants
       FROM dedup WHERE inspection_number IS NOT NULL AND violation_date BETWEEN '2006-01-01' AND '{CUTOFF}'
       GROUP BY inspection_number""")
    d = con.sql('SELECT * FROM inspections ORDER BY inspection_date, inspection_number').df()
    ambiguous = (d[['status_variants', 'object_variants', 'bureau_variants', 'category_variants', 'date_variants']]>1).any(axis=1)
    checks['ambiguous_inspections'] = int(ambiguous.sum())
    checks['missing_inspection_status'] = int(d['inspection_status'].isna().sum())
    checks['inspection_conflict_counts'] = {c: int((d[c]>1).sum()) for c in ['status_variants','object_variants','bureau_variants','category_variants','date_variants']}
    all_status = d['inspection_status'].fillna('UNKNOWN').value_counts().to_dict()
    d['ambiguous'] = ambiguous
    d.to_parquet(OUT / 'inspections.parquet', index=False)
    labeled = d.loc[~ambiguous & d['inspection_status'].isin(['FAILED', 'PASSED'])].copy()
    labeled['target_failed'] = labeled['inspection_status'].eq('FAILED').astype('int8')
    dt = labeled['inspection_date'].dt
    labeled['year'] = dt.year
    labeled['month'] = dt.month
    labeled['month_sin'] = np.sin(2*np.pi*dt.month/12)
    labeled['month_cos'] = np.cos(2*np.pi*dt.month/12)
    labeled['dayofweek'] = dt.dayofweek
    for c in CAT:
        labeled[c] = labeled[c].fillna('UNKNOWN').astype(str)
    labeled.to_parquet(OUT / 'model_data.parquet', index=False)
    raw_status = con.sql('SELECT coalesce(inspection_status,\'UNKNOWN\') AS status, count(*) AS n FROM raw GROUP BY 1').df().set_index('status')['n'].to_dict()
    summary = {'rows': int(n), 'columns': len(cols), **checks,
               'date_parse_failures': date_parse, 'raw_status_counts': raw_status,
               'inspection_status_counts': all_status, 'aggregated_inspections': len(d),
               'labeled_inspections': len(labeled), 'failed_share': float(labeled.target_failed.mean()),
               'csv_mb': csv.stat().st_size/1e6, 'parquet_mb': pq.stat().st_size/1e6,
               'convert_seconds': convert_s}
    for k, v in list(summary.items()):
        if isinstance(v, np.integer): summary[k] = int(v)
    save_json(REP / 'quality.json', summary)
    con.close()
    return d, labeled, missing, summary


def explore(d, labeled, missing):
    FIG.mkdir(parents=True, exist_ok=True)
    plt.rcParams.update({'font.family': 'DejaVu Sans', 'font.size': 10, 'axes.spines.top': False,
                         'axes.spines.right': False, 'axes.labelcolor': DARK, 'text.color': DARK,
                         'axes.edgecolor': GRAY, 'figure.facecolor': PAPER, 'axes.facecolor': PAPER,
                         'savefig.facecolor': PAPER})
    insights = []
    def finish(fig, name, title, conclusion):
        fig.suptitle(title, fontsize=14, fontweight='bold', x=.03, ha='left')
        if name != 'eda07_geo':
            fig.tight_layout(rect=[0, 0, 1, .92])
        fig.savefig(FIG / f'{name}.png', dpi=180)
        svg = FIG / f'{name}.svg'
        fig.savefig(svg)
        svg.write_text('\n'.join(line.rstrip() for line in svg.read_text().splitlines())+'\n')
        plt.close(fig)
        insights.append({'id': name, 'title': title, 'conclusion': conclusion})
    fig, ax = plt.subplots(figsize=(9, 4.3))
    m = missing.sort_values('missing_pct').tail(10)
    ax.barh(m.column, m.missing_pct, color=ORANGE)
    ax.set(xlabel='Пропуски, %', xlim=(0,100))
    top = m.iloc[-1]
    finish(fig, 'eda01_missing', '1. Полнота исходных полей', f"Максимальная доля пропусков — {top.missing_pct:.1f}% в {top.column}. Даты статуса, комментарии и SSA не включены в модель; координаты заполняются медианой только по train с индикатором пропуска.")
    yearly = d.groupby(d.inspection_date.dt.year).size()
    yearly.rename('inspections').to_csv(REP/'annual_counts.csv')
    fig, ax = plt.subplots(figsize=(9, 4.3))
    ax.bar(yearly.index, yearly.values, color=[ORANGE if y==2026 else DARK for y in yearly.index])
    ax.set(xlabel='Год (2026 — неполный)', ylabel='Число инспекций в журнале')
    finish(fig, 'eda02_year', '2. Объём наблюдений по годам', f"Максимум в архиве — {int(yearly.idxmax())} год: {int(yearly.max()):,} инспекций. 2026 год неполный и не используется как тест. Изменение объёма отражает и практику регистрации, поэтому не доказывает изменение качества зданий.")
    counts=d.inspection_status.fillna('UNKNOWN').value_counts()
    fig, ax=plt.subplots(figsize=(9,4.3))
    bars=ax.bar(counts.index, counts.values, color=[ORANGE if x=='FAILED' else DARK for x in counts.index])
    ax.bar_label(bars,labels=[f'{v/len(d):.1%}' for v in counts.values],padding=4)
    ax.set(ylabel='Инспекции (до исключения конфликтов)',ylim=(0,counts.max()*1.2))
    finish(fig,'eda03_status','3. Результаты на уровне инспекций',f"В бинарной выборке FAILED составляет {labeled.target_failed.mean():.1%}. CLOSED, HOLD и UNKNOWN исключены: отсутствие FAILED не означает PASSED. Accuracy без метрик классов может скрывать пропуски проблемных проверок.")
    n=d.n_violations
    bins=pd.cut(n,[0,1,2,5,10,20,np.inf],labels=['1','2','3–5','6–10','11–20','>20'])
    hist=bins.value_counts(sort=False)
    fig,ax=plt.subplots(figsize=(9,4.3));ax.bar(hist.index.astype(str),hist.values,color=DARK)
    ax.set(xlabel='Записей нарушения на инспекцию',ylabel='Инспекции')
    finish(fig,'eda04_grain','4. Почему строки нельзя считать инспекциями',f"Медиана — {n.median():.0f} записи, 95-й перцентиль — {n.quantile(.95):.0f}, максимум — {n.max():.0f}. Доля FAILED по строкам перевзвешивает инспекции с большим числом нарушений. Число нарушений текущей проверки используется только в EDA.")
    bureau=labeled.groupby('department_bureau').target_failed.agg(['size','mean']).sort_values('size',ascending=False).head(10).sort_values('mean')
    p=bureau['mean']; nn=bureau['size'];z=1.96
    center=(p+z*z/(2*nn))/(1+z*z/nn); half=z*np.sqrt(p*(1-p)/nn+z*z/(4*nn*nn))/(1+z*z/nn)
    fig,ax=plt.subplots(figsize=(9,4.8));ax.errorbar(p,bureau.index,xerr=[p-(center-half),(center+half)-p],fmt='o',color=ORANGE,capsize=3)
    ax.set(xlabel='Доля FAILED; описательные 95% интервалы Уилсона',xlim=(-.03,1.03))
    bureau.to_csv(REP/'bureau_rates.csv')
    finish(fig,'eda05_bureau','5. Результат зависит от службы',f"Среди 10 крупнейших служб доля FAILED меняется от {p.min():.1%} до {p.max():.1%}. Служба — кандидат в признаки, но различие может отражать процедуры учёта. Интервалы описательные и не учитывают зависимость повторных инспекций одного объекта.")
    season=labeled[labeled.year.between(2006,2025)].groupby('month').target_failed.agg(['size','mean'])
    fig,(a,b)=plt.subplots(1,2,figsize=(9,4.3));a.bar(season.index,season['size'],color=DARK);b.plot(season.index,season['mean'],color=ORANGE,marker='o')
    for ax in [a,b]:ax.set(xlabel='Месяц',xticks=[1,3,6,9,12])
    a.set(ylabel='Инспекции, 2006–2025');b.set(ylabel='Доля FAILED',ylim=(0,1))
    finish(fig,'eda06_season','6. Сезонность наблюдений и результата',f"В полных годах месячная доля FAILED составляет {season['mean'].min():.1%}–{season['mean'].max():.1%}. Месяц кодируется циклически; это ассоциация, смешанная с составом служб и лет, а не причинный эффект сезона.")
    from src.geography import inspection_map
    coords=d.dropna(subset=['longitude','latitude'])
    fig=inspection_map(d)
    finish(fig,'eda07_geo','7. Пространственная концентрация инспекций',f"У {len(d)-len(coords):,} инспекций нет пригодных координат. Карта показывает интенсивность наблюдения в архиве, без нормировки на число зданий. География может помочь ранжированию, но не позволяет объявить район более опасным.")
    save_json(REP/'eda_insights.json',insights)
    return insights


def model(labeled):
    train=labeled[labeled.year<=2023];val=labeled[labeled.year==2024];test=labeled[labeled.year==2025]
    for name,part in [('train',train),('validation',val),('test',test)]:
        assert len(part)>0 and part.target_failed.nunique()==2, f'Empty/one-class {name}'
    assert train.inspection_date.max()<val.inspection_date.min()<test.inspection_date.min()
    assert not set(FEATURES)&{'inspection_status','n_violations','n_codes','inspection_waived','target_failed'}
    X=train[FEATURES];y=train.target_failed
    prep=ColumnTransformer([
        ('cat',Pipeline([('missing',SimpleImputer(strategy='constant',fill_value='UNKNOWN')),('encode',OneHotEncoder(handle_unknown='ignore'))]),CAT),
        ('num',Pipeline([('missing',SimpleImputer(strategy='median',add_indicator=True)),('scale',StandardScaler())]),NUM)])
    models={'Dummy prior':DummyClassifier(strategy='prior'),
            'Logistic regression':Pipeline([('prepare',prep),('model',LogisticRegression(max_iter=2000,random_state=42,solver='lbfgs'))])}
    rows=[];matrices={};thresholds={}
    for name,est in models.items():
        t=time.perf_counter();est.fit(X,y);elapsed=time.perf_counter()-t
        pv=est.predict_proba(val[FEATURES])[:,1]
        # Only validation labels choose the operating threshold, FN cost is 5, FP cost 1.
        grid=np.linspace(0,1,201)
        costs=[5*int(((pv<th)&(val.target_failed==1)).sum())+int(((pv>=th)&(val.target_failed==0)).sum()) for th in grid]
        th=float(grid[int(np.argmin(costs))]);thresholds[name]=th
        for split,part in [('validation',val),('test',test)]:
            p=est.predict_proba(part[FEATURES])[:,1];pred=(p>=th).astype(int);yy=part.target_failed
            tn,fp,fn,tp=confusion_matrix(yy,pred,labels=[0,1]).ravel()
            row={'model':name,'split':split,'n':len(part),'prevalence':float(yy.mean()),'threshold':th,
                 'accuracy':accuracy_score(yy,pred),'balanced_accuracy':balanced_accuracy_score(yy,pred),
                 'precision':precision_score(yy,pred,zero_division=0),'recall':recall_score(yy,pred),
                 'f1':f1_score(yy,pred),'roc_auc':roc_auc_score(yy,p),'average_precision':average_precision_score(yy,p),
                 'brier':brier_score_loss(yy,p),'cost_per_inspection':float((5*fn+fp)/len(part)),
                 'tn':int(tn),'fp':int(fp),'fn':int(fn),'tp':int(tp),'fit_seconds':elapsed}
            rows.append(row)
            if split=='test':
                matrices[name]=[[int(tn),int(fp)],[int(fn),int(tp)]]
                if name=='Logistic regression':
                    pd.DataFrame({'inspection_number':part.inspection_number,'target':yy,'probability':p,'prediction':pred}).to_parquet(OUT/'test_predictions.parquet',index=False)
    pd.DataFrame(rows).to_csv(REP/'baseline_metrics.csv',index=False)
    splits={name:{'n':len(part),'failed_share':float(part.target_failed.mean()),'date_from':str(part.inspection_date.min().date()),'date_to':str(part.inspection_date.max().date())} for name,part in [('train',train),('validation',val),('test',test)]}
    logistic_test=next(r for r in rows if r['model']=='Logistic regression' and r['split']=='test')
    dummy_test=next(r for r in rows if r['model']=='Dummy prior' and r['split']=='test')
    success=(logistic_test['average_precision']>=dummy_test['average_precision']+.05 and logistic_test['recall']>=.8 and logistic_test['precision']>=logistic_test['prevalence'])
    overlap=len(set(train.property_group)&set(test.property_group))/test.property_group.nunique()
    result={'features':FEATURES,'splits':splits,'excluded_2026':int((labeled.year==2026).sum()),'thresholds':thresholds,'fn_cost':5,'fp_cost':1,
            'metrics':rows,'confusion_matrices_test':matrices,'success_criterion_met':bool(success),
            'test_objects_seen_in_train_share':overlap,
            'evaluation_protocol':'Frozen models fitted on <=2023; threshold chosen on 2024 only; evaluated on 2025. No refitting on validation. Historical snapshot lacks point-in-time field versions.'}
    save_json(REP/'baseline.json',result)
    fig,axes=plt.subplots(1,2,figsize=(9,3.8))
    for ax,(name,matrix) in zip(axes,matrices.items()):
        mat=np.array(matrix);ax.imshow(mat,cmap='Oranges');ax.set(xticks=[0,1],yticks=[0,1],xticklabels=['PASSED','FAILED'],yticklabels=['PASSED','FAILED'],xlabel='Прогноз',ylabel='Факт',title=name)
        for i in range(2):
            for j in range(2):ax.text(j,i,f'{mat[i,j]:,}',ha='center',va='center',color=DARK)
    fig.tight_layout();fig.savefig(FIG/'baseline_confusion.png',dpi=180);plt.close(fig)
    return result


def run():
    print('Download / verify snapshot',flush=True);csv=download()
    print('Audit and aggregate full archive',flush=True);d,labeled,missing,q=audit_and_prepare(csv)
    print(json.dumps(q,ensure_ascii=False,default=str),flush=True)
    print('Generate 7 EDA figures',flush=True);explore(d,labeled,missing)
    print('Train and evaluate baselines',flush=True);m=model(labeled)
    save_json(REP/'environment.json',{'python':platform.python_version(),'platform':platform.platform(),'packages':{n:importlib.metadata.version(n) for n in ['pandas','pyarrow','duckdb','scikit-learn','matplotlib','python-pptx','python-docx','requests','nbformat','nbclient']}})
    print(pd.DataFrame(m['metrics']).to_string(index=False),flush=True)


def run_snapshot():
    """Recompute EDA and baseline using the compact committed snapshot."""
    d=pd.read_parquet(OUT/'inspections.parquet')
    labeled=pd.read_parquet(OUT/'model_data.parquet')
    missing=pd.read_csv(REP/'missingness.csv')
    explore(d,labeled,missing)
    result=model(labeled)
    print(pd.DataFrame(result['metrics']).to_string(index=False),flush=True)
    return result


if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--download-only',action='store_true')
    parser.add_argument('--from-snapshot',action='store_true',help='Use committed compact data; no download or raw audit.')
    args=parser.parse_args()
    if args.download_only: download()
    elif args.from_snapshot: run_snapshot()
    else: run()
