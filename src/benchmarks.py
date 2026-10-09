"""Optional fresh local benchmarks; historical figures remain in the appendix."""
import gc
import json
import platform
import statistics
import time

import duckdb
import pandas as pd
import psutil

from src.pipeline import RAW, REP, save_json


def measure(fn, runs):
    timings=[]
    for _ in range(runs):
        gc.collect();start=time.perf_counter();result=fn();timings.append(time.perf_counter()-start)
        del result
    return {'median_seconds':statistics.median(timings),'runs_seconds':timings}


def run():
    csv=RAW/'building_violations.csv';pq=RAW/'building_violations.parquet'
    if not csv.exists() or not pq.exists():
        raise FileNotFoundError('First run raw/raw.ipynb or python -m src.pipeline.')
    con=duckdb.connect()
    result={'platform':platform.platform(),'python':platform.python_version(),'ram_total_gb':psutil.virtual_memory().total/1e9,
            'csv_mb':csv.stat().st_size/1e6,'parquet_mb':pq.stat().st_size/1e6,
            'note':'Local wall-clock times, mixed warm OS cache; CSV parsing versus typed Parquet reading. No cold-cache control.'}
    result['read_csv']=measure(lambda:pd.read_csv(csv,low_memory=False),3)
    result['read_parquet']=measure(lambda:pd.read_parquet(pq),5)
    result['duckdb_group']=measure(lambda:con.sql(f"SELECT inspection_status,count(*) FROM read_parquet('{pq.as_posix()}') GROUP BY 1").fetchall(),15)
    save_json(REP/'benchmarks_current.json',result)
    con.close();return result


if __name__=='__main__':
    print(json.dumps(run(),ensure_ascii=False,indent=2))
