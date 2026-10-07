#!/usr/bin/env python3
"""Lê a aba EFETIVIDADE do Excel e gera SQL idempotente (apaga e reinsere os dias presentes no arquivo).
Uso: python etl/etl.py data/arquivo.xlsx data/load.sql"""
import sys, pandas as pd
COLS = ['data','dia_semana','empresa','cod_empresa','cod_posto','posto','cargo','qe_d','qe_n','faltas_d','faltas_n',
        'res_d','res_n','efet_d','pct_d','efet_n','pct_n','colaborador_falta','reserva_cobertura','obs','obs_extra']
NUM = ['qe_d','qe_n','faltas_d','faltas_n','res_d','res_n','efet_d','efet_n','cod_empresa','cod_posto']
TXT = ['dia_semana','empresa','posto','cargo','colaborador_falta','reserva_cobertura','obs','obs_extra']

def load(path):
    df = pd.read_excel(path, sheet_name='EFETIVIDADE', header=None, skiprows=3).iloc[:, :21]  # pula título + 2 linhas de cabeçalho
    df.columns = COLS
    df = df[pd.to_datetime(df.data, errors='coerce').notna()].copy()
    df['data'] = pd.to_datetime(df.data).dt.strftime('%Y-%m-%d')
    for c in NUM: df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0).astype(int)  # % é calculado na consulta
    for c in TXT: df[c] = df[c].map(lambda v: None if pd.isna(v) or str(v).strip()=='' else ' '.join(str(v).split()))
    df['seq'] = df.groupby(['data','cod_posto','posto','cargo']).cumcount()
    return df.drop(columns=['pct_d','pct_n'])

def q(v):
    if v is None: return 'NULL'
    return str(v) if isinstance(v, int) else "'" + str(v).replace("'", "''") + "'"

def to_sql(df, batch=50):
    cols = list(df.columns)
    rows = [[None if x is None or (isinstance(x, float) and x != x) else (int(x) if str(type(x)).startswith("<class 'numpy.int") else x) for x in r] for r in df.astype(object).itertuples(index=False)]
    out = ["DELETE FROM efetividade WHERE data IN (" + ",".join(q(d) for d in sorted(df.data.unique())) + ");"]
    for i in range(0, len(rows), batch):
        vals = ",\n".join("(" + ",".join(q(x) for x in r) + ")" for r in rows[i:i+batch])
        out.append(f"INSERT INTO efetividade ({','.join(cols)}) VALUES\n{vals};")
    return "\n".join(out)

if __name__ == '__main__':
    df = load(sys.argv[1]); open(sys.argv[2], 'w', encoding='utf-8').write(to_sql(df))
    print(f"{len(df)} linhas | {df.data.nunique()} dias | {df.data.min()} a {df.data.max()} -> {sys.argv[2]}")
