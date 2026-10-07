CREATE TABLE IF NOT EXISTS efetividade (
  data TEXT NOT NULL, dia_semana TEXT, empresa TEXT NOT NULL,
  cod_empresa INTEGER, cod_posto INTEGER NOT NULL, posto TEXT NOT NULL, cargo TEXT NOT NULL,
  seq INTEGER NOT NULL DEFAULT 0,          -- desempata linhas repetidas no mesmo dia/posto/cargo
  qe_d INTEGER DEFAULT 0, qe_n INTEGER DEFAULT 0,
  faltas_d INTEGER DEFAULT 0, faltas_n INTEGER DEFAULT 0,
  res_d INTEGER DEFAULT 0, res_n INTEGER DEFAULT 0,
  efet_d INTEGER DEFAULT 0, efet_n INTEGER DEFAULT 0,
  colaborador_falta TEXT, reserva_cobertura TEXT, obs TEXT, obs_extra TEXT,
  PRIMARY KEY (data, cod_posto, posto, cargo, seq)
);
CREATE INDEX IF NOT EXISTS ix_data ON efetividade(data);
CREATE INDEX IF NOT EXISTS ix_empresa ON efetividade(empresa, data);

CREATE TABLE IF NOT EXISTS sync_state (data TEXT PRIMARY KEY, hash TEXT NOT NULL, atualizado_em TEXT);
