-- ================================================================
-- KPI Panel — setup tabelle autenticazione
-- Incolla questo nell'SQL Editor di Supabase
-- ================================================================

-- Utenti
CREATE TABLE IF NOT EXISTS kpi_users (
  id            UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  email         TEXT UNIQUE NOT NULL,
  name          TEXT NOT NULL,
  password_hash TEXT NOT NULL,
  role          TEXT NOT NULL DEFAULT 'admin', -- 'admin' | 'viewer'
  dashboard     TEXT NOT NULL DEFAULT 'rs-italia', -- 'rs-italia' | 'optimedia'
  is_active     BOOLEAN DEFAULT true,
  created_at    TIMESTAMPTZ DEFAULT NOW(),
  last_login    TIMESTAMPTZ
);

-- Log accessi
CREATE TABLE IF NOT EXISTS kpi_access_logs (
  id         UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  user_id    UUID REFERENCES kpi_users(id),
  email      TEXT,
  ip         TEXT,
  user_agent TEXT,
  action     TEXT, -- 'login' | 'logout' | 'access' | 'failed_login'
  created_at TIMESTAMPTZ DEFAULT NOW()
);

-- Indici
CREATE INDEX IF NOT EXISTS kpi_logs_user_id    ON kpi_access_logs(user_id);
CREATE INDEX IF NOT EXISTS kpi_logs_created_at ON kpi_access_logs(created_at DESC);

-- Row Level Security (solo il service_role può leggere/scrivere)
ALTER TABLE kpi_users       ENABLE ROW LEVEL SECURITY;
ALTER TABLE kpi_access_logs ENABLE ROW LEVEL SECURITY;

-- Nessun accesso pubblico (solo service_role via API server-side)
CREATE POLICY "no_public_access_users" ON kpi_users       FOR ALL USING (false);
CREATE POLICY "no_public_access_logs"  ON kpi_access_logs FOR ALL USING (false);

-- ================================================================
-- MIGRAZIONE multi-tenant (esegui solo se la tabella esiste già)
-- ================================================================
ALTER TABLE kpi_users ADD COLUMN IF NOT EXISTS dashboard TEXT NOT NULL DEFAULT 'rs-italia';

-- Crea utente Optimedia (poi cambia la password con add_user.js):
-- node --env-file=.env scripts/add_user.js optimedia@example.com "Optimedia Admin" PASSWORD admin optimedia
