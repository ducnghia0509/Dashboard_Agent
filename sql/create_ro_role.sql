-- Tao role Postgres READ-ONLY danh rieng cho luong qa (dashboard_qa).
-- KHONG tu chay script nay - user tu chay thu cong tren Postgres that (psql / pgAdmin / Neon SQL editor)
-- roi dien connection string cua role nay vao DATABASE_URL_RO trong DashBoard_Agent/.env.
--
-- Doi 'CHANGE_ME_STRONG_PASSWORD' truoc khi chay.

CREATE ROLE dashboard_qa_ro LOGIN PASSWORD 'CHANGE_ME_STRONG_PASSWORD';

GRANT CONNECT ON DATABASE tc_dashboard TO dashboard_qa_ro;
GRANT USAGE ON SCHEMA public TO dashboard_qa_ro;
GRANT SELECT ON ALL TABLES IN SCHEMA public TO dashboard_qa_ro;

-- Dam bao cac bang tao SAU nay cung tu dong duoc SELECT (khong can chay lai script).
ALTER DEFAULT PRIVILEGES IN SCHEMA public GRANT SELECT ON TABLES TO dashboard_qa_ro;

-- Gioi han them (khuyen nghi, tuy Postgres/hosting co ho tro):
-- ALTER ROLE dashboard_qa_ro SET statement_timeout = '5s';
-- ALTER ROLE dashboard_qa_ro CONNECTION LIMIT 5;
