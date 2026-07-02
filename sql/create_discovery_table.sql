-- DU PHONG: neu sau nay muon chuyen DISCOVERY MEMORY tu JSON (memory/discoveries/*.json)
-- sang Postgres, dung schema nay + sua servers/common/memory.py de doc/ghi bang thay vi file.
-- KHONG chay trong lan trien khai nay (dang dung JSON, xem plan.md quyet dinh #3).

CREATE TABLE IF NOT EXISTS template_discoveries (
    id BIGSERIAL PRIMARY KEY,
    file_name TEXT NOT NULL,
    fingerprint TEXT NOT NULL,
    sheets JSONB NOT NULL DEFAULT '[]',
    columns_per_sheet JSONB NOT NULL DEFAULT '{}',
    detected_report_type TEXT,
    header_row INTEGER,
    mapping JSONB NOT NULL DEFAULT '{}',
    period TEXT,
    confidence DOUBLE PRECISION DEFAULT 0,
    anomalies JSONB NOT NULL DEFAULT '[]',
    created_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS ix_template_discoveries_fp ON template_discoveries(fingerprint);
CREATE INDEX IF NOT EXISTS ix_template_discoveries_rt ON template_discoveries(detected_report_type);
CREATE INDEX IF NOT EXISTS ix_template_discoveries_file ON template_discoveries(file_name);
