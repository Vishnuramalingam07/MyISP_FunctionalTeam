-- ============================================================
-- MyISP Tools – Local PostgreSQL Database Schema
-- ============================================================

-- 1. Authorized Users
CREATE TABLE IF NOT EXISTS authorized_users (
    id          BIGSERIAL PRIMARY KEY,
    username    TEXT NOT NULL UNIQUE,
    email       TEXT,
    role        TEXT DEFAULT 'member',
    created_at  TIMESTAMPTZ DEFAULT now()
);

-- 2. Team Members
CREATE TABLE IF NOT EXISTS team_members (
    id          BIGSERIAL PRIMARY KEY,
    name        TEXT NOT NULL,
    lead        TEXT NOT NULL DEFAULT '',
    location    TEXT NOT NULL DEFAULT '',
    active      BOOLEAN NOT NULL DEFAULT TRUE,
    created_at  TIMESTAMPTZ DEFAULT now()
);

-- 3. Attendance Records
CREATE TABLE IF NOT EXISTS attendance_records (
    id          BIGSERIAL PRIMARY KEY,
    member_name TEXT NOT NULL,
    lead_name   TEXT NOT NULL DEFAULT '',
    location    TEXT NOT NULL DEFAULT '',
    year        SMALLINT NOT NULL,
    month       SMALLINT NOT NULL CHECK (month BETWEEN 1 AND 12),
    day         SMALLINT NOT NULL CHECK (day BETWEEN 1 AND 31),
    status      TEXT NOT NULL DEFAULT '',
    updated_at  TIMESTAMPTZ DEFAULT now(),
    UNIQUE (member_name, year, month, day)
);

CREATE INDEX IF NOT EXISTS idx_attendance_year_month
    ON attendance_records (year, month);

CREATE INDEX IF NOT EXISTS idx_attendance_member
    ON attendance_records (member_name);

-- 4. Attendance Logs
CREATE TABLE IF NOT EXISTS attendance_logs (
    id           BIGSERIAL PRIMARY KEY,
    saved_at     TIMESTAMPTZ DEFAULT now(),
    user_id      TEXT NOT NULL DEFAULT '',
    lead_name    TEXT NOT NULL DEFAULT '',
    member_name  TEXT NOT NULL DEFAULT '',
    location     TEXT NOT NULL DEFAULT '',
    month_name   TEXT NOT NULL DEFAULT '',
    year         SMALLINT NOT NULL,
    day          SMALLINT,
    old_value    TEXT DEFAULT '',
    new_value    TEXT DEFAULT '',
    changed      TEXT DEFAULT '',
    sheet_name   TEXT DEFAULT '',
    client_ip    TEXT DEFAULT ''
);

CREATE INDEX IF NOT EXISTS idx_logs_saved_at
    ON attendance_logs (saved_at DESC);
