-- ============================================================
-- Middlesex Frame Tracker — Supabase Database Setup
-- Run this in: Supabase Dashboard → SQL Editor → New query
-- ============================================================

-- Frame installation tracking (Graboyes field entries)
CREATE TABLE IF NOT EXISTS frame_updates (
  id                INTEGER PRIMARY KEY,
  shipment_received DATE,
  install_started   DATE,
  install_completed DATE,
  graboyes_notes    TEXT,
  updated_at        TIMESTAMPTZ DEFAULT NOW()
);

-- Cassette tracking
CREATE TABLE IF NOT EXISTS cassette_updates (
  drawing        TEXT PRIMARY KEY,
  qty_fabricated INTEGER,
  date_shipped   DATE,
  date_installed DATE,
  updated_at     TIMESTAMPTZ DEFAULT NOW()
);

-- Enable Row Level Security
ALTER TABLE frame_updates    ENABLE ROW LEVEL SECURITY;
ALTER TABLE cassette_updates ENABLE ROW LEVEL SECURITY;

-- Allow read/write from the app without login
CREATE POLICY "team access" ON frame_updates
  FOR ALL USING (true) WITH CHECK (true);

CREATE POLICY "team access" ON cassette_updates
  FOR ALL USING (true) WITH CHECK (true);
