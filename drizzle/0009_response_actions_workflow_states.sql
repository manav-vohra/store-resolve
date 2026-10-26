-- 0009: Fix response_actions CHECK constraints to match the approval-gated
-- reply workflow (worker/response-workflow.ts) and create response_templates.
--
-- Problem: migration 0006 created response_actions with CHECK constraints
-- that reject every value the workflow writes:
--   response_kind: only RECEIPT_ACK/SUBSTANTIVE_RESPONSE, code writes
--     ACKNOWLEDGMENT/RESOLUTION (and arbitrary kinds)
--   policy_state: only the old Phase C states, code writes
--     DRAFT/PENDING_APPROVAL/APPROVED/REJECTED/SENDING/SENT/SEND_FAILED/
--     NEEDS_RECONCILIATION
--   content_kind: only DETERMINISTIC_TEMPLATE/AI_DRAFT/HUMAN_EDITED,
--     code writes 'draft-body'
-- Also, the UNIQUE(complaint_id, response_kind, policy_version) constraint
-- blocks Dad from creating a second draft for the same complaint.
--
-- Fix: recreate the table with expanded policy_state CHECK (old + new
-- states), no response_kind/content_kind CHECK, and no UNIQUE constraint.
-- All columns, foreign keys, and unrelated constraints are preserved.
-- response_actions is empty in production (verified 2026-10-25), so the
-- copy step moves zero rows; it is kept for safety regardless.
-- Also creates response_templates (the content of deferred 0007) so the
-- template endpoints work; 0007 itself stays unapplied to avoid confusion.

PRAGMA foreign_keys = OFF;

CREATE TABLE response_actions_new (
  id TEXT PRIMARY KEY,
  complaint_id TEXT NOT NULL REFERENCES complaints(id),
  source_message_id TEXT REFERENCES mail_source_messages(id),
  response_kind TEXT NOT NULL,
  policy_version TEXT NOT NULL,
  policy_state TEXT NOT NULL CHECK(policy_state IN (
    'DRAFT', 'PENDING_APPROVAL', 'APPROVED', 'REJECTED',
    'SENDING', 'SENT', 'SEND_FAILED', 'NEEDS_RECONCILIATION',
    'NOT_EVALUATED', 'NO_RESPONSE', 'ACK_ELIGIBLE', 'ACK_ALREADY_SENT', 'ACK_BLOCKED',
    'DRAFT_REQUIRED', 'OWNER_APPROVAL_REQUIRED', 'READY_TO_SEND'
  )),
  content_kind TEXT,
  content_ref TEXT,
  template_version TEXT,
  model_name TEXT,
  model_snapshot TEXT,
  prompt_version TEXT,
  approved_by TEXT REFERENCES users(id),
  approved_at TEXT,
  created_at TEXT NOT NULL,
  updated_at TEXT NOT NULL
);

INSERT INTO response_actions_new
  SELECT id, complaint_id, source_message_id, response_kind, policy_version,
    policy_state, content_kind, content_ref, template_version, model_name,
    model_snapshot, prompt_version, approved_by, approved_at, created_at, updated_at
  FROM response_actions;

DROP TABLE response_actions;

ALTER TABLE response_actions_new RENAME TO response_actions;

CREATE TABLE IF NOT EXISTS response_templates (
  id TEXT PRIMARY KEY,
  name TEXT NOT NULL,
  response_kind TEXT NOT NULL,
  subject_template TEXT,
  body_template TEXT NOT NULL,
  version INTEGER NOT NULL DEFAULT 1,
  active INTEGER NOT NULL DEFAULT 1,
  created_at TEXT NOT NULL,
  updated_at TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_response_templates_kind_active
  ON response_templates(response_kind, active);

PRAGMA foreign_keys = ON;
