-- Stage A: Web Push subscriptions (additive; organization-scoped for future
-- franchise expansion). One row per (user, device endpoint); revoked rows are
-- kept for audit and never reused for sending.
CREATE TABLE push_subscriptions (
  id TEXT PRIMARY KEY,
  organization_id TEXT NOT NULL REFERENCES organizations(id),
  user_id TEXT NOT NULL REFERENCES users(id),
  endpoint TEXT NOT NULL UNIQUE,
  p256dh_key TEXT NOT NULL,
  auth_key TEXT NOT NULL,
  device_label TEXT,
  created_at TEXT NOT NULL,
  last_used_at TEXT,
  revoked_at TEXT
);
CREATE INDEX idx_push_subscriptions_user ON push_subscriptions(user_id);
