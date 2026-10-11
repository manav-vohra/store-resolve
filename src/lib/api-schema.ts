import { z } from 'zod'

export const severitySchema = z.enum(['LOW', 'MEDIUM', 'HIGH', 'CRITICAL'])
export const complaintInputSchema = z.object({
  externalCaseId: z.string().trim().min(1).max(100),
  storeNumber: z.string().trim().min(1).max(30),
  subject: z.string().trim().min(1).max(300),
  complaintText: z.string().trim().min(1).max(20_000),
  category: z.string().trim().min(1).max(100),
  severity: severitySchema,
})
export const actionSchema = z.object({
  action: z.enum([
    'ACKNOWLEDGE',
    'START_INVESTIGATION',
    'CONTACT_CUSTOMER',
    'SUBMIT_RESOLUTION',
    'CLOSE',
    'REOPEN',
  ]),
  data: z.record(z.string(), z.unknown()).default({}),
})
export const configSchema = z.object({
  mode: z.enum(['MOCK', 'FAMILY_PILOT', 'SINGLE_STORE_PILOT', 'FULL']),
  externalNotificationsEnabled: z.boolean(),
  pilotStoreId: z.string().optional(),
  emailIngestionEnabled: z.boolean().optional().default(false),
  emailAckEnabled: z.boolean().optional().default(false),
  managerAckDeadlineMinutes: z.number().int().min(5).max(10_080).optional().default(30),
  managerResolutionTargetHours: z.number().int().min(1).max(720).optional().default(24),
  escalationIntervalMinutes: z.number().int().min(15).max(10_080).optional().default(60),
  signalWireReconcileAfterMinutes: z.number().int().min(5).max(1_440).optional().default(10),
  emailLookbackDays: z.number().int().min(1).max(365).optional().default(30),
})
export const contactSchema = z.object({
  name: z.string().trim().min(1).max(100),
  email: z.union([z.literal(''), z.string().email().max(254)]),
  phone: z.union([z.literal(''), z.string().regex(/^\+[1-9]\d{7,14}$/, 'Use E.164 format')]),
  smsEnabled: z.boolean(),
  active: z.boolean(),
})
export const testNotificationSchema = z.object({
  recipientUserId: z.enum(['father', 'uncle', 'grandfather', 'pilot-admin', 'dad']),
  confirmed: z.literal(true),
})
const base64url = z
  .string()
  .min(8)
  .max(256)
  .regex(/^[A-Za-z0-9_-]+$/, 'Must be base64url')
export const pushSubscribeSchema = z.object({
  endpoint: z.string().url().startsWith('https://').max(2000),
  keys: z.object({ p256dh: base64url, auth: base64url }),
  deviceLabel: z.string().trim().min(1).max(80).optional(),
})
export const pushUnsubscribeSchema = z.object({
  endpoint: z.string().url().startsWith('https://').max(2000),
})
export const signalWireCallbackSchema = z.union([
  z.object({ id: z.string().uuid(), project_id: z.string().uuid() }),
  z.object({ MessageSid: z.string().uuid() }),
])
export const signalWireReconcileSchema = z.object({
  providerMessageId: z.string().uuid(),
})
export const storeAdminSchema = z.object({
  number: z.string().trim().min(1).max(30),
  name: z.string().trim().min(1).max(150),
  address: z.string().trim().min(1).max(250),
  city: z.string().trim().min(1).max(100),
  state: z.string().trim().min(2).max(50),
  postalCode: z.string().trim().min(3).max(20),
  phone: z.union([z.literal(''), z.string().regex(/^\+[1-9]\d{7,14}$/)]),
  active: z.boolean(),
  managerId: z.string().trim().max(100).optional(),
  aliases: z.array(z.string().trim().min(3).max(250)).max(25).default([]),
})
export const userAdminSchema = z.object({
  name: z.string().trim().min(1).max(100).optional(),
  email: z.string().email().max(254).optional(),
  phone: z.union([z.literal(''), z.string().regex(/^\+[1-9]\d{7,14}$/)]).optional(),
  role: z.enum(['OWNER', 'VIEW_ONLY', 'STORE_MANAGER', 'ADMIN']).optional(),
  active: z.boolean().optional(),
  smsEnabled: z.boolean().optional(),
  complaintNotificationsEnabled: z.boolean().optional(),
  timezone: z.string().trim().min(1).max(100).optional(),
  storeIds: z.array(z.string().trim().min(1).max(100)).max(50).optional(),
})

// Phase C: family review queue
export const reviewResolveSchema = z.object({
  action: z.enum(['dismiss', 'link']),
  complaintId: z.string().trim().min(1).max(100).optional(),
  note: z.string().trim().max(2000).optional(),
})
export const reviewAssignSchema = z.object({
  assigneeUserId: z.string().trim().min(1).max(100),
})

// Phase C: response templates + approval-gated workflow
export const responseTemplateSchema = z.object({
  name: z.string().trim().min(1).max(200),
  responseKind: z.string().trim().min(1).max(100),
  subjectTemplate: z.string().trim().max(500).optional(),
  bodyTemplate: z.string().trim().min(1).max(20000),
})
export const responseDraftSchema = z.object({
  responseKind: z.string().trim().min(1).max(100).optional(),
  templateId: z.string().trim().min(1).max(100).optional(),
  templateVars: z.record(z.string(), z.string().max(2000)).optional(),
  body: z.string().trim().min(1).max(20000).optional(),
})
export const responseApprovalSchema = z.object({})
