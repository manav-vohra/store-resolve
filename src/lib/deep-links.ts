// Shareable deep links for complaints: `#/complaints/:id` and
// `#/complaints/:id/reply` (opens the complaint scrolled to the Reply section).
//
// Used by Web Push notification taps (`notificationclick` →
// `clients.openWindow(url)`). The hash keeps the URL stable behind Cloudflare
// Access and requires no router — the app parses it on load and on
// `hashchange`.
export const parseComplaintHash = (hash: string): { id: string; section?: string } | undefined => {
  const match = /^#\/complaints\/([A-Za-z0-9-]+)(?:\/(reply))?\/?$/.exec(hash)
  if (!match) return undefined
  return { id: match[1], section: match[2] }
}

export const complaintHashFor = (id: string): string => `#/complaints/${id}`

export const complaintReplyHashFor = (id: string): string => `#/complaints/${id}/reply`
