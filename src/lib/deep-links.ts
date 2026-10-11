// Shareable deep links for complaints: `#/complaints/:id`.
//
// Used by Web Push notification taps (`notificationclick` →
// `clients.openWindow(url)`). The hash keeps the URL stable behind Cloudflare
// Access and requires no router — the app parses it on load and on
// `hashchange`.
export const parseComplaintHash = (hash: string): string | undefined => {
  const match = /^#\/complaints\/([A-Za-z0-9-]+)\/?$/.exec(hash)
  return match?.[1]
}

export const complaintHashFor = (id: string): string => `#/complaints/${id}`
