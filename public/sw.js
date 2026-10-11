/* StoreResolve service worker — Stage A.
 *
 * Responsibilities: display Web Push notifications, handle notification taps
 * (deep link into the app), and maintain the app badge where supported.
 * Dependency-free. Notification content is built server-side under the
 * privacy rule: no customer PII in title/body — case IDs and store names only.
 */

const DEFAULT_ICON = '/icons/icon-192.png'
const BADGE_ICON = '/icons/icon-192.png'

/**
 * @param {number} count outstanding actionable tasks; 0 clears the badge.
 */
async function updateBadge(count) {
  try {
    if (typeof count !== 'number' || !('setAppBadge' in navigator)) return
    if (count > 0) await navigator.setAppBadge(count)
    else if ('clearAppBadge' in navigator) await navigator.clearAppBadge()
  } catch {
    // Badging is a supplement, never the accountability mechanism.
  }
}

self.addEventListener('push', (event) => {
  let data = {}
  try {
    data = event.data ? event.data.json() : {}
  } catch {
    data = {}
  }
  const title = typeof data.title === 'string' && data.title ? data.title : 'StoreResolve'
  const body = typeof data.body === 'string' ? data.body : ''
  const url = typeof data.url === 'string' && data.url.startsWith('/') ? data.url : '/'
  const tag = typeof data.tag === 'string' ? data.tag : undefined

  event.waitUntil(
    (async () => {
      await self.registration.showNotification(title, {
        body,
        icon: DEFAULT_ICON,
        badge: BADGE_ICON,
        tag,
        data: { url },
        // renotify only for explicit re-alerts sharing a tag
        renotify: false,
      })
      if (typeof data.badgeCount === 'number') await updateBadge(data.badgeCount)
    })(),
  )
})

self.addEventListener('notificationclick', (event) => {
  event.notification.close()
  const rawUrl = event.notification.data && event.notification.data.url
  const url = typeof rawUrl === 'string' && rawUrl.startsWith('/') ? rawUrl : '/'
  event.waitUntil(
    (async () => {
      const windows = await self.clients.matchAll({ type: 'window', includeUncontrolled: true })
      for (const client of windows) {
        // Reuse an open app tab when one exists.
        if ('focus' in client) {
          await client.focus()
          if ('navigate' in client) {
            try {
              await client.navigate(url)
            } catch {
              // Cross-origin or uncontrolled client; fall through to openWindow.
            }
            return
          }
          return
        }
      }
      await self.clients.openWindow(url)
    })(),
  )
})

// Allow the page to ask the worker to refresh the badge (e.g. after the user
// opens the app and works their queue).
self.addEventListener('message', (event) => {
  const data = event.data || {}
  if (data.type === 'storeresolve.badge' && typeof data.count === 'number') {
    event.waitUntil(updateBadge(data.count))
  }
})
