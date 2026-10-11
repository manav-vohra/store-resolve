// Client-side Web Push helpers (Stage A).
//
// All subscription traffic goes to the Access-authenticated same-origin API,
// so a subscription is always bound to the logged-in user server-side.

export type PushStatus =
  | 'unsupported'
  | 'denied'
  | 'subscribed'
  | 'unsubscribed'
  | 'prompt'

export function isPushSupported(): boolean {
  return (
    typeof window !== 'undefined' &&
    'serviceWorker' in navigator &&
    'PushManager' in window &&
    'Notification' in window
  )
}

/** True when launched from the Home Screen (standalone display mode). */
export function isStandalone(): boolean {
  if (typeof window === 'undefined') return false
  return (
    window.matchMedia('(display-mode: standalone)').matches ||
    (window.navigator as unknown as { standalone?: boolean }).standalone === true
  )
}

export async function getPushStatus(): Promise<PushStatus> {
  if (!isPushSupported()) return 'unsupported'
  if (Notification.permission === 'denied') return 'denied'
  try {
    const registration = await navigator.serviceWorker.ready
    const subscription = await registration.pushManager.getSubscription()
    if (subscription) return 'subscribed'
    return Notification.permission === 'granted' ? 'unsubscribed' : 'prompt'
  } catch {
    return 'unsupported'
  }
}

function urlBase64ToUint8Array(base64url: string): Uint8Array {
  const normalized = base64url.replace(/-/g, '+').replace(/_/g, '/')
  const padded = normalized + '='.repeat((4 - (normalized.length % 4)) % 4)
  const binary = atob(padded)
  const bytes = new Uint8Array(binary.length)
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i)
  return bytes
}

function bufferToBase64url(buffer: ArrayBuffer | null): string {
  if (!buffer) throw new Error('Missing subscription key material')
  const bytes = new Uint8Array(buffer)
  let binary = ''
  for (let i = 0; i < bytes.length; i++) binary += String.fromCharCode(bytes[i])
  return btoa(binary).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '')
}

export interface PushKeys {
  endpoint: string
  p256dh: string
  auth: string
}

/** Request permission, subscribe with the server VAPID key, return raw keys. */
export async function subscribeForPush(applicationServerKey: string): Promise<PushKeys> {
  if (!isPushSupported()) throw new Error('Push notifications are not supported on this device')
  const permission = await Notification.requestPermission()
  if (permission !== 'granted') throw new Error('Notification permission was not granted')
  const registration = await navigator.serviceWorker.ready
  const existing = await registration.pushManager.getSubscription()
  const subscription =
    existing ??
    (await registration.pushManager.subscribe({
      userVisibleOnly: true,
      applicationServerKey: urlBase64ToUint8Array(applicationServerKey),
    }))
  return {
    endpoint: subscription.endpoint,
    p256dh: bufferToBase64url(subscription.getKey('p256dh')),
    auth: bufferToBase64url(subscription.getKey('auth')),
  }
}

/** Remove the local push subscription (server revocation happens via API). */
export async function unsubscribeLocalPush(): Promise<string | undefined> {
  if (!isPushSupported()) return undefined
  const registration = await navigator.serviceWorker.ready
  const subscription = await registration.pushManager.getSubscription()
  const endpoint = subscription?.endpoint
  if (subscription) await subscription.unsubscribe()
  return endpoint
}

/** Ask the service worker to refresh the app badge (best effort). */
export function refreshAppBadge(count: number): void {
  if (!('serviceWorker' in navigator)) return
  navigator.serviceWorker.ready
    .then((registration) => {
      registration.active?.postMessage({ type: 'storeresolve.badge', count })
    })
    .catch(() => {})
}
