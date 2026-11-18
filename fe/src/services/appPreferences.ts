import type { PageKey } from '../components/Sidebar'

export type AppPreferences = {
  defaultPage: PageKey
}

const STORAGE_KEY = 'fss-control-center:preferences'
const pageKeys: PageKey[] = [
  'dashboard',
  'build',
  'update',
  'tanpa-rumus',
  'history',
  'settings',
  'about',
]

export function getAppPreferences(): AppPreferences {
  try {
    const stored = window.localStorage.getItem(STORAGE_KEY)
    if (!stored) return { defaultPage: 'dashboard' }

    const parsed: unknown = JSON.parse(stored)
    if (
      parsed !== null &&
      typeof parsed === 'object' &&
      'defaultPage' in parsed &&
      pageKeys.includes(parsed.defaultPage as PageKey)
    ) {
      return { defaultPage: parsed.defaultPage as PageKey }
    }

    throw new Error('Format preferensi aplikasi tidak valid.')
  } catch (error) {
    console.error('Gagal membaca preferensi aplikasi.', error)
    return { defaultPage: 'dashboard' }
  }
}

export function saveAppPreferences(preferences: AppPreferences): void {
  window.localStorage.setItem(STORAGE_KEY, JSON.stringify(preferences))
}
