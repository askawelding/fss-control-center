export type AppTask = {
  id: number
  feature: string
  stage: string
  status: 'running' | 'completed' | 'failed'
  startedAt: number
  endedAt?: number
  message?: string
}

const HISTORY_STORAGE_KEY = 'fss-control-center:task-history'
const MAX_HISTORY_ITEMS = 100

type TaskHandle = {
  update: (stage: string) => void
  complete: (message?: string) => void
  fail: (message: string) => void
}

let nextTaskId = 1
let currentTasks: AppTask[] = []
let taskHistory = readTaskHistory()
const listeners = new Set<() => void>()

function readTaskHistory(): AppTask[] {
  try {
    const stored = window.localStorage.getItem(HISTORY_STORAGE_KEY)
    if (!stored) return []
    const parsed: unknown = JSON.parse(stored)
    if (!Array.isArray(parsed)) throw new Error('Format riwayat proses tidak valid.')
    return parsed.filter((item): item is AppTask =>
      item !== null &&
      typeof item === 'object' &&
      typeof item.id === 'number' &&
      typeof item.feature === 'string' &&
      typeof item.stage === 'string' &&
      (item.status === 'completed' || item.status === 'failed') &&
      typeof item.startedAt === 'number' &&
      typeof item.endedAt === 'number',
    ).slice(0, MAX_HISTORY_ITEMS)
  } catch (error) {
    console.error('Gagal membaca riwayat proses lokal.', error)
    return []
  }
}

function publish(next: AppTask[]) {
  currentTasks = next
  listeners.forEach((listener) => listener())
}

function publishHistory(task: AppTask) {
  taskHistory = [task, ...taskHistory.filter((item) => item.id !== task.id)]
    .slice(0, MAX_HISTORY_ITEMS)
  try {
    window.localStorage.setItem(HISTORY_STORAGE_KEY, JSON.stringify(taskHistory))
  } catch (error) {
    console.error('Gagal menyimpan riwayat proses lokal.', error)
  }
  listeners.forEach((listener) => listener())
}

export function subscribeTaskStatus(listener: () => void) {
  listeners.add(listener)
  return () => listeners.delete(listener)
}

export function getTaskStatusSnapshot() {
  return currentTasks
}

export function getTaskHistorySnapshot() {
  return taskHistory
}

export function startAppTask(feature: string, stage: string): TaskHandle {
  const id = nextTaskId++
  const task: AppTask = {
    id,
    feature,
    stage,
    status: 'running',
    startedAt: Date.now(),
  }
  // Simpan semua pekerjaan aktif dan tiga status terakhir agar banner tidak
  // tertimpa ketika pengguna memulai pekerjaan lain dari menu berbeda.
  const active = currentTasks.filter((item) => item.status === 'running')
  const recentCompleted = currentTasks.filter((item) => item.status !== 'running').slice(-3)
  publish([...active, ...recentCompleted, task])

  const update = (nextStage: string) => {
    const existing = currentTasks.find((item) => item.id === id)
    if (!existing || existing.status !== 'running') return
    publish(currentTasks.map((item) => item.id === id ? { ...item, stage: nextStage } : item))
  }

  const complete = (message = 'Proses selesai.') => {
    const existing = currentTasks.find((item) => item.id === id)
    if (!existing || existing.status !== 'running') return
    const completedTask = { ...existing, status: 'completed' as const, endedAt: Date.now(), message }
    publish(currentTasks.map((item) => item.id === id
      ? completedTask
      : item))
    publishHistory(completedTask)
  }

  const fail = (message: string) => {
    const existing = currentTasks.find((item) => item.id === id)
    if (!existing || existing.status !== 'running') return
    const failedTask = { ...existing, status: 'failed' as const, endedAt: Date.now(), message }
    publish(currentTasks.map((item) => item.id === id
      ? failedTask
      : item))
    publishHistory(failedTask)
  }

  return { update, complete, fail }
}

export function dismissAppTask(id: number) {
  if (currentTasks.some((item) => item.id === id && item.status !== 'running')) {
    publish(currentTasks.filter((item) => item.id !== id))
  }
}
