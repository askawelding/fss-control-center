export type AppTask = {
  id: number
  feature: string
  stage: string
  status: 'running' | 'completed' | 'failed'
  startedAt: number
  endedAt?: number
  message?: string
}

type TaskHandle = {
  update: (stage: string) => void
  complete: (message?: string) => void
  fail: (message: string) => void
}

let nextTaskId = 1
let currentTasks: AppTask[] = []
const listeners = new Set<() => void>()

function publish(next: AppTask[]) {
  currentTasks = next
  listeners.forEach((listener) => listener())
}

export function subscribeTaskStatus(listener: () => void) {
  listeners.add(listener)
  return () => listeners.delete(listener)
}

export function getTaskStatusSnapshot() {
  return currentTasks
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
    publish(currentTasks.map((item) => item.id === id
      ? { ...item, status: 'completed', endedAt: Date.now(), message }
      : item))
  }

  const fail = (message: string) => {
    const existing = currentTasks.find((item) => item.id === id)
    if (!existing || existing.status !== 'running') return
    publish(currentTasks.map((item) => item.id === id
      ? { ...item, status: 'failed', endedAt: Date.now(), message }
      : item))
  }

  return { update, complete, fail }
}

export function dismissAppTask(id: number) {
  if (currentTasks.some((item) => item.id === id && item.status !== 'running')) {
    publish(currentTasks.filter((item) => item.id !== id))
  }
}
