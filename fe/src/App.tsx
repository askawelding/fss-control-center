import { useEffect, useState, useSyncExternalStore } from 'react'

import {
  ArrowRight,
  CheckCircle2,
  Clock3,
  FileSpreadsheet,
  Files,
  FolderOpen,
  RefreshCw,
  Server,
} from 'lucide-react'

import Sidebar, { type PageKey } from './components/Sidebar'
import Topbar from './components/Topbar'
import BuildFSSPage from './pages/BuildFSSPage'
import SummaryCard, {
  type SummaryCardProps,
} from './components/SummaryCard'
import UpdateSCHDataPage from './pages/UpdateSCHDataPage'
import HistoryPage from './pages/HistoryPage'
import SettingsPage from './pages/SettingsPage'
import AboutPage from './pages/AboutPage'
import TanpaRumusPage from './pages/TanpaRumusPage'
import { getBackendHealth } from './services/api'
import {
  dismissAppTask,
  getTaskStatusSnapshot,
  subscribeTaskStatus,
} from './services/taskStatus'


import './App.css'

// Data untuk kartu status Dashboard.
const dashboardStats: SummaryCardProps[] = [
  {
    tag: 'MASTER',
    label: 'Master FSS terakhir',
    value: 'Master Belum tersedia',
    note: 'File master belum dibuat.',
    icon: FileSpreadsheet,
    variant: 'orange',
  },
  {
    tag: 'AKTIVITAS',
    label: 'Proses terakhir',
    value: 'Belum ada riwayat',
    note: 'Belum ada proses yang tercatat.',
    icon: Clock3,
    variant: 'purple',
  },
]

// Informasi judul dan deskripsi setiap halaman.
const pageDetails: Record<
  PageKey,
  { title: string; description: string }
> = {
  dashboard: {
    title: 'Dashboard',
    description: 'Ringkasan aktivitas FSS Control Center.',
  },
  build: {
    title: 'Build Data FSS',
    description: 'Gabungkan file FSS menjadi satu master Excel.',
  },
  update: {
    title: 'Update Schedule DATA',
    description: 'Perbarui data FSS pada workbook schedule.',
  },
  'tanpa-rumus': {
    title: 'Tanpa Rumus',
    description: 'Finalisasi workbook delivery dan kelompok receiving.',
  },
  history: {
    title: 'Riwayat Proses',
    description: 'Lihat catatan proses yang sudah dijalankan.',
  },
  settings: {
    title: 'Pengaturan',
    description: 'Atur preferensi aplikasi lokal.',
  },
  about: {
    title: 'Tentang',
    description: 'Informasi tentang aplikasi dan pengembang.',
  },
}

function formatTaskDuration(totalSeconds: number): string {
  const seconds = Math.max(0, Math.floor(totalSeconds))
  if (seconds < 60) return `${seconds} detik`
  const minutes = Math.floor(seconds / 60)
  const remainingSeconds = seconds % 60
  return `${minutes} menit ${remainingSeconds} detik`
}

// Halaman Dashboard.
function DashboardPage({
  onNavigate,
  isBackendOnline,
}: {
  onNavigate: (page: PageKey) => void
  isBackendOnline: boolean
}) {
  return (
    <main className="page-content">
      <section className="page-intro">
        <div>
          <p className="eyebrow">WELDING DEPARTMENT</p>

          <h1>Dashboard</h1>

          <p className="page-description">
            Pantau kesiapan aplikasi dan mulai pekerjaan FSS dari sini.
          </p>
        </div>

        <div className="demo-badge">
          <span className="demo-dot" />
          Mode pengembangan
        </div>
      </section>

      <section className="welcome-banner">
        <div className="welcome-content">
          <p className="banner-eyebrow">WORKSPACE OVERVIEW</p>

          <h2>Semua proses FSS, dalam satu tempat.</h2>

          <p>
            Gabungkan data customer dan perbarui schedule Excel
            tanpa copy-paste manual.
          </p>
        </div>

        <div className="banner-mark">
          <FileSpreadsheet size={44} strokeWidth={1.5} />
        </div>
      </section>

      {/* Kartu dibuat otomatis berdasarkan data dashboardStats. */}
      <section
        className="status-grid"
        aria-label="Status aplikasi"
      >
        {dashboardStats.map((stat) => (
          <SummaryCard key={stat.tag} {...stat} />
        ))}

        <SummaryCard
          tag="SISTEM"
          label="Backend lokal"
            value={isBackendOnline ? 'Backend aktif' : 'Backend belum terhubung'}
            note={
            isBackendOnline
              ? 'API merespons dengan baik.'
              : 'Periksa apakah server backend berjalan.'
          }
          icon={Server}
          variant={isBackendOnline ? 'orange' : 'gray'}
        />        
      </section>

      <section className="work-section">
        <div className="section-heading">
          <div>
            <h2>Mulai pekerjaan</h2>

            <p>Pilih proses yang ingin kamu kerjakan.</p>
          </div>
        </div>

        <div className="action-grid">
          <button
            type="button"
            className="action-card"
            onClick={() => onNavigate('build')}
          >
            <div className="action-icon icon-orange">
              <Files size={23} />
            </div>

            <div className="action-copy">
              <span className="action-kicker">PROSES 01</span>

              <h3>Build Data FSS</h3>

              <p>
                Gabungkan file Firm dan Forecast menjadi
                satu master Excel.
              </p>

              <span className="action-link">
                Buka fitur <ArrowRight size={16} />
              </span>
            </div>
          </button>

          <button
            type="button"
            className="action-card"
            onClick={() => onNavigate('update')}
          >
            <div className="action-icon icon-purple">
              <RefreshCw size={23} />
            </div>

            <div className="action-copy">
              <span className="action-kicker">PROSES 02</span>

              <h3>Update Schedule DATA</h3>

              <p>
                Perbarui data FSS dan tinjau perubahan
                pada workbook schedule.
              </p>

              <span className="action-link">
                Buka fitur <ArrowRight size={16} />
              </span>
            </div>
          </button>
        </div>
      </section>

      <section className="activity-panel">
        <div className="section-heading">
          <div>
            <h2>Aktivitas terbaru</h2>

            <p>Ringkasan pekerjaan yang sudah dilakukan.</p>
          </div>

          <button
            type="button"
            className="text-button"
            onClick={() => onNavigate('history')}
          >
            Lihat riwayat <ArrowRight size={15} />
          </button>
        </div>

        <div className="empty-state">
          <div className="empty-icon">
            <FolderOpen size={26} />
          </div>

          <h3>Belum ada riwayat proses</h3>

          <p>
            Aktivitas akan muncul di sini setelah
            proses dijalankan dan backend terhubung.
          </p>
        </div>
      </section>
    </main>
  )
}

// Halaman sementara
function PlaceholderPage({
  page,
  onNavigate,
}: {
  page: Exclude<PageKey, 'dashboard'>
  onNavigate: (page: PageKey) => void
}) {
  const details = pageDetails[page]

  return (
    <main className="page-content">
      <section className="page-intro">
        <div>
          <p className="eyebrow">WELDING DEPARTMENT</p>

          <h1>{details.title}</h1>

          <p className="page-description">
            {details.description}
          </p>
        </div>
      </section>

      <section className="placeholder-panel">
        <div className="placeholder-icon">
          <FolderOpen size={32} />
        </div>

        <h2>Halaman ini belum dibuat</h2>

        <p>
          Kita akan membangunnya sedikit demi sedikit
          setelah kerangka aplikasi selesai.
        </p>

        <button
          type="button"
          className="primary-button"
          onClick={() => onNavigate('dashboard')}
        >
          Kembali ke Dashboard
          <ArrowRight size={17} />
        </button>
      </section>
    </main>
  )
}

// Komponen utama
export default function App() {
  const [activePage, setActivePage] =
    useState<PageKey>('dashboard')

  // Halaman yang pernah dibuka tetap mounted ketika berpindah sidebar.
  // Dengan demikian state form, file terpilih, dan request fetch yang sedang
  // berjalan tidak hilang saat pengguna membuka fitur lain.
  const [visitedPages, setVisitedPages] = useState<Set<PageKey>>(() => new Set(['dashboard']))
  
  const [isSidebarVisible, setIsSidebarVisible] =
    useState(true)

  const [isBackendOnline, setIsBackendOnline] = useState(false)
  const tasks = useSyncExternalStore(subscribeTaskStatus, getTaskStatusSnapshot, getTaskStatusSnapshot)
  const hasRunningTasks = tasks.some((item) => item.status === 'running')
  const [clockNow, setClockNow] = useState(Date.now())

  function navigateToPage(page: PageKey) {
    setActivePage(page)
    setVisitedPages((current) => {
      if (current.has(page)) return current
      const next = new Set(current)
      next.add(page)
      return next
    })
  }

  useEffect(() => {
    if (!hasRunningTasks) return
    const interval = window.setInterval(() => setClockNow(Date.now()), 500)
    return () => window.clearInterval(interval)
  }, [hasRunningTasks])

    useEffect(() => {
      let isMounted = true

  const checkBackend = async () => {
    const health = await getBackendHealth()

    if (isMounted) {
      setIsBackendOnline(health?.status === 'ok')
    }
  }

  void checkBackend()

  const intervalId = window.setInterval(() => {
    void checkBackend()
    }, 5000)

  return () => {
    isMounted = false
    window.clearInterval(intervalId)
    }
  },  [])

  const currentTitle = pageDetails[activePage].title

  return (
    <div className="app-shell">
      {isSidebarVisible && (
      <Sidebar
        activePage={activePage}
        onNavigate={navigateToPage}
        isBackendOnline={isBackendOnline}
      />
      )}

      <div className="main-area">
        <Topbar 
          title={currentTitle}
          isSidebarVisible={isSidebarVisible}
          onToggleSidebar={() =>
            setIsSidebarVisible((visible) => !visible)
          }
        />

        {tasks.length > 0 && (
          <div className="global-task-list">
            {tasks.map((task) => (
              <section
                key={task.id}
                className={`global-task-banner ${task.status}`}
                role="status"
                aria-live="polite"
              >
                <div className="global-task-indicator">
                  {task.status === 'running' ? (
                    <RefreshCw className="spin" size={18} />
                  ) : task.status === 'completed' ? (
                    <CheckCircle2 size={18} />
                  ) : (
                    <Server size={18} />
                  )}
                </div>
                <div className="global-task-copy">
                  <strong>{task.feature}: {task.status === 'running' ? task.stage : task.message}</strong>
                  <span>
                    {task.status === 'running'
                      ? `Sedang berjalan · ${formatTaskDuration((clockNow - task.startedAt) / 1000)}`
                      : `Durasi ${formatTaskDuration(((task.endedAt ?? clockNow) - task.startedAt) / 1000)}`}
                  </span>
                </div>
                {task.status !== 'running' && (
                  <button type="button" aria-label="Tutup status proses" onClick={() => dismissAppTask(task.id)}>×</button>
                )}
              </section>
            ))}
          </div>
        )}

        {visitedPages.has('dashboard') && <div hidden={activePage !== 'dashboard'}><DashboardPage onNavigate={navigateToPage} isBackendOnline={isBackendOnline} /></div>}
        {visitedPages.has('build') && <div hidden={activePage !== 'build'}><BuildFSSPage /></div>}
        {visitedPages.has('update') && <div hidden={activePage !== 'update'}><UpdateSCHDataPage /></div>}
        {visitedPages.has('tanpa-rumus') && <div hidden={activePage !== 'tanpa-rumus'}><TanpaRumusPage /></div>}
        {visitedPages.has('history') && <div hidden={activePage !== 'history'}><HistoryPage /></div>}
        {visitedPages.has('settings') && <div hidden={activePage !== 'settings'}><SettingsPage /></div>}
        {visitedPages.has('about') && <div hidden={activePage !== 'about'}><AboutPage onNavigate={navigateToPage} /></div>}
        {(['dashboard', 'build', 'update', 'tanpa-rumus', 'history', 'settings', 'about'] as PageKey[]).includes(activePage) ? null : (
          <PlaceholderPage page={activePage as Exclude<PageKey, 'dashboard'>} onNavigate={navigateToPage} />
        )}
      </div>
    </div>
  )
}