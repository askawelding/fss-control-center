import {
  Activity,
  Files,
  History,
  Info,
  LayoutDashboard,
  RefreshCw,
  Rocket,
  Settings,
  type LucideIcon,
} from 'lucide-react'
import wIconWhite from '../assets/w_icon_white.svg'

export type PageKey =
  | 'dashboard'
  | 'build'
  | 'update'
  | 'tanpa-rumus'
  | 'history'
  | 'settings'
  | 'about'

type MenuItem = {
  id: PageKey
  label: string
  icon: LucideIcon
}

type SidebarProps = {
  activePage: PageKey
  onNavigate: (page: PageKey) => void
  isBackendOnline: boolean
}

const menuItems: MenuItem[] = [
  {
    id: 'dashboard',
    label: 'Dashboard',
    icon: LayoutDashboard,
  },
  {
    id: 'build',
    label: 'Build Data FSS',
    icon: Files,
  },
  {
    id: 'update',
    label: 'Update Schedule DATA',
    icon: RefreshCw,
  },
  {
    id: 'tanpa-rumus',
    label: 'Tanpa Rumus',
    icon: Rocket,
  },
  {
    id: 'history',
    label: 'Riwayat Proses',
    icon: History,
  },
  {
    id: 'settings',
    label: 'Pengaturan',
    icon: Settings,
  },
  {
    id: 'about',
    label: 'Tentang Aplikasi',
    icon: Info,
  },
]

export default function Sidebar({
  activePage,
  onNavigate,
  isBackendOnline,
}: SidebarProps) {
  return (
    <aside className="sidebar">
      <div className="brand">
        <div className="brand-logo">
          <img src={wIconWhite} alt="FSS Control Center" />
        </div>

        <div className="brand-text">
          <strong>FSS Control Center</strong>
          <span>Welding Department</span>
        </div>
      </div>

      <p className="nav-caption">MENU UTAMA</p>

      <nav className="sidebar-nav" aria-label="Menu utama">
        {menuItems.map((item) => {
          const Icon = item.icon

          return (
            <button
              type="button"
              key={item.id}
              className={`nav-item ${
                activePage === item.id ? 'active' : ''
              }`}
              onClick={() => onNavigate(item.id)}
            >
              <Icon size={19} strokeWidth={1.9} />

              <span className="nav-item-label">
                {item.label}
              </span>
            </button>
          )
        })}
      </nav>

      <div className="sidebar-bottom">
        <div className="sidebar-status-icon">
          <Activity size={17} />
        </div>

      <div className="sidebar-status-copy">
        <strong>
          {isBackendOnline ? 'Backend aktif' : 'Backend belum terhubung'}
        </strong>
        <span>
          {isBackendOnline ? 'API merespons dengan baik' : 'Periksa server lokal'}
        </span>
      </div>

        <span className="sidebar-status-dot" />
      </div>

      <div className="sidebar-version">
        FSS CONTROL CENTER <span>v0.1</span>
      </div>
    </aside>
  )
}