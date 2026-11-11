import {
  ChevronRight,
  PanelLeftClose,
  PanelLeftOpen,
} from 'lucide-react'

type TopbarProps = {
  title: string
  isSidebarVisible: boolean
  onToggleSidebar: () => void
}

export default function Topbar({
  title,
  isSidebarVisible,
  onToggleSidebar,
}: TopbarProps) {
  return (
    <header className="topbar">
      <div className="topbar-left">
        <button
          type="button"
          className="sidebar-toggle"
          onClick={onToggleSidebar}
          aria-label={
            isSidebarVisible
              ? 'Sembunyikan sidebar'
              : 'Tampilkan sidebar'
          }
          title={
            isSidebarVisible
              ? 'Sembunyikan sidebar'
              : 'Tampilkan sidebar'
          }
        >
          {isSidebarVisible ? (
            <PanelLeftClose size={19} />
          ) : (
            <PanelLeftOpen size={19} />
          )}
        </button>

        <div className="breadcrumb">
          <span>Workspace</span>

          <ChevronRight size={15} />

          <strong>{title}</strong>
        </div>
      </div>

      <div className="topbar-right">
        <div className="mode-label">
          <span className="mode-dot" />
          <span>Mode pengembangan</span>
        </div>

        <div className="profile-avatar">G</div>

        <div className="profile-meta">
          <strong>Galih</strong>
          <span>Developer</span>
        </div>
      </div>
    </header>
  )
}