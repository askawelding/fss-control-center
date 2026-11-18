import {
  ChevronRight,
  PanelLeftClose,
  PanelLeftOpen,
} from 'lucide-react'

type TopbarProps = {
  title: string
  isSidebarVisible: boolean
  isBackendOnline: boolean
  onToggleSidebar: () => void
}

export default function Topbar({
  title,
  isSidebarVisible,
  isBackendOnline,
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
        <div className="status-indicator">
          <span 
            className={`status-dot ${isBackendOnline ? 'online' : 'offline'}`} 
            title={isBackendOnline ? 'Backend Terhubung' : 'Backend Terputus'}
          />
          <span className="status-text">
            {isBackendOnline ? 'Terhubung' : 'Terputus'}
          </span>
        </div>
      </div>
    </header>
  )
}