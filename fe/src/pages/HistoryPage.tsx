import { useState } from 'react'

import {
  AlertTriangle,
  CheckCircle2,
  Clock3,
  Eye,
  FileSpreadsheet,
  FilterX,
  Search,
  XCircle,
} from 'lucide-react'

type ProcessType = 'Build Data FSS' | 'Update Schedule DATA'
type ProcessStatus = 'Berhasil' | 'Peringatan' | 'Gagal'

type StatusFilter = 'Semua' | ProcessStatus
type TypeFilter = 'Semua' | ProcessType

type HistoryRecord = {
  id: string
  date: string
  type: ProcessType
  inputFiles: string
  outputFile: string
  status: ProcessStatus
  duration: string
  rowsProcessed: string
  message: string
  warnings: string[]
  outputLocation: string
}

// Seluruh data di bawah hanya untuk latihan frontend.
const mockHistory: HistoryRecord[] = [
  {
    id: 'JOB-0004',
    date: '28 Okt 2026, 10:15',
    type: 'Build Data FSS',
    inputFiles: '11 file Excel',
    outputFile: '20261028.xlsx',
    status: 'Berhasil',
    duration: '01:42',
    rowsProcessed: '12.450',
    message: 'Contoh proses Build Data FSS berhasil.',
    warnings: [],
    outputLocation: 'Contoh folder output lokal',
  },
  {
    id: 'JOB-0003',
    date: '27 Okt 2026, 15:20',
    type: 'Update Schedule DATA',
    inputFiles: '2 file Excel',
    outputFile: 'Schedule_UPDATED.xlsx',
    status: 'Peringatan',
    duration: '03:18',
    rowsProcessed: '8.320',
    message: 'Contoh proses selesai dengan peringatan.',
    warnings: [
      'Ada baris yang memerlukan pemeriksaan lebih lanjut.',
    ],
    outputLocation: 'Contoh folder output lokal',
  },
  {
    id: 'JOB-0002',
    date: '27 Okt 2026, 11:05',
    type: 'Build Data FSS',
    inputFiles: '8 file Excel',
    outputFile: '20261027.xlsx',
    status: 'Gagal',
    duration: '00:12',
    rowsProcessed: '—',
    message: 'Contoh validasi header gagal.',
    warnings: [
      'Header wajib tidak cocok dengan struktur yang diharapkan.',
    ],
    outputLocation: 'Belum ada output valid',
  },
  {
    id: 'JOB-0001',
    date: '26 Okt 2026, 14:30',
    type: 'Update Schedule DATA',
    inputFiles: '2 file Excel',
    outputFile: 'Schedule_261026.xlsx',
    status: 'Berhasil',
    duration: '02:54',
    rowsProcessed: '9.120',
    message: 'Contoh proses Update Schedule berhasil.',
    warnings: [],
    outputLocation: 'Contoh folder output lokal',
  },
]

export default function HistoryPage() {
  const [searchQuery, setSearchQuery] = useState('')
  const [statusFilter, setStatusFilter] =
    useState<StatusFilter>('Semua')
  const [typeFilter, setTypeFilter] =
    useState<TypeFilter>('Semua')
  const [selectedRecord, setSelectedRecord] =
    useState<HistoryRecord | null>(null)

  const filteredRecords = mockHistory.filter((record) => {
    const query = searchQuery.trim().toLowerCase()

    const matchesSearch =
      !query ||
      record.id.toLowerCase().includes(query) ||
      record.inputFiles.toLowerCase().includes(query) ||
      record.outputFile.toLowerCase().includes(query) ||
      record.type.toLowerCase().includes(query)

    const matchesStatus =
      statusFilter === 'Semua' ||
      record.status === statusFilter

    const matchesType =
      typeFilter === 'Semua' ||
      record.type === typeFilter

    return matchesSearch && matchesStatus && matchesType
  })

  function clearFilters() {
    setSearchQuery('')
    setStatusFilter('Semua')
    setTypeFilter('Semua')
    setSelectedRecord(null)
  }

  return (
    <main className="page-content history-page">
      <section className="page-intro">
        <div>
          <p className="eyebrow">WELDING DEPARTMENT</p>

          <h1>Riwayat Proses</h1>

          <p className="page-description">
            Periksa aktivitas, status, durasi, dan hasil
            pengolahan file sebelumnya.
          </p>
        </div>

        <div className="history-demo-badge">
          DATA SIMULASI
        </div>
      </section>

      <section className="history-notice">
        <AlertTriangle size={18} />

        <p>
          Riwayat yang terlihat saat ini adalah data latihan.
          Belum terhubung ke SQLite atau proses Excel sebenarnya.
        </p>
      </section>

      <section className="history-stats">
        <article className="history-stat-card">
          <div className="history-stat-icon neutral">
            <Clock3 size={19} />
          </div>

          <span>Total contoh proses</span>
          <strong>{mockHistory.length}</strong>
        </article>

        <article className="history-stat-card">
          <div className="history-stat-icon success">
            <CheckCircle2 size={19} />
          </div>

          <span>Berhasil</span>

          <strong>
            {mockHistory.filter(
              (record) => record.status === 'Berhasil',
            ).length}
          </strong>
        </article>

        <article className="history-stat-card">
          <div className="history-stat-icon warning">
            <AlertTriangle size={19} />
          </div>

          <span>Peringatan</span>

          <strong>
            {mockHistory.filter(
              (record) => record.status === 'Peringatan',
            ).length}
          </strong>
        </article>

        <article className="history-stat-card">
          <div className="history-stat-icon danger">
            <XCircle size={19} />
          </div>

          <span>Gagal</span>

          <strong>
            {mockHistory.filter(
              (record) => record.status === 'Gagal',
            ).length}
          </strong>
        </article>
      </section>

      <section className="history-panel">
        <div className="history-panel-heading">
          <div>
            <h2>Semua proses</h2>

            <p>
              Cari file atau saring riwayat berdasarkan jenis
              dan statusnya.
            </p>
          </div>
        </div>

        <div className="history-filters">
          <label className="history-search">
            <Search size={18} />

            <input
              type="search"
              placeholder="Cari file atau ID proses..."
              value={searchQuery}
              onChange={(event) =>
                setSearchQuery(event.currentTarget.value)
              }
            />
          </label>

          <select
            className="history-select"
            aria-label="Filter jenis proses"
            value={typeFilter}
            onChange={(event) =>
              setTypeFilter(event.currentTarget.value as TypeFilter)
            }
          >
            <option value="Semua">Semua jenis proses</option>
            <option value="Build Data FSS">Build Data FSS</option>
            <option value="Update Schedule DATA">
              Update Schedule DATA
            </option>
          </select>

          <select
            className="history-select"
            aria-label="Filter status"
            value={statusFilter}
            onChange={(event) =>
              setStatusFilter(
                event.currentTarget.value as StatusFilter,
              )
            }
          >
            <option value="Semua">Semua status</option>
            <option value="Berhasil">Berhasil</option>
            <option value="Peringatan">Peringatan</option>
            <option value="Gagal">Gagal</option>
          </select>

          <button
            type="button"
            className="history-reset-button"
            onClick={clearFilters}
            title="Reset filter"
          >
            <FilterX size={17} />
            Reset
          </button>
        </div>

        <div className="history-result-count">
          Menampilkan {filteredRecords.length} dari{' '}
          {mockHistory.length} contoh proses
        </div>

        <div className="history-table-wrap">
          <table className="history-table">
            <thead>
              <tr>
                <th>Waktu / ID</th>
                <th>Jenis proses</th>
                <th>File input</th>
                <th>Status</th>
                <th>Durasi</th>
                <th>Aksi</th>
              </tr>
            </thead>

            <tbody>
              {filteredRecords.map((record) => (
                <tr key={record.id}>
                  <td>
                    <div className="history-date">
                      <strong>{record.date}</strong>
                      <span>{record.id}</span>
                    </div>
                  </td>

                  <td>
                    <div className="history-process-type">
                      <FileSpreadsheet size={17} />

                      <span>{record.type}</span>
                    </div>
                  </td>

                  <td>{record.inputFiles}</td>

                  <td>
                    <span
                      className={`history-status history-status-${record.status.toLowerCase()}`}
                    >
                      {record.status}
                    </span>
                  </td>

                  <td>{record.duration}</td>

                  <td>
                    <button
                      type="button"
                      className="history-detail-button"
                      onClick={() =>
                        setSelectedRecord(
                          selectedRecord?.id === record.id
                            ? null
                            : record,
                        )
                      }
                    >
                      <Eye size={15} />

                      {selectedRecord?.id === record.id
                        ? 'Tutup'
                        : 'Detail'}
                    </button>
                  </td>
                </tr>
              ))}

              {filteredRecords.length === 0 && (
                <tr>
                  <td colSpan={6}>
                    <div className="history-empty">
                      <Search size={25} />

                      <strong>Riwayat tidak ditemukan</strong>

                      <span>
                        Coba gunakan kata kunci atau filter lain.
                      </span>

                      <button
                        type="button"
                        className="history-empty-reset"
                        onClick={clearFilters}
                      >
                        Hapus filter
                      </button>
                    </div>
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </section>

      {selectedRecord && (
        <section className="history-detail-panel">
          <div className="history-detail-heading">
            <div>
              <p className="eyebrow">DETAIL PROSES</p>

              <h2>{selectedRecord.id}</h2>

              <p>{selectedRecord.message}</p>
            </div>

            <span
              className={`history-status history-status-${selectedRecord.status.toLowerCase()}`}
            >
              {selectedRecord.status}
            </span>
          </div>

          <div className="history-detail-grid">
            <div>
              <span>Jenis proses</span>
              <strong>{selectedRecord.type}</strong>
            </div>

            <div>
              <span>Waktu proses</span>
              <strong>{selectedRecord.date}</strong>
            </div>

            <div>
              <span>Durasi</span>
              <strong>{selectedRecord.duration}</strong>
            </div>

            <div>
              <span>Jumlah baris</span>
              <strong>{selectedRecord.rowsProcessed}</strong>
            </div>

            <div>
              <span>File input</span>
              <strong>{selectedRecord.inputFiles}</strong>
            </div>

            <div>
              <span>File output</span>
              <strong>{selectedRecord.outputFile}</strong>
            </div>

            <div>
              <span>Lokasi output</span>
              <strong>{selectedRecord.outputLocation}</strong>
            </div>
          </div>

          <div className="history-warning-list">
            <h3>Catatan dan peringatan</h3>

            {selectedRecord.warnings.length === 0 ? (
              <p className="history-no-warning">
                Tidak ada peringatan pada contoh riwayat ini.
              </p>
            ) : (
              selectedRecord.warnings.map((warning) => (
                <div className="history-warning-item" key={warning}>
                  <AlertTriangle size={16} />
                  <span>{warning}</span>
                </div>
              ))
            )}
          </div>
        </section>
      )}
    </main>
  )
}