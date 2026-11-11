import { useState, type ChangeEvent } from 'react'
import {
  AlertCircle,
  CheckCircle2,
  FileSpreadsheet,
  FileUp,
  RefreshCw,
  ShieldCheck,
  Eye,
  AlertTriangle,
  Download,
  X,
} from 'lucide-react'
import {
  previewSCHReconciliation,
  updateSCHWorkbooks,
  validateSCHFiles,
  type SCHReconciliationFile,
  type SCHReconciliationResponse,
  type SCHValidationResponse,
} from '../services/UpdateSchApi'
import { startAppTask } from '../services/taskStatus'
import './UpdateSCHDataPage.css'

const MAX_SCHEDULE_FILES = 20

const COUNT_LABELS = [
  'Firm lama dipertahankan',
  'Firm terbaru sudah ada (tidak ditambahkan)',
  'Firm dari master masuk ke bagian Firm',
  'Forecast lama akan diganti',
  'Forecast terbaru akan ditambahkan',
]

function formatFileSize(bytes: number) {
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`
}

function formatNumber(value?: number) {
  return new Intl.NumberFormat('id-ID').format(value ?? 0)
}

function statusLabel(file: SCHReconciliationFile) {
  if (file.status === 'blocked') return 'Diblokir'
  if (file.status === 'review') return 'Perlu diperiksa'
  return 'Siap'
}

export default function UpdateSCHDataPage() {
  const [masterFile, setMasterFile] = useState<File | null>(null)
  const [files, setFiles] = useState<File[]>([])
  const [isReviewing, setIsReviewing] = useState(false)
  const [reviewStage, setReviewStage] = useState<'idle' | 'validating' | 'reconciling' | 'complete' | 'failed'>('idle')
  const [isUpdating, setIsUpdating] = useState(false)
  const [error, setError] = useState('')
  const [success, setSuccess] = useState('')
  const [validation, setValidation] = useState<SCHValidationResponse | null>(null)
  const [preview, setPreview] = useState<SCHReconciliationResponse | null>(null)

  const selectedCount = files.length

  function clearProcessResults() {
    setValidation(null)
    setPreview(null)
    setError('')
    setSuccess('')
    setReviewStage('idle')
  }

  function handleMasterChange(file?: File) {
    setMasterFile(file ?? null)
    clearProcessResults()
  }

  function handleScheduleFilesChange(event: ChangeEvent<HTMLInputElement>) {
    const selected = Array.from(event.currentTarget.files ?? [])
    event.currentTarget.value = ''

    if (selected.length > MAX_SCHEDULE_FILES) {
      setError(`Maksimal ${MAX_SCHEDULE_FILES} workbook dapat dipilih sekaligus.`)
      return
    }
    setFiles(selected)
    clearProcessResults()
  }

  function removeFile(indexToRemove: number) {
    setFiles((current) => current.filter((_, index) => index !== indexToRemove))
    clearProcessResults()
  }

  async function handleReview() {
    if (!masterFile) {
      setError('Pilih Data FSS master terlebih dahulu.')
      return
    }
    if (selectedCount < 1 || selectedCount > MAX_SCHEDULE_FILES) {
      setError(`Pilih 1 sampai ${MAX_SCHEDULE_FILES} workbook schedule.`)
      return
    }

    setIsReviewing(true)
    setReviewStage('validating')
    setError('')
    setSuccess('')
    setValidation(null)
    setPreview(null)
    const task = startAppTask('Update Schedule', 'Memvalidasi struktur workbook schedule...')

    try {
      // Tahap 1: validasi struktur setiap workbook lewat backend.
      const structure = await validateSCHFiles(files)
      setValidation(structure)

      if (!structure.valid || structure.total_files !== selectedCount) {
        setReviewStage('failed')
        setError(
          `Validasi struktur selesai: ${structure.valid_files} dari ${selectedCount} workbook valid. ` +
          'Periksa pesan pada kartu file yang gagal; rekonsiliasi belum dijalankan.'
        )
        task.complete('Validasi struktur selesai; periksa file yang belum valid.')
        return
      }

      // Tahap 2: cocokkan identitas order, konflik, dan status Firm/Forecast.
      setReviewStage('reconciling')
      task.update('Mencocokkan identitas order dengan Data FSS master...')
      const reconciliation = await previewSCHReconciliation(masterFile, files)
      setPreview(reconciliation)

      if (reconciliation.total_files !== selectedCount || reconciliation.valid_files !== selectedCount) {
        setReviewStage('failed')
        setError(
          `Rekonsiliasi selesai: ${reconciliation.valid_files} dari ${selectedCount} workbook berhasil dianalisis. ` +
          'Periksa pesan pada kartu file yang gagal.'
        )
        task.complete('Rekonsiliasi selesai; periksa workbook yang belum berhasil dianalisis.')
        return
      }

      if (reconciliation.master.ambiguous_rows > 0) {
        setError(
          `Master Data FSS memiliki ${formatNumber(reconciliation.master.ambiguous_rows)} baris ambigu. ` +
          'Update belum dapat dijalankan.'
        )
        task.complete('Rekonsiliasi selesai; Data FSS master memiliki baris ambigu.')
        return
      }

      const readyCount = reconciliation.files.filter((file) =>
        file.valid && file.safe_to_update && file.status === 'ready' &&
        file.reconciliation?.blocking_count === 0 && file.reconciliation?.review_count === 0
      ).length

      setReviewStage('complete')
      if (readyCount === selectedCount) {
        setSuccess(`Tinjauan backend selesai. ${selectedCount} workbook lolos validasi dan rekonsiliasi. Update siap dijalankan.`)
      } else {
        setError(`Tinjauan backend selesai. ${readyCount} dari ${selectedCount} workbook siap; periksa konflik atau temuan pada hasil preview.`)
      }
      task.complete('Validasi dan rekonsiliasi Update Schedule selesai.')
    } catch (err) {
      setReviewStage('failed')
      const message = err instanceof Error ? err.message : 'Tinjauan pilihan gagal menghubungi backend.'
      setError(message === 'Failed to fetch'
        ? 'Tidak dapat menghubungi backend. Pastikan terminal backend aktif di http://127.0.0.1:8000 lalu coba Tinjau Pilihan lagi.'
        : message)
      task.fail(message)
    } finally {
      setIsReviewing(false)
    }
  }

  const canUpdate = Boolean(
    preview &&
    preview.total_files === selectedCount &&
    preview.valid_files === selectedCount &&
    preview.files.length === selectedCount &&
    preview.master.ambiguous_rows === 0 &&
    preview.files.every((file) =>
      file.valid &&
      file.safe_to_update &&
      file.status === 'ready' &&
      file.reconciliation &&
      file.reconciliation.blocking_count === 0 &&
      file.reconciliation.review_count === 0,
    ),
  )

  async function handleUpdate() {
    if (!masterFile || !canUpdate || selectedCount < 1) {
      setError('Preview belum aman. Jalankan Tinjau Pilihan ulang sebelum update.')
      return
    }
    const accepted = window.confirm(
      `Proses ini akan membuat ${selectedCount} file Excel hasil update sebagai salinan baru dalam satu ZIP.\n\n` +
      'File sumber tidak ditimpa. Hanya workbook yang dipilih yang akan diproses. Pastikan Microsoft Excel desktop sudah terpasang dan tutup workbook sumber yang sedang terbuka.\n\nLanjutkan proses?',
    )
    if (!accepted) return

    setIsUpdating(true)
    setError('')
    setSuccess('')
    const task = startAppTask('Update Schedule', `Memproses ${selectedCount} workbook schedule...`)
    try {
      const blob = await updateSCHWorkbooks(masterFile, files)
      const url = URL.createObjectURL(blob)
      const link = document.createElement('a')
      link.href = url
      link.download = 'FSS_Control_Center_Update_SCH.zip'
      document.body.appendChild(link)
      link.click()
      link.remove()
      window.setTimeout(() => URL.revokeObjectURL(url), 1000)
      setSuccess(`Update berhasil. ZIP berisi ${selectedCount} workbook hasil dan manifest validasi. File sumber tidak ditimpa.`)
      task.complete(`Update selesai. ${selectedCount} workbook hasil sudah diunduh.`)
    } catch (err) {
      const message = err instanceof Error ? err.message : 'Proses Update SCH gagal dijalankan.'
      setError(message)
      task.fail(message)
    } finally {
      setIsUpdating(false)
    }
  }

  return (
    <div className="update-sch-page">
      <header className="update-sch-header">
        <div>
          <span className="update-sch-eyebrow">PENGOLAHAN WORKBOOK</span>
          <h1>Update SCH Data</h1>
          <p>Bandingkan Data FSS master dengan satu atau lebih workbook schedule yang kamu pilih.</p>
        </div>
        <div className="update-sch-header-icon">
          <FileSpreadsheet size={26} />
        </div>
      </header>

      <section className="update-sch-info">
        <ShieldCheck size={19} />
        <div>
          <strong>{reviewStage === 'validating' ? 'Backend sedang memvalidasi struktur workbook' : reviewStage === 'reconciling' ? 'Struktur valid — backend sedang memeriksa order dan konflik' : reviewStage === 'complete' ? 'Tinjauan backend selesai' : reviewStage === 'failed' ? 'Tinjauan gagal' : 'Menunggu Tinjau Pilihan'}</strong>
          <span aria-live="polite">
            {reviewStage === 'validating'
              ? 'Memeriksa worksheet DATE, header 12 field, rentang data, dan formula helper.'
              : reviewStage === 'reconciling'
                ? 'Membandingkan identitas order master dengan workbook, memeriksa konflik, dan menghitung perubahan Firm/Forecast.'
                : reviewStage === 'complete'
                  ? 'Backend sudah menyelesaikan validasi dan rekonsiliasi. Jika seluruh file berstatus Siap, tekan Proses Update & Download untuk membuat ZIP.'
                  : reviewStage === 'failed'
                    ? 'Periksa pesan error di bawah ini. Belum ada workbook yang ditulis.'
                    : 'Klik Tinjau Pilihan untuk menjalankan validasi dan rekonsiliasi backend. Tahap ini hanya membaca file dan belum membuat download.'}
          </span>
        </div>
      </section>

      <section className="update-sch-section">
        <div className="update-sch-section-heading">
          <div>
            <h2>1. Pilih Data FSS master</h2>
            <p>Gunakan file Excel hasil Build Data FSS terbaru, misalnya 20261018.xlsx.</p>
          </div>
        </div>

        <label className={`update-sch-file-card update-sch-master-card ${masterFile ? 'is-selected' : ''}`}>
          <input
            type="file"
            accept=".xlsx,.xlsm"
            onChange={(event) => handleMasterChange(event.target.files?.[0])}
          />
          <div className="update-sch-file-icon">
            {masterFile ? <CheckCircle2 size={21} /> : <FileSpreadsheet size={21} />}
          </div>
          <div className="update-sch-file-copy">
            <strong>{masterFile ? masterFile.name : 'Data FSS Master'}</strong>
            <span>{masterFile ? 'File master dipilih' : 'Pilih workbook hasil Build Data FSS (.xlsx/.xlsm)'}</span>
            {masterFile && <small>{formatFileSize(masterFile.size)}</small>}
          </div>
          <span className="update-sch-file-action">{masterFile ? 'Ganti file' : 'Pilih master'}</span>
        </label>
      </section>

      <section className="update-sch-section">
        <div className="update-sch-section-heading">
          <div>
            <h2>2. Pilih workbook schedule</h2>
            <p>Pilih bebas 1–{MAX_SCHEDULE_FILES} file XLSX/XLSM. Tidak wajib memilih empat template tertentu.</p>
          </div>
          <span className="update-sch-count">{selectedCount}/{MAX_SCHEDULE_FILES}</span>
        </div>

        <label className={`update-sch-file-card ${selectedCount ? 'is-selected' : ''}`}>
          <input
            type="file"
            accept=".xlsx,.xlsm"
            multiple
            onChange={handleScheduleFilesChange}
          />
          <div className="update-sch-file-icon">
            {selectedCount ? <CheckCircle2 size={21} /> : <FileUp size={21} />}
          </div>
          <div className="update-sch-file-copy">
            <strong>{selectedCount ? `${selectedCount} workbook dipilih` : 'Pilih workbook schedule'}</strong>
            <span>{selectedCount ? 'Klik untuk mengganti daftar file' : 'Bisa pilih 1, 4, 5, 15, atau sampai 20 file sekaligus'}</span>
            <small>XLSX/XLSM · Maksimal 50 MB per file</small>
          </div>
          <span className="update-sch-file-action">{selectedCount ? 'Ganti pilihan' : 'Pilih file'}</span>
        </label>

        {files.length > 0 && (
          <div className="update-sch-selected-files">
            {files.map((file, index) => (
              <div className="update-sch-selected-file" key={`${file.name}-${file.size}-${index}`}>
                <div>
                  <strong>{index + 1}. {file.name}</strong>
                  <small>{formatFileSize(file.size)}</small>
                </div>
                <button
                  type="button"
                  className="update-sch-file-remove"
                  onClick={() => removeFile(index)}
                  disabled={isReviewing || isUpdating}
                  aria-label={`Hapus ${file.name}`}
                  title="Hapus file"
                >
                  <X size={16} />
                </button>
              </div>
            ))}
          </div>
        )}

        {error && (
          <div className="update-sch-alert is-error">
            <AlertCircle size={18} />
            <span>{error}</span>
          </div>
        )}
        {success && (
          <div className="update-sch-alert is-success">
            <CheckCircle2 size={18} />
            <span>{success}</span>
          </div>
        )}

        <div className="update-sch-actions">
          <button
            type="button"
            className="update-sch-primary-button"
            onClick={handleReview}
            disabled={!masterFile || selectedCount < 1 || selectedCount > MAX_SCHEDULE_FILES || isReviewing || isUpdating}
          >
            {isReviewing ? <RefreshCw className="update-sch-spin" size={17} /> : <Eye size={17} />}
            {isReviewing ? 'Backend sedang memeriksa...' : 'Tinjau Pilihan'}
          </button>
          <span>{selectedCount} file dipilih · Maksimal {MAX_SCHEDULE_FILES} file. Tinjauan menjalankan validasi struktur dan rekonsiliasi backend.</span>
        </div>
      </section>

      {validation && (
        <section className="update-sch-results">
          <div className="update-sch-section-heading">
            <div>
              <h2>3. Hasil validasi struktur</h2>
              <p>{validation.valid_files} valid · {validation.invalid_files} perlu diperiksa dari {validation.total_files} file</p>
            </div>
            <span className={`update-sch-status ${validation.valid ? 'is-valid' : 'is-warning'}`}>
              {validation.valid ? 'Struktur valid' : 'Perlu pemeriksaan'}
            </span>
          </div>

          <div className="update-sch-result-grid">
            {validation.files.map((file, index) => (
              <article className="update-sch-result-card" key={file.role}>
                <div className="update-sch-result-title">
                  <div>
                    <strong>{file.label || `Workbook ${index + 1}`}</strong>
                    <span>{file.filename}</span>
                  </div>
                  {file.valid ? <CheckCircle2 className="is-good" size={20} /> : <AlertCircle className="is-bad" size={20} />}
                </div>
                <p className={file.valid ? 'update-sch-message' : 'update-sch-message is-bad'}>{file.message}</p>

                {file.valid && (
                  <>
                    <div className="update-sch-metrics">
                      <div><span>Jumlah data</span><strong>{formatNumber(file.data_rows)}</strong></div>
                      <div><span>Header</span><strong>Baris {file.header_row}</strong></div>
                      <div><span>Sheet</span><strong>{file.sheet_name}</strong></div>
                    </div>
                    <details className="update-sch-details">
                      <summary>Lihat posisi 12 field data</summary>
                      <div className="update-sch-column-map">
                        {Object.entries(file.column_positions ?? {}).map(([name, column]) => (
                          <span key={name}><b>{column}</b>{name}</span>
                        ))}
                      </div>
                    </details>
                    {file.warnings?.map((warning) => (
                      <div className="update-sch-warning" key={warning}>{warning}</div>
                    ))}
                  </>
                )}
              </article>
            ))}
          </div>
          <p className="update-sch-readonly-note">{validation.note}</p>
        </section>
      )}

      {preview && (
        <section className="update-sch-results">
          <div className="update-sch-section-heading">
            <div>
              <h2>4. Preview rekonsiliasi</h2>
              <p>
                Master {preview.master.filename} · {formatNumber(preview.master.total_rows)} order ·
                {' '}{formatNumber(preview.master.firm_rows)} Firm · {formatNumber(preview.master.forecast_rows)} Forecast
              </p>
            </div>
            <span className={`update-sch-status ${preview.valid ? 'is-valid' : 'is-warning'}`}>
              {preview.valid_files}/{preview.total_files} workbook valid
            </span>
          </div>

          <div className="update-sch-recon-grid">
            {preview.files.map((file, index) => {
              const counts = file.reconciliation?.counts ?? {}
              const statusClass = file.status === 'ready' ? 'is-valid' : file.status === 'blocked' ? 'is-blocked' : 'is-warning'
              return (
                <article className="update-sch-recon-card" key={file.role}>
                  <div className="update-sch-recon-card-heading">
                    <div>
                      <strong>{file.label || `Workbook ${index + 1}`}</strong>
                      <span>{file.filename}</span>
                    </div>
                    <span className={`update-sch-status ${statusClass}`}>{file.valid ? statusLabel(file) : 'Gagal dianalisis'}</span>
                  </div>
                  {file.valid && file.reconciliation ? (
                    <>
                      <div className="update-sch-recon-overview">
                        <div><span>Firm lama</span><strong>{formatNumber(file.firm_rows)}</strong></div>
                        <div><span>Forecast lama</span><strong>{formatNumber(file.forecast_rows)}</strong></div>
                        <div><span>Baris ambigu</span><strong>{formatNumber(file.ambiguous_rows)}</strong></div>
                      </div>
                      <div className="update-sch-recon-counts">
                        {COUNT_LABELS.map((label) => (
                          <div key={label}>
                            <span>{label}</span>
                            <strong>{formatNumber(counts[label])}</strong>
                          </div>
                        ))}
                      </div>
                      <div className="update-sch-recon-extra">
                        <span>Firm baru murni (di luar transisi) <b>{formatNumber(file.reconciliation.firm_new_non_transition_count)}</b></span>
                        <span>Transisi Forecast → Firm <b>{formatNumber(file.reconciliation.transition_count)}</b></span>
                        <span>Forecast lama dikeluarkan <b>{formatNumber(file.reconciliation.dropped_forecast_count)}</b></span>
                        <span>Konflik pemblokir <b>{formatNumber(file.reconciliation.blocking_count)}</b></span>
                      </div>
                      <div className="update-sch-recon-message is-info">
                        Jumlah "Firm dari master masuk ke bagian Firm" sudah mencakup baris transisi Forecast → Firm. Angka transisi adalah bagian dari jumlah tersebut, bukan tambahan di luar jumlah Firm.
                      </div>
                      {file.reconciliation.blocking_messages.map((message) => (
                        <div className="update-sch-recon-message is-danger" key={message}>
                          <AlertTriangle size={15} />{message}
                        </div>
                      ))}
                      {file.reconciliation.review_messages.map((message) => (
                        <div className="update-sch-recon-message is-warning" key={message}>
                          <AlertCircle size={15} />{message}
                        </div>
                      ))}
                      {file.reconciliation.info.map((message) => (
                        <div className="update-sch-recon-message is-info" key={message}>{message}</div>
                      ))}
                      {file.reconciliation.entries_preview.length > 0 && (
                        <details className="update-sch-details">
                          <summary>Lihat detail temuan ({formatNumber(file.reconciliation.entries_total)})</summary>
                          <div className="update-sch-entry-list">
                            {file.reconciliation.entries_preview.slice(0, 12).map((entry, entryIndex) => (
                              <div className="update-sch-entry" key={`${entry.category}-${entry.schedule_row ?? 'x'}-${entry.master_row ?? 'x'}-${entryIndex}`}>
                                <strong>{entry.category}</strong>
                                <span>{entry.message}</span>
                              </div>
                            ))}
                            {file.reconciliation.entries_truncated > 0 && (
                              <small>{formatNumber(file.reconciliation.entries_truncated)} temuan lainnya tidak ditampilkan pada preview.</small>
                            )}
                          </div>
                        </details>
                      )}
                    </>
                  ) : (
                    <p className="update-sch-message is-bad">{file.message}</p>
                  )}
                </article>
              )
            })}
          </div>
          <div className="update-sch-update-action">
            <div>
              <strong>{canUpdate ? `Semua pemeriksaan ${selectedCount} workbook lolos. Update siap dijalankan.` : 'Update belum tersedia.'}</strong>
              <span>
                Proses akan membuat {selectedCount} salinan Excel baru, memeriksa formula helper, menghitung ulang dengan Microsoft Excel, lalu mengunduh satu ZIP.
              </span>
            </div>
            <button
              type="button"
              className="update-sch-primary-button"
              onClick={handleUpdate}
              disabled={!canUpdate || isUpdating || isReviewing}
            >
              {isUpdating ? <RefreshCw className="update-sch-spin" size={17} /> : <Download size={17} />}
              {isUpdating ? `Memproses ${selectedCount} workbook...` : 'Proses Update & Download'}
            </button>
          </div>
          <p className="update-sch-readonly-note">{preview.note}</p>
        </section>
      )}
    </div>
  )
}
