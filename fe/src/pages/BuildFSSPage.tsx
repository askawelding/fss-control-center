import { useRef, useState, type DragEvent } from 'react'

import {
  AlertCircle,
  ArrowLeft,
  ArrowRight,
  CheckCircle2,
  FileSpreadsheet,
  FileUp,
  Info,
  RefreshCw,
  Trash2,
  X,
} from 'lucide-react'
import {
  validateFSSFiles,
  buildFSS,
  previewFSS,
  type FSSBuildSummary,
  type FSSFileValidation,
  type FSSPreviewResponse,
} from '../services/api'
import { startAppTask } from '../services/taskStatus'

type BuildStep = 1 | 2

type ValidationState = {
  valid: boolean
  files: FSSFileValidation[]
  totalFiles: number
  validFiles: number
  invalidFiles: number
}

function formatFileSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`

  if (bytes < 1024 * 1024) {
    return `${(bytes / 1024).toFixed(1)} KB`
  }

  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`
}

function getFileId(file: File): string {
  return `${file.name}-${file.size}-${file.lastModified}`
}

export default function BuildFSSPage() {
  const [releaseDate, setReleaseDate] = useState(() => {
  const today = new Date()
  const year = today.getFullYear()
  const month = String(today.getMonth() + 1).padStart(2, '0')
  const day = String(today.getDate()).padStart(2, '0')

  return `${year}-${month}-${day}`
  })

  const [isBuilding, setIsBuilding] = useState(false)
  const [buildError, setBuildError] = useState('')
  const [buildSummary, setBuildSummary] =
  useState<FSSBuildSummary | null>(null)
  const [files, setFiles] = useState<File[]>([])
  const [message, setMessage] = useState('')
  const [currentStep, setCurrentStep] = useState<BuildStep>(1)
  const [isDragging, setIsDragging] = useState(false)
  const [validation, setValidation] =
    useState<ValidationState | null>(null)

  const [isValidating, setIsValidating] = useState(false)

  const [validationError, setValidationError] = useState('')
  const [previewData, setPreviewData] =
  useState<FSSPreviewResponse | null>(null)

  const [isPreviewing, setIsPreviewing] = useState(false)
  const [previewError, setPreviewError] = useState('')

  const fileInputRef = useRef<HTMLInputElement>(null)

  function handleSelectedFiles(fileList: FileList | null) {
    if (!fileList) return

    const selectedFiles = Array.from(fileList)

    const validFiles = selectedFiles.filter((file) =>
      /\.(xlsx|xlsm|tsv)$/i.test(file.name),
    )

    const invalidCount = selectedFiles.length - validFiles.length

    // Hindari file yang sama masuk ke daftar lebih dari sekali.
    const existingIds = new Set(files.map(getFileId))

    const newFiles = validFiles.filter((file) => {
      const id = getFileId(file)

      if (existingIds.has(id)) {
        return false
      }

      existingIds.add(id)
      return true
    })

    setFiles((currentFiles) => [...currentFiles, ...newFiles])

    if (invalidCount > 0) {
      setMessage(
        `${invalidCount} file dilewati karena bukan file .xlsx, .xlsm, atau .tsv.`,
      )
    } else if (newFiles.length === 0 && validFiles.length > 0) {
      setMessage('File tersebut sudah ada di dalam daftar.')
    } else if (newFiles.length === 0) {
      setMessage('Tidak ada file baru yang ditambahkan.')
    } else {
      setMessage(`${newFiles.length} file berhasil ditambahkan.`)
    }
  }

  function handleDragOver(event: DragEvent<HTMLDivElement>) {
    event.preventDefault()
    setIsDragging(true)
  }

  function handleDragLeave(event: DragEvent<HTMLDivElement>) {
    event.preventDefault()

    const nextTarget = event.relatedTarget

    if (
      !(nextTarget instanceof Node) ||
      !event.currentTarget.contains(nextTarget)
    ) {
      setIsDragging(false)
    }
  }

  function handleDrop(event: DragEvent<HTMLDivElement>) {
    event.preventDefault()
    setIsDragging(false)

    handleSelectedFiles(event.dataTransfer.files)
  }

  function removeFile(fileToRemove: File) {
    const targetId = getFileId(fileToRemove)

    setFiles((currentFiles) =>
      currentFiles.filter((file) => getFileId(file) !== targetId),
    )

    setMessage('File berhasil dikeluarkan dari daftar.')
  }

  function clearFiles() {
    setFiles([])
    setMessage('Daftar file berhasil dikosongkan.')
  }

    async function handleBuildFSS() {
    if (!validation || !validation.valid) {
      setBuildError('Pastikan semua header file valid terlebih dahulu.')
      return
    }

    const forecastCount = validation.files.filter(
      (file) => file.detected_type === 'Forecast',
    ).length

    const firmCount = validation.files.filter(
      (file) => file.detected_type === 'Firm',
    ).length

    if (forecastCount !== 1 || firmCount < 1) {
      setBuildError(
        'Pilih minimal satu file Firm dan tepat satu file Forecast.',
      )
      return
    }

    if (!releaseDate) {
      setBuildError('Pilih tanggal rilis terlebih dahulu.')
      return
    }

    setIsBuilding(true)
    setBuildError('')
    setBuildSummary(null)
    const task = startAppTask('Build Data FSS', 'Menggabungkan file Firm dan Forecast menjadi master...')

    try {
      const result = await buildFSS(
        files,
        releaseDate.replaceAll('-', ''),
      )

      const url = URL.createObjectURL(result.blob)
      const link = document.createElement('a')

      link.href = url
      link.download = result.summary?.filename || 'FSS.xlsx'

      document.body.appendChild(link)
      link.click()
      link.remove()

      URL.revokeObjectURL(url)

      setBuildSummary(result.summary)
      task.complete('Data FSS master berhasil dibuat dan dikirim ke unduhan.')
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Gagal membangun file FSS.'
      setBuildError(
        message,
      )
      task.fail(message)
    } finally {
      setIsBuilding(false)
    }
  }

  async function goToReview() {
  if (files.length === 0) {
    setMessage('Pilih minimal satu file Excel terlebih dahulu.')
    return
  }

  setIsValidating(true)
  setValidationError('')
  setPreviewError('')
  setValidation(null)
  setPreviewData(null)
  setMessage('Sedang memeriksa file melalui backend...')
  const task = startAppTask('Build Data FSS', 'Memvalidasi header file Firm dan Forecast...')

  try {
    const result = await validateFSSFiles(files)

    const validationResult: ValidationState = {
      valid: result.invalid_files === 0,
      files: result.files,
      totalFiles: result.total_files,
      validFiles: result.valid_files,
      invalidFiles: result.invalid_files,
    }

    setValidation(validationResult)
    setCurrentStep(2)

    if (result.invalid_files > 0) {
      setMessage(
        `${result.invalid_files} file perlu diperbaiki sebelum diproses.`,
      )
      task.complete('Pemeriksaan selesai; ada file yang perlu diperbaiki.')
      return
    }

    setMessage('Header valid. Sedang mengambil preview data asli...')
    task.update('Membaca preview data dan menghitung jumlah baris...')

    setIsPreviewing(true)

    try {
      const preview = await previewFSS(
        files,
        releaseDate.replaceAll('-', ''),
      )

      setPreviewData(preview)
      setMessage('Preview data asli berhasil dimuat.')
      task.complete('Validasi dan preview Build Data FSS selesai.')
    } catch (error) {
      setPreviewError(
        error instanceof Error
          ? error.message
          : 'Gagal mengambil preview data.',
      )
      setMessage('')
      task.fail(error instanceof Error ? error.message : 'Gagal mengambil preview data.')
    } finally {
      setIsPreviewing(false)
    }
  } catch (error) {
    const errorMessage =
      error instanceof Error
        ? error.message
        : 'Terjadi kesalahan saat memeriksa file.'

    setValidationError(errorMessage)
    setMessage('')
    task.fail(errorMessage)
  } finally {
    setIsValidating(false)
  }
}


  const totalSize = files.reduce(
    (total, file) => total + file.size,
    0,
  )

  return (
    <main className="page-content build-page">
      <section className="page-intro">
        <div>
          <p className="eyebrow">WELDING DEPARTMENT</p>

          <h1>Build Data FSS</h1>

          <p className="page-description">
            Gabungkan file FSS customer menjadi satu master Excel.
          </p>
        </div>

        <div className="build-step-count">
          LANGKAH {currentStep} DARI 2
          <span>
            {currentStep === 1
              ? 'PEMILIHAN FILE'
              : 'RINGKASAN & PREVIEW'}
          </span>
        </div>
      </section>

      <div className="build-steps">
        <div
          className={`build-step ${
            currentStep === 1 ? 'active' : 'completed'
          }`}
        >
          <span className="build-step-number">
            {currentStep === 1 ? '1' : <CheckCircle2 size={16} />}
          </span>

          <div>
            <strong>Pilih file</strong>
            <span>File sumber Excel</span>
          </div>
        </div>

        <div className="build-step-line" />

        <div
          className={`build-step ${
            currentStep === 2 ? 'active' : ''
          }`}
        >
          <span className="build-step-number">2</span>

          <div>
            <strong>Ringkasan & Preview</strong>
            <span>Periksa persiapan output</span>
          </div>
        </div>
      </div>

      {currentStep === 1 ? (
        <section className="build-upload-panel">
          <div className="build-panel-heading">
            <div className="build-file-icon">
              <FileSpreadsheet size={23} />
            </div>

            <div>
              <h2>Pilih file sumber</h2>

              <p>
                Pilih seluruh file Firm dan Forecast untuk satu rilis customer.
              </p>
            </div>
          </div>

          <input
            ref={fileInputRef}
            className="visually-hidden-input"
            type="file"
            accept=".xlsx,.xlsm,.tsv"
            multiple
            onChange={(event) => {
              handleSelectedFiles(event.currentTarget.files)
              event.currentTarget.value = ''
            }}
          />

          <div
            className={`build-dropzone ${
              isDragging ? 'dragging' : ''
            }`}
            onDragEnter={handleDragOver}
            onDragOver={handleDragOver}
            onDragLeave={handleDragLeave}
            onDrop={handleDrop}
          >
            <div className="build-drop-icon">
              <FileUp size={29} />
            </div>

            <h3>
              {isDragging
                ? 'Lepaskan file di sini'
                : 'Tambahkan file Excel'}
            </h3>

            <p>
              Pilih file dari komputer kamu atau seret beberapa
              file Excel langsung ke area ini.
            </p>

            <button
              type="button"
              className="primary-button"
              onClick={() => fileInputRef.current?.click()}
            >
              <FileUp size={17} />
              Pilih File Excel
            </button>

            <span className="build-format-note">
              Format yang didukung: .tsv, .xlsx, atau .xlsm
            </span>
          </div>

          <div className="build-list-heading">
            <div>
              <h3>Daftar file terpilih</h3>

              <p>
                {files.length} file · {formatFileSize(totalSize)}
              </p>
            </div>

            {files.length > 0 && (
              <button
                type="button"
                className="build-clear-button"
                onClick={clearFiles}
              >
                <Trash2 size={15} />
                Kosongkan
              </button>
            )}
          </div>

          {files.length === 0 ? (
            <div className="build-empty-state">
              <FileSpreadsheet size={27} />

              <p>Belum ada file yang dipilih.</p>

              <span>
                File yang kamu pilih akan muncul di bagian ini.
              </span>
            </div>
          ) : (
            <div className="build-file-list">
              {files.map((file) => (
                <div
                  className="build-file-row"
                  key={getFileId(file)}
                >
                  <div className="build-file-row-icon">
                    <FileSpreadsheet size={19} />
                  </div>

                  <div className="build-file-details">
                    <strong>{file.name}</strong>
                    <span>{formatFileSize(file.size)}</span>
                  </div>

                  <span className="build-file-status">
                    Dipilih
                  </span>

                  <button
                    type="button"
                    className="build-remove-button"
                    aria-label={`Hapus ${file.name}`}
                    title="Hapus file"
                    onClick={() => removeFile(file)}
                  >
                    <X size={17} />
                  </button>
                </div>
              ))}
            </div>
          )}

          {message && (
            <div className="build-message" role="status">
              <Info size={16} />
              <span>{message}</span>
            </div>
          )}

          <div className="build-footer">
            <span>{files.length} file siap ditinjau</span>

            <button
              type="button"
              className="primary-button"
              disabled={files.length === 0 || isValidating}
              onClick={goToReview}
            >
              {isValidating ? 'Memeriksa file...' : 'Lanjut ke Preview'}
                {isValidating ? (
                  <RefreshCw size={17} className="spin" />
                  ) : (
              <ArrowRight size={17} />
                )}
            </button>
          </div>
        </section>
      ) : (
        <>
          <section className="build-upload-panel">
            <div className="build-panel-heading">
              <div className="build-file-icon">
                <CheckCircle2 size={23} />
              </div>

              <div>
                <h2>Ringkasan file sumber</h2>

                <p>
                  Periksa daftar file sebelum backend melakukan validasi Excel.
                </p>
              </div>
            </div>

            <div className="build-summary-grid">
              <div className="build-summary-card">
                <span>Jumlah file</span>
                <strong>{validation?.totalFiles ?? files.length}</strong>
              </div>

              <div className="build-summary-card">
                <span>Total ukuran file</span>
                <strong>{formatFileSize(totalSize)}</strong>
              </div>

              <div className="build-summary-card">
                <span>Status header</span>
                <strong
                  className={
                    validation?.valid
                      ? 'build-file-status'
                      : 'build-pending-status'
                  }
                >
                  {!validation
                    ? 'Belum diperiksa'
                    : validation.valid
                      ? 'Header valid'
                      : 'Perlu diperbaiki'}
                </strong>
              </div>
            </div>

            <div className="build-list-heading">
              <div>
                <h3>File yang akan diperiksa</h3>
                <p>Daftar file yang dipilih sebelumnya.</p>
              </div>
            </div>

            
            <div className="build-file-list">
              {files.map((file) => {
                const result = validation?.files.find(
                  (item) => item.filename === file.name,
                )

                return (
                  <div className="build-file-row" key={getFileId(file)}>
                    <div className="build-file-row-icon">
                      <FileSpreadsheet size={19} />
                    </div>

                    <div className="build-file-details">
                      <strong>{file.name}</strong>
                      <span>
                        {result?.detected_type
                          ? `Tipe: ${result.detected_type}`
                          : result?.format ?? formatFileSize(file.size)}
                      </span>
                      {result && (
                        <span>
                          {result.message}
                        </span>
                      )}
                    </div>

                    <span
                      className={
                        result
                          ? result.valid
                            ? 'build-file-status'
                            : 'build-pending-status'
                          : 'build-pending-status'
                      }
                    >
                      {result
                        ? result.valid
                          ? 'Header valid'
                          : 'Perlu diperbaiki'
                        : 'Belum diperiksa'}
                    </span>
                  </div>
                )
              })}
            </div>

          </section>
          
          {validationError && (
            <div className="build-message" role="alert">
              <AlertCircle size={16} />
              <span>{validationError}</span>
            </div>
          )}
          
          <section className="build-upload-panel">
            <div className="build-panel-heading">
              <div className="build-file-icon">
                <FileSpreadsheet size={23} />
              </div>

              <div>
                <h2>Tanggal rilis FSS</h2>
                <p>
                  Pilih tanggal rilis untuk nama file Excel hasil Build.
                </p>
              </div>
            </div>

            <div className="build-summary-card">
              <label htmlFor="fss-release-date">
                Tanggal rilis
              </label>

              <input
                id="fss-release-date"
                type="date"
                value={releaseDate}
                onChange={(event) => setReleaseDate(event.target.value)}
                disabled={isBuilding}
              />

              <p>
                Tanggal ini akan digunakan untuk menentukan nama file output.
              </p>
            </div>
          </section>


          <section className="build-preview-panel">
            <div className="build-preview-heading">
              <div>
                <h2>Preview data master</h2>
                <p>
                  Menampilkan data asli hasil pembacaan file oleh backend Python.
                </p>
              </div>

              <span className="build-demo-label">
                {previewData ? 'DATA ASLI' : 'PREVIEW'}
              </span>
            </div>

            {isPreviewing && (
              <div className="build-message" role="status">
                <RefreshCw size={16} className="spin" />
                <span>Sedang memproses dan mengambil preview data...</span>
              </div>
            )}

            {previewError && (
              <div className="build-message" role="alert">
                <AlertCircle size={16} />
                <span>{previewError}</span>
              </div>
            )}

            {previewData && (
              <>
                <div className="build-summary-grid">
                  <div className="build-summary-card">
                    <span>Baris Firm</span>
                    <strong>{previewData.firm_rows.toLocaleString('id-ID')}</strong>
                  </div>

                  <div className="build-summary-card">
                    <span>Baris Forecast</span>
                    <strong>{previewData.forecast_rows.toLocaleString('id-ID')}</strong>
                  </div>

                  <div className="build-summary-card">
                    <span>Total baris</span>
                    <strong>{previewData.total_rows.toLocaleString('id-ID')}</strong>
                  </div>

                  <div className="build-summary-card">
                    <span>Baris ditampilkan</span>
                    <strong>{previewData.preview_count.toLocaleString('id-ID')}</strong>
                  </div>
                </div>

                <div className="build-preview-table-wrap">
                  <table className="build-preview-table">
                    <thead>
                      <tr>
                        <th>Tipe</th>
                        {previewData.columns.map((column, index) => (
                          <th key={`${column}-${index}`}>{column}</th>
                        ))}
                      </tr>
                    </thead>

                    <tbody>
                      {previewData.rows.map((row, rowIndex) => (
                        <tr key={`${row.type}-${rowIndex}`}>
                          <td>
                            <strong>{row.type}</strong>
                          </td>

                          {row.values.map((value, valueIndex) => (
                            <td key={`${rowIndex}-${valueIndex}`}>
                              {value ?? '—'}
                            </td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>

                <p className="build-preview-caption">
                  Preview menampilkan maksimal 5 baris Firm dan 5 baris Forecast.
                  Data lengkap tetap diproses saat Build FSS.
                </p>
              </>
            )}

            {!previewData && !isPreviewing && !previewError && (
              <div className="build-empty-state">
                <FileSpreadsheet size={27} />
                <p>Preview belum tersedia.</p>
                <span>Silakan periksa kembali file yang dipilih.</span>
              </div>
            )}
          </section>
          
          
            {buildError && (
              <div className="build-message" role="alert">
                <AlertCircle size={16} />
                <span>{buildError}</span>
              </div>
            )}

            {buildSummary && (
              <section className="build-upload-panel">
                <div className="build-panel-heading">
                  <div className="build-file-icon">
                    <CheckCircle2 size={23} />
                  </div>

                  <div>
                    <h2>FSS berhasil dibangun</h2>
                    <p>File Excel telah dikirim untuk diunduh.</p>
                  </div>
                </div>

                <div className="build-summary-grid">
                  <div className="build-summary-card">
                    <span>File output</span>
                    <strong>{buildSummary.filename}</strong>
                  </div>

                  <div className="build-summary-card">
                    <span>Baris Firm</span>
                    <strong>{buildSummary.firm_rows}</strong>
                  </div>

                  <div className="build-summary-card">
                    <span>Baris Forecast</span>
                    <strong>{buildSummary.forecast_rows_after_filter}</strong>
                  </div>

                  <div className="build-summary-card">
                    <span>Peringatan</span>
                    <strong>{buildSummary.warning_count}</strong>
                  </div>
                </div>
              </section>
            )}

          
          <div className="build-footer">
            <button
              type="button"
              className="build-back-button"
              onClick={() => setCurrentStep(1)}
              disabled={isBuilding}
            >
              <ArrowLeft size={17} />
              Kembali ke Pilih File
            </button>

            <button
              type="button"
              className="primary-button"
              onClick={handleBuildFSS}
              disabled={
                isBuilding ||
                isPreviewing ||
                !previewData ||
                !validation?.valid ||
                validation.files.filter(
                  (file) => file.detected_type === 'Forecast',
                ).length !== 1 ||
                validation.files.filter(
                  (file) => file.detected_type === 'Firm',
                ).length < 1
              }
            >
              {isBuilding ? 'Sedang membangun...' : 'Build FSS'}
              <FileSpreadsheet size={17} />
            </button>
          </div>

        </>
      )}
    </main>
  )
}