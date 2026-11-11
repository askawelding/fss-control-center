import { useEffect, useMemo, useRef, useState } from 'react'
import {
  AlertTriangle,
  ArrowRight,
  CheckCircle2,
  Clock3,
  FileSpreadsheet,
  FileUp,
  Info,
  LoaderCircle,
  Pencil,
  Plus,
  Rocket,
  Save,
  Settings2,
  Trash2,
  X,
} from 'lucide-react'
import {
  createReceivingGroup,
  deleteReceivingGroup,
  getReceivingGroups,
  getTanpaRumusHistory,
  previewTanpaRumusDelivery,
  previewTanpaRumusSource,
  processTanpaRumusDelivery,
  updateReceivingGroup,
  type ReceivingGroup,
  type TanpaRumusDeliveryPreview,
  type TanpaRumusDeliverySummary,
  type TanpaRumusSourcePreview,
  type TanpaRumusRun,
} from '../services/api'
import './TanpaRumusPage.css'
import { startAppTask } from '../services/taskStatus'

type FeedbackType = 'info' | 'success' | 'warning' | 'error'

type PlannedOutput = {
  id: string
  name: string
  description: string
}

function formatFileSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`
}

function formatDuration(seconds?: number): string {
  if (seconds === undefined || !Number.isFinite(seconds)) return '—'
  if (seconds < 60) return `${seconds.toFixed(1)} detik`
  const minutes = Math.floor(seconds / 60)
  const remaining = Math.round(seconds % 60)
  return `${minutes} menit ${remaining} detik`
}

function getOutputFileName(file: File, label: string): string {
  const extension = /\.xlsm$/i.test(file.name) ? '.xlsm' : '.xlsx'
  const baseName = file.name.replace(/\.(xlsx|xlsm)$/i, '')
  const safeLabel = label.replace(/[<>:"/\\|?*]/g, '-').trim()
  return `${baseName} (${safeLabel})${extension}`
}

function normalizeCodes(value: string): string[] {
  return [...new Set(value.split(/[,\n;]+/).map((code) => code.trim().toUpperCase()).filter(Boolean))]
}

function downloadBlob(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob)
  const anchor = document.createElement('a')
  anchor.href = url
  anchor.download = filename
  document.body.appendChild(anchor)
  anchor.click()
  anchor.remove()
  window.setTimeout(() => URL.revokeObjectURL(url), 30_000)
}

export default function TanpaRumusPage() {
  const [mmkiFile, setMmkiFile] = useState<File | null>(null)
  const [ckdFile, setCkdFile] = useState<File | null>(null)

  const [makeMmki, setMakeMmki] = useState(false)
  const [makeCkdAll, setMakeCkdAll] = useState(false)
  const [makeCkdGroups, setMakeCkdGroups] = useState(false)

  const [groups, setGroups] = useState<ReceivingGroup[]>([])
  const [groupsLoading, setGroupsLoading] = useState(true)
  const [groupsSaving, setGroupsSaving] = useState(false)
  const [selectedGroupIds, setSelectedGroupIds] = useState<string[]>([])

  const [formOpen, setFormOpen] = useState(false)
  const [editingGroupId, setEditingGroupId] = useState<string | null>(null)
  const [groupName, setGroupName] = useState('')
  const [groupCodes, setGroupCodes] = useState('')
  const [outputLabel, setOutputLabel] = useState('')

  const [preview, setPreview] = useState<TanpaRumusDeliveryPreview | null>(null)
  const [mmkiPreview, setMmkiPreview] = useState<TanpaRumusSourcePreview | null>(null)
  const [showReview, setShowReview] = useState(false)
  const [feedback, setFeedback] = useState('')
  const [feedbackType, setFeedbackType] = useState<FeedbackType>('info')
  const [processing, setProcessing] = useState(false)
  const [reviewing, setReviewing] = useState(false)
  const [reviewElapsed, setReviewElapsed] = useState(0)
  const [processElapsed, setProcessElapsed] = useState(0)
  const [processStage, setProcessStage] = useState('')
  const [lastSummary, setLastSummary] = useState<TanpaRumusDeliverySummary | null>(null)
  const [history, setHistory] = useState<TanpaRumusRun[]>([])
  const [historyLoading, setHistoryLoading] = useState(true)

  const mmkiInputRef = useRef<HTMLInputElement>(null)
  const ckdInputRef = useRef<HTMLInputElement>(null)
  const selectedGroups = useMemo(
    () => groups.filter((group) => selectedGroupIds.includes(group.id)),
    [groups, selectedGroupIds],
  )
  const hasMmkiSource = Boolean(mmkiFile)
  const hasCkdSource = Boolean(ckdFile)

  useEffect(() => {
    if (!reviewing && !processing) return
    const startedAt = Date.now()
    const update = () => {
      const seconds = (Date.now() - startedAt) / 1000
      if (reviewing) setReviewElapsed(seconds)
      if (processing) setProcessElapsed(seconds)
    }
    update()
    const timer = window.setInterval(update, 250)
    return () => window.clearInterval(timer)
  }, [reviewing, processing])

  const plannedOutputs = useMemo<PlannedOutput[]>(() => {
    const outputs: PlannedOutput[] = []
    if (makeMmki && mmkiFile) {
      outputs.push({
        id: 'mmki-all',
        name: getOutputFileName(mmkiFile, 'TANPA RUMUS'),
        description: 'MMKI Delivery keseluruhan',
      })
    }
    if (makeCkdAll && ckdFile) {
      outputs.push({
        id: 'ckd-all',
        name: getOutputFileName(ckdFile, 'TANPA RUMUS'),
        description: 'CKD Delivery keseluruhan',
      })
    }
    if (makeCkdGroups && ckdFile) {
      selectedGroups.forEach((group) => {
        outputs.push({
          id: group.id,
          name: getOutputFileName(ckdFile, group.outputLabel),
          description: `CKD Delivery · ${group.name}`,
        })
      })
    }
    return outputs
  }, [makeMmki, makeCkdAll, makeCkdGroups, mmkiFile, ckdFile, selectedGroups])

  useEffect(() => {
    let active = true
    Promise.all([getReceivingGroups(), getTanpaRumusHistory()])
      .then(([items, runs]) => {
        if (!active) return
        setGroups(items)
        setSelectedGroupIds(items.map((group) => group.id))
        setHistory(runs)
      })
      .catch((error: Error) => {
        if (active) showMessage(error.message, 'error')
      })
      .finally(() => {
        if (active) {
          setGroupsLoading(false)
          setHistoryLoading(false)
        }
      })
    return () => {
      active = false
    }
    // Hanya muat konfigurasi saat halaman dibuka.
  }, [])

  function showMessage(text: string, type: FeedbackType = 'info') {
    setFeedback(text)
    setFeedbackType(type)
  }

  function invalidateReview() {
    setShowReview(false)
    setPreview(null)
    setMmkiPreview(null)
    setLastSummary(null)
  }

  function handleMmkiFile(file: File | null) {
    if (!file) return
    if (!/\.(xlsx|xlsm)$/i.test(file.name)) {
      showMessage('File MMKI harus berformat .xlsx atau .xlsm.', 'error')
      return
    }
    if (file.size > 50 * 1024 * 1024) {
      showMessage('Ukuran file MMKI melebihi batas 50 MB.', 'error')
      return
    }
    setMmkiFile(file)
    if (!ckdFile) {
      setMakeMmki(true)
      setMakeCkdAll(false)
      setMakeCkdGroups(false)
    }
    invalidateReview()
    showMessage('Workbook MMKI dipilih.', 'success')
  }

  function handleCkdFile(file: File | null) {
    if (!file) return
    if (!/\.(xlsx|xlsm)$/i.test(file.name)) {
      showMessage('File CKD harus berformat .xlsx atau .xlsm.', 'error')
      return
    }
    if (file.size > 50 * 1024 * 1024) {
      showMessage('Ukuran file CKD melebihi batas 50 MB.', 'error')
      return
    }
    setCkdFile(file)
    if (!mmkiFile) {
      setMakeMmki(false)
      setMakeCkdAll(true)
      setMakeCkdGroups(false)
    }
    invalidateReview()
    showMessage('Workbook CKD dipilih. Lakukan pemeriksaan kode receiving.', 'success')
  }

  function openNewGroupForm() {
    setEditingGroupId(null)
    setGroupName('')
    setGroupCodes('')
    setOutputLabel('')
    setFormOpen(true)
    showMessage('Isi kode receiving yang akan digabung menjadi satu file.', 'info')
  }

  function openEditGroupForm(group: ReceivingGroup) {
    setEditingGroupId(group.id)
    setGroupName(group.name)
    setGroupCodes(group.codes.join(', '))
    setOutputLabel(group.outputLabel)
    setFormOpen(true)
  }

  async function reloadGroups(selectNewId?: string) {
    const items = await getReceivingGroups()
    setGroups(items)
    if (selectNewId) {
      setSelectedGroupIds((current) => current.includes(selectNewId) ? current : [...current, selectNewId])
    } else {
      setSelectedGroupIds((current) => current.filter((id) => items.some((group) => group.id === id)))
    }
    invalidateReview()
  }

  async function saveGroup() {
    const cleanName = groupName.trim()
    const cleanLabel = outputLabel.trim()
    const codes = normalizeCodes(groupCodes)
    if (!cleanName || !cleanLabel || codes.length === 0) {
      showMessage('Nama kelompok, kode receiving, dan label nama file wajib diisi.', 'error')
      return
    }
    const invalid = codes.filter((code) => !/^[A-Z0-9]+$/.test(code))
    if (invalid.length) {
      showMessage(`Format kode tidak valid: ${invalid.join(', ')}.`, 'error')
      return
    }

    setGroupsSaving(true)
    try {
      const payload = { name: cleanName, codes, outputLabel: cleanLabel }
      const saved = editingGroupId
        ? await updateReceivingGroup(editingGroupId, payload)
        : await createReceivingGroup(payload)
      await reloadGroups(saved.id)
      setFormOpen(false)
      showMessage(`Kelompok ${saved.name} berhasil disimpan permanen.`, 'success')
    } catch (error) {
      showMessage(error instanceof Error ? error.message : 'Kelompok tidak dapat disimpan.', 'error')
    } finally {
      setGroupsSaving(false)
    }
  }

  async function deleteGroup(group: ReceivingGroup) {
    const confirmed = window.confirm(
      `Nonaktifkan kelompok "${group.name}"? Riwayat proses lama tetap menyimpan konfigurasi sebelumnya.`,
    )
    if (!confirmed) return
    setGroupsSaving(true)
    try {
      await deleteReceivingGroup(group.id)
      await reloadGroups()
      showMessage(`Kelompok ${group.name} dinonaktifkan.`, 'success')
      if (editingGroupId === group.id) setFormOpen(false)
    } catch (error) {
      showMessage(error instanceof Error ? error.message : 'Kelompok tidak dapat dihapus.', 'error')
    } finally {
      setGroupsSaving(false)
    }
  }

  function toggleGroup(groupId: string, checked: boolean) {
    setSelectedGroupIds((current) => checked
      ? (current.includes(groupId) ? current : [...current, groupId])
      : current.filter((id) => id !== groupId))
    invalidateReview()
  }

  async function reloadHistory() {
    try {
      setHistory(await getTanpaRumusHistory())
    } catch (error) {
      showMessage(error instanceof Error ? error.message : 'Riwayat Tanpa Rumus tidak dapat dibaca.', 'warning')
    }
  }

  async function handleReview() {
    if (processing || reviewing || groupsLoading) return
    if (makeMmki && !hasMmkiSource) {
      showMessage('Pilih workbook MMKI terlebih dahulu.', 'error')
      return
    }
    if ((makeCkdAll || makeCkdGroups) && !hasCkdSource) {
      showMessage('Pilih workbook CKD terlebih dahulu.', 'error')
      return
    }
    if (!makeMmki && !makeCkdAll && !makeCkdGroups) {
      showMessage('Pilih minimal satu jenis output.', 'error')
      return
    }
    if (makeMmki && !mmkiFile) {
      showMessage('Pilih workbook MMKI terlebih dahulu.', 'error')
      return
    }
    if ((makeCkdAll || makeCkdGroups) && !ckdFile) {
      showMessage('Pilih workbook CKD terlebih dahulu.', 'error')
      return
    }
    if (makeCkdGroups && selectedGroups.length === 0) {
      showMessage('Pilih minimal satu kelompok receiving.', 'error')
      return
    }

    setShowReview(false)
    setPreview(null)
    setMmkiPreview(null)
    setLastSummary(null)
    setReviewElapsed(0)
    setReviewing(true)
    const task = startAppTask('Tanpa Rumus', 'Memeriksa workbook sumber dan struktur sheet...')
    showMessage('Memeriksa workbook sumber...', 'info')

    try {
      if (makeMmki) {
        if (!mmkiFile) throw new Error('Pilih workbook MMKI terlebih dahulu.')
        task.update('Memeriksa struktur workbook MMKI dan worksheet bulanan...')
        const mmkiResult = await previewTanpaRumusSource(mmkiFile)
        setMmkiPreview(mmkiResult)
        setShowReview(true)
      }
      if (makeCkdAll || makeCkdGroups) {
        if (!ckdFile) throw new Error('Pilih workbook CKD terlebih dahulu.')
        task.update('Membaca sheet Date dan mengidentifikasi Receiving Area...')
        const result = await previewTanpaRumusDelivery(ckdFile)
        setPreview(result)
        setShowReview(true)
        if (result.unmapped_codes.length > 0 && makeCkdGroups) {
          showMessage(
            `Ada kode negara belum dipetakan: ${result.unmapped_codes.join(', ')}. Tambahkan ke kelompok sebelum membuat output CKD per kelompok.`,
            'warning',
          )
          task.complete('Pemeriksaan selesai; ada kode Receiving Area yang belum dipetakan.')
          return
        }
        if (result.blank_receiving_rows > 0 && makeCkdGroups) {
          showMessage(
            `Ada ${result.blank_receiving_rows} baris data tanpa Receiving Area. Periksa sheet Date agar order tidak salah masuk/tertinggal.`,
            'warning',
          )
          task.complete('Pemeriksaan selesai; ada baris Receiving Area yang kosong.')
          return
        }
      }
      showMessage('Pemeriksaan selesai. Periksa daftar output sebelum diproses.', 'success')
      task.complete('Pemeriksaan workbook selesai.')
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Pemeriksaan workbook gagal.'
      showMessage(message, 'error')
      task.fail(message)
    } finally {
      setReviewing(false)
    }
  }

  async function handleProcess() {
    if (processing || reviewing) return
    if (!showReview) {
      await handleReview()
      return
    }
    if (makeCkdGroups && preview?.unmapped_codes.length) {
      showMessage('Lengkapi mapping kode receiving yang belum dikenal sebelum memproses.', 'error')
      return
    }
    if (makeCkdGroups && preview?.blank_receiving_rows) {
      showMessage('Perbaiki baris dengan Receiving Area kosong sebelum memproses output kelompok.', 'error')
      return
    }

    setProcessing(true)
    setProcessElapsed(0)
    setLastSummary(null)
    setProcessStage('Menyiapkan salinan workbook sumber...')
    const task = startAppTask('Tanpa Rumus', 'Menyalin workbook dan menyiapkan Excel...')
    showMessage('Proses berjalan di backend lokal. Kamu boleh berpindah menu; jangan hentikan server backend.', 'info')
    try {
      setProcessStage('Excel membuka workbook dan menyiapkan perhitungan...')
      task.update('Excel dan Python sedang memproses output yang dipilih...')
      const result = await processTanpaRumusDelivery({
        mmkiFile,
        ckdFile,
        makeMmki,
        makeCkdAll,
        makeCkdGroups,
        selectedGroupIds,
      })
      downloadBlob(result.blob, 'FSS_Control_Center_Tanpa_Rumus.zip')
      setLastSummary(result.summary)
      await reloadHistory()
      setProcessStage('Selesai. ZIP hasil sudah dikirim ke unduhan browser.')
      task.complete(`Selesai: ${result.summary?.succeeded ?? 0} output berhasil dibuat.`)
      showMessage(
        result.summary
          ? `${result.summary.succeeded} output berhasil dibuat${result.summary.failed ? `, ${result.summary.failed} gagal` : ''}. ZIP hasil sedang diunduh.`
          : 'Proses selesai. ZIP hasil sedang diunduh.',
        result.summary?.failed ? 'warning' : 'success',
      )
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Proses Tanpa Rumus gagal.'
      setProcessStage('Proses gagal.')
      showMessage(message, 'error')
      task.fail(message)
    } finally {
      setProcessing(false)
    }
  }

  const hasUnmappedCodes = Boolean(preview?.unmapped_codes.length)
  const cannotProcessGroups = makeCkdGroups && Boolean(
    hasUnmappedCodes || (preview?.blank_receiving_rows ?? 0) > 0,
  )

  return (
    <main className="page-content without-formula-page">
      <section className="page-intro">
        <div>
          <p className="eyebrow">WELDING DEPARTMENT</p>
          <h1>Tanpa Rumus</h1>
          <p className="page-description">
            Buat file delivery final dari workbook hasil Update Schedule: bekukan rumus pada sheet bulanan dan hapus sheet Date.
          </p>
        </div>
        <div className="without-formula-badge"><Rocket size={15} /> Finalisasi workbook</div>
      </section>

      <section className="without-formula-notice">
        <Info size={19} />
        <div>
          <strong>Setiap output dibuat dari sumber asli yang sama</strong>
          <p>
            File sumber tidak diubah. Untuk CKD per kelompok, aplikasi membuat salinan baru dari CKD asli untuk setiap kelompok,
            menyaring data FSS pada sheet Date, menghitung rumus, membekukan sheet bulanan menjadi nilai, lalu menghapus Date.
          </p>
        </div>
      </section>

      <section className="without-formula-panel">
        <div className="without-formula-heading">
          <div className="without-formula-section-icon orange"><FileSpreadsheet size={21} /></div>
          <div>
            <h2>1. Workbook sumber</h2>
            <p>Pilih file delivery yang sudah melewati fitur Update Schedule.</p>
          </div>
        </div>

        <div className="without-formula-file-grid">
          <article className="without-formula-file-card">
            <span className="without-formula-label">INPUT MMKI</span>
            <h3>MMKI Delivery Schedule</h3>
            <p>Untuk membuat satu file MMKI keseluruhan tanpa rumus pada sheet bulanan.</p>
            <input
              ref={mmkiInputRef}
              className="visually-hidden-input"
              type="file"
              accept=".xlsx,.xlsm"
              onChange={(event) => {
                handleMmkiFile(event.currentTarget.files?.[0] ?? null)
                event.currentTarget.value = ''
              }}
            />
            {mmkiFile ? (
              <div className="without-formula-selected-file">
                <FileSpreadsheet size={18} />
                <div><strong>{mmkiFile.name}</strong><span>{formatFileSize(mmkiFile.size)}</span></div>
                <button type="button" aria-label="Hapus file MMKI" onClick={() => { setMmkiFile(null); setMakeMmki(false); invalidateReview() }}><X size={16} /></button>
              </div>
            ) : (
              <div className="without-formula-empty-file"><FileUp size={22} /><span>Belum ada file dipilih</span></div>
            )}
            <button type="button" className="without-formula-select-button" onClick={() => mmkiInputRef.current?.click()}>
              <FileUp size={16} /> {mmkiFile ? 'Ganti File MMKI' : 'Pilih File MMKI'}
            </button>
          </article>

          <article className="without-formula-file-card">
            <span className="without-formula-label">INPUT CKD</span>
            <h3>CKD Delivery Schedule</h3>
            <p>Untuk CKD keseluruhan dan/atau satu file untuk setiap kelompok receiving terpilih.</p>
            <input
              ref={ckdInputRef}
              className="visually-hidden-input"
              type="file"
              accept=".xlsx,.xlsm"
              onChange={(event) => {
                handleCkdFile(event.currentTarget.files?.[0] ?? null)
                event.currentTarget.value = ''
              }}
            />
            {ckdFile ? (
              <div className="without-formula-selected-file">
                <FileSpreadsheet size={18} />
                <div><strong>{ckdFile.name}</strong><span>{formatFileSize(ckdFile.size)}</span></div>
                <button type="button" aria-label="Hapus file CKD" onClick={() => { setCkdFile(null); setMakeCkdAll(false); setMakeCkdGroups(false); invalidateReview() }}><X size={16} /></button>
              </div>
            ) : (
              <div className="without-formula-empty-file"><FileUp size={22} /><span>Belum ada file dipilih</span></div>
            )}
            <button type="button" className="without-formula-select-button" onClick={() => ckdInputRef.current?.click()}>
              <FileUp size={16} /> {ckdFile ? 'Ganti File CKD' : 'Pilih File CKD'}
            </button>
          </article>
        </div>
      </section>

      <section className="without-formula-panel">
        <div className="without-formula-heading">
          <div className="without-formula-section-icon purple"><Rocket size={21} /></div>
          <div>
            <h2>2. Jenis output</h2>
            <p>Pilih satu atau beberapa hasil yang akan dibuat.</p>
          </div>
        </div>
        <div className="without-formula-options">
          <label className="without-formula-option">
            <input type="checkbox" checked={makeMmki} disabled={!hasMmkiSource || processing || reviewing} onChange={(event) => { setMakeMmki(event.currentTarget.checked); invalidateReview() }} />
            <div className="without-formula-option-copy"><strong>MMKI keseluruhan</strong><span>{hasMmkiSource ? 'Bekukan formula pada sheet bulanan MMKI, lalu hapus sheet Date. Tidak ada filter receiving.' : 'Pilih workbook MMKI pada langkah 1 untuk mengaktifkan output ini.'}</span></div>
          </label>
          <label className="without-formula-option">
            <input type="checkbox" checked={makeCkdAll} disabled={!hasCkdSource || processing || reviewing} onChange={(event) => { setMakeCkdAll(event.currentTarget.checked); invalidateReview() }} />
            <div className="without-formula-option-copy"><strong>CKD keseluruhan</strong><span>{hasCkdSource ? 'Bekukan formula pada sheet bulanan CKD, lalu hapus sheet Date tanpa memisahkan receiving.' : 'Pilih workbook CKD pada langkah 1 untuk mengaktifkan output ini.'}</span></div>
          </label>
          <label className="without-formula-option">
            <input type="checkbox" checked={makeCkdGroups} disabled={!hasCkdSource || processing || reviewing} onChange={(event) => { setMakeCkdGroups(event.currentTarget.checked); invalidateReview() }} />
            <div className="without-formula-option-copy"><strong>CKD per kelompok receiving</strong><span>{hasCkdSource ? 'Satu file per kelompok. Semua kode dalam satu kelompok digabung pada satu output, bukan satu file per kode.' : 'Pilih workbook CKD pada langkah 1 untuk mengaktifkan output ini.'}</span></div>
          </label>
        </div>
      </section>

      <section className="without-formula-panel">
        <div className="without-formula-heading group-heading">
          <div className="without-formula-section-icon orange"><Settings2 size={21} /></div>
          <div><h2>3. Kelompok receiving</h2><p>Tambahkan satu atau beberapa kode ke dalam satu kelompok output.</p></div>
          <button type="button" className="without-formula-add-button" onClick={openNewGroupForm} disabled={groupsSaving || processing || reviewing}><Plus size={16} /> Tambah kelompok</button>
        </div>

        {groupsLoading ? (
          <div className="tanpa-loading-inline"><LoaderCircle className="spin" size={19} /> Membaca konfigurasi kelompok dari database lokal...</div>
        ) : (
          <div className="without-formula-group-list">
            {groups.map((group) => {
              const isSelected = selectedGroupIds.includes(group.id)
              return (
                <article key={group.id} className={`without-formula-group-card ${isSelected ? 'selected' : ''}`}>
                  <label className="without-formula-group-check">
                    <input type="checkbox" checked={isSelected} disabled={!hasCkdSource || !makeCkdGroups || groupsSaving || processing || reviewing} onChange={(event) => toggleGroup(group.id, event.currentTarget.checked)} />
                  </label>
                  <div className="without-formula-group-main">
                    <div className="without-formula-group-title"><strong>{group.name}</strong><span>{group.codes.length} kode</span></div>
                    <div className="without-formula-code-list">{group.codes.map((code) => <span key={code}>{code}</span>)}</div>
                    <p className="without-formula-output-label">Output: {group.outputLabel}</p>
                  </div>
                  <div className="without-formula-group-actions">
                    <button type="button" aria-label={`Edit ${group.name}`} title="Edit kelompok" disabled={groupsSaving || processing || reviewing} onClick={() => openEditGroupForm(group)}><Pencil size={16} /></button>
                    <button type="button" aria-label={`Hapus ${group.name}`} title="Nonaktifkan kelompok" disabled={groupsSaving || processing || reviewing} onClick={() => void deleteGroup(group)}><Trash2 size={16} /></button>
                  </div>
                </article>
              )
            })}
            {groups.length === 0 && <div className="without-formula-empty-groups"><Info size={23} /><p>Belum ada kelompok receiving.</p><span>Tambahkan kelompok pertama untuk mulai mengatur output CKD.</span></div>}
          </div>
        )}

        <div className="without-formula-config-note"><Info size={16} /><p>Konfigurasi disimpan permanen di SQLite lokal. Satu kode hanya boleh berada pada satu kelompok aktif. Pengaturan yang digunakan setiap proses ikut dicatat dalam laporan.</p></div>

        {formOpen && (
          <div className="without-formula-form">
            <div className="without-formula-form-heading">
              <div><h3>{editingGroupId ? 'Edit kelompok receiving' : 'Tambah kelompok receiving'}</h3><p>Masukkan beberapa kode jika semuanya harus menjadi satu file output.</p></div>
              <button type="button" aria-label="Tutup form" onClick={() => setFormOpen(false)}><X size={18} /></button>
            </div>
            <div className="without-formula-form-grid">
              <label className="without-formula-field"><span>Nama kelompok</span><input value={groupName} onChange={(event) => setGroupName(event.currentTarget.value)} placeholder="Contoh: Singapore Delivery" /></label>
              <label className="without-formula-field"><span>Label nama file output</span><input value={outputLabel} onChange={(event) => setOutputLabel(event.currentTarget.value)} placeholder="Contoh: 7R12 & 7R13 SINGAPORE" /></label>
              <label className="without-formula-field full-width"><span>Kode receiving</span><input value={groupCodes} onChange={(event) => setGroupCodes(event.currentTarget.value)} placeholder="Contoh: 7R12, 7R13" /><small>Pisahkan dengan koma. Satu kode tidak boleh dipakai di dua kelompok.</small></label>
            </div>
            <div className="without-formula-form-actions">
              <button type="button" className="without-formula-cancel-button" onClick={() => setFormOpen(false)} disabled={groupsSaving}>Batal</button>
              <button type="button" className="primary-button" onClick={() => void saveGroup()} disabled={groupsSaving}>
                {groupsSaving ? <LoaderCircle className="spin" size={16} /> : <Save size={16} />} Simpan kelompok
              </button>
            </div>
          </div>
        )}
      </section>

      {reviewing && <div className="tanpa-processing-panel" role="status"><LoaderCircle className="spin" size={22} /><div><strong>Sedang memeriksa workbook sumber...</strong><span>Berjalan {formatDuration(reviewElapsed)}. Backend sedang membaca struktur workbook dan data Receiving Area.</span></div></div>}
      {feedback && <div className={`without-formula-feedback ${feedbackType}`} role="status"><Info size={17} /><span>{feedback}</span></div>}

      <section className="without-formula-footer">
        <div><span className="without-formula-footer-label">RENCANA OUTPUT</span><strong>{plannedOutputs.length} file</strong><p>Setiap kelompok receiving menghasilkan satu file gabungan. Setiap output dibuat ulang dari workbook sumber asli.</p></div>
        <button type="button" className="primary-button" onClick={() => void handleReview()} disabled={processing || reviewing || groupsLoading || plannedOutputs.length === 0}>
          {reviewing ? <LoaderCircle className="spin" size={17} /> : <ArrowRight size={17} />}
          {reviewing ? `Memeriksa · ${formatDuration(reviewElapsed)}` : 'Periksa workbook'}
        </button>
      </section>

      {showReview && (
        <section className="without-formula-review">
          <div className="without-formula-review-heading"><CheckCircle2 size={22} /><div><h2>Hasil pemeriksaan dan rencana output</h2><p>Pastikan semua kode negara sudah masuk ke kelompok sebelum membuat file final.</p></div></div>
          {mmkiPreview && (
            <div className="tanpa-preview-summary">
              <div><span>Workbook MMKI</span><strong>{mmkiPreview.filename}</strong></div>
              <div><span>Sheet sumber</span><strong>{mmkiPreview.date_sheet}</strong></div>
              <div><span>Sheet bulanan</span><strong>{mmkiPreview.month_sheets.length}</strong></div>
              <div><span>Waktu pemeriksaan MMKI</span><strong>{formatDuration(mmkiPreview.processing_seconds)}</strong></div>
            </div>
          )}
          {preview && (
            <div className="tanpa-preview-summary">
              <div><span>Sheet sumber</span><strong>{preview.date_sheet}</strong></div>
              <div><span>Header</span><strong>Baris {preview.header_row}</strong></div>
              <div><span>Baris data terbaca</span><strong>{preview.row_count.toLocaleString('id-ID')}</strong></div>
              <div><span>Sheet bulanan</span><strong>{preview.month_sheets.length}</strong></div>
              <div><span>Waktu pemeriksaan</span><strong>{formatDuration(preview.processing_seconds)}</strong></div>
            </div>
          )}
          {preview && (
            <div className="tanpa-code-preview">
              <div className="tanpa-preview-heading"><strong>Kode Receiving Area pada sheet Date</strong><span>{preview.receiving_areas.length} kode unik</span></div>
              <div className="tanpa-code-table-wrap">
                <table className="tanpa-code-table"><thead><tr><th>Kode</th><th>Jumlah order</th><th>Status mapping</th></tr></thead><tbody>
                  {preview.receiving_areas.map((area) => {
                    const mapped = !area.is_country_code || !preview.unmapped_codes.includes(area.code)
                    return <tr key={area.code}><td><strong>{area.code}</strong></td><td>{area.count.toLocaleString('id-ID')}</td><td>{area.is_country_code ? (mapped ? <span className="tanpa-status-mapped">Sudah dipetakan</span> : <span className="tanpa-status-unmapped">Belum dipetakan</span>) : <span className="tanpa-status-neutral">Di luar kode negara 7R</span>}</td></tr>
                  })}
                </tbody></table>
              </div>
              {preview.unmapped_codes.length > 0 && (
                <div className="tanpa-unmapped-warning"><AlertTriangle size={18} /><div><strong>Kode negara baru belum dipetakan: {preview.unmapped_codes.join(', ')}</strong><p>Edit kelompok yang ada atau buat kelompok baru, lalu periksa ulang workbook. Kode baru tidak akan diabaikan secara diam-diam.</p></div></div>
              )}
              {preview.blank_receiving_rows > 0 && <div className="tanpa-unmapped-warning"><AlertTriangle size={18} /><div><strong>{preview.blank_receiving_rows} baris mempunyai Receiving Area kosong</strong><p>Lengkapi data pada sheet Date sebelum membuat output per negara.</p></div></div>}
            </div>
          )}
          <div className="without-formula-review-list">
            {plannedOutputs.map((output) => <div className="without-formula-review-row" key={`${output.id}-${output.name}`}><FileSpreadsheet size={18} /><div><strong>{output.name}</strong><span>{output.description}</span></div></div>)}
            {plannedOutputs.length === 0 && <p>Belum ada output yang direncanakan.</p>}
          </div>
          <div className="tanpa-process-actions">
            <div className="without-formula-review-note"><Info size={18} /><p>Setiap kelompok diproses dari file CKD sumber yang sama, bukan dari hasil kelompok sebelumnya. Formula pada sheet bulanan dihitung setelah filter Date, dibekukan menjadi nilai, kemudian sheet Date dihapus pada salinan output.</p></div>
            <button type="button" className="primary-button" onClick={() => void handleProcess()} disabled={processing || reviewing || plannedOutputs.length === 0 || cannotProcessGroups || (makeCkdGroups && selectedGroups.length === 0)}>
              {processing ? <LoaderCircle className="spin" size={17} /> : <Rocket size={17} />}
              {processing ? 'Sedang memproses...' : 'Buat file Tanpa Rumus'}
            </button>
          </div>
        </section>
      )}

      {processing && <div className="tanpa-processing-panel" role="status"><LoaderCircle className="spin" size={22} /><div><strong>{processStage || 'Memproses workbook...'}</strong><span>Berjalan {formatDuration(processElapsed)}. Proses tetap berjalan saat kamu berpindah menu; jangan hentikan backend.</span></div></div>}
      {processStage && !processing && <p className="tanpa-process-stage">{processStage}</p>}
      {lastSummary && (
        <section className="tanpa-result-panel">
          <div><CheckCircle2 size={20} /><strong>Ringkasan proses terakhir</strong></div>
          <p>{lastSummary.succeeded} berhasil · {lastSummary.failed} gagal · {lastSummary.total_outputs} output direncanakan · Total {formatDuration(lastSummary.processing_seconds)}{lastSummary.preview_seconds !== undefined ? ` · Preview server ${formatDuration(lastSummary.preview_seconds)}` : ''}.</p>
          {lastSummary.outputs.map((output) => <span key={output.name} className={output.status === 'succeeded' ? 'success' : 'failed'}>{output.name}: {output.status === 'succeeded' ? `berhasil · ${formatDuration(output.processing_seconds)}` : `gagal · ${formatDuration(output.processing_seconds)}${output.error ? ` · ${output.error}` : ''}`}</span>)}
        </section>
      )}

      <section className="without-formula-panel tanpa-history-panel">
        <div className="without-formula-heading">
          <div className="without-formula-section-icon purple"><Clock3 size={21} /></div>
          <div><h2>Riwayat Tanpa Rumus</h2><p>Riwayat nyata dari proses yang tersimpan di database lokal.</p></div>
        </div>
        {historyLoading ? (
          <div className="tanpa-loading-inline"><LoaderCircle className="spin" size={18} /> Membaca riwayat proses...</div>
        ) : history.length === 0 ? (
          <div className="without-formula-empty-groups"><Info size={22} /><p>Belum ada proses Tanpa Rumus.</p><span>Riwayat akan muncul setelah proses pertama dijalankan.</span></div>
        ) : (
          <div className="tanpa-history-list">
            {history.map((run) => {
              const outputs = Array.isArray(run.report.files) ? run.report.files : []
              const successful = run.status === 'succeeded'
              return (
                <article className="tanpa-history-item" key={run.id}>
                  <div className={`tanpa-history-status ${successful ? 'success' : run.status === 'partial' ? 'warning' : 'failed'}`}>
                    {successful ? <CheckCircle2 size={17} /> : <AlertTriangle size={17} />}
                  </div>
                  <div className="tanpa-history-content">
                    <div className="tanpa-history-title"><strong>{run.sourceFiles.join(' · ')}</strong><span>{run.createdAt}</span></div>
                    <p>{run.report.succeeded ?? 0} berhasil · {run.report.failed ?? 0} gagal · {run.report.total_outputs ?? outputs.length} output · Total {formatDuration(run.report.processing_seconds)}</p>
                    {outputs.length > 0 && <div className="tanpa-history-output-list">{outputs.map((output, index) => <span key={`${run.id}-${index}`} className={output.status === 'succeeded' ? 'success' : 'failed'}>{output.output ?? output.description ?? `Output ${index + 1}`} · {formatDuration(output.processing_seconds)}</span>)}</div>}
                  </div>
                </article>
              )
            })}
          </div>
        )}
      </section>
    </main>
  )
}
