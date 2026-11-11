
export type BackendHealth = {
  status: string
  app: string
  message: string
  version: string
}

export type FSSFileValidation = {
  filename: string
  valid: boolean
  format?: string
  size_bytes?: number
  detected_type?: 'Firm' | 'Forecast'
  column_count?: number
  missing_columns?: string[]
  unknown_headers?: string[]
  duplicate_headers?: string[]
  row_preview?: string[]
  header_row?: number
  header_sheet?: string
  sheets?: string[]
  message: string
}

export type FSSValidationResponse = {
  status: string
  total_files: number
  valid_files: number
  invalid_files: number
  files: FSSFileValidation[]
}

export type FSSBuildSummary = {
  filename: string
  firm_files: string[]
  forecast_file: string
  firm_rows: number
  forecast_rows_before_filter: number
  forecast_rows_removed: number
  forecast_rows_after_filter: number
  total_rows: number
  duplicate_rows_warned: number
  warning_count: number
}

export type FSSPreviewRow = {
  type: 'Firm' | 'Forecast'
  values: (string | number | boolean | null)[]
}

export type FSSPreviewResponse = {
  status: string
  filename: string
  columns: string[]
  rows: FSSPreviewRow[]
  preview_count: number
  firm_rows: number
  forecast_rows: number
  total_rows: number
}

const API_BASE_URL = 'http://127.0.0.1:8000'

async function getErrorMessage(response: Response): Promise<string> {
  try {
    const data: { detail?: string } = await response.json()
    return data.detail || `Permintaan gagal (${response.status}).`
  } catch {
    return `Permintaan gagal (${response.status}).`
  }
}

export async function getBackendHealth(): Promise<BackendHealth | null> {
  try {
    const response = await fetch(`${API_BASE_URL}/api/health`)

    if (!response.ok) {
      return null
    }

    return (await response.json()) as BackendHealth
  } catch {
    return null
  }
}

export async function validateFSSFiles(
  files: File[],
): Promise<FSSValidationResponse> {
  const formData = new FormData()

  files.forEach((file) => {
    formData.append('files', file)
  })

  const response = await fetch(`${API_BASE_URL}/api/fss/validate`, {
    method: 'POST',
    body: formData,
  })

  if (!response.ok) {
    throw new Error(await getErrorMessage(response))
  }

  return (await response.json()) as FSSValidationResponse
}

export async function buildFSS(
  files: File[],
  releaseDate: string,
): Promise<{ blob: Blob; summary: FSSBuildSummary | null }> {
  const formData = new FormData()

  files.forEach((file) => {
    formData.append('files', file)
  })

  formData.append('release_date', releaseDate)

  const response = await fetch(`${API_BASE_URL}/api/fss/build`, {
    method: 'POST',
    body: formData,
  })

  if (!response.ok) {
    throw new Error(await getErrorMessage(response))
  }

  const blob = await response.blob()
  const summaryHeader = response.headers.get('X-FSS-Summary')

  let summary: FSSBuildSummary | null = null

  if (summaryHeader) {
    try {
      summary = JSON.parse(
        decodeURIComponent(summaryHeader),
      ) as FSSBuildSummary
    } catch {
      // File tetap dapat diunduh meskipun ringkasan tidak terbaca.
    }
  }

  return { blob, summary }
}

export async function previewFSS(
  files: File[],
  releaseDate: string,
): Promise<FSSPreviewResponse> {
  const formData = new FormData()

  files.forEach((file) => formData.append('files', file))
  formData.append('release_date', releaseDate)

  const response = await fetch(`${API_BASE_URL}/api/fss/preview`, {
    method: 'POST',
    body: formData,
  })

  if (!response.ok) {
    throw new Error(await getErrorMessage(response))
  }

  return (await response.json()) as FSSPreviewResponse
}
// ---------------------------------------------------------------------------
// Tanpa Rumus — konfigurasi receiving, preview Date, dan finalisasi delivery
// ---------------------------------------------------------------------------

export type ReceivingGroup = {
  id: string
  name: string
  codes: string[]
  outputLabel: string
  active: boolean
  createdAt?: string
  updatedAt?: string
}

export type ReceivingAreaPreview = {
  code: string
  count: number
  is_country_code: boolean
}

export type TanpaRumusDeliveryPreview = {
  date_sheet: string
  header_row: number
  data_start_row: number
  last_data_row: number
  receiving_column: number
  fss_columns: Record<string, number>
  month_sheets: string[]
  receiving_areas: ReceivingAreaPreview[]
  country_receiving_codes: { code: string; count: number; mapped: boolean }[]
  unmapped_codes: string[]
  blank_receiving_rows: number
  row_count: number
  processing_seconds?: number
}

export type TanpaRumusSourcePreview = {
  filename: string
  extension: string
  date_sheet: string
  month_sheets: string[]
  sheet_names: string[]
  processing_seconds: number
}

export type TanpaRumusDeliverySummary = {
  total_outputs: number
  succeeded: number
  failed: number
  status: 'succeeded' | 'partial' | 'failed'
  processing_seconds?: number
  preview_seconds?: number
  outputs: {
    name: string
    status: string
    description: string
    processing_seconds?: number
    calculation_seconds?: number
    calculation_mode?: string
    error?: string
  }[]
}

export type TanpaRumusRun = {
  id: string
  createdAt: string
  sourceFiles: string[]
  selectedOutputs: { kind: string; group_id?: string; name?: string; codes?: string[]; outputLabel?: string }[]
  groups: ReceivingGroup[]
  status: 'succeeded' | 'partial' | 'failed' | string
  report: {
    total_outputs?: number
    succeeded?: number
    failed?: number
    processing_seconds?: number
    preview_seconds?: number
    files?: {
      output?: string
      description?: string
      status?: string
      error?: string
      processing_seconds?: number
      calculation_seconds?: number
    }[]
    [key: string]: unknown
  }
}

export async function getTanpaRumusHistory(limit = 20): Promise<TanpaRumusRun[]> {
  const response = await fetch(`${API_BASE_URL}/api/tanpa-rumus/history?limit=${limit}`)
  if (!response.ok) throw new Error(await getTanpaRumusError(response))
  const payload = await response.json() as { items: TanpaRumusRun[] }
  return payload.items
}

async function getTanpaRumusError(response: Response): Promise<string> {
  try {
    const payload = await response.json() as { detail?: string }
    return payload.detail || `Permintaan gagal (${response.status}).`
  } catch {
    return `Permintaan gagal (${response.status}).`
  }
}

export async function getReceivingGroups(): Promise<ReceivingGroup[]> {
  const response = await fetch(`${API_BASE_URL}/api/tanpa-rumus/groups`)
  if (!response.ok) throw new Error(await getTanpaRumusError(response))
  const payload = await response.json() as { groups: ReceivingGroup[] }
  return payload.groups
}

export async function createReceivingGroup(
  payload: Pick<ReceivingGroup, 'name' | 'codes' | 'outputLabel'>,
): Promise<ReceivingGroup> {
  const response = await fetch(`${API_BASE_URL}/api/tanpa-rumus/groups`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  })
  if (!response.ok) throw new Error(await getTanpaRumusError(response))
  return ((await response.json()) as { group: ReceivingGroup }).group
}

export async function updateReceivingGroup(
  groupId: string,
  payload: Pick<ReceivingGroup, 'name' | 'codes' | 'outputLabel'>,
): Promise<ReceivingGroup> {
  const response = await fetch(`${API_BASE_URL}/api/tanpa-rumus/groups/${encodeURIComponent(groupId)}`, {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  })
  if (!response.ok) throw new Error(await getTanpaRumusError(response))
  return ((await response.json()) as { group: ReceivingGroup }).group
}

export async function deleteReceivingGroup(groupId: string): Promise<void> {
  const response = await fetch(`${API_BASE_URL}/api/tanpa-rumus/groups/${encodeURIComponent(groupId)}`, {
    method: 'DELETE',
  })
  if (!response.ok) throw new Error(await getTanpaRumusError(response))
}

export async function previewTanpaRumusDelivery(
  ckdFile: File,
): Promise<TanpaRumusDeliveryPreview> {
  const form = new FormData()
  form.append('ckd_file', ckdFile)
  const response = await fetch(`${API_BASE_URL}/api/tanpa-rumus/preview-delivery`, {
    method: 'POST',
    body: form,
  })
  if (!response.ok) throw new Error(await getTanpaRumusError(response))
  return (await response.json()) as TanpaRumusDeliveryPreview
}

export async function previewTanpaRumusSource(
  workbookFile: File,
): Promise<TanpaRumusSourcePreview> {
  const form = new FormData()
  form.append('workbook_file', workbookFile)
  const response = await fetch(`${API_BASE_URL}/api/tanpa-rumus/preview-source`, {
    method: 'POST',
    body: form,
  })
  if (!response.ok) throw new Error(await getTanpaRumusError(response))
  return (await response.json()) as TanpaRumusSourcePreview
}

export async function processTanpaRumusDelivery(args: {
  mmkiFile: File | null
  ckdFile: File | null
  makeMmki: boolean
  makeCkdAll: boolean
  makeCkdGroups: boolean
  selectedGroupIds: string[]
}): Promise<{ blob: Blob; summary: TanpaRumusDeliverySummary | null }> {
  const form = new FormData()
  if (args.mmkiFile) form.append('mmki_file', args.mmkiFile)
  if (args.ckdFile) form.append('ckd_file', args.ckdFile)
  form.append('make_mmki', String(args.makeMmki))
  form.append('make_ckd_all', String(args.makeCkdAll))
  form.append('make_ckd_groups', String(args.makeCkdGroups))
  form.append('selected_group_ids', JSON.stringify(args.selectedGroupIds))

  const response = await fetch(`${API_BASE_URL}/api/tanpa-rumus/process-delivery`, {
    method: 'POST',
    body: form,
  })
  if (!response.ok) throw new Error(await getTanpaRumusError(response))

  const blob = await response.blob()
  let summary: TanpaRumusDeliverySummary | null = null
  const header = response.headers.get('X-Tanpa-Rumus-Summary')
  if (header) {
    try {
      summary = JSON.parse(decodeURIComponent(header)) as TanpaRumusDeliverySummary
    } catch {
      // ZIP tetap dapat diunduh walau ringkasan header tidak dapat dibaca.
    }
  }
  return { blob, summary }
}
