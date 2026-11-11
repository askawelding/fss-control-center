export type SCHFileRole = string

export type SCHPreviewRow = Record<string, string | number | boolean | null>

export type SCHFileValidation = {
  role: SCHFileRole
  label: string
  filename: string
  valid: boolean
  message: string
  sheet_name?: string
  header_row?: number
  data_start_row?: number
  data_rows?: number
  last_data_row?: number
  worksheet_max_row?: number
  scanned_rows?: number
  column_positions?: Record<string, string>
  helper_columns_between_data_fields?: string[]
  helper_header_positions?: Record<string, string>
  helper_formula_samples?: Record<string, {
    sampled_cells: number
    formula_cells_found: number
  }>
  preview?: SCHPreviewRow[]
  warnings: string[]
}

export type SCHValidationResponse = {
  status: 'ok' | 'needs_review'
  valid: boolean
  total_files: number
  valid_files: number
  invalid_files: number
  files: SCHFileValidation[]
  note: string
}

export type SCHReconciliationEntry = {
  category: string
  message: string
  schedule_row: number | null
  master_row: number | null
  old_value: string
  new_value: string
  blocking: boolean
}

export type SCHReconciliationFile = {
  role: SCHFileRole
  label?: string
  filename: string
  valid: boolean
  status?: 'ready' | 'review' | 'blocked'
  safe_to_update?: boolean
  sheet_name?: string
  header_row?: number
  data_rows?: number
  firm_rows?: number
  forecast_rows?: number
  ambiguous_rows?: number
  scanned_rows?: number
  column_positions?: Record<string, string>
  reconciliation?: {
    counts: Record<string, number>
    transition_count: number
    firm_new_non_transition_count: number
    dropped_forecast_count: number
    blocking_count: number
    review_count: number
    info: string[]
    review_messages: string[]
    blocking_messages: string[]
    entries_total: number
    entries_preview: SCHReconciliationEntry[]
    entries_truncated: number
  }
  message?: string
}

export type SCHReconciliationResponse = {
  status: 'ok' | 'needs_review'
  valid: boolean
  master: {
    filename: string
    total_rows: number
    firm_rows: number
    forecast_rows: number
    ambiguous_rows: number
  }
  total_files: number
  valid_files: number
  invalid_files: number
  files: SCHReconciliationFile[]
  note: string
}

const API_BASE_URL = (import.meta.env.VITE_API_BASE_URL || 'http://127.0.0.1:8000').replace(/\/+$/, '')

async function getErrorMessage(response: Response, fallback: string) {
  try {
    const payload = await response.json()
    if (typeof payload.detail === 'string') return payload.detail
    if (Array.isArray(payload.detail)) {
      return payload.detail.map((item: unknown) => {
        if (item && typeof item === 'object' && 'msg' in item) return String((item as { msg: unknown }).msg)
        return JSON.stringify(item)
      }).join('; ') || fallback
    }
  } catch {
    // Response bukan JSON; gunakan pesan default.
  }
  return fallback
}

function appendScheduleFiles(formData: FormData, files: File[]) {
  files.forEach((file) => formData.append('schedule_files', file, file.name))
}

export async function validateSCHFiles(
  files: File[],
): Promise<SCHValidationResponse> {
  const formData = new FormData()
  appendScheduleFiles(formData, files)

  const response = await fetch(`${API_BASE_URL}/api/schedule/validate`, {
    method: 'POST',
    body: formData,
  })
  if (!response.ok) {
    throw new Error(await getErrorMessage(response, 'Gagal memeriksa workbook SCH.'))
  }
  return (await response.json()) as SCHValidationResponse
}

export async function previewSCHReconciliation(
  masterFile: File,
  files: File[],
): Promise<SCHReconciliationResponse> {
  const formData = new FormData()
  formData.append('master_file', masterFile, masterFile.name)
  appendScheduleFiles(formData, files)

  const response = await fetch(`${API_BASE_URL}/api/schedule/preview`, {
    method: 'POST',
    body: formData,
  })
  if (!response.ok) {
    throw new Error(await getErrorMessage(response, 'Gagal membuat preview rekonsiliasi SCH.'))
  }
  return (await response.json()) as SCHReconciliationResponse
}

export async function updateSCHWorkbooks(
  masterFile: File,
  files: File[],
): Promise<Blob> {
  const formData = new FormData()
  formData.append('master_file', masterFile, masterFile.name)
  appendScheduleFiles(formData, files)
  formData.append('confirm_update', 'true')

  const response = await fetch(`${API_BASE_URL}/api/schedule/update`, {
    method: 'POST',
    body: formData,
  })
  if (!response.ok) {
    throw new Error(await getErrorMessage(response, 'Proses Update SCH gagal dijalankan.'))
  }
  return await response.blob()
}
