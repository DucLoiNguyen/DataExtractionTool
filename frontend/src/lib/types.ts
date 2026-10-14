export type Mode = 'mode1' | 'mode2'

export interface SheetInfo {
  sheet: string
  emails: number
  overlap_with_previous: number
  suggest_skip: boolean
}

export interface UploadResponse {
  file_id: string
  name: string
  size: number
  sheets: SheetInfo[]
  suggestion: string
  error: string | null
}

export interface RecordRow {
  name: string
  email: string
  phone: string
  unit: string
  to_chuc: string
  source: string
}

export interface IssueRow {
  source: string
  stt: string
  name: string
  email: string
  phone: string
  unit: string
  reason: string
  note: string
}

export interface PipelineStats {
  total_read: number
  missing_email: number
  invalid_email: number
  duplicates: number
  final: number
}

export interface ExtractResult {
  job_id: string
  mode: Mode
  stats: PipelineStats
  records: RecordRow[]
  issues: IssueRow[]
  warnings: string[]
  files: { source: string; total_read: number; final_count: number; error: string | null }[]
  has_issues_file: boolean
}

export type JobState =
  | { status: 'running' }
  | { status: 'done'; result: ExtractResult }
  | { status: 'error'; error: string }

/** Mot dong trong danh sach nguon: 1 tep, hoac 1 sheet cua tep Excel nhieu sheet. */
export interface SourceRow {
  key: string
  fileId: string
  fileName: string
  sheet: string | null
  emails: number | null
  skip: boolean
  overlap: number
  toChuc: string
}

export interface JobRequest {
  mode: Mode
  sources: { file_id: string; sheet: string | null; to_chuc: string }[]
  dedupe: boolean
  password: string
}
