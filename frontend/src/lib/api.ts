import type { JobRequest, JobState, UploadResponse } from './types'

/** Rut thong bao loi tu phan hoi FastAPI ({detail: "..."} hoac danh sach loi kiem tra). */
async function errorMessage(res: Response): Promise<string> {
  try {
    const body = await res.json()
    if (typeof body.detail === 'string') return body.detail
    if (Array.isArray(body.detail)) return body.detail.map((d: { msg?: string }) => d.msg).join('; ')
  } catch {
    /* khong phai JSON */
  }
  return `Lỗi ${res.status}`
}

async function json<T>(res: Response): Promise<T> {
  if (!res.ok) throw new Error(await errorMessage(res))
  return res.json() as Promise<T>
}

export const api = {
  config: () => fetch('/api/config').then((r) => json<{ default_password: string; max_upload_mb: number }>(r)),

  upload(file: File): Promise<UploadResponse> {
    const fd = new FormData()
    fd.append('file', file)
    return fetch('/api/uploads', { method: 'POST', body: fd }).then((r) => json<UploadResponse>(r))
  },

  async deleteUpload(fileId: string) {
    await fetch(`/api/uploads/${fileId}`, { method: 'DELETE' }).catch(() => undefined)
  },

  createJob(req: JobRequest): Promise<{ job_id: string }> {
    return fetch('/api/jobs', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(req),
    }).then((r) => json<{ job_id: string }>(r))
  },

  getJob: (jobId: string): Promise<JobState> => fetch(`/api/jobs/${jobId}`).then((r) => json<JobState>(r)),

  downloadUrl: (jobId: string, kind: 'result' | 'issues') => `/api/download/${jobId}/${kind}`,
}

/** Hoi ket qua cho den khi xong. Dung signal de huy khi nguoi dung doi che do / roi trang. */
export async function waitForJob(jobId: string, signal: AbortSignal, intervalMs = 500): Promise<JobState> {
  for (;;) {
    if (signal.aborted) throw new DOMException('Đã huỷ', 'AbortError')
    const state = await api.getJob(jobId)
    if (state.status !== 'running') return state
    await new Promise((resolve) => setTimeout(resolve, intervalMs))
  }
}
