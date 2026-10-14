import { AnimatePresence } from 'motion/react'
import { AlertCircle, Loader2, Play } from 'lucide-react'
import { useCallback, useEffect, useRef, useState } from 'react'
import { ResultView } from '@/components/ResultView'
import { SourceList } from '@/components/SourceList'
import { UploadZone } from '@/components/UploadZone'
import { Button } from '@/components/ui/button'
import { Checkbox } from '@/components/ui/checkbox'
import { Input } from '@/components/ui/input'
import { Tabs, TabsList, TabsTrigger } from '@/components/ui/tabs'
import { api, waitForJob } from '@/lib/api'
import type { ExtractResult, Mode, SourceRow, UploadResponse } from '@/lib/types'

const HINTS: Record<Mode, string> = {
  mode1: 'Chế độ 1 xử lý 1 tệp mỗi lần · PDF, Word (.docx/.doc), Excel (.xlsx/.xls), CSV',
  mode2: 'Có thể chọn nhiều tệp; tệp Excel nhiều sheet được tách theo sheet · PDF, Word, Excel, CSV',
}

function rowsFromUpload(up: UploadResponse, mode: Mode, toChuc: string): SourceRow[] {
  const base = { fileId: up.file_id, fileName: up.name, toChuc }
  if (mode === 'mode2' && up.sheets.length) {
    return up.sheets.map((s) => ({
      ...base,
      key: `${up.file_id}:${s.sheet}`,
      sheet: s.sheet,
      emails: s.emails,
      skip: s.suggest_skip,
      overlap: s.overlap_with_previous,
    }))
  }
  return [{ ...base, key: up.file_id, sheet: null, emails: null, skip: false, overlap: 0 }]
}

export default function App() {
  const [mode, setMode] = useState<Mode>('mode1')
  const [sources, setSources] = useState<SourceRow[]>([])
  const [uploading, setUploading] = useState(0)
  const [running, setRunning] = useState(false)
  const [result, setResult] = useState<ExtractResult | null>(null)
  const [errors, setErrors] = useState<string[]>([])
  const [bulk, setBulk] = useState('')
  const [password, setPassword] = useState('')
  const [dedupe, setDedupe] = useState(true)
  const abortRef = useRef<AbortController | null>(null)

  useEffect(() => {
    api.config().then((c) => setPassword(c.default_password)).catch(() => undefined)
    return () => abortRef.current?.abort()
  }, [])

  const discard = useCallback((rows: SourceRow[]) => {
    for (const id of new Set(rows.map((r) => r.fileId))) void api.deleteUpload(id)
  }, [])

  const changeMode = (next: Mode) => {
    if (next === mode) return
    abortRef.current?.abort()
    discard(sources)
    setMode(next)
    setSources([])
    setResult(null)
    setErrors([])
    setRunning(false)
  }

  const addFiles = async (picked: File[]) => {
    const files = mode === 'mode1' ? picked.slice(0, 1) : picked
    setErrors([])
    setResult(null)
    if (mode === 'mode1') {
      discard(sources)
      setSources([])
    }
    setUploading((n) => n + files.length)
    await Promise.all(
      files.map(async (file) => {
        try {
          const up = await api.upload(file)
          if (up.error) setErrors((e) => [...e, `${file.name}: ${up.error}`])
          setSources((prev) => [...prev, ...rowsFromUpload(up, mode, bulk.trim() || up.suggestion)])
        } catch (err) {
          setErrors((e) => [...e, `${file.name}: ${(err as Error).message}`])
        } finally {
          setUploading((n) => n - 1)
        }
      }),
    )
  }

  const patchSource = (key: string, patch: Partial<SourceRow>) =>
    setSources((prev) => prev.map((s) => (s.key === key ? { ...s, ...patch } : s)))

  const removeSource = (key: string) => {
    const gone = sources.find((s) => s.key === key)
    const rest = sources.filter((s) => s.key !== key)
    setSources(rest)
    if (gone && !rest.some((s) => s.fileId === gone.fileId)) void api.deleteUpload(gone.fileId)
  }

  const applyBulk = () => setSources((prev) => prev.map((s) => (s.skip ? s : { ...s, toChuc: bulk.trim() })))

  const active = sources.filter((s) => !s.skip)

  const run = async () => {
    abortRef.current?.abort()
    const ctrl = new AbortController()
    abortRef.current = ctrl
    setRunning(true)
    setErrors([])
    setResult(null)
    try {
      const { job_id } = await api.createJob({
        mode,
        sources: active.map((s) => ({ file_id: s.fileId, sheet: s.sheet, to_chuc: s.toChuc })),
        dedupe,
        password,
      })
      const state = await waitForJob(job_id, ctrl.signal)
      if (state.status === 'done') setResult(state.result)
      else if (state.status === 'error') setErrors([state.error])
    } catch (err) {
      if ((err as Error).name !== 'AbortError') setErrors([(err as Error).message])
    } finally {
      if (abortRef.current === ctrl) setRunning(false)
    }
  }

  return (
    <div className="mx-auto max-w-[1180px] px-4 pb-16 pt-6 md:px-6">
      <header className="mb-5">
        <h1 className="m-0 font-serif text-3xl font-normal">Tạo danh sách tài khoản</h1>
        <p className="mt-1 text-muted">
          Nền tảng Trí tuệ nhân tạo hỗ trợ công vụ — trích xuất Họ tên / Email / SĐT từ công văn, Word, Excel, CSV.
        </p>
      </header>

      <Tabs value={mode} onValueChange={(v) => changeMode(v as Mode)} className="mb-4">
        <TabsList aria-label="Chế độ">
          <TabsTrigger value="mode1">Chế độ 1 · Tên, Email, SĐT</TabsTrigger>
          <TabsTrigger value="mode2">Chế độ 2 · Thêm Đơn vị / Tổ chức</TabsTrigger>
        </TabsList>
      </Tabs>

      <section className="mb-4 rounded-xl border border-border bg-surface p-4" aria-labelledby="h-src">
        <h2 id="h-src" className="mb-2.5 text-[15px] font-semibold">1. Chọn tệp nguồn</h2>
        <UploadZone multiple={mode === 'mode2'} hint={HINTS[mode]} busy={uploading > 0} onFiles={addFiles} />
        <SourceList mode={mode} sources={sources} onChange={patchSource} onRemove={removeSource} />
        {mode === 'mode2' && sources.length > 0 && (
          <div className="mt-4 flex flex-wrap items-end gap-2">
            <label className="min-w-52 flex-1 text-[13px] text-muted">
              Gán Tổ chức mặc định cho tất cả nguồn
              <Input
                className="mt-1 text-text"
                value={bulk}
                placeholder="vd: Cần Thơ, Lào Cai, Viện Năng lượng nguyên tử Việt Nam"
                onChange={(e) => setBulk(e.target.value)}
              />
            </label>
            <Button onClick={applyBulk}>Gán</Button>
          </div>
        )}
      </section>

      <section className="mb-4 rounded-xl border border-border bg-surface p-4" aria-labelledby="h-opt">
        <h2 id="h-opt" className="mb-2.5 text-[15px] font-semibold">2. Tuỳ chọn</h2>
        <div className="flex flex-wrap items-end gap-4">
          <label className="min-w-52 flex-1 text-[13px] text-muted">
            Mật khẩu mặc định
            <Input className="mt-1 text-text" value={password} onChange={(e) => setPassword(e.target.value)} />
          </label>
          <label className="flex h-9 cursor-pointer items-center gap-2 text-text">
            <Checkbox checked={dedupe} onCheckedChange={(c) => setDedupe(c === true)} aria-label="Lọc trùng email" />
            Lọc trùng email
          </label>
          <Button variant="primary" onClick={run} disabled={running || uploading > 0 || active.length === 0}>
            {running ? <Loader2 className="size-4 animate-spin motion-reduce:animate-none" aria-hidden /> : <Play className="size-4" aria-hidden />}
            {running ? 'Đang trích xuất…' : 'Trích xuất'}
          </Button>
        </div>
        <div className="mt-2 min-h-5 text-[13px] text-muted" aria-live="polite">
          {uploading > 0 && 'Đang đọc tệp…'}
          {running && 'Đang trích xuất… (PDF lớn có thể mất vài chục giây)'}
          {!running && uploading === 0 && result && 'Hoàn tất trích xuất.'}
        </div>
      </section>

      {errors.map((msg, i) => (
        <div key={i} role="alert" className="mb-3 flex items-start gap-2 whitespace-pre-wrap rounded-lg bg-danger-bg px-3 py-2.5 text-sm text-danger">
          <AlertCircle className="mt-0.5 size-4 shrink-0" aria-hidden />
          <span>{msg}</span>
        </div>
      ))}

      <AnimatePresence>{result && <ResultView key={result.job_id} result={result} />}</AnimatePresence>
    </div>
  )
}
