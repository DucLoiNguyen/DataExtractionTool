import { m } from 'motion/react'
import { AlertTriangle, Download } from 'lucide-react'
import { useMemo, useState } from 'react'
import { Button } from '@/components/ui/button'
import { Tabs, TabsList, TabsTrigger } from '@/components/ui/tabs'
import { api } from '@/lib/api'
import type { ExtractResult } from '@/lib/types'
import { fmt } from '@/lib/utils'
import { CountUp } from './CountUp'
import { ResultTable, type Column } from './ResultTable'

const MONO = true

function previewColumns(mode: ExtractResult['mode']): Column[] {
  if (mode === 'mode1') {
    return [
      { key: 'name', label: 'Họ và tên', min: 180, fr: 2 },
      { key: 'email', label: 'Email', min: 220, fr: 2.5, mono: MONO },
      { key: 'phone', label: 'Điện thoại', min: 120, fr: 1, mono: MONO },
    ]
  }
  return [
    { key: 'name', label: 'Tên tài khoản', min: 170, fr: 1.6 },
    { key: 'email', label: 'Email', min: 220, fr: 2, mono: MONO },
    { key: 'phone', label: 'Số điện thoại', min: 120, fr: 0.9, mono: MONO },
    { key: 'unit', label: 'Đơn vị', min: 220, fr: 2.2 },
    { key: 'to_chuc', label: 'Tổ chức', min: 170, fr: 1.4 },
  ]
}

function issueColumns(result: ExtractResult): Column[] {
  const cols: Column[] = [
    { key: 'stt', label: 'STT', min: 64, fr: 0.4 },
    { key: 'name', label: 'Họ tên (gốc)', min: 160, fr: 1.5 },
    { key: 'email', label: 'Email (gốc)', min: 200, fr: 2, mono: MONO },
    { key: 'phone', label: 'SĐT (gốc)', min: 110, fr: 0.9, mono: MONO },
  ]
  if (result.mode === 'mode2') cols.push({ key: 'unit', label: 'Đơn vị (gốc)', min: 180, fr: 1.6 })
  cols.push({ key: 'reason', label: 'Lý do', min: 140, fr: 1.2 })
  if (result.issues.some((i) => i.note)) cols.push({ key: 'note', label: 'Ghi chú', min: 240, fr: 2.4 })
  if (result.issues.some((i) => i.source)) cols.unshift({ key: 'source', label: 'Tệp nguồn', min: 160, fr: 1.4 })
  return cols
}

export function ResultView({ result }: { result: ExtractResult }) {
  const [tab, setTab] = useState<'prev' | 'iss'>('prev')
  const s = result.stats
  const columns = useMemo(() => (tab === 'prev' ? previewColumns(result.mode) : issueColumns(result)), [tab, result])
  const rows = tab === 'prev' ? result.records : result.issues
  const fileErrors = result.files.filter((f) => f.error).map((f) => `${f.source}: ${f.error}`)
  const warnings = [...fileErrors, ...result.warnings]

  return (
    <m.section
      initial={{ opacity: 0, y: 8 }}
      animate={{ opacity: 1, y: 0 }}
      transition={{ duration: 0.22, ease: 'easeOut' }}
      className="flex flex-col gap-4"
      aria-label="Kết quả trích xuất"
    >
      <div className="rounded-xl border border-border bg-surface p-4">
        <div className="mb-3 flex flex-wrap gap-2">
          <Button asChild variant="ink">
            <a href={api.downloadUrl(result.job_id, 'result')} download>
              <Download className="size-4" aria-hidden />
              Tải tệp kết quả (.xlsx)
            </a>
          </Button>
          {result.has_issues_file && (
            <Button asChild>
              <a href={api.downloadUrl(result.job_id, 'issues')} download>
                <Download className="size-4" aria-hidden />
                Tải danh sách cần kiểm tra (.xlsx)
              </a>
            </Button>
          )}
        </div>
        <div className="flex flex-wrap items-center gap-x-2 gap-y-1" aria-live="polite">
          <span>{fmt(s.total_read)} dòng đọc được</span>
          <span className="text-muted" aria-hidden>→</span>
          <span className="text-warn">−{fmt(s.missing_email)} thiếu email</span>
          <span className="text-muted" aria-hidden>→</span>
          <span className="text-warn">−{fmt(s.invalid_email)} sai định dạng</span>
          <span className="text-muted" aria-hidden>→</span>
          <span className="text-danger">−{fmt(s.duplicates)} trùng email</span>
          <span className="text-muted" aria-hidden>→</span>
          <span>
            <CountUp value={s.final} className="font-serif text-[28px] text-accent" /> dòng kết quả
          </span>
        </div>

        {warnings.map((w, i) => (
          <div key={i} className="mt-2 flex items-start gap-2 rounded-lg bg-warn-bg px-3 py-2 text-sm text-warn">
            <AlertTriangle className="mt-0.5 size-4 shrink-0" aria-hidden />
            <span className="min-w-0 break-words">{w}</span>
          </div>
        ))}

      </div>

      <div className="rounded-xl border border-border bg-surface p-4">
        <Tabs value={tab} onValueChange={(v) => setTab(v as 'prev' | 'iss')}>
          <TabsList className="mb-3 gap-1">
            <TabsTrigger value="prev" className="whitespace-nowrap py-1.5">
              Xem trước ({fmt(result.records.length)})
            </TabsTrigger>
            <TabsTrigger value="iss" className="gap-1.5 whitespace-nowrap py-1.5">
              Cần kiểm tra ({fmt(result.issues.length)})
              {result.issues.length > 0 && <AlertTriangle className="size-3.5" aria-hidden />}
            </TabsTrigger>
          </TabsList>
        </Tabs>
        <m.div key={tab} initial={{ opacity: 0 }} animate={{ opacity: 1 }} transition={{ duration: 0.15 }}>
          <ResultTable resetKey={tab} columns={columns} rows={rows as unknown as Record<string, string>[]} />
        </m.div>
      </div>
    </m.section>
  )
}
