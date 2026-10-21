import { m } from 'motion/react'
import { FileSpreadsheet, ListChecks, UploadCloud } from 'lucide-react'

const fade = {
  initial: { opacity: 0 },
  animate: { opacity: 1 },
  exit: { opacity: 0 },
  transition: { duration: 0.15 },
}

export function EmptyState() {
  const steps = [
    { icon: UploadCloud, text: 'Kéo thả tệp danh sách vào khung bên trái.' },
    { icon: ListChecks, text: 'Kiểm tra nguồn, Tổ chức, rồi bấm Trích xuất.' },
    { icon: FileSpreadsheet, text: 'Xem kết quả tại đây và tải tệp Excel.' },
  ]
  return (
    <m.section {...fade} className="rounded-xl border border-dashed border-border px-6 py-10 text-muted">
      <h2 className="m-0 mb-3 text-[15px] font-semibold text-text">Kết quả sẽ hiện ở đây</h2>
      <ol className="m-0 flex list-none flex-col gap-3 p-0">
        {steps.map(({ icon: Icon, text }, i) => (
          <li key={i} className="flex items-center gap-3">
            <span className="flex size-8 shrink-0 items-center justify-center rounded-full bg-surface-2 text-text">
              <Icon className="size-4" strokeWidth={1.75} aria-hidden />
            </span>
            {text}
          </li>
        ))}
      </ol>
    </m.section>
  )
}

/** Phan tram: so nguon da xong that + uoc luong cho nguon dang doc (tiem can, khong bao gio cham muc ke tiep). */
export function estimatePercent(done: number, total: number, elapsed: number): number {
  const slice = 1 / Math.max(total, 1)
  const creep = 0.95 * (1 - Math.exp(-elapsed / 8))
  return Math.min(99, Math.round((done / Math.max(total, 1) + creep * slice) * 100))
}

export function RunningCard({
  elapsed,
  sources,
  progress,
}: {
  elapsed: number
  sources: number
  progress: { done: number; total: number } | null
}) {
  const total = progress?.total ?? Math.max(sources, 1)
  const done = progress?.done ?? 0
  const pct = estimatePercent(done, total, elapsed)
  return (
    <m.section {...fade} className="rounded-xl border border-border bg-surface p-6" aria-live="polite">
      <div className="flex items-baseline justify-between gap-3">
        <h2 className="m-0 text-[15px] font-semibold">Đang trích xuất {sources > 1 ? `${sources} nguồn` : 'tệp'}…</h2>
        <span className="font-serif text-2xl tabular-nums">{pct}%</span>
      </div>
      <p className="mb-4 mt-1 text-[13px] text-muted">
        {total > 1 ? `Đã xong ${done}/${total} nguồn · ` : ''}Đã chạy {elapsed} giây. PDF lớn có thể mất vài chục giây, bạn có thể để trang này mở.
      </p>
      <div
        className="h-1.5 overflow-hidden rounded-full bg-surface-2"
        role="progressbar"
        aria-label="Tiến độ xử lý"
        aria-valuemin={0}
        aria-valuemax={100}
        aria-valuenow={pct}
      >
        <div className="h-full rounded-full bg-accent transition-[width] duration-500 motion-reduce:transition-none" style={{ width: `${pct}%` }} />
      </div>
    </m.section>
  )
}
