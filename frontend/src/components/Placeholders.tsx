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

export function RunningCard({ elapsed, sources }: { elapsed: number; sources: number }) {
  return (
    <m.section {...fade} className="rounded-xl border border-border bg-surface p-6" aria-live="polite">
      <h2 className="m-0 text-[15px] font-semibold">Đang trích xuất {sources > 1 ? `${sources} nguồn` : 'tệp'}…</h2>
      <p className="mb-4 mt-1 text-[13px] text-muted">
        Đã chạy {elapsed} giây. PDF lớn có thể mất vài chục giây, bạn có thể để trang này mở.
      </p>
      <div className="h-1.5 overflow-hidden rounded-full bg-surface-2" role="progressbar" aria-label="Đang xử lý">
        <div className="h-full w-1/3 rounded-full bg-accent motion-safe:animate-[slide_1.4s_ease-in-out_infinite]" />
      </div>
    </m.section>
  )
}
