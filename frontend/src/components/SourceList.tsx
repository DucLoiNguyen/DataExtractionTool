import { AnimatePresence, m } from 'motion/react'
import { X } from 'lucide-react'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { fmt } from '@/lib/utils'
import { cn } from '@/lib/utils'
import type { Mode, SourceRow } from '@/lib/types'

interface Props {
  mode: Mode
  sources: SourceRow[]
  onChange: (key: string, patch: Partial<SourceRow>) => void
  onRemove: (key: string) => void
}

export function SourceList({ mode, sources, onChange, onRemove }: Props) {
  if (!sources.length) return null
  const m2 = mode === 'mode2'
  return (
    <ul className="mt-4 flex flex-col gap-2" aria-label="Danh sách nguồn">
      <AnimatePresence initial={false}>
        {sources.map((s) => (
          <m.li
            key={s.key}
            initial={{ opacity: 0, y: -6 }}
            animate={{ opacity: 1, y: 0 }}
            exit={{ opacity: 0 }}
            transition={{ duration: 0.18, ease: 'easeOut' }}
            className="grid items-center gap-x-3 gap-y-2 rounded-lg border border-border bg-surface-2/50 px-3 py-2 md:grid-cols-[minmax(0,1fr)_auto_minmax(0,16rem)_auto]"
          >
            <div className="min-w-0">
              <div className={cn('truncate', s.skip && 'text-muted line-through')} title={s.fileName}>
                {s.fileName}
                {s.sheet != null && <span className="text-muted"> › {s.sheet}</span>}
              </div>
              {s.skip && (
                <div className="text-[13px] text-muted">
                  Bỏ qua: trùng {fmt(s.overlap)}/{fmt(s.emails ?? 0)} email với sheet trước
                </div>
              )}
            </div>
            {m2 && (
              <span className="text-[13px] text-muted">{s.emails == null ? '—' : `${fmt(s.emails)} email`}</span>
            )}
            {m2 && (
              <Input
                value={s.toChuc}
                disabled={s.skip}
                placeholder="Tổ chức mặc định"
                aria-label={`Tổ chức mặc định cho ${s.fileName}${s.sheet ? ` › ${s.sheet}` : ''}`}
                onChange={(e) => onChange(s.key, { toChuc: e.target.value })}
              />
            )}
            <div className={cn('flex items-center gap-1.5', !m2 && 'md:col-start-4')}>
              {m2 && s.sheet != null && (
                <Button size="sm" onClick={() => onChange(s.key, { skip: !s.skip })}>
                  {s.skip ? 'Dùng sheet' : 'Bỏ sheet'}
                </Button>
              )}
              <Button size="sm" variant="ghost" aria-label={`Xoá ${s.fileName}`} onClick={() => onRemove(s.key)}>
                <X className="size-4" aria-hidden />
                Xoá
              </Button>
            </div>
          </m.li>
        ))}
      </AnimatePresence>
    </ul>
  )
}
