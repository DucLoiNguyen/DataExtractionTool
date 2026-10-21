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
    <ul className="mt-3 flex flex-col gap-2" aria-label="Danh sách nguồn">
      <AnimatePresence initial={false}>
        {sources.map((s) => (
          <m.li
            key={s.key}
            initial={{ opacity: 0, y: -6 }}
            animate={{ opacity: 1, y: 0 }}
            exit={{ opacity: 0 }}
            transition={{ duration: 0.18, ease: 'easeOut' }}
            className="rounded-lg border border-border bg-surface-2/50 px-3 py-2"
          >
            <div className="flex items-start gap-2">
              <div className="min-w-0 flex-1">
                <div className={cn('truncate', s.skip && 'text-muted line-through')} title={s.fileName}>
                  {s.fileName}
                  {s.sheet != null && <span className="text-muted"> › {s.sheet}</span>}
                </div>
                <div className="text-[13px] text-muted">
                  {s.skip
                    ? `Bỏ qua: trùng ${fmt(s.overlap)}/${fmt(s.emails ?? 0)} email với sheet trước`
                    : m2 && s.emails != null
                      ? `${fmt(s.emails)} email`
                      : null}
                </div>
              </div>
              {m2 && s.sheet != null && (
                <Button size="sm" onClick={() => onChange(s.key, { skip: !s.skip })}>
                  {s.skip ? 'Dùng sheet' : 'Bỏ sheet'}
                </Button>
              )}
              <Button size="sm" variant="ghost" className="px-2" aria-label={`Xoá ${s.fileName}`} onClick={() => onRemove(s.key)}>
                <X className="size-4" aria-hidden />
              </Button>
            </div>
            {m2 && (
              <Input
                className="mt-2"
                value={s.commune}
                disabled={s.skip}
                placeholder="Xã/phường mặc định (nếu đơn vị không ghi tên xã)"
                aria-label={`Xã/phường mặc định cho ${s.fileName}${s.sheet ? ` › ${s.sheet}` : ''}`}
                onChange={(e) => onChange(s.key, { commune: e.target.value })}
              />
            )}
            {m2 && (
              <Input
                className="mt-2"
                value={s.toChuc}
                disabled={s.skip}
                placeholder="Tổ chức mặc định"
                aria-label={`Tổ chức mặc định cho ${s.fileName}${s.sheet ? ` › ${s.sheet}` : ''}`}
                onChange={(e) => onChange(s.key, { toChuc: e.target.value })}
              />
            )}
          </m.li>
        ))}
      </AnimatePresence>
    </ul>
  )
}
