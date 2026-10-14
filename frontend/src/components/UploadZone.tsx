import { UploadCloud } from 'lucide-react'
import { useRef, useState, type DragEvent, type KeyboardEvent } from 'react'
import { cn } from '@/lib/utils'

interface Props {
  multiple: boolean
  hint: string
  busy: boolean
  onFiles: (files: File[]) => void
}

const ACCEPT = '.pdf,.docx,.doc,.xlsx,.xls,.csv'

export function UploadZone({ multiple, hint, busy, onFiles }: Props) {
  const inputRef = useRef<HTMLInputElement>(null)
  const [over, setOver] = useState(false)

  const open = () => inputRef.current?.click()
  const onKey = (e: KeyboardEvent) => {
    if (e.key === 'Enter' || e.key === ' ') {
      e.preventDefault()
      open()
    }
  }
  const onDrop = (e: DragEvent) => {
    e.preventDefault()
    setOver(false)
    const files = [...e.dataTransfer.files]
    if (files.length) onFiles(files)
  }

  return (
    <div
      role="button"
      tabIndex={0}
      aria-label="Chọn hoặc kéo thả tệp nguồn"
      aria-busy={busy}
      onClick={open}
      onKeyDown={onKey}
      onDragOver={(e) => {
        e.preventDefault()
        setOver(true)
      }}
      onDragLeave={() => setOver(false)}
      onDrop={onDrop}
      className={cn(
        'flex cursor-pointer flex-col items-center gap-1.5 rounded-xl border-2 border-dashed border-border px-4 py-6 text-center text-muted',
        'transition-colors duration-150 hover:border-accent/60 motion-reduce:transition-none',
        over && 'border-accent bg-accent/10',
      )}
    >
      <UploadCloud className="size-6" strokeWidth={1.75} aria-hidden />
      <span className="text-text">Kéo thả tệp vào đây hoặc bấm để chọn</span>
      <span className="text-[13px]">{hint}</span>
      <input
        ref={inputRef}
        type="file"
        hidden
        multiple={multiple}
        accept={ACCEPT}
        onChange={(e) => {
          const files = [...(e.target.files ?? [])]
          e.target.value = ''
          if (files.length) onFiles(files)
        }}
      />
    </div>
  )
}
