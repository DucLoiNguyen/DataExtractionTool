import {
  flexRender,
  getCoreRowModel,
  getFilteredRowModel,
  getSortedRowModel,
  useReactTable,
  type ColumnDef,
  type SortingState,
} from '@tanstack/react-table'
import { useVirtualizer } from '@tanstack/react-virtual'
import { ArrowDown, ArrowUp, Search } from 'lucide-react'
import { useDeferredValue, useMemo, useRef, useState } from 'react'
import { Input } from '@/components/ui/input'
import { compareText, fold } from '@/lib/text'
import { cn, fmt } from '@/lib/utils'

export interface Column {
  key: string
  label: string
  /** be rong toi thieu (px) va ti le gian (fr) */
  min: number
  fr: number
  mono?: boolean
}

type Row = Record<string, string> & { _s: string }

interface Props {
  columns: Column[]
  rows: Record<string, string>[]
  /** doi gia tri nay (vd. "prev"/"iss") de dat lai tim kiem + sap xep khi doi bang */
  resetKey: string
}

const ROW_H = 36

/**
 * Bang ket qua: vai nghin dong (Son La ~6000) nhung chi ve ~30 dong dang thay (ao hoa hang).
 * Chuoi tim kiem da bo dau duoc tinh 1 lan khi nhan du lieu, khong tinh lai moi lan go.
 */
export function ResultTable({ columns, rows, resetKey }: Props) {
  return <TableInner key={resetKey} columns={columns} rows={rows} />
}

function TableInner({ columns, rows }: Omit<Props, 'resetKey'>) {
  const [query, setQuery] = useState('')
  const deferred = useDeferredValue(query)
  const [sorting, setSorting] = useState<SortingState>([])
  const scrollRef = useRef<HTMLDivElement>(null)

  const data = useMemo<Row[]>(
    () => rows.map((r) => ({ ...r, _s: fold(columns.map((c) => r[c.key]).join(' ')) })),
    [rows, columns],
  )

  const defs = useMemo<ColumnDef<Row>[]>(
    () =>
      columns.map((c) => ({
        id: c.key,
        accessorKey: c.key,
        header: c.label,
        sortingFn: (a, b, id) => compareText(a.getValue(id), b.getValue(id)),
      })),
    [columns],
  )

  // TanStack Table tra ham khong memo hoa duoc; ket qua chi dung trong thanh phan nay nen an toan
  // eslint-disable-next-line react-hooks/incompatible-library
  const table = useReactTable({
    data,
    columns: defs,
    state: { sorting, globalFilter: fold(deferred.trim()) },
    onSortingChange: setSorting,
    globalFilterFn: (row, _id, value: string) => row.original._s.includes(value),
    getCoreRowModel: getCoreRowModel(),
    getFilteredRowModel: getFilteredRowModel(),
    getSortedRowModel: getSortedRowModel(),
  })

  const tableRows = table.getRowModel().rows
  const virtual = useVirtualizer({
    count: tableRows.length,
    getScrollElement: () => scrollRef.current,
    estimateSize: () => ROW_H,
    overscan: 12,
  })

  const template = columns.map((c) => `minmax(${c.min}px, ${c.fr}fr)`).join(' ')
  const minWidth = columns.reduce((sum, c) => sum + c.min, 0)

  return (
    <div>
      <div className="mb-2 flex flex-wrap items-center justify-between gap-2">
        <div className="relative w-full max-w-xs">
          <Search className="pointer-events-none absolute left-2.5 top-2.5 size-4 text-muted" aria-hidden />
          <Input
            className="pl-8"
            placeholder="Tìm kiếm (không cần gõ dấu)…"
            aria-label="Tìm kiếm trong bảng"
            value={query}
            onChange={(e) => setQuery(e.target.value)}
          />
        </div>
        <span className="text-[13px] text-muted" aria-live="polite">
          {fmt(tableRows.length)}
          {tableRows.length !== rows.length && ` / ${fmt(rows.length)}`} dòng
        </span>
      </div>

      <div ref={scrollRef} className="max-h-[520px] overflow-auto rounded-lg border border-border" role="table">
        <div style={{ minWidth }} role="rowgroup">
          {table.getHeaderGroups().map((hg) => (
            <div
              key={hg.id}
              role="row"
              className="sticky top-0 z-10 grid border-b border-border bg-surface-2"
              style={{ gridTemplateColumns: template }}
            >
              {hg.headers.map((h) => {
                const dir = h.column.getIsSorted()
                return (
                  <button
                    key={h.id}
                    role="columnheader"
                    aria-sort={dir === 'asc' ? 'ascending' : dir === 'desc' ? 'descending' : 'none'}
                    onClick={h.column.getToggleSortingHandler()}
                    className="flex h-9 cursor-pointer items-center gap-1 px-2.5 text-left text-[13px] font-semibold text-muted hover:text-text"
                  >
                    <span className="truncate">{flexRender(h.column.columnDef.header, h.getContext())}</span>
                    {dir === 'asc' && <ArrowUp className="size-3.5 shrink-0" aria-hidden />}
                    {dir === 'desc' && <ArrowDown className="size-3.5 shrink-0" aria-hidden />}
                  </button>
                )
              })}
            </div>
          ))}
        </div>

        <div style={{ height: virtual.getTotalSize(), minWidth, position: 'relative' }} role="rowgroup">
          {virtual.getVirtualItems().map((v) => {
            const row = tableRows[v.index]
            return (
              <div
                key={row.id}
                role="row"
                className="absolute left-0 grid w-full border-b border-border/70 text-sm hover:bg-surface-2/60"
                style={{ gridTemplateColumns: template, height: ROW_H, transform: `translateY(${v.start}px)` }}
              >
                {columns.map((c) => {
                  const val = row.original[c.key] ?? ''
                  return (
                    <div
                      key={c.key}
                      role="cell"
                      title={val}
                      className={cn('flex items-center px-2.5', c.mono && 'font-mono text-[13px]')}
                    >
                      <span className="truncate">{val}</span>
                    </div>
                  )
                })}
              </div>
            )
          })}
        </div>

        {!tableRows.length && (
          <p className="px-3 py-6 text-center text-muted">{rows.length ? 'Không có dòng nào khớp.' : 'Không có dòng nào.'}</p>
        )}
      </div>
    </div>
  )
}
