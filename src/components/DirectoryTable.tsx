/* eslint-disable react-hooks/refs, react-hooks/incompatible-library */

import { startTransition, useEffect, useState } from 'react'
import {
  DndContext,
  PointerSensor,
  closestCenter,
  useSensor,
  useSensors,
  type DragEndEvent,
} from '@dnd-kit/core'
import {
  SortableContext,
  arrayMove,
  horizontalListSortingStrategy,
  useSortable,
} from '@dnd-kit/sortable'
import { CSS } from '@dnd-kit/utilities'
import {
  flexRender,
  getCoreRowModel,
  getFilteredRowModel,
  getSortedRowModel,
  useReactTable,
  type Column,
  type ColumnDef,
  type ColumnFiltersState,
  type HeaderContext,
  type SortingState,
  type VisibilityState,
} from '@tanstack/react-table'
import type { DirectoryPayload, DirectoryUser } from '../shared/types'
import { UserPhoto } from './UserPhoto'

type ColumnMeta = {
  filterVariant?: 'text' | 'photo' | 'none'
  headerClassName?: string
}

function getColumnMeta(column: Column<DirectoryUser, unknown> | HeaderContext<DirectoryUser, unknown>['header']['column']) {
  return (column.columnDef.meta ?? {}) as ColumnMeta
}

const DEFAULT_COLUMN_ORDER = [
  'actions',
  'profilePicture',
  'displayName',
  'status',
  'givenName',
  'surname',
  'jobTitle',
  'department',
  'email',
  'reportsTo',
  'organization',
  'lastSeen',
  'timeZone',
  'workLocation',
  'officeLocation',
]

const DEFAULT_VISIBILITY: VisibilityState = {
  givenName: false,
  surname: false,
  officeLocation: false,
}

function formatDateTime(value: string | null) {
  if (!value) {
    return ''
  }

  return new Intl.DateTimeFormat(undefined, {
    dateStyle: 'medium',
    timeStyle: 'short',
  }).format(new Date(value))
}

function getStatusTone(status: string) {
  const normalized = status.toLowerCase()

  if (normalized.includes('available')) return 'available'
  if (normalized.includes('busy') || normalized.includes('meeting')) return 'busy'
  if (normalized.includes('away')) return 'away'
  if (normalized.includes('offline') || normalized.includes('off work')) return 'offline'
  if (normalized.includes('out of office')) return 'ooo'
  return 'unknown'
}

function ActionButtons({ user }: { user: DirectoryUser }) {
  const canOpen = Boolean(user.email)

  return (
    <div className="action-buttons">
      <button
        className="icon-button"
        disabled={!canOpen}
        onClick={() => {
          if (user.email) {
            void window.teamspy.teams.openAction('chat', user.email)
          }
        }}
        title={canOpen ? `Start a Teams chat with ${user.displayName}` : 'No email available'}
        type="button"
      >
        <svg viewBox="0 0 24 24" aria-hidden="true">
          <path d="M4 5.5A2.5 2.5 0 0 1 6.5 3h11A2.5 2.5 0 0 1 20 5.5v7A2.5 2.5 0 0 1 17.5 15H10l-4.5 4v-4.3A2.5 2.5 0 0 1 4 12.5z" />
        </svg>
      </button>
      <button
        className="icon-button"
        disabled={!canOpen}
        onClick={() => {
          if (user.email) {
            void window.teamspy.teams.openAction('call', user.email)
          }
        }}
        title={canOpen ? `Start a Teams call with ${user.displayName}` : 'No email available'}
        type="button"
      >
        <svg viewBox="0 0 24 24" aria-hidden="true">
          <path d="M6.7 4.2a2 2 0 0 1 2.2-.4l2 1a2 2 0 0 1 1 2.4l-.7 2a1.8 1.8 0 0 0 .4 1.8l1.4 1.4a1.8 1.8 0 0 0 1.8.4l2-.7a2 2 0 0 1 2.4 1l1 2a2 2 0 0 1-.4 2.2l-1 1a3 3 0 0 1-2.7.8c-2.8-.5-5.8-2.3-8.4-4.9C7.3 12.6 5.5 9.6 5 6.8a3 3 0 0 1 .8-2.7z" />
        </svg>
      </button>
    </div>
  )
}

function HeaderCell({ header }: { header: HeaderContext<DirectoryUser, unknown>['header'] }) {
  const sortable = useSortable({ id: header.column.id })
  const canSort = header.column.getCanSort()
  const sortState = header.column.getIsSorted()

  return (
    <th
      ref={sortable.setNodeRef}
      className={`table-header ${getColumnMeta(header.column).headerClassName ?? ''}`}
      style={{
        transform: CSS.Transform.toString(sortable.transform),
        transition: sortable.transition,
        opacity: sortable.isDragging ? 0.7 : 1,
      }}
    >
      <div className="header-shell">
        <button
          className="drag-handle"
          {...sortable.attributes}
          {...sortable.listeners}
          aria-label={`Reorder ${String(header.column.columnDef.header)}`}
          type="button"
        >
          <span />
          <span />
        </button>
        <button
          className={`header-button ${canSort ? 'is-sortable' : ''}`}
          onClick={canSort ? header.column.getToggleSortingHandler() : undefined}
          type="button"
        >
          <span>{flexRender(header.column.columnDef.header, header.getContext())}</span>
          {canSort ? (
            <span className="sort-indicator">
              {sortState === 'asc' ? '↑' : sortState === 'desc' ? '↓' : '↕'}
            </span>
          ) : null}
        </button>
      </div>
    </th>
  )
}

function FilterField({ column }: { column: Column<DirectoryUser, unknown> }) {
  const variant = getColumnMeta(column).filterVariant ?? 'text'

  if (!column.getCanFilter() || variant === 'none') {
    return <div className="filter-spacer" />
  }

  if (variant === 'photo') {
    return (
      <select
        className="filter-input"
        onChange={(event) => {
          const value = event.target.value
          startTransition(() => {
            column.setFilterValue(value === 'all' ? '' : value)
          })
        }}
        value={(column.getFilterValue() as string | undefined) ?? 'all'}
      >
        <option value="all">All</option>
        <option value="with-photo">With photo</option>
        <option value="without-photo">Without photo</option>
      </select>
    )
  }

  return (
    <input
      className="filter-input"
      onChange={(event) => {
        const value = event.target.value
        startTransition(() => {
          column.setFilterValue(value)
        })
      }}
      placeholder="Filter"
      type="text"
      value={(column.getFilterValue() as string | undefined) ?? ''}
    />
  )
}

function defaultTextFilter(
  row: { getValue: (columnId: string) => unknown },
  columnId: string,
  filterValue: string,
) {
  const raw = row.getValue(columnId)
  return String(raw ?? '')
    .toLowerCase()
    .includes(filterValue.toLowerCase())
}

function photoFilter(
  row: { getValue: (columnId: string) => unknown },
  columnId: string,
  filterValue: string,
) {
  if (!filterValue) {
    return true
  }

  const hasPhoto = Boolean(row.getValue(columnId))
  return filterValue === 'with-photo' ? hasPhoto : !hasPhoto
}

function makeColumns(): ColumnDef<DirectoryUser>[] {
  return [
    {
      id: 'actions',
      header: 'Action',
      cell: ({ row }) => <ActionButtons user={row.original} />,
      enableColumnFilter: false,
      enableSorting: false,
      accessorFn: () => '',
      meta: { filterVariant: 'none', headerClassName: 'narrow-col' },
    },
    {
      id: 'profilePicture',
      header: 'Picture',
      accessorFn: (row) => row.id,
      cell: ({ row }) => (
        <div className="avatar-frame">
          <UserPhoto user={row.original} />
        </div>
      ),
      filterFn: photoFilter,
      enableSorting: false,
      meta: { filterVariant: 'photo', headerClassName: 'narrow-col' },
    },
    {
      accessorKey: 'displayName',
      header: 'User Name',
      filterFn: defaultTextFilter,
    },
    {
      accessorKey: 'status',
      header: 'Status',
      filterFn: defaultTextFilter,
      cell: ({ getValue }) => {
        const value = String(getValue() ?? '')
        return (
          <span className={`status-pill tone-${getStatusTone(value)}`}>{value || 'Unknown'}</span>
        )
      },
    },
    {
      accessorKey: 'givenName',
      header: 'First Name',
      filterFn: defaultTextFilter,
    },
    {
      accessorKey: 'surname',
      header: 'Last Name',
      filterFn: defaultTextFilter,
    },
    {
      accessorKey: 'jobTitle',
      header: 'Job Title',
      filterFn: defaultTextFilter,
    },
    {
      accessorKey: 'department',
      header: 'Department',
      filterFn: defaultTextFilter,
    },
    {
      accessorKey: 'email',
      header: 'Email',
      filterFn: defaultTextFilter,
      cell: ({ getValue }) => <span className="mono-cell">{String(getValue() ?? '')}</span>,
    },
    {
      accessorKey: 'reportsTo',
      header: 'Reports To',
      filterFn: defaultTextFilter,
    },
    {
      accessorKey: 'organization',
      header: 'Organization',
      filterFn: defaultTextFilter,
    },
    {
      accessorKey: 'lastSeen',
      header: 'Last Seen',
      filterFn: defaultTextFilter,
      cell: ({ getValue }) => (
        <span>{formatDateTime((getValue() as string | null) ?? null) || '—'}</span>
      ),
    },
    {
      accessorKey: 'timeZone',
      header: 'Timezone',
      filterFn: defaultTextFilter,
    },
    {
      accessorKey: 'workLocation',
      header: 'Work Location',
      filterFn: defaultTextFilter,
    },
    {
      accessorKey: 'officeLocation',
      header: 'Office Location',
      filterFn: defaultTextFilter,
    },
  ]
}

function StatsBar({ users }: { users: DirectoryUser[] }) {
  const summary = users.reduce(
    (accumulator, user) => {
      const tone = getStatusTone(user.status)
      accumulator.total += 1
      if (tone === 'available') accumulator.available += 1
      if (tone === 'busy') accumulator.busy += 1
      if (tone === 'away') accumulator.away += 1
      return accumulator
    },
    { total: 0, available: 0, busy: 0, away: 0 },
  )

  return (
    <div className="stats-strip">
      <div>
        <span>Total</span>
        <strong>{summary.total}</strong>
      </div>
      <div>
        <span>Available</span>
        <strong>{summary.available}</strong>
      </div>
      <div>
        <span>Busy</span>
        <strong>{summary.busy}</strong>
      </div>
      <div>
        <span>Away</span>
        <strong>{summary.away}</strong>
      </div>
    </div>
  )
}

function storageKey(accountId: string) {
  return `teamspy:table-prefs:${accountId}`
}

type StoredTablePrefs = {
  columnFilters?: ColumnFiltersState
  columnOrder?: string[]
  columnVisibility?: VisibilityState
  sorting?: SortingState
}

export function DirectoryTable({
  payload,
  loading,
  onRefresh,
  accountId,
}: {
  payload: DirectoryPayload
  loading: boolean
  onRefresh: () => void
  accountId: string
}) {
  const columns = makeColumns()
  const [sorting, setSorting] = useState<SortingState>([{ id: 'displayName', desc: false }])
  const [columnFilters, setColumnFilters] = useState<ColumnFiltersState>([])
  const [columnOrder, setColumnOrder] = useState(DEFAULT_COLUMN_ORDER)
  const [columnVisibility, setColumnVisibility] =
    useState<VisibilityState>(DEFAULT_VISIBILITY)
  const sensors = useSensors(useSensor(PointerSensor, { activationConstraint: { distance: 8 } }))

  useEffect(() => {
    const raw = localStorage.getItem(storageKey(accountId))

    if (!raw) {
      setSorting([{ id: 'displayName', desc: false }])
      setColumnFilters([])
      setColumnOrder(DEFAULT_COLUMN_ORDER)
      setColumnVisibility(DEFAULT_VISIBILITY)
      return
    }

    try {
      const parsed = JSON.parse(raw) as StoredTablePrefs

      setSorting(parsed.sorting ?? [{ id: 'displayName', desc: false }])
      setColumnFilters(parsed.columnFilters ?? [])
      setColumnOrder(parsed.columnOrder ?? DEFAULT_COLUMN_ORDER)
      setColumnVisibility(parsed.columnVisibility ?? DEFAULT_VISIBILITY)
    } catch {
      setSorting([{ id: 'displayName', desc: false }])
      setColumnFilters([])
      setColumnOrder(DEFAULT_COLUMN_ORDER)
      setColumnVisibility(DEFAULT_VISIBILITY)
    }
  }, [accountId])

  useEffect(() => {
    localStorage.setItem(
      storageKey(accountId),
      JSON.stringify({
        sorting,
        columnFilters,
        columnOrder,
        columnVisibility,
      } satisfies StoredTablePrefs),
    )
  }, [accountId, columnFilters, columnOrder, columnVisibility, sorting])

  const table = useReactTable({
    data: payload.users,
    columns,
    state: {
      sorting,
      columnFilters,
      columnOrder,
      columnVisibility,
    },
    onSortingChange: setSorting,
    onColumnFiltersChange: setColumnFilters,
    onColumnOrderChange: setColumnOrder,
    onColumnVisibilityChange: setColumnVisibility,
    getCoreRowModel: getCoreRowModel(),
    getFilteredRowModel: getFilteredRowModel(),
    getSortedRowModel: getSortedRowModel(),
  })

  const handleDragEnd = (event: DragEndEvent) => {
    const { active, over } = event

    if (!over || active.id === over.id) {
      return
    }

    setColumnOrder((current) => {
      const oldIndex = current.indexOf(String(active.id))
      const newIndex = current.indexOf(String(over.id))
      return arrayMove(current, oldIndex, newIndex)
    })
  }

  return (
    <section className="table-panel">
      <div className="panel-toolbar">
        <div>
          <p className="eyebrow">Directory</p>
          <h2>Presence board</h2>
          <p className="subtle">
            Updated {formatDateTime(payload.loadedAt)}. Filters apply per column.
          </p>
        </div>
        <div className="toolbar-actions">
          <details className="column-menu">
            <summary>Columns</summary>
            <div className="column-menu-list">
              {table
                .getAllLeafColumns()
                .filter((column) => column.id !== 'actions')
                .map((column) => (
                  <label key={column.id}>
                    <input
                      checked={column.getIsVisible()}
                      onChange={column.getToggleVisibilityHandler()}
                      type="checkbox"
                    />
                    <span>{String(column.columnDef.header)}</span>
                  </label>
                ))}
            </div>
          </details>
          <button className="refresh-button" onClick={onRefresh} type="button">
            {loading ? 'Refreshing…' : 'Refresh'}
          </button>
        </div>
      </div>

      <StatsBar users={table.getFilteredRowModel().rows.map((row) => row.original)} />

      {payload.notes.length > 0 ? (
        <div className="notes-list">
          {payload.notes.map((note) => (
            <p key={note}>{note}</p>
          ))}
        </div>
      ) : null}

      <div className="table-wrap">
        <DndContext
          collisionDetection={closestCenter}
          onDragEnd={handleDragEnd}
          sensors={sensors}
        >
          <table>
            <thead>
              {table.getHeaderGroups().map((headerGroup) => (
                <tr key={headerGroup.id}>
                  <SortableContext
                    items={columnOrder}
                    strategy={horizontalListSortingStrategy}
                  >
                    {headerGroup.headers.map((header) => (
                      <HeaderCell header={header} key={header.id} />
                    ))}
                  </SortableContext>
                </tr>
              ))}
              {table.getHeaderGroups().map((headerGroup) => (
                <tr className="filter-row" key={`${headerGroup.id}-filters`}>
                  {headerGroup.headers.map((header) => (
                    <th className="filter-header" key={`${header.id}-filter`}>
                      {header.isPlaceholder ? null : (
                        <FilterField column={header.column} />
                      )}
                    </th>
                  ))}
                </tr>
              ))}
            </thead>
            <tbody>
              {table.getRowModel().rows.map((row) => (
                <tr key={row.id}>
                  {row.getVisibleCells().map((cell) => (
                    <td key={cell.id}>
                      {flexRender(cell.column.columnDef.cell, cell.getContext())}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </DndContext>
      </div>
    </section>
  )
}
