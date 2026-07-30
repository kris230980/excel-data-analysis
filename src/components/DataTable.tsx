import React, { useState, useMemo } from 'react';
import {
  ColumnDef,
  flexRender,
  getCoreRowModel,
  getSortedRowModel,
  getPaginationRowModel,
  SortingState,
  useReactTable,
} from '@tanstack/react-table';
import { ChevronDown, ChevronUp, ChevronsLeft, ChevronLeft, ChevronRight, ChevronsRight } from 'lucide-react';

interface DataTableProps {
  columnsData: any[]; // The safeUploadedData
}

export function DataTable({ columnsData }: DataTableProps) {
  const [sorting, setSorting] = useState<SortingState>([]);

  // Convert columnar data to row-based data
  const { data, columns } = useMemo(() => {
    if (!columnsData || columnsData.length === 0) {
      return { data: [], columns: [] };
    }

    const rowCount = columnsData[0].values?.length || 0;
    const rowData: any[] = [];

    for (let i = 0; i < rowCount; i++) {
      const row: any = { id: i };
      columnsData.forEach((col) => {
        row[col.name] = col.values[i];
      });
      rowData.push(row);
    }

    const tableColumns: ColumnDef<any>[] = columnsData.map((col) => ({
      accessorKey: col.name,
      header: col.name,
      cell: (info) => {
        const val = info.getValue();
        if (val instanceof Date) return val.toLocaleDateString();
        return val !== null && val !== undefined ? String(val) : '-';
      },
    }));

    return { data: rowData, columns: tableColumns };
  }, [columnsData]);

  const table = useReactTable({
    data,
    columns,
    getCoreRowModel: getCoreRowModel(),
    onSortingChange: setSorting,
    getSortedRowModel: getSortedRowModel(),
    getPaginationRowModel: getPaginationRowModel(),
    state: {
      sorting,
    },
    initialState: {
      pagination: { pageSize: 15 },
    }
  });

  if (columns.length === 0) {
    return <div className="text-center p-8 text-muted-foreground">No data available to display in table.</div>;
  }

  return (
    <div className="space-y-4">
      <div className="rounded-xl border border-border/50 glass overflow-hidden shadow-sm">
        <div className="overflow-x-auto">
          <table className="w-full text-sm text-left">
            <thead className="bg-muted/50 text-muted-foreground border-b border-border/50">
              {table.getHeaderGroups().map((headerGroup) => (
                <tr key={headerGroup.id}>
                  {headerGroup.headers.map((header) => {
                    return (
                      <th
                        key={header.id}
                        className="px-4 py-3 font-medium cursor-pointer select-none hover:bg-muted/80 transition-colors"
                        onClick={header.column.getToggleSortingHandler()}
                      >
                        <div className="flex items-center gap-1">
                          {flexRender(
                            header.column.columnDef.header,
                            header.getContext()
                          )}
                          {{
                            asc: <ChevronUp className="w-4 h-4 text-primary" />,
                            desc: <ChevronDown className="w-4 h-4 text-primary" />,
                          }[header.column.getIsSorted() as string] ?? <span className="w-4 h-4" />}
                        </div>
                      </th>
                    );
                  })}
                </tr>
              ))}
            </thead>
            <tbody>
              {table.getRowModel().rows?.length ? (
                table.getRowModel().rows.map((row) => (
                  <tr
                    key={row.id}
                    className="border-b border-border/30 last:border-0 hover:bg-muted/30 transition-colors"
                  >
                    {row.getVisibleCells().map((cell) => (
                      <td key={cell.id} className="px-4 py-2.5 whitespace-nowrap">
                        {flexRender(cell.column.columnDef.cell, cell.getContext())}
                      </td>
                    ))}
                  </tr>
                ))
              ) : (
                <tr>
                  <td colSpan={columns.length} className="h-24 text-center">
                    No results.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
      
      {/* Pagination Controls */}
      <div className="flex items-center justify-between px-2">
        <div className="text-sm text-muted-foreground">
          Showing {table.getState().pagination.pageIndex * table.getState().pagination.pageSize + 1} to{' '}
          {Math.min(
            (table.getState().pagination.pageIndex + 1) * table.getState().pagination.pageSize,
            table.getFilteredRowModel().rows.length
          )}{' '}
          of {table.getFilteredRowModel().rows.length} rows
        </div>
        <div className="flex items-center space-x-2">
          <button
            className="p-1 rounded hover:bg-muted disabled:opacity-50 transition-colors"
            onClick={() => table.setPageIndex(0)}
            disabled={!table.getCanPreviousPage()}
          >
            <ChevronsLeft className="w-5 h-5" />
          </button>
          <button
            className="p-1 rounded hover:bg-muted disabled:opacity-50 transition-colors"
            onClick={() => table.previousPage()}
            disabled={!table.getCanPreviousPage()}
          >
            <ChevronLeft className="w-5 h-5" />
          </button>
          <span className="text-sm font-medium px-2 py-1 bg-muted/50 rounded-md">
            Page {table.getState().pagination.pageIndex + 1} of {table.getPageCount()}
          </span>
          <button
            className="p-1 rounded hover:bg-muted disabled:opacity-50 transition-colors"
            onClick={() => table.nextPage()}
            disabled={!table.getCanNextPage()}
          >
            <ChevronRight className="w-5 h-5" />
          </button>
          <button
            className="p-1 rounded hover:bg-muted disabled:opacity-50 transition-colors"
            onClick={() => table.setPageIndex(table.getPageCount() - 1)}
            disabled={!table.getCanNextPage()}
          >
            <ChevronsRight className="w-5 h-5" />
          </button>
        </div>
      </div>
    </div>
  );
}
