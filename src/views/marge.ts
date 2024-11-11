import readXlsxFile from 'read-excel-file'
import type { Row, Input } from 'read-excel-file'
import writeXlsxFile from 'write-excel-file'

// File path.
// const file = path.resolve('src', 'input.xlsx')
export async function compareSheet(file: File) {
  const listOfSheet = await readAllSheets(file)
  const todayRows = listOfSheet[0]
  const previousRows = listOfSheet[1]

  const newItems: string[] = findOutNewItem(todayRows, previousRows)
  const onlineItems: string[] = findOutOnlineItem(todayRows, previousRows)

  const mergedRowsReason = mergeReason(todayRows, previousRows)
  const header = mergedRowsReason.shift()
  mergedRowsReason.sort((a, b) => {
    if (a[1] == b[1]) {
      return a[0] < b[0] ? -1 : 1
    } else {
      return a[1] < b[1] ? -1 : 1
    }
  })
  if (header) mergedRowsReason.unshift(header)


  console.log('today offline items: ')
  console.log(JSON.stringify(newItems))

  console.log('today online items: ')
  console.log(JSON.stringify(onlineItems))

  return {
    newItems,
    onlineItems,
    mergedRowsReason,
  }
}

function readAllSheets(file: File) {
  const allOfSheet = Array.from({ length: 2 }).map((_, i) => readXlsxFile(file, { sheet: i + 1 }))
  return Promise.allSettled(allOfSheet).then((results) =>
    results.filter((r) => r.status === 'fulfilled').map((r) => r.value),
  )
}

export function writeToFile(rows) {
  const HEADER_ROW = [
    {
      value: '监控点名称',
      fontWeight: 'bold',
    },
    {
      value: '所在区域',
      fontWeight: 'bold',
    },
    {
      value: 'ip',
      fontWeight: 'bold',
    },
    {
      value: '在线状态',
      fontWeight: 'bold',
    },
    {
      value: '录制状态',
      fontWeight: 'bold',
    },
    {
      value: '状态持续时长',
      fontWeight: 'bold',
    },
    {
      value: '巡检时间',
      fontWeight: 'bold',
    },
    {
      value: '原因',
      fontWeight: 'bold',
    },
  ]

  const dataRows = rows.map((r) => {
    return r.map((cell) => {
      const isDate = typeof cell == 'object'
      return {
        type: isDate ? Date : String,
        value: cell,
        format: isDate ? 'yyyy-MM-dd HH:mm' : undefined,
      }
    })
  })

  const data = [
    // HEADER_ROW,
    ...dataRows,
  ]

  writeXlsxFile(data, {
    // columns, // (optional) column widths, etc.
    fileName: 'output.xlsx',
  })

  // console.log(data)
}

// today offline items:
function findOutNewItem(todayRows: Row[], previousRows: Row[]) {
  const todayName = todayRows.map((t) => t[0])
  const previousName = previousRows.map((t) => t[0])

  const newItems = todayName.filter((t) => !previousName.includes(t))
  return newItems as string[]
}

// today online items:
function findOutOnlineItem(todayRows: Row[], previousRows: Row[]) {
  const todayName = todayRows.map((t) => t[0])
  const previousName = previousRows.map((t) => t[0])

  const onlineItems = previousName.filter((t) => !todayName.includes(t))
  return onlineItems as string[]
}

function mergeReason(todayRows: Row[], previousRows: Row[]) {
  todayRows.forEach((t) => {
    const previousRow = previousRows.find((p) => p[0] === t[0])
    if (previousRow && previousRow[7] && !t[7]) {
      t[7] = previousRow[7]
    }
  })

  return todayRows
}

/*
in excel file,  using visual basic to do this:
 1. read data from sheet 9.11 and sheet 9.10
 2. those sheet has same structure,
 2. the first column is identifier, if it contains multiple same identifier, find out the new item by compare the date in column G, then discard the old items
 3. the date format in column G is yyyy-mm-dd hh:mm,
4. write the result to new sheet
*/
