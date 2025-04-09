import { createEvents, EventAttributes } from 'ics'
import * as XLSX from 'sheetjs'

const months = [
  'januar',
  'februar',
  'marts',
  'april',
  'maj',
  'juni',
  'juli',
  'august',
  'september',
  'oktober',
  'november',
  'december',
]

const ignoredTurns = /^(DATO|FRI|FERIE|BAGBAG)$/i
const initials = /(^|\s)SG(\s|$)/i

export function main(path: string): EventAttributes[] {
  const ws = XLSX.readFile(path)
  const rows = XLSX.utils.sheet_to_csv(ws.Sheets['Ark1']).split('\n')
  const columns = rows[1].split(',')
  const dateColumn = columns.indexOf('DATO')

  const m = rows[0].match(/(\w+) (\d+)/)
  if (!m) throw new Error('Could not parse month')
  const month = months.indexOf(m[1].toLowerCase()) + 1
  const year = parseInt(m[2])

  const events: EventAttributes[] = []
  for (const row of rows.slice(2)) {
    const values = row.split(',')
    const day = parseInt(values[dateColumn])
    if (isNaN(day)) break
    let hours = 0

    let start = Temporal.ZonedDateTime.from({
      timeZone: 'Europe/Copenhagen',
      year,
      month,
      day,
      hour: 8,
      minute: 15,
    }).withTimeZone('UTC')
    if ([6, 7].includes(start.dayOfWeek)) {
      start = start.add(Temporal.Duration.from({ minutes: 45 }))
    }

    for (let i = dateColumn; i < columns.length; i++) {
      if (columns[i].match(ignoredTurns)) continue
      if (!values[i].match(initials)) continue
      let minutes = 0

      switch (columns[i]) {
        case 'BA.VA':
          switch (start.dayOfWeek) {
            case 5:
              hours = 24
              minutes = 45
              break
            case 7:
              hours = 23
              minutes = 15
              break
            default:
              hours = 24
          }
          break
        case 'STUEGANG':
          if (hours > 8) continue
        default:
          hours = 8
      }

      events.push({
        start: [start.year, start.month, start.day, start.hour, start.minute],
        startInputType: 'utc',
        duration: { hours, minutes },
        title: columns[i],
      })
    }
  }

  return events
}

if (import.meta.main) {
  const extension = /.xlsx?$/
  if (!Deno.args[0]?.match(extension)) {
    console.error('Usage: vagtplan.exe <excel file>')
    Deno.exit(1)
  }
  Deno.writeTextFileSync(
    Deno.args[0].replace(extension, '.ics'),
    createEvents(main(Deno.args[0])).value ?? '',
  )
}
