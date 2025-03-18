package main

import (
	"log"
	"os"
	"regexp"
	"slices"
	"strconv"
	"strings"
	"time"

	ics "github.com/arran4/golang-ical"
	"github.com/google/uuid"
	"github.com/xuri/excelize/v2"
)

var months = []string{
	"januar",
	"februar",
	"marts",
	"april",
	"maj",
	"juni",
	"juli",
	"august",
	"september",
	"oktober",
	"november",
	"december",
}

var ignoredTurns = regexp.MustCompile(`^(DATO|FRI|FERIE|BAGBAG)$`)
var initials = regexp.MustCompile(`(^|\s)SG(\s|$)`)

func main() {
	if len(os.Args) != 2 {
		log.Fatal("Usage: vagtplan.exe <xlsx file>")
	}

	ws, _ := excelize.OpenFile(os.Args[1])
	rows, _ := ws.GetRows("Ark1")
	columns := rows[1]
	dateColumn := slices.Index(columns, "DATO")

	m := regexp.MustCompile(`(\w+) (\d+)`).FindStringSubmatch(rows[0][5])
	if m == nil {
		log.Fatal("Could not parse month")
	}
	month := slices.Index(months, strings.ToLower(m[1])) + 1
	year, _ := strconv.Atoi(m[2])

	cal := ics.NewCalendar()
	loc, _ := time.LoadLocation("Europe/Copenhagen")
	for _, row := range rows[2:] {
		day, err := strconv.Atoi(row[dateColumn])
		if err != nil {
			break
		}
		hours := 0

		start := time.Date(year, time.Month(month), day, 8, 15, 0, 0, loc).UTC()
		if start.Weekday() == time.Saturday || start.Weekday() == time.Sunday {
			start = start.Add(time.Minute * 45)
		}

		for i, cell := range row {
			if ignoredTurns.MatchString(columns[i]) {
				continue
			}
			if !initials.MatchString(cell) {
				continue
			}
			minutes := 0

			switch columns[i] {
			case "BA.VA":
				switch start.Weekday() {
				case time.Friday:
					hours = 24
					minutes = 45
				case time.Sunday:
					hours = 23
					minutes = 15
				default:
					hours = 24
				}
			case "STUEGANG":
				if hours > 8 {
					continue
				}
				hours = 8
			default:
				hours = 8
			}

			event := cal.AddEvent(uuid.NewString())
			event.SetSummary(columns[i])
			event.SetStartAt(start)
			event.SetEndAt(start.Add(time.Duration(hours)*time.Hour + time.Duration(minutes)*time.Minute))
		}
	}

	icsFile := strings.Replace(os.Args[1], ".xlsx", ".ics", 1)
	file, _ := os.Create(icsFile)
	defer file.Close()

	file.Write([]byte(cal.Serialize()))
}
