// Package xlsx writes to xlsx files using github.com/xuri/excelize/v2
//
// The struct field's xlsx tag contains its heading name and comma-separated options
// fields with no xlsx tag are not written to the xlsx
//
// Examples of struct field tags and their meanings:
//
//	// Field is written under column heading "Name" in the xlsx
//	Field string `xls:"Name"`
//
//	// Field is written under column heading "Name" in the xlsx as a whole number
//	Field string `xls:"Number,{\"number_format\":1}"`
//
// see https://github.com/qax-os/excelize/blob/master/styles.go for format styles
// only number_format is support for now
package xlsx

import (
	"fmt"
	"log"
	"math"
	"reflect"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

// Xlsx is our type
type Xlsx struct {
	x            *excelize.File
	columnWidths map[string][]int
	wroteHeader  map[string]bool
	rowNum       map[string]int
	headings     map[string][]string
	styles       map[string][]*excelize.Style
}

// NewXlsx creates new Xlsx
func NewXlsx() *Xlsx {
	return &Xlsx{
		x:            excelize.NewFile(),
		columnWidths: make(map[string][]int),
		wroteHeader:  make(map[string]bool),
		rowNum:       make(map[string]int),
		headings:     make(map[string][]string),
		styles:       make(map[string][]*excelize.Style),
	}
}

// CreateSheet adds a new sheet to the workbook named sheetName
func (xlsx *Xlsx) CreateSheet(sheetName string) error {
	_, err := xlsx.x.NewSheet(sheetName)
	return err
}

// RemoveSheet delete the sheet named sheetName from the workbook
func (xlsx *Xlsx) RemoveSheet(sheetName string) error {
	return xlsx.x.DeleteSheet(sheetName)
}

func innerGetRowHeadings(row interface{}) []string {
	fields := reflect.ValueOf(row)
	t := reflect.TypeOf(row)

	// iterate over all available fields and read the tag value to get what to put in the header
	headings := make([]string, 0, t.NumField())
	for i := 0; i < t.NumField(); i++ {
		field := t.Field(i)
		f := fields.Field(i)
		k := f.Kind()
		_, isTime := f.Interface().(time.Time)

		if isTime || (k != reflect.Struct) {
			tag := field.Tag.Get("xls")
			if tag != "" {
				hs := strings.Split(tag, ",")
				headings = append(headings, hs[0])
			}
		} else {
			headings = append(headings, innerGetRowHeadings(f.Interface())...)
		}
	}

	return headings
}

func (xlsx *Xlsx) getRowHeadings(sheetName string, row interface{}) []string {
	if xlsx.headings[sheetName] != nil {
		return xlsx.headings[sheetName]
	}

	headings := innerGetRowHeadings(row)

	// cache for later
	xlsx.headings[sheetName] = headings

	return headings
}

func innerGetRowStyles(row interface{}) []*excelize.Style {
	fields := reflect.ValueOf(row)
	t := reflect.TypeOf(row)

	// iterate over all available fields and read the tag value to get the style
	styles := make([]*excelize.Style, 0, t.NumField())
	for i := 0; i < t.NumField(); i++ {
		field := t.Field(i)
		f := fields.Field(i)
		k := f.Kind()
		_, isTime := f.Interface().(time.Time)

		if isTime || (k != reflect.Struct) {
			tag := field.Tag.Get("xls")
			if tag != "" {
				hs := strings.Split(tag, ",")
				if len(hs) > 1 {
					// only support number format for now
					var n int
					_, err := fmt.Fscanf(strings.NewReader(hs[1]), "{\"number_format\":%d}", &n)
					if err != nil {
						log.Printf("%+v", err)
						return nil
					}

					styles = append(styles, &excelize.Style{NumFmt: n})
				} else {
					styles = append(styles, nil)
				}
			}
		} else {
			styles = append(styles, innerGetRowStyles(f.Interface())...)
		}
	}

	return styles
}

func (xlsx *Xlsx) getRowStyles(sheetName string, row interface{}) []*excelize.Style {
	if xlsx.styles[sheetName] != nil {
		return xlsx.styles[sheetName]
	}

	styles := innerGetRowStyles(row)

	// cache for later
	xlsx.styles[sheetName] = styles

	return styles
}

func innerGetRowData(row interface{}) []interface{} {
	fields := reflect.ValueOf(row)
	t := reflect.TypeOf(row)

	// iterate over all available fields and read the value
	values := make([]interface{}, 0)
	for i := 0; i < t.NumField(); i++ {
		field := t.Field(i)
		f := fields.Field(i)
		k := f.Kind()
		_, isTime := f.Interface().(time.Time)

		if isTime || (k != reflect.Struct) {
			tag := field.Tag.Get("xls")
			if tag != "" {
				if f.Kind() == reflect.Pointer {
					v := f.Elem()
					if v.IsValid() {
						values = append(values, f.Elem().Interface())
					} else {
						values = append(values, nil)
					}
				} else {
					values = append(values, f.Interface())
				}
			}
		} else {
			values = append(values, innerGetRowData(f.Interface())...)
		}
	}

	return values
}

func (xlsx *Xlsx) getRowData(row interface{}) []interface{} {
	return innerGetRowData(row)
}

// WriteRow appends row contents to sheet named sheetName, creates a header row if this is the first row written to xlsx
func (xlsx *Xlsx) WriteRow(sheetName string, row interface{}) error {
	t := reflect.TypeOf(row)
	if t.Kind() != reflect.Struct {
		err := fmt.Errorf("wrong kind %s", t.Kind())
		log.Printf("%+v", err)
		return err
	}

	rowData := xlsx.getRowData(row)
	if !xlsx.wroteHeader[sheetName] {
		columns := xlsx.getRowHeadings(sheetName, row)

		// accumulate columnWidths
		xlsx.columnWidths[sheetName] = make([]int, len(rowData))

		// header styling
		styleHeader, err := xlsx.x.NewStyle(&excelize.Style{Font: &excelize.Font{Bold: true}})
		if err != nil {
			log.Printf("%+v", err)
			return err
		}

		// write out header row
		xlsx.rowNum[sheetName]++
		for i, columnName := range columns {
			n, err := excelize.CoordinatesToCellName(i+1, xlsx.rowNum[sheetName])
			if err != nil {
				log.Printf("%+v", err)
				return err
			}

			err = xlsx.x.SetCellValue(sheetName, n, columnName)
			if err != nil {
				log.Printf("%+v", err)
				return err
			}

			xlsx.columnWidths[sheetName][i] = max(xlsx.columnWidths[sheetName][i], len(fmt.Sprintf("%v", columnName)))

			err = xlsx.x.SetCellStyle(sheetName, n, n, styleHeader)
			if err != nil {
				log.Printf("%+v", err)
				return err
			}
		}

		// freeze header row
		err = xlsx.x.SetPanes(sheetName, &excelize.Panes{
			Freeze:      true,
			Split:       false,
			XSplit:      0,
			YSplit:      1,
			TopLeftCell: "A2",
			ActivePane:  "bottomLeft",
			Selection: []excelize.Selection{
				{SQRef: "A2", ActiveCell: "A2", Pane: "bottomLeft"},
			},
		})
		if err != nil {
			log.Printf("%+v", err)
			return err
		}

		xlsx.wroteHeader[sheetName] = true
	}

	// write out cells
	xlsx.rowNum[sheetName]++
	styles := xlsx.getRowStyles(sheetName, row)
	for i, value := range rowData {
		n, err := excelize.CoordinatesToCellName(i+1, xlsx.rowNum[sheetName])
		if err != nil {
			log.Printf("%+v", err)
			return err
		}

		err = xlsx.x.SetCellValue(sheetName, n, value)
		if err != nil {
			log.Printf("%+v", err)
			return err
		}

		xlsx.columnWidths[sheetName][i] = max(xlsx.columnWidths[sheetName][i], len(fmt.Sprintf("%v", value)))

		// styling
		if styles[i] != nil {
			style, err := xlsx.x.NewStyle(styles[i])
			if err != nil {
				log.Printf("%+v", err)
				return err
			}

			err = xlsx.x.SetCellStyle(sheetName, n, n, style)
			if err != nil {
				log.Printf("%+v", err)
				return err
			}
		}
	}

	return nil
}

// closeSheet finializes a sheet that was written to with WriteRow
func (xlsx *Xlsx) closeSheet(sheetName string) error {
	// set all column widths as best we can
	for i, width := range xlsx.columnWidths[sheetName] {
		n, err := excelize.ColumnNumberToName(i + 1)
		if err != nil {
			log.Printf("%+v", err)
			return err
		}

		// add padding for filter selection, extra for filter button
		err = xlsx.x.SetColWidth(sheetName, n, n, char2width(width+3))
		if err != nil {
			log.Printf("%+v", err)
			return err
		}
	}

	// turn on autofilter for each column with data in it
	n, err := excelize.ColumnNumberToName(len(xlsx.columnWidths[sheetName]))
	if err != nil {
		log.Printf("%+v", err)
		return err
	}
	err = xlsx.x.AutoFilter(sheetName, fmt.Sprintf("A1:%s1", n), []excelize.AutoFilterOptions{})
	if err != nil {
		log.Printf("%+v", err)
		return err
	}

	return nil
}

// Close finializes xlsx and writes to fn
func (xlsx *Xlsx) Close(fn string) error {
	// close all the sheets written to with WriteRow
	for sheetName := range xlsx.wroteHeader {
		err := xlsx.closeSheet(sheetName)
		if err != nil {
			log.Printf("%+v", err)
			return err
		}
	}

	// save xlsx to file
	err := xlsx.x.SaveAs(fn)
	if err != nil {
		log.Printf("%+v", err)
		return err
	}

	return nil
}

// char2width calculates the cell width based on number of characters, best guess
// from the Open XML SDK specs
func char2width(chr int) float64 {
	return (math.Round((float64(chr)*7 + 5) / 7 * 256.0)) / 256.0
}

// max returns the maximum of integers a or b
func max(a, b int) int {
	if a > b {
		return a
	}
	return b
}
