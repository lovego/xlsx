package xlsx

import (
	"reflect"
	"strings"

	"github.com/lovego/errs"
	"github.com/lovego/strs"
	"github.com/shopspring/decimal"
	"github.com/tealeg/xlsx"
)

type Sheet struct {
	Name    string
	Data    interface{}
	Columns []Column
}

type Column struct {
	Label string  `json:"label" c:"显示名"`
	Prop  string  `json:"prop"  c:"数据字段名"`
	Width float64 `json:"width" c:"宽度"`
}

func (s *Sheet) Generate(file *xlsx.File) error {
	sheet, err := file.AddSheet(s.Name)
	if err != nil {
		return err
	}
	s.generateHeader(sheet)
	if err := s.generateBody(sheet); err != nil {
		return err
	}
	return nil
}

func (s *Sheet) generateHeader(sheet *xlsx.Sheet) {
	row := sheet.AddRow()
	for i := range s.Columns {
		cell := row.AddCell()
		cell.SetStyle(defaultHeaderStyle)
		cell.SetString(s.Columns[i].Label)
		sheet.SetColWidth(i, i+1, s.Columns[i].Width)
	}
}

func (s *Sheet) generateBody(sheet *xlsx.Sheet) error {
	data := reflect.ValueOf(s.Data)
	if data.Kind() == reflect.Ptr {
		data = data.Elem()
	}
	if data.Len() == 0 {
		return nil
	}

	var fieldNames = make([][]string, len(s.Columns))
	for i := range s.Columns {
		names := strings.SplitN(s.Columns[i].Prop, ".", -1)
		for i := range names {
			names[i] = strs.FirstLetterToUpper(names[i])
		}
		fieldNames[i] = names
	}

	for i := 0; i < data.Len(); i++ {
		rowData := data.Index(i)
		row := sheet.AddRow()
		for i, fieldName := range fieldNames {
			if v, ok := GetValue(rowData, fieldName); !ok {
				return errs.New(`xlsx-err`, `xlsx: no such field: `+s.Columns[i].Prop)
			} else {
				cell := row.AddCell()
				cell.SetValue(v)

				if _, ok := v.(decimal.Decimal); ok {
					cell.NumFmt = "general"
					cell.SetFormula("")
				}
			}
		}
	}
	return nil
}

var defaultHeaderStyle = getDefaultHeaderStyle()

func getDefaultHeaderStyle() *xlsx.Style {
	s := xlsx.NewStyle()
	s.Fill = xlsx.Fill{
		PatternType: "solid",
		FgColor:     "00B7DEE8",
		BgColor:     "FF000000",
	}
	return s
}
