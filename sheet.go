package xlsx

import (
	"reflect"

	"github.com/lovego/errs"
	"github.com/shopspring/decimal"
	"github.com/tealeg/xlsx"
)

type Sheet struct {
	Name    string
	Data    interface{}
	Columns []Column
}

func (s *Sheet) Generate(file *xlsx.File) error {
	sheet, err := file.AddSheet(s.Name)
	if err != nil {
		return err
	}
	s.GenerateHeader(sheet)
	if err := s.GenerateBody(sheet); err != nil {
		return err
	}
	return nil
}

func (s *Sheet) GenerateHeader(sheet *xlsx.Sheet) {
	row := sheet.AddRow()
	for i := range s.Columns {
		cell := row.AddCell()
		cell.SetStyle(defaultHeaderStyle)
		cell.SetString(s.Columns[i].Label)
		sheet.SetColWidth(i, i, s.Columns[i].Width)
	}
}

func (s *Sheet) GenerateBody(sheet *xlsx.Sheet) error {
	data := reflect.ValueOf(s.Data)
	if data.Kind() == reflect.Ptr {
		data = data.Elem()
	}
	if data.Len() == 0 {
		return nil
	}

	for i := 0; i < data.Len(); i++ {
		rowData := data.Index(i)
		row := sheet.AddRow()
		for j := range s.Columns {
			if v, ok := s.Columns[j].GetValue(rowData); !ok {
				return errs.New(`xlsx-err`, `xlsx: no such field: `+s.Columns[j].Prop)
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
