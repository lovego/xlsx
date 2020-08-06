package xlsx_test

import (
	"fmt"
	"time"

	"github.com/lovego/date"
	"github.com/lovego/xlsx"
	"github.com/shopspring/decimal"
)

func ExampleFile_WriteFile() {
	data := []struct {
		String  string
		Bool    bool
		Time    time.Time
		Date    date.Date
		Decimal decimal.Decimal
	}{
		{
			String: "中国", Bool: true, Time: time.Now(),
			Date: date.Date{time.Now()}, Decimal: decimal.New(12399, -2),
		},
		{
			String: "世界", Bool: false, Time: time.Time{},
			Date: date.Date{}, Decimal: decimal.New(123, 0),
		},
	}
	columns := []xlsx.Column{
		{Label: "字符串", Prop: "string", Width: 8},
		{Label: "布尔", Prop: "bool", Width: 6},
		{Label: "时间", Prop: "time", Width: 20},
		{Label: "日期", Prop: "date", Width: 12},
		{Label: "十进制数", Prop: "decimal", Width: 10},
	}

	fmt.Println(xlsx.WriteFile("test", xlsx.Sheet{
		Name:    "工作簿1",
		Data:    data,
		Columns: columns,
	}))

	// Output:
	// <nil>
}
