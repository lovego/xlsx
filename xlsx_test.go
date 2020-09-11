package xlsx_test

import (
	"fmt"
	"time"

	"github.com/lovego/date"
	"github.com/lovego/xlsx"
	"github.com/shopspring/decimal"
)

type TestStruct struct {
	String  string
	Bool    bool
	Layer   TestStruct2
	Decimal decimal.Decimal
}
type TestStruct2 struct {
	Time time.Time
	Date date.Date
}

func (t TestStruct) Method() string {
	return "方法"
}

func (t *TestStruct) PtrMethod() string {
	return "指针方法"
}

func ExampleFile_WriteFile() {
	data := []TestStruct{
		{
			String: "中国", Bool: true,
			Layer:   TestStruct2{Time: time.Now(), Date: date.Today()},
			Decimal: decimal.New(12399, -2),
		},
		{
			String: "世界", Bool: false,
			Layer:   TestStruct2{Time: time.Time{}, Date: date.Date{}},
			Decimal: decimal.New(123, 0),
		},
	}
	columns := []xlsx.Column{
		{Label: "字符串", Prop: "string", Width: 8},
		{Label: "布尔", Prop: "bool", Width: 6},
		{Label: "时间", Prop: "layer.time", Width: 20},
		{Label: "日期", Prop: "layer.date", Width: 12},
		{Label: "十进制数", Prop: "decimal", Width: 10},
		{Label: "方法", Prop: "method", Width: 10},
		{Label: "指针方法", Prop: "ptrMethod", Width: 15},
	}

	fmt.Println(xlsx.WriteFile("test", xlsx.Sheet{
		Name:    "工作簿1",
		Data:    data,
		Columns: columns,
	}))

	// Output:
	// <nil>
}
