package xlsx_test

import (
	"fmt"
	"reflect"

	"github.com/lovego/date"
	"github.com/lovego/xlsx"
)

func ExampleNonPtrType_GetValue() {
	day, _ := date.New("2020-08-07")
	ts := TestStruct{String: "中国", Layer: TestStruct2{Date: *day}}
	v := reflect.ValueOf(ts)
	fmt.Println(xlsx.GetValue(v, []string{"String"}))
	fmt.Println(xlsx.GetValue(v, []string{"Layer", "Date"}))
	fmt.Println(xlsx.GetValue(v, []string{"Method"}))
	fmt.Println(xlsx.GetValue(v, []string{"PtrMethod"}))

	v = reflect.ValueOf(&ts).Elem()
	fmt.Println(xlsx.GetValue(v, []string{"PtrMethod"}))

	// Output:
	// 中国 true
	// 2020-08-07 true
	// 方法 true
	// <nil> false
	// 指针方法 true
}

func ExamplePtrType_GetValue() {
	day, _ := date.New("2020-08-07")
	ts := TestStruct{String: "中国", Layer: TestStruct2{Date: *day}}
	v := reflect.ValueOf(&ts)
	fmt.Println(xlsx.GetValue(v, []string{"String"}))
	fmt.Println(xlsx.GetValue(v, []string{"Layer", "Date"}))
	fmt.Println(xlsx.GetValue(v, []string{"Method"}))
	fmt.Println(xlsx.GetValue(v, []string{"PtrMethod"}))

	// Output:
	// 中国 true
	// 2020-08-07 true
	// 方法 true
	// 指针方法 true
}

// https://golang.org/ref/spec#Method_sets
// The method set of the pointer type *T is the set of all methods declared with receiver *T or T (that is, it also contains the method set of T).
func ExampleNonPtrType_MethodByName() {
	t := reflect.TypeOf(TestStruct{})
	method, _ := t.MethodByName("Method")
	fmt.Println(method.Type)
	method, _ = t.MethodByName("PtrMethod")
	fmt.Println(method.Type)

	// Output:
	// func(xlsx_test.TestStruct) string
	// <nil>
}

func ExamplePtrType_MethodByName() {
	t := reflect.TypeOf(&TestStruct{})
	method, _ := t.MethodByName("Method")
	fmt.Println(method.Type)
	method, _ = t.MethodByName("PtrMethod")
	fmt.Println(method.Type)

	// Output:
	// func(*xlsx_test.TestStruct) string
	// func(*xlsx_test.TestStruct) string
}
