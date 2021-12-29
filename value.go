package xlsx

import (
	"reflect"
	"time"

	valuePkg "github.com/lovego/value"
)

func GetValue(value reflect.Value, names []string) (interface{}, bool) {
	value = valuePkg.Get(value, names)
	if value.IsValid() {
		return format(value), true
	}
	return nil, false
}

func format(value reflect.Value) interface{} {
	switch value.Kind() {
	case reflect.Ptr, reflect.Interface:
		if value.IsNil() {
			return ``
		}
		value = value.Elem()
	}
	ifc := value.Interface()

	switch v := ifc.(type) {
	case bool:
		if v {
			return "是"
		} else {
			return "否"
		}
	case time.Time:
		return formatTime(v)
	}

	return ifc
}

func formatTime(t time.Time) string {
	if t.IsZero() {
		return ""
	}
	return t.Format(`2006-01-02 15:04:05`)
}
