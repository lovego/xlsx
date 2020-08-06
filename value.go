package xlsx

import (
	"reflect"
	"strings"
	"time"

	"github.com/lovego/errs"
)

func getValue(value reflect.Value, fieldName string) (interface{}, error) {
	names := strings.SplitN(fieldName, ".", -1)
	for _, name := range names {
		if value = value.FieldByName(name); !value.IsValid() {
			return nil, errs.New(`xlsx-err`, `xlsx: no such field: `+fieldName)
		}
	}
	return format(value), nil
}

func format(value reflect.Value) interface{} {
	switch value.Kind() {
	case reflect.Ptr, reflect.Interface:
		if value.IsNil() {
			return ``
		}
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
	case *time.Time:
		return formatTime(*v)
	}
	return ifc
}

func formatTime(t time.Time) string {
	if t.IsZero() {
		return ""
	}
	return t.Format(`2006-01-02 15:04:05`)
}
