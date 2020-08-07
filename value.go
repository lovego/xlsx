package xlsx

import (
	"reflect"
	"time"
)

func GetValue(value reflect.Value, names []string) (interface{}, bool) {
	for _, name := range names {
		if v := tryField(value, name); v.IsValid() {
			value = v
		} else if v := tryMethod(value, name); v.IsValid() {
			value = v
		} else if value.Kind() != reflect.Ptr && value.CanAddr() {
			if v := tryMethod(value.Addr(), name); v.IsValid() {
				value = v
			} else {
				return nil, false
			}
		} else {
			return nil, false
		}
	}
	return format(value), true
}

func tryField(value reflect.Value, name string) reflect.Value {
	if value.Kind() == reflect.Ptr {
		value = value.Elem()
	}
	return value.FieldByName(name)
}

func tryMethod(value reflect.Value, name string) reflect.Value {
	if v := value.MethodByName(name); v.IsValid() && v.Type().NumOut() == 1 {
		return v.Call(nil)[0]
	}
	return reflect.Value{}
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
