package xlsx

import (
	"reflect"
	"strings"

	"github.com/lovego/strs"
)

type Column struct {
	Label     string                            `json:"label" c:"标签（显示名称）"`
	Prop      string                            `json:"prop"  c:"属性（字段名称）"`
	Width     float64                           `json:"width" c:"宽度"`
	Getter    func(row interface{}) interface{} `json:"-"`
	propParts []string
}

func (c *Column) GetValue(row reflect.Value) (interface{}, bool) {
	if c.Getter != nil {
		return c.Getter(row.Interface()), true
	}
	if c.propParts == nil {
		c.propParts = strings.SplitN(c.Prop, ".", -1)
		for i := range c.propParts {
			c.propParts[i] = strs.FirstLetterToUpper(c.propParts[i])
		}
	}
	return GetValue(row, c.propParts)
}

func SetGetters(columns []Column, getters map[string]func(row interface{}) interface{}) {
	for i := range columns {
		if getter := getters[columns[i].Prop]; getter != nil {
			columns[i].Getter = getter
		}
	}
}
