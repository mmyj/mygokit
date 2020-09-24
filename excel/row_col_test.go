package excel_test

import (
	"mygokit/excel"
	"testing"

	"github.com/go-playground/assert/v2"
)

func TestMetaInit(t *testing.T) {
	type colType1 struct {
		School string `col:"school"`
		Name   string `col:"姓名"`
	}
	col1 := colType1{
		School: "1234",
		Name:   "5678",
	}
	m := excel.ReflectRomMate(&col1)
	assert.Equal(t, m.Column(&col1.Name).Tag(), "姓名")
	assert.Equal(t, excel.GetAllColumnValue(&col1), []interface{}{"1234", "5678"})
}
