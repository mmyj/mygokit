package excel_test

import (
	"mygokit/excel"
	"testing"

	"github.com/go-playground/assert/v2"
)

func TestSaveTo(t *testing.T) {
	type colType1 struct {
		School string `col:"school"`
		Name   string `col:"姓名"`
	}
	col1 := &colType1{
		School: "1234",
		Name:   "5678",
	}
	err := excel.SaveExcelTo("test.xlsx", "Sheet1", []interface{}{col1})
	assert.Equal(t, err, nil)
}

func TestAppend(t *testing.T) {
	type colType1 struct {
		School string `col:"school"`
		Name   string `col:"姓名"`
	}
	col1 := &colType1{
		School: "1234",
		Name:   "5678",
	}
	excel, err := excel.NewExcel("Sheet1", col1)
	assert.Equal(t, err, nil)
	err = excel.Append([]interface{}{col1, col1, col1, col1})
	assert.Equal(t, err, nil)
	err = excel.Append([]interface{}{col1, col1, col1, col1})
	assert.Equal(t, err, nil)
	err = excel.SaveTo("test.xlsx")
	assert.Equal(t, err, nil)
}
