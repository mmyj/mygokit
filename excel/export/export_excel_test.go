package export_test

import (
	"github.com/go-playground/assert/v2"
	"mygokit/excel/export"
	"testing"
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
	err := export.SaveExcelTo("test.xlsx", "Sheet1", []interface{}{col1})
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
	excel, err := export.NewExcel("Sheet1", col1)
	assert.Equal(t, err, nil)
	err = excel.Append([]interface{}{col1, col1, col1, col1})
	assert.Equal(t, err, nil)
	err = excel.Append([]interface{}{col1, col1, col1, col1})
	assert.Equal(t, err, nil)
	err = excel.SaveTo("test.xlsx")
	assert.Equal(t, err, nil)
}
