package excel

import (
	"fmt"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
)

func SaveExcelTo(filePath, sheet string, rows []interface{}) error {
	streamWriter, err := NewExcel(sheet, rows[0])
	if err != nil {
		return fmt.Errorf("%w: NewExcel failed", err)
	}
	err = streamWriter.Append(rows)
	if err != nil {
		return fmt.Errorf("%w: Append failed", err)
	}
	err = streamWriter.SaveTo(filePath)
	if err != nil {
		return fmt.Errorf("%w: SaveExcelTo failed", err)
	}
	return nil
}

type Writer struct {
	sw       *excelize.StreamWriter
	rowCount int
}

func NewExcel(sheet string, row interface{}) (*Writer, error) {
	file := excelize.NewFile()
	streamWriter, err := file.NewStreamWriter(sheet)
	if err != nil {
		return nil, fmt.Errorf("%w: NewStreamWriter failed", err)
	}
	err = streamWriter.SetRow("A1", GetAllColumnNameInterface(row))
	if err != nil {
		return nil, fmt.Errorf("%w: SetRow failed", err)
	}
	return &Writer{sw: streamWriter, rowCount: 2}, nil
}

func (w *Writer) Append(rows []interface{}) error {
	for _, row := range rows {
		cell, _ := excelize.CoordinatesToCellName(1, w.rowCount)
		if err := w.sw.SetRow(cell, GetAllColumnValue(row)); err != nil {
			return fmt.Errorf("%w: SetRow failed", err)
		}
		w.rowCount++
	}
	return nil
}

func (w *Writer) SaveTo(filePath string) error {
	err := w.sw.Flush()
	if err != nil {
		return fmt.Errorf("%w: Flush failed", err)
	}
	err = w.sw.File.SaveAs(filePath)
	if err != nil {
		return fmt.Errorf("%w: SaveAs failed", err)
	}
	return nil
}
