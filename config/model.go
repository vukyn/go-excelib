package config

import (
	"fmt"

	"github.com/vukyn/go-excelib/constants"
	"github.com/xuri/excelize/v2"
)

type Excelib struct {
	File   *excelize.File
	Stream *excelize.StreamWriter
}

type ExportConfig struct {
	FileName  string
	Title     string
	SheetName string
	TableName string
	HasIndex  bool
	IndexName string
}

type TableConfig struct {
	NumFields    int
	NumRows      int
	EndColumnKey string
	EndRowIndex  int
	LastCell     string
	LastCellRow  string
	LastCellCol  string
}

func (t *TableConfig) ResetTableConfig() {
	t.EndColumnKey = string(constants.EXCEL_COLUMN[t.NumFields])
	t.EndRowIndex = t.NumRows + constants.OFFSET_COLUMN
	t.LastCell = fmt.Sprintf("%v%v", string(constants.EXCEL_COLUMN[t.NumFields]), t.NumRows+constants.OFFSET_COLUMN)
	t.LastCellRow = fmt.Sprintf("A%v", t.NumRows+constants.OFFSET_COLUMN)
	t.LastCellCol = fmt.Sprintf("%v%v", string(constants.EXCEL_COLUMN[t.NumFields]), constants.OFFSET_COLUMN)
}
