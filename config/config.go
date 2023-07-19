package config

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

const (
	DEFAULT_FILE_NAME  = "export"
	DEFAULT_TITLE      = "Export Example Excel Title"
	DEFAULT_SHEET_NAME = "Sheet1"
	DEFAULT_TABLE_NAME = "Table1"
	DEFAULT_INDEX_NAME = "No."
	EXCEL_COLUMN       = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
	OFFSET_COLUMN      = 6
	MAX_ROW            = 1000000
)

type Excelib struct {
	File   *excelize.File
	Stream *excelize.StreamWriter
}

type ExportConfig struct {
	FileName        string
	Title           string
	SheetName       string
	TableName       string
	HasIndex        bool
	IndexName       string
	HasFooter       bool
	ShowFirstColumn bool
	ShowLastColumn  bool
}

type TableConfig struct {
	NumFields      int
	NumRows        int
	StartColumnKey string
	EndColumnKey   string
	StartRowIndex  int
	EndRowIndex    int
	LastCell       string
	LastCellRow    string
	LastCellCol    string
}

func (t *TableConfig) ResetTableConfig() {
	t.StartColumnKey = string(EXCEL_COLUMN[0])
	t.EndColumnKey = string(EXCEL_COLUMN[t.NumFields])
	t.StartRowIndex = OFFSET_COLUMN
	t.EndRowIndex = t.NumRows + OFFSET_COLUMN
	t.LastCell = fmt.Sprintf("%v%v", string(EXCEL_COLUMN[t.NumFields]), t.NumRows+OFFSET_COLUMN)
	t.LastCellRow = fmt.Sprintf("A%v", t.NumRows+OFFSET_COLUMN)
	t.LastCellCol = fmt.Sprintf("%v%v", string(EXCEL_COLUMN[t.NumFields]), OFFSET_COLUMN)
}
