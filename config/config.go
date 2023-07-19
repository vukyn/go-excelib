package config

import (
	"fmt"
)

const (
	DEFAULT_TITLE       = "Export Example Excel Title"
	DEFAULT_SHEET_NAME  = "Sheet1"
	DEFAULT_TABLE_NAME  = "Table1"
	DEFAULT_INDEX_NAME  = "No."
	DEFAULT_TABLE_STYLE = "TableStyleLight9"
	EXCEL_COLUMN        = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
	OFFSET              = 6
	MAX_ROW             = 1000000
)

type ExportConfig struct {
	Title     string
	SheetName string
	TableName string
	// Index column with auto increment
	HasIndex bool
	// Index column name
	IndexName string
	// Footer row with function
	HasFooter bool
	// Additional description row below header row
	HasDescription bool
	// Display special formatting for the first column
	ShowFirstColumn bool
	// Display special formatting for the last column
	ShowLastColumn bool

	// StyleName: The built-in table style names
	//
	//	TableStyleLight1 - TableStyleLight21
	//	TableStyleMedium1 - TableStyleMedium28
	//	TableStyleDark1 - TableStyleDark11
	TableStyle string
}

type TableConfig struct {
	// Total fields in the table (include index field, skip ignore field)
	NumFields int
	// Total rows in the table (ignore header row)
	NumRows        int
	StartColumnKey string
	EndColumnKey   string
	StartRowIndex  int
	EndRowIndex    int
	// FirstCell = StartColumnKey + StartRowIndex
	FirstCell string
	// LastCell = EndColumnKey + EndRowIndex
	LastCell string
	// LastCellRow = StartColumnKey + EndRowIndex
	LastCellRow string
	// LastCellCol = EndColumnKey + StartRowIndex
	LastCellCol string
}

func (t *TableConfig) ResetTableConfig() {
	t.StartColumnKey = string(EXCEL_COLUMN[0])
	t.EndColumnKey = string(EXCEL_COLUMN[t.NumFields])
	t.StartRowIndex = OFFSET
	t.EndRowIndex = t.NumRows + t.StartRowIndex
	t.FirstCell = fmt.Sprintf("%v%v", t.StartColumnKey, t.StartRowIndex)
	t.LastCell = fmt.Sprintf("%v%v", t.EndColumnKey, t.EndRowIndex)
	t.LastCellRow = fmt.Sprintf("%v%v", t.StartColumnKey, t.EndRowIndex)
	t.LastCellCol = fmt.Sprintf("%v%v", t.EndColumnKey, t.StartRowIndex)
}
