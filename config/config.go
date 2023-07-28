package config

import (
	"fmt"
	"strings"
	"time"
)

type ExportConfig struct {
	Title     string
	TableName string
	fileName  string
	Type      ExportType
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

func (e *ExportConfig) SetFileName(path, name string) string {
	filepath := path + "/{time}_{file}.xlsx"
	filepath = strings.ReplaceAll(filepath, "{time}", time.Now().Format("2006_01_02_15_04_05"))
	filepath = strings.ReplaceAll(filepath, "{file}", name)
	e.fileName = filepath
	return e.fileName
}

func (e *ExportConfig) GetFileName() string {
	return e.fileName
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
	t.StartColumnKey = EXCEL_COLUMN[0]
	t.EndColumnKey = EXCEL_COLUMN[t.NumFields]
	t.StartRowIndex = OFFSET
	t.EndRowIndex = t.NumRows + t.StartRowIndex
	t.FirstCell = fmt.Sprintf("%v%v", t.StartColumnKey, t.StartRowIndex)
	t.LastCell = fmt.Sprintf("%v%v", t.EndColumnKey, t.EndRowIndex)
	t.LastCellRow = fmt.Sprintf("%v%v", t.StartColumnKey, t.EndRowIndex)
	t.LastCellCol = fmt.Sprintf("%v%v", t.EndColumnKey, t.StartRowIndex)
}
