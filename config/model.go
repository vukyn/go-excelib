package config

import "github.com/xuri/excelize/v2"

type ExportConfig struct {
	FileName  string
	Title     string
	SheetName string
	TableName string
	HasIndex  bool
	IndexName string
}

type Excelib struct {
	File   *excelize.File
	Stream *excelize.StreamWriter
}
