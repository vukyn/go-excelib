package main

import (
	"fmt"
	"reflect"
	"strings"
	"time"

	"github.com/vukyn/go-excelib/config"
	"github.com/vukyn/go-excelib/constants"
	"github.com/xuri/excelize/v2"
)

func ExportExcel(objs interface{}, config *config.ExportConfig) error {

	// Validate
	values := reflect.ValueOf(objs)
	if values.Kind() != reflect.Slice {
		return fmt.Errorf("ExportExcel: objs must be a slice")
	}
	if values.Len() == 0 {
		return fmt.Errorf("ExportExcel: objs must be not empty")
	}
	if values.Index(0).Kind() != reflect.Struct {
		return fmt.Errorf("ExportExcel: objs must be a slice of struct")
	}
	validateConfig(config)
	// End validate

	// define value
	numRows := values.Len()
	numFields := values.Index(0).NumField() - 1
	tbConfig := &tableConfig{}
	tbConfig.resetTableConfig(numFields, numRows)
	fmt.Printf("%v-%v-%v", tbConfig.lastCell, tbConfig.lastCellCol, tbConfig.lastCellRow)

	f := excelize.NewFile()
	index, err := f.NewSheet(config.SheetName)
	if err != nil {
		return err
	}
	f.SetActiveSheet(index)

	// Set default value cell
	if err := f.SetCellValue(config.SheetName, "A1", config.Title); err != nil {
		return err
	}
	if err := f.MergeCell(config.SheetName, "A1", fmt.Sprintf("%v1", tbConfig.endColumnKey)); err != nil {
		return err
	}
	timeSearch := fmt.Sprintf("Thời gian: %s", time.Now().Format("02/01/2006 15:04:05"))
	if err := f.SetCellValue(config.SheetName, "G4", timeSearch); err != nil {
		return err
	}
	if err := f.MergeCell(config.SheetName, "G4", "J4"); err != nil {
		return err
	}

	// Set header
	headers := []string{}
	if config.HasIndex {
		headers = append(headers, config.IndexName)
		numFields++
		tbConfig.resetTableConfig(numFields, numRows)
	}
	for i := 0; i < values.Index(0).NumField(); i++ {
		fieldName := values.Index(0).Type().Field(i).Tag.Get("field")
		if fieldName == "-" {
			numFields--
			tbConfig.resetTableConfig(numFields, numRows)
			continue
		}
		if fieldName == "" {
			fieldName = values.Index(0).Type().Field(i).Name
		}
		headers = append(headers, fieldName)
	}
	if err := f.SetSheetRow(config.SheetName, "A6", &headers); err != nil {
		return err
	}

	// Set data
	for i := 0; i < values.Len(); i++ {
		row := []interface{}{}
		if config.HasIndex {
			row = append(row, i+1)
		}
		for j := 0; j < values.Index(i).NumField(); j++ {
			if values.Index(i).Type().Field(j).Tag.Get("field") == "-" {
				continue
			}
			row = append(row, values.Index(i).Field(j).Interface())
		}
		if err := f.SetSheetRow(config.SheetName, fmt.Sprintf("A%v", i+7), &row); err != nil {
			return err
		}
	}

	// Set style
	styleBoldCenter, err := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{
			Bold: true,
		},
		Alignment: &excelize.Alignment{
			Horizontal: "center",
			Vertical:   "center",
		},
	})
	if err != nil {
		return err
	}
	if err := f.SetCellStyle(config.SheetName, "A6", tbConfig.lastCellCol, styleBoldCenter); err != nil {
		return err
	}

	// Set table
	refRange := fmt.Sprintf("A6:%v", tbConfig.lastCell)
	if err := f.AddTable(config.SheetName, &excelize.Table{
		Range:     refRange,
		Name:      config.TableName,
		StyleName: "TableStyleLight9",
	}); err != nil {
		return err
	}

	if err := f.SaveAs(generateExportPath(config.FileName)); err != nil {
		return err
	}

	return nil
}

func generateExportPath(filename string) string {
	pathExport := "tmp/{time}_{file}.xlsx"
	pathExport = strings.ReplaceAll(pathExport, "{time}", time.Now().Format("2006_01_02_15_04_05"))
	pathExport = strings.ReplaceAll(pathExport, "{file}", filename)
	return pathExport
}

type tableConfig struct {
	endColumnKey string
	endRowIndex  int
	lastCell     string
	lastCellRow  string
	lastCellCol  string
}

func (t *tableConfig) resetTableConfig(numFields int, numRows int) {
	t.endColumnKey = string(constants.EXCEL_COLUMN[numFields])
	t.endRowIndex = numRows + constants.OFFSET_COLUMN
	t.lastCell = fmt.Sprintf("%v%v", string(constants.EXCEL_COLUMN[numFields]), numRows+constants.OFFSET_COLUMN)
	t.lastCellRow = fmt.Sprintf("A%v", numRows+constants.OFFSET_COLUMN)
	t.lastCellCol = fmt.Sprintf("%v%v", string(constants.EXCEL_COLUMN[numFields]), constants.OFFSET_COLUMN)
}

func validateConfig(cfg *config.ExportConfig) {
	if cfg == nil {
		cfg = &config.ExportConfig{}
	}
	if cfg.FileName == "" {
		cfg.FileName = constants.DEFAULT_FILE_NAME
	}
	if cfg.Title == "" {
		cfg.Title = constants.DEFAULT_TITLE
	}
	if cfg.SheetName == "" {
		cfg.SheetName = constants.DEFAULT_SHEET_NAME
	}
	if cfg.TableName == "" {
		cfg.TableName = constants.DEFAULT_TABLE_NAME
	}
	if cfg.IndexName == "" {
		cfg.IndexName = constants.DEFAULT_INDEX_NAME
	}
}
