package excelib

import (
	"fmt"
	"reflect"
	"strings"
	"time"

	"github.com/vukyn/go-excelib/config"
	"github.com/xuri/excelize/v2"
)

func validateConfig(cfg *config.ExportConfig) {
	if cfg == nil {
		cfg = &config.ExportConfig{}
	}
	if cfg.Title == "" {
		cfg.Title = config.DEFAULT_TITLE
	}
	if cfg.SheetName == "" {
		cfg.SheetName = config.DEFAULT_SHEET_NAME
	}
	if cfg.TableName == "" {
		cfg.TableName = config.DEFAULT_TABLE_NAME
	}
	if cfg.IndexName == "" {
		cfg.IndexName = config.DEFAULT_INDEX_NAME
	}
}

func setMetadata(f *excelize.File, cfg *config.ExportConfig, tbCfg *config.TableConfig) error {
	if err := f.SetCellValue(cfg.SheetName, "A1", cfg.Title); err != nil {
		return err
	}
	if err := f.MergeCell(cfg.SheetName, "A1", fmt.Sprintf("%v1", tbCfg.EndColumnKey)); err != nil {
		return err
	}
	if err := f.SetCellValue(cfg.SheetName, "A2", fmt.Sprintf("Thời gian: %s", time.Now().Format("02/01/2006 15:04:05"))); err != nil {
		return err
	}
	if err := f.MergeCell(cfg.SheetName, "A2", "D2"); err != nil {
		return err
	}
	return nil
}

func setHeader(f *excelize.File, cfg *config.ExportConfig, tbCfg *config.TableConfig, values reflect.Value) error {
	headers := []string{}
	if cfg.HasIndex {
		headers = append(headers, cfg.IndexName)
		tbCfg.NumFields++
		tbCfg.ResetTableConfig()
	}
	for i := 0; i < values.Index(0).NumField(); i++ {
		fieldName := values.Index(0).Type().Field(i).Tag.Get("field")
		if fieldName == "-" {
			tbCfg.NumFields--
			tbCfg.ResetTableConfig()
			continue
		}
		if fieldName == "" {
			fieldName = values.Index(0).Type().Field(i).Name
		}
		headers = append(headers, fieldName)
	}
	if err := f.SetSheetRow(cfg.SheetName, tbCfg.FirstCell, &headers); err != nil {
		return err
	}
	return nil
}

func setDescription(f *excelize.File, cfg *config.ExportConfig, tbCfg *config.TableConfig, values reflect.Value) error {
	descriptions := []string{}
	tbCfg.NumRows++
	tbCfg.ResetTableConfig()
	if cfg.HasIndex {
		descriptions = append(descriptions, "")
	}

	for i := 0; i < values.Index(0).NumField(); i++ {
		fieldName := values.Index(0).Type().Field(i).Tag.Get("field")
		desc := values.Index(0).Type().Field(i).Tag.Get("description")
		if fieldName == "-" {
			continue
		}
		descriptions = append(descriptions, desc)
	}
	if err := f.SetSheetRow(cfg.SheetName, fmt.Sprintf("%v%v", tbCfg.StartColumnKey, tbCfg.StartRowIndex+1), &descriptions); err != nil {
		return err
	}
	return nil
}

func setBody(f *excelize.File, cfg *config.ExportConfig, tbCfg *config.TableConfig, values reflect.Value) error {
	startRows := tbCfg.StartRowIndex + 1 // Skip header row
	if cfg.HasDescription {
		startRows++ // Skip description row
	}
	for i := 0; i < values.Len(); i++ {
		row := []interface{}{}
		if cfg.HasIndex {
			row = append(row, i+1)
		}
		for j := 0; j < values.Index(i).NumField(); j++ {
			if values.Index(i).Type().Field(j).Tag.Get("field") == "-" {
				continue
			}
			row = append(row, values.Index(i).Field(j).Interface())
		}
		if err := f.SetSheetRow(cfg.SheetName, fmt.Sprintf("%v%v", tbCfg.StartColumnKey, i+startRows), &row); err != nil {
			return err
		}
	}
	return nil
}

func setFooter(f *excelize.File, cfg *config.ExportConfig, tbCfg *config.TableConfig, values reflect.Value) error {
	lastRowIndex := tbCfg.EndRowIndex + 1
	if cfg.HasIndex {
		if err := f.SetCellValue(cfg.SheetName, fmt.Sprintf("%v%v", tbCfg.StartColumnKey, lastRowIndex), "Total"); err != nil {
			return err
		}
	}

	skipFields := 0
	if cfg.HasIndex {
		skipFields--
	}
	formulaType := excelize.STCellFormulaTypeDataTable
	for i := 0; i < values.Index(0).NumField(); i++ {
		fieldName := values.Index(0).Type().Field(i).Tag.Get("field")
		if fieldName == "-" {
			skipFields++
			continue
		}
		cell := fmt.Sprintf("%v%v", string(config.EXCEL_COLUMN[i-skipFields]), lastRowIndex)
		switch strings.ToLower(values.Index(0).Type().Field(i).Tag.Get("footer")) {
		case "sum":
			if err := f.SetCellFormula(cfg.SheetName, cell, fmt.Sprintf("Sum(%v[%v])", cfg.TableName, fieldName), excelize.FormulaOpts{Type: &formulaType}); err != nil {
				return err
			}
		case "max":
			if err := f.SetCellFormula(cfg.SheetName, cell, fmt.Sprintf("Max(%v[%v])", cfg.TableName, fieldName), excelize.FormulaOpts{Type: &formulaType}); err != nil {
				return err
			}
		case "min":
			if err := f.SetCellFormula(cfg.SheetName, cell, fmt.Sprintf("Min(%v[%v])", cfg.TableName, fieldName), excelize.FormulaOpts{Type: &formulaType}); err != nil {
				return err
			}
		case "average":
			if err := f.SetCellFormula(cfg.SheetName, cell, fmt.Sprintf("Average(%v[%v])", cfg.TableName, fieldName), excelize.FormulaOpts{Type: &formulaType}); err != nil {
				return err
			}
		}
	}
	return nil
}

func setStyle(f *excelize.File, cfg *config.ExportConfig, tbCfg *config.TableConfig) error {
	boldCenter := config.BoldCenter(f)
	if err := f.SetCellStyle(cfg.SheetName, "A1", "A1", boldCenter); err != nil {
		return err
	}
	if err := f.SetCellStyle(cfg.SheetName, tbCfg.FirstCell, tbCfg.LastCellCol, boldCenter); err != nil {
		return err
	}
	if cfg.HasDescription {
		if err := f.SetCellStyle(cfg.SheetName, fmt.Sprintf("%v%v", tbCfg.StartColumnKey, tbCfg.StartRowIndex+1), tbCfg.LastCellCol, boldCenter); err != nil {
			return err
		}
	}
	if cfg.HasFooter {
		lastCell := fmt.Sprintf("%v%v", tbCfg.EndColumnKey, tbCfg.EndRowIndex+1)
		lastCellRow := fmt.Sprintf("%v%v", tbCfg.StartColumnKey, tbCfg.EndRowIndex+1)
		if err := f.SetCellStyle(cfg.SheetName, lastCellRow, lastCell, boldCenter); err != nil {
			return err
		}
	}
	return nil
}

func setTable(f *excelize.File, cfg *config.ExportConfig, tbCfg *config.TableConfig) error {
	refRange := fmt.Sprintf("%v:%v", tbCfg.FirstCell, tbCfg.LastCell)
	if err := f.AddTable(cfg.SheetName, &excelize.Table{
		Range:           refRange,
		Name:            cfg.TableName,
		StyleName:       "TableStyleLight9",
		ShowFirstColumn: cfg.ShowFirstColumn,
		ShowLastColumn:  cfg.ShowLastColumn,
	}); err != nil {
		return err
	}
	return nil
}
