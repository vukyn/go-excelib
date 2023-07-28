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
	if cfg.TableName == "" {
		cfg.TableName = config.DEFAULT_TABLE_NAME
	}
	if cfg.IndexName == "" {
		cfg.IndexName = config.DEFAULT_INDEX_NAME
	}
	if cfg.TableStyle == "" {
		cfg.TableStyle = config.DEFAULT_TABLE_STYLE
	}
}

func (e *Excelib) recalculateConfig(tbCfg *config.TableConfig, values reflect.Value) {
	if e.cfg.HasIndex {
		tbCfg.NumFields++
	}
	if e.cfg.HasDescription {
		tbCfg.NumRows++
	}
	for i := 0; i < values.Index(0).NumField(); i++ {
		fieldName := values.Index(0).Type().Field(i).Tag.Get("field")
		if fieldName == "-" {
			tbCfg.NumFields--
		}
	}
	tbCfg.ResetTableConfig()
}

func (e *Excelib) setMetadata(tbCfg *config.TableConfig) error {
	sheetName := e.File.GetSheetName(e.File.GetActiveSheetIndex())
	if err := e.File.SetCellValue(sheetName, "A1", e.cfg.Title); err != nil {
		return err
	}
	if err := e.File.MergeCell(sheetName, "A1", fmt.Sprintf("%v1", tbCfg.EndColumnKey)); err != nil {
		return err
	}
	if err := e.File.SetCellValue(sheetName, "A2", fmt.Sprintf("Thời gian: %s", time.Now().Format("02/01/2006 15:04:05"))); err != nil {
		return err
	}
	if err := e.File.MergeCell(sheetName, "A2", "D2"); err != nil {
		return err
	}
	return nil
}

func (e *Excelib) setStreamMetadata(tbCfg *config.TableConfig) error {
	boldCenter := config.BoldCenter(e.File)
	if err := e.Stream.SetRow("A1", []interface{}{excelize.Cell{Value: e.cfg.Title, StyleID: boldCenter}}); err != nil {
		return err
	}
	if err := e.Stream.MergeCell("A1", fmt.Sprintf("%v1", tbCfg.EndColumnKey)); err != nil {
		return err
	}
	if err := e.Stream.SetRow("A2", []interface{}{excelize.Cell{Value: fmt.Sprintf("Thời gian: %s", time.Now().Format("02/01/2006 15:04:05"))}}); err != nil {
		return err
	}
	if err := e.Stream.MergeCell("A2", "D2"); err != nil {
		return err
	}
	return nil
}

func (e *Excelib) setHeader(tbCfg *config.TableConfig, values reflect.Value) error {
	headers := []string{}
	if e.cfg.HasIndex {
		headers = append(headers, e.cfg.IndexName)
	}
	for i := 0; i < values.Index(0).NumField(); i++ {
		fieldName := values.Index(0).Type().Field(i).Tag.Get("field")
		if fieldName == "-" {
			continue
		}
		if fieldName == "" {
			fieldName = values.Index(0).Type().Field(i).Name
		}
		headers = append(headers, fieldName)
	}
	sheetName := e.File.GetSheetName(e.File.GetActiveSheetIndex())
	if err := e.File.SetSheetRow(sheetName, tbCfg.FirstCell, &headers); err != nil {
		return err
	}
	return nil
}

func (e *Excelib) setStreamHeader(tbCfg *config.TableConfig, values reflect.Value) error {
	headers := []interface{}{}
	boldCenter := config.BoldCenter(e.File)
	if e.cfg.HasIndex {
		headers = append(headers, excelize.Cell{Value: e.cfg.IndexName, StyleID: boldCenter})
	}
	for i := 0; i < values.Index(0).NumField(); i++ {
		fieldName := values.Index(0).Type().Field(i).Tag.Get("field")
		if fieldName == "-" {
			continue
		}
		if fieldName == "" {
			fieldName = values.Index(0).Type().Field(i).Name
		}
		headers = append(headers, excelize.Cell{Value: fieldName, StyleID: boldCenter})
	}
	if err := e.Stream.SetRow(tbCfg.FirstCell, headers); err != nil {
		return err
	}
	return nil
}

func (e *Excelib) setDescription(tbCfg *config.TableConfig, values reflect.Value) error {
	descriptions := []string{}
	tbCfg.NumRows++
	tbCfg.ResetTableConfig()
	if e.cfg.HasIndex {
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
	sheetName := e.File.GetSheetName(e.File.GetActiveSheetIndex())
	if err := e.File.SetSheetRow(sheetName, fmt.Sprintf("%v%v", tbCfg.StartColumnKey, tbCfg.StartRowIndex+1), &descriptions); err != nil {
		return err
	}
	return nil
}

func (e *Excelib) setStreamDescription(tbCfg *config.TableConfig, values reflect.Value) error {
	descriptions := []interface{}{}
	boldCenter := config.BoldCenter(e.File)
	if e.cfg.HasIndex {
		descriptions = append(descriptions, excelize.Cell{Value: ""})
	}

	for i := 0; i < values.Index(0).NumField(); i++ {
		fieldName := values.Index(0).Type().Field(i).Tag.Get("field")
		desc := values.Index(0).Type().Field(i).Tag.Get("description")
		if fieldName == "-" {
			continue
		}
		descriptions = append(descriptions, excelize.Cell{Value: desc, StyleID: boldCenter})
	}
	if err := e.Stream.SetRow(fmt.Sprintf("%v%v", tbCfg.StartColumnKey, tbCfg.StartRowIndex+1), descriptions); err != nil {
		return err
	}
	return nil
}

func (e *Excelib) setBody(tbCfg *config.TableConfig, values reflect.Value) error {
	startRows := tbCfg.StartRowIndex + 1 // Skip header row
	if e.cfg.HasDescription {
		startRows++ // Skip description row
	}
	for i := 0; i < values.Len(); i++ {
		row := []interface{}{}
		if e.cfg.HasIndex {
			row = append(row, i+1)
		}
		for j := 0; j < values.Index(i).NumField(); j++ {
			if values.Index(i).Type().Field(j).Tag.Get("field") == "-" {
				continue
			}
			row = append(row, values.Index(i).Field(j).Interface())
		}
		sheetName := e.File.GetSheetName(e.File.GetActiveSheetIndex())
		if err := e.File.SetSheetRow(sheetName, fmt.Sprintf("%v%v", tbCfg.StartColumnKey, i+startRows), &row); err != nil {
			return err
		}
	}
	return nil
}

func (e *Excelib) setStreamBody(tbCfg *config.TableConfig, values reflect.Value) error {
	center := config.Center(e.File)
	startRows := tbCfg.StartRowIndex + 1 // Skip header row
	if e.cfg.HasDescription {
		startRows++ // Skip description row
	}
	for i := 0; i < values.Len(); i++ {
		row := []interface{}{}
		if e.cfg.HasIndex {
			row = append(row, excelize.Cell{Value: i + 1, StyleID: center})
		}
		for j := 0; j < values.Index(i).NumField(); j++ {
			if values.Index(i).Type().Field(j).Tag.Get("field") == "-" {
				continue
			}
			row = append(row, excelize.Cell{Value: values.Index(i).Field(j).Interface()})
		}
		if err := e.Stream.SetRow(fmt.Sprintf("%v%v", tbCfg.StartColumnKey, i+startRows), row); err != nil {
			return err
		}
	}
	return nil
}

func (e *Excelib) setFooter(tbCfg *config.TableConfig, values reflect.Value) error {
	lastRowIndex := tbCfg.EndRowIndex + 1
	sheetName := e.File.GetSheetName(e.File.GetActiveSheetIndex())
	if e.cfg.HasIndex {
		if err := e.File.SetCellValue(sheetName, fmt.Sprintf("%v%v", tbCfg.StartColumnKey, lastRowIndex), "Total"); err != nil {
			return err
		}
	}

	skipFields := 0
	if e.cfg.HasIndex {
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
			if err := e.File.SetCellFormula(sheetName, cell, fmt.Sprintf("Sum(%v[%v])", e.cfg.TableName, fieldName), excelize.FormulaOpts{Type: &formulaType}); err != nil {
				return err
			}
		case "max":
			if err := e.File.SetCellFormula(sheetName, cell, fmt.Sprintf("Max(%v[%v])", e.cfg.TableName, fieldName), excelize.FormulaOpts{Type: &formulaType}); err != nil {
				return err
			}
		case "min":
			if err := e.File.SetCellFormula(sheetName, cell, fmt.Sprintf("Min(%v[%v])", e.cfg.TableName, fieldName), excelize.FormulaOpts{Type: &formulaType}); err != nil {
				return err
			}
		case "average":
			if err := e.File.SetCellFormula(sheetName, cell, fmt.Sprintf("Average(%v[%v])", e.cfg.TableName, fieldName), excelize.FormulaOpts{Type: &formulaType}); err != nil {
				return err
			}
		}
	}
	return nil
}

func (e *Excelib) setStreamFooter(tbCfg *config.TableConfig, values reflect.Value) error {
	footers := []interface{}{}
	boldCenter := config.BoldCenter(e.File)
	if e.cfg.HasIndex {
		footers = append(footers, excelize.Cell{Value: "Total", StyleID: boldCenter})
	}

	skipFields := 0
	if e.cfg.HasIndex {
		skipFields--
	}
	for i := 0; i < values.Index(0).NumField(); i++ {
		fieldName := values.Index(0).Type().Field(i).Tag.Get("field")
		if fieldName == "-" {
			skipFields++
			continue
		}
		switch strings.ToLower(values.Index(0).Type().Field(i).Tag.Get("footer")) {
		case "sum":
			footers = append(footers, excelize.Cell{Formula: fmt.Sprintf("Sum(%v[%v])", e.cfg.TableName, fieldName), StyleID: boldCenter})
		case "max":
			footers = append(footers, excelize.Cell{Formula: fmt.Sprintf("Max(%v[%v])", e.cfg.TableName, fieldName), StyleID: boldCenter})
		case "min":
			footers = append(footers, excelize.Cell{Formula: fmt.Sprintf("Min(%v[%v])", e.cfg.TableName, fieldName), StyleID: boldCenter})
		case "average":
			footers = append(footers, excelize.Cell{Formula: fmt.Sprintf("Average(%v[%v])", e.cfg.TableName, fieldName), StyleID: boldCenter})
		default:
			footers = append(footers, nil)
		}
	}
	if err := e.Stream.SetRow(fmt.Sprintf("%v%v", tbCfg.StartColumnKey, tbCfg.EndRowIndex+1), footers); err != nil {
		return err
	}
	return nil
}

func (e *Excelib) setStyle(tbCfg *config.TableConfig) error {
	center := config.Center(e.File)
	boldCenter := config.BoldCenter(e.File)
	sheetName := e.File.GetSheetName(e.File.GetActiveSheetIndex())
	if err := e.File.SetCellStyle(sheetName, "A1", "A1", boldCenter); err != nil {
		return err
	}
	if err := e.File.SetCellStyle(sheetName, tbCfg.FirstCell, tbCfg.LastCellCol, boldCenter); err != nil {
		return err
	}
	if e.cfg.HasIndex {
		if err := e.File.SetCellStyle(sheetName, fmt.Sprintf("%v%v", tbCfg.StartColumnKey, tbCfg.StartRowIndex+1), fmt.Sprintf("%v%v", tbCfg.StartColumnKey, tbCfg.EndRowIndex+1), center); err != nil {
			return err
		}
	}
	if e.cfg.HasDescription {
		if err := e.File.SetCellStyle(sheetName, fmt.Sprintf("%v%v", tbCfg.StartColumnKey, tbCfg.StartRowIndex+1), tbCfg.LastCellCol, boldCenter); err != nil {
			return err
		}
	}
	if e.cfg.HasFooter {
		lastCell := fmt.Sprintf("%v%v", tbCfg.EndColumnKey, tbCfg.EndRowIndex+1)
		lastCellRow := fmt.Sprintf("%v%v", tbCfg.StartColumnKey, tbCfg.EndRowIndex+1)
		if err := e.File.SetCellStyle(sheetName, lastCellRow, lastCell, boldCenter); err != nil {
			return err
		}
	}
	return nil
}

func (e *Excelib) setTable(tbCfg *config.TableConfig) error {
	sheetName := e.File.GetSheetName(e.File.GetActiveSheetIndex())
	refRange := fmt.Sprintf("%v:%v", tbCfg.FirstCell, tbCfg.LastCell)
	if err := e.File.AddTable(sheetName, &excelize.Table{
		Range:           refRange,
		Name:            e.cfg.TableName,
		StyleName:       e.cfg.TableStyle,
		ShowFirstColumn: e.cfg.ShowFirstColumn,
		ShowLastColumn:  e.cfg.ShowLastColumn,
	}); err != nil {
		return err
	}
	return nil
}

func (e *Excelib) setStreamTable(tbCfg *config.TableConfig) error {
	refRange := fmt.Sprintf("%v:%v", tbCfg.FirstCell, tbCfg.LastCell)
	if err := e.Stream.AddTable(&excelize.Table{
		Range:           refRange,
		Name:            e.cfg.TableName,
		StyleName:       e.cfg.TableStyle,
		ShowFirstColumn: e.cfg.ShowFirstColumn,
		ShowLastColumn:  e.cfg.ShowLastColumn,
	}); err != nil {
		return err
	}
	return nil
}
