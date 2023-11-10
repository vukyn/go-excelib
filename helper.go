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

func validateObj(objs []interface{}) (reflect.Value, error) {
	values := reflect.ValueOf(objs[0])
	if len(objs) == 0 {
		return values, fmt.Errorf("objs must be not empty")
	}
	if len(objs) > config.MAX_ROW {
		return values, fmt.Errorf("objs must be less than %v records", config.MAX_ROW)
	}
	if values.Kind() != reflect.Struct {
		return values, fmt.Errorf("objs must be a slice of struct")
	}
	return values, nil
}

func (e *excelib) recalculateConfig(values reflect.Value) {
	if e.exportCfg.HasIndex {
		e.tbCfg.NumFields++
	}
	if e.exportCfg.HasDescription {
		e.tbCfg.NumRows++
	}
	for i := 0; i < values.NumField(); i++ {
		fieldName := values.Type().Field(i).Tag.Get("field")
		if fieldName == "-" {
			e.tbCfg.NumFields--
		}
	}
	e.tbCfg.ResetTableConfig()
}

func (e *excelib) setMetadata() error {
	sheetName := e.file.GetSheetName(e.file.GetActiveSheetIndex())
	if err := e.file.SetCellValue(sheetName, "A1", e.exportCfg.Title); err != nil {
		return err
	}
	if err := e.file.MergeCell(sheetName, "A1", fmt.Sprintf("%v1", e.tbCfg.EndColumnKey)); err != nil {
		return err
	}
	if err := e.file.SetCellValue(sheetName, "A2", fmt.Sprintf("Thời gian: %s", time.Now().Format("02/01/2006 15:04:05"))); err != nil {
		return err
	}
	if err := e.file.MergeCell(sheetName, "A2", "D2"); err != nil {
		return err
	}
	return nil
}

func (e *excelib) setStreamMetadata() error {
	boldCenter := config.BoldCenter(e.file)
	if err := e.stream.SetRow("A1", []interface{}{excelize.Cell{Value: e.exportCfg.Title, StyleID: boldCenter}}); err != nil {
		return err
	}
	if err := e.stream.MergeCell("A1", fmt.Sprintf("%v1", e.tbCfg.EndColumnKey)); err != nil {
		return err
	}
	if err := e.stream.SetRow("A2", []interface{}{excelize.Cell{Value: fmt.Sprintf("Thời gian: %s", time.Now().Format("02/01/2006 15:04:05"))}}); err != nil {
		return err
	}
	if err := e.stream.MergeCell("A2", "D2"); err != nil {
		return err
	}
	return nil
}

func (e *excelib) setHeader(values reflect.Value) error {
	headers := []string{}
	if e.exportCfg.HasIndex {
		headers = append(headers, e.exportCfg.IndexName)
	}
	for i := 0; i < values.NumField(); i++ {
		fieldName := values.Type().Field(i).Tag.Get("field")
		if fieldName == "-" {
			continue
		}
		if fieldName == "" {
			fieldName = values.Type().Field(i).Name
		}
		headers = append(headers, fieldName)
	}
	sheetName := e.file.GetSheetName(e.file.GetActiveSheetIndex())
	if err := e.file.SetSheetRow(sheetName, e.tbCfg.FirstCell, &headers); err != nil {
		return err
	}
	return nil
}

func (e *excelib) setStreamHeader(values reflect.Value) {
	headers := []interface{}{}
	boldCenter := config.BoldCenter(e.file)
	if e.exportCfg.HasIndex {
		headers = append(headers, excelize.Cell{Value: e.exportCfg.IndexName, StyleID: boldCenter})
	}
	for i := 0; i < values.NumField(); i++ {
		fieldName := values.Type().Field(i).Tag.Get("field")
		if fieldName == "-" {
			continue
		}
		if fieldName == "" {
			fieldName = values.Type().Field(i).Name
		}
		headers = append(headers, excelize.Cell{Value: fieldName, StyleID: boldCenter})
	}
	e.headers = headers
}

func (e *excelib) writeStreamHeader() error {
	if err := e.stream.SetRow(e.tbCfg.FirstCell, e.headers); err != nil {
		return err
	}
	return nil
}

func (e *excelib) setDescription(values reflect.Value) error {
	descriptions := []string{}
	e.tbCfg.NumRows++
	e.tbCfg.ResetTableConfig()
	if e.exportCfg.HasIndex {
		descriptions = append(descriptions, "")
	}

	for i := 0; i < values.NumField(); i++ {
		fieldName := values.Type().Field(i).Tag.Get("field")
		desc := values.Type().Field(i).Tag.Get("description")
		if fieldName == "-" {
			continue
		}
		descriptions = append(descriptions, desc)
	}
	sheetName := e.file.GetSheetName(e.file.GetActiveSheetIndex())
	if err := e.file.SetSheetRow(sheetName, fmt.Sprintf("%v%v", e.tbCfg.StartColumnKey, e.tbCfg.StartRowIndex+1), &descriptions); err != nil {
		return err
	}
	return nil
}

func (e *excelib) setStreamDescription(values reflect.Value) {
	descriptions := []interface{}{}
	boldCenter := config.BoldCenter(e.file)
	if e.exportCfg.HasIndex {
		descriptions = append(descriptions, excelize.Cell{Value: ""})
	}

	for i := 0; i < values.NumField(); i++ {
		fieldName := values.Type().Field(i).Tag.Get("field")
		desc := values.Type().Field(i).Tag.Get("description")
		if fieldName == "-" {
			continue
		}
		descriptions = append(descriptions, excelize.Cell{Value: desc, StyleID: boldCenter})
	}
	e.descriptions = descriptions
}

func (e *excelib) writeStreamDescription() error {
	if err := e.stream.SetRow(fmt.Sprintf("%v%v", e.tbCfg.StartColumnKey, e.tbCfg.StartRowIndex+1), e.descriptions); err != nil {
		return err
	}
	return nil
}

func (e *excelib) setBody(objs []interface{}) error {
	startRows := e.tbCfg.StartRowIndex + 1 // Skip header row
	if e.exportCfg.HasDescription {
		startRows++ // Skip description row
	}
	for i := 0; i < len(objs); i++ {
		row := []interface{}{}
		if e.exportCfg.HasIndex {
			row = append(row, i+1)
		}
		value := reflect.ValueOf(objs[i])
		for j := 0; j < value.NumField(); j++ {
			if value.Type().Field(j).Tag.Get("field") == "-" {
				continue
			}
			row = append(row, value.Field(j).Interface())
		}
		sheetName := e.file.GetSheetName(e.file.GetActiveSheetIndex())
		if err := e.file.SetSheetRow(sheetName, fmt.Sprintf("%v%v", e.tbCfg.StartColumnKey, i+startRows), &row); err != nil {
			return err
		}
	}
	return nil
}

func (e *excelib) setStreamBody(objs []interface{}) error {
	center := config.Center(e.file)
	startRows := e.tbCfg.StartRowIndex + 1 // Skip header row
	if e.exportCfg.HasDescription {
		startRows++ // Skip description row
	}
	for i := 0; i < len(objs); i++ {
		row := []interface{}{}
		if e.exportCfg.HasIndex {
			row = append(row, excelize.Cell{Value: i + 1 + e.tbCfg.NumRows, StyleID: center})
		}
		value := reflect.ValueOf(objs[i])
		for j := 0; j < value.NumField(); j++ {
			if value.Type().Field(j).Tag.Get("field") == "-" {
				continue
			}
			row = append(row, excelize.Cell{Value: value.Field(j).Interface()})
		}
		if err := e.stream.SetRow(fmt.Sprintf("%v%v", e.tbCfg.StartColumnKey, i+startRows+e.tbCfg.NumRows), row); err != nil {
			return err
		}
	}
	e.tbCfg.NumRows += len(objs)
	e.tbCfg.ResetTableConfig()
	return nil
}

func (e *excelib) setFooter(values reflect.Value) error {
	lastRowIndex := e.tbCfg.EndRowIndex + 1
	sheetName := e.file.GetSheetName(e.file.GetActiveSheetIndex())

	skipFields := 0
	if e.exportCfg.HasIndex {
		skipFields--
	}
	counter := values.NumField()
	formulaType := excelize.STCellFormulaTypeDataTable
	for i := 0; i < values.NumField(); i++ {
		fieldName := values.Type().Field(i).Tag.Get("field")
		if fieldName == "-" {
			skipFields++
			counter--
			continue
		}
		cell := fmt.Sprintf("%v%v", string(config.EXCEL_COLUMN[i-skipFields]), lastRowIndex)
		switch strings.ToLower(values.Type().Field(i).Tag.Get("footer")) {
		case "sum":
			if err := e.file.SetCellFormula(sheetName, cell, fmt.Sprintf("Sum(%v[%v])", e.exportCfg.TableName, fieldName), excelize.FormulaOpts{Type: &formulaType}); err != nil {
				return err
			}
		case "max":
			if err := e.file.SetCellFormula(sheetName, cell, fmt.Sprintf("Max(%v[%v])", e.exportCfg.TableName, fieldName), excelize.FormulaOpts{Type: &formulaType}); err != nil {
				return err
			}
		case "min":
			if err := e.file.SetCellFormula(sheetName, cell, fmt.Sprintf("Min(%v[%v])", e.exportCfg.TableName, fieldName), excelize.FormulaOpts{Type: &formulaType}); err != nil {
				return err
			}
		case "average":
			if err := e.file.SetCellFormula(sheetName, cell, fmt.Sprintf("Average(%v[%v])", e.exportCfg.TableName, fieldName), excelize.FormulaOpts{Type: &formulaType}); err != nil {
				return err
			}
		default:
			counter--
		}
	}

	if e.exportCfg.HasIndex && counter > 0 {
		if err := e.file.SetCellValue(sheetName, fmt.Sprintf("%v%v", e.tbCfg.StartColumnKey, lastRowIndex), "Total"); err != nil {
			return err
		}
	}
	return nil
}

func (e *excelib) setStreamFooter(values reflect.Value) {
	footers := []interface{}{}
	boldCenter := config.BoldCenter(e.file)

	counter := values.NumField()
	for i := 0; i < values.NumField(); i++ {
		fieldName := values.Type().Field(i).Tag.Get("field")
		if fieldName == "-" {
			counter--
			continue
		}
		switch strings.ToLower(values.Type().Field(i).Tag.Get("footer")) {
		case "sum":
			footers = append(footers, excelize.Cell{Formula: fmt.Sprintf("Sum(%v[%v])", e.exportCfg.TableName, fieldName), StyleID: boldCenter})
		case "max":
			footers = append(footers, excelize.Cell{Formula: fmt.Sprintf("Max(%v[%v])", e.exportCfg.TableName, fieldName), StyleID: boldCenter})
		case "min":
			footers = append(footers, excelize.Cell{Formula: fmt.Sprintf("Min(%v[%v])", e.exportCfg.TableName, fieldName), StyleID: boldCenter})
		case "average":
			footers = append(footers, excelize.Cell{Formula: fmt.Sprintf("Average(%v[%v])", e.exportCfg.TableName, fieldName), StyleID: boldCenter})
		default:
			counter--
			footers = append(footers, nil)
		}
	}
	if e.exportCfg.HasIndex && counter > 0 {
		footers = append([]interface{}{excelize.Cell{Value: "Total", StyleID: boldCenter}}, footers...)
	}
	e.footers = footers
}

func (e *excelib) writeStreamFooter() error {
	if err := e.stream.SetRow(fmt.Sprintf("%v%v", e.tbCfg.StartColumnKey, e.tbCfg.EndRowIndex+1), e.footers); err != nil {
		return err
	}
	return nil
}

func (e *excelib) setStyle() error {
	center := config.Center(e.file)
	boldCenter := config.BoldCenter(e.file)
	sheetName := e.file.GetSheetName(e.file.GetActiveSheetIndex())
	if err := e.file.SetCellStyle(sheetName, "A1", "A1", boldCenter); err != nil {
		return err
	}
	if err := e.file.SetCellStyle(sheetName, e.tbCfg.FirstCell, e.tbCfg.LastCellCol, boldCenter); err != nil {
		return err
	}
	if e.exportCfg.HasIndex {
		if err := e.file.SetCellStyle(sheetName, fmt.Sprintf("%v%v", e.tbCfg.StartColumnKey, e.tbCfg.StartRowIndex+1), fmt.Sprintf("%v%v", e.tbCfg.StartColumnKey, e.tbCfg.EndRowIndex+1), center); err != nil {
			return err
		}
	}
	if e.exportCfg.HasDescription {
		if err := e.file.SetCellStyle(sheetName, fmt.Sprintf("%v%v", e.tbCfg.StartColumnKey, e.tbCfg.StartRowIndex+1), e.tbCfg.LastCellCol, boldCenter); err != nil {
			return err
		}
	}
	if e.exportCfg.HasFooter {
		lastCell := fmt.Sprintf("%v%v", e.tbCfg.EndColumnKey, e.tbCfg.EndRowIndex+1)
		lastCellRow := fmt.Sprintf("%v%v", e.tbCfg.StartColumnKey, e.tbCfg.EndRowIndex+1)
		if err := e.file.SetCellStyle(sheetName, lastCellRow, lastCell, boldCenter); err != nil {
			return err
		}
	}
	return nil
}

func (e *excelib) setTable() error {
	sheetName := e.file.GetSheetName(e.file.GetActiveSheetIndex())
	refRange := fmt.Sprintf("%v:%v", e.tbCfg.FirstCell, e.tbCfg.LastCell)
	if err := e.file.AddTable(sheetName, &excelize.Table{
		Range:           refRange,
		Name:            e.exportCfg.TableName,
		StyleName:       e.exportCfg.TableStyle,
		ShowFirstColumn: e.exportCfg.ShowFirstColumn,
		ShowLastColumn:  e.exportCfg.ShowLastColumn,
	}); err != nil {
		return err
	}
	return nil
}

func (e *excelib) setStreamTable() error {
	refRange := fmt.Sprintf("%v:%v", e.tbCfg.FirstCell, e.tbCfg.LastCell)
	if err := e.stream.AddTable(&excelize.Table{
		Range:           refRange,
		Name:            e.exportCfg.TableName,
		StyleName:       e.exportCfg.TableStyle,
		ShowFirstColumn: e.exportCfg.ShowFirstColumn,
		ShowLastColumn:  e.exportCfg.ShowLastColumn,
	}); err != nil {
		return err
	}
	return nil
}
