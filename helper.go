package excelib

import (
	"fmt"
	"reflect"
	"strings"
	"time"

	"github.com/vukyn/go-excelib/config"
	"github.com/vukyn/go-excelib/constants"
	"github.com/xuri/excelize/v2"
)

func generateExportPath(filename string) string {
	pathExport := "tmp/{time}_{file}.xlsx"
	pathExport = strings.ReplaceAll(pathExport, "{time}", time.Now().Format("2006_01_02_15_04_05"))
	pathExport = strings.ReplaceAll(pathExport, "{file}", filename)
	return pathExport
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

func setMetadata(f *excelize.File, cfg *config.ExportConfig, tbConfig *config.TableConfig) error {
	if err := f.SetCellValue(cfg.SheetName, "A1", cfg.Title); err != nil {
		return err
	}
	if err := f.MergeCell(cfg.SheetName, "A1", fmt.Sprintf("%v1", tbConfig.EndColumnKey)); err != nil {
		return err
	}
	if err := f.SetCellValue(cfg.SheetName, "G4", fmt.Sprintf("Thời gian: %s", time.Now().Format("02/01/2006 15:04:05"))); err != nil {
		return err
	}
	if err := f.MergeCell(cfg.SheetName, "G4", "J4"); err != nil {
		return err
	}
	return nil
}

func setHeader(f *excelize.File, cfg *config.ExportConfig, tbConfig *config.TableConfig, values reflect.Value) error {
	headers := []string{}
	if cfg.HasIndex {
		headers = append(headers, cfg.IndexName)
		tbConfig.NumFields++
		tbConfig.ResetTableConfig()
	}
	for i := 0; i < values.Index(0).NumField(); i++ {
		fieldName := values.Index(0).Type().Field(i).Tag.Get("field")
		if fieldName == "-" {
			tbConfig.NumFields--
			tbConfig.ResetTableConfig()
			continue
		}
		if fieldName == "" {
			fieldName = values.Index(0).Type().Field(i).Name
		}
		headers = append(headers, fieldName)
	}
	if err := f.SetSheetRow(cfg.SheetName, "A6", &headers); err != nil {
		return err
	}
	return nil
}

func setBody(f *excelize.File, cfg *config.ExportConfig, tbConfig *config.TableConfig, values reflect.Value) error {
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
		if err := f.SetSheetRow(cfg.SheetName, fmt.Sprintf("A%v", i+7), &row); err != nil {
			return err
		}
	}
	return nil
}

func setFooter(f *excelize.File, cfg *config.ExportConfig, tbConfig *config.TableConfig, values reflect.Value) error {
	return nil
}

func setStyle(f *excelize.File, cfg *config.ExportConfig, tbConfig *config.TableConfig) error {
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
	if err := f.SetCellStyle(cfg.SheetName, "A6", tbConfig.LastCellCol, styleBoldCenter); err != nil {
		return err
	}
	return nil
}

func setTable(f *excelize.File, cfg *config.ExportConfig, tbConfig *config.TableConfig) error {
	refRange := fmt.Sprintf("A6:%v", tbConfig.LastCell)
	if err := f.AddTable(cfg.SheetName, &excelize.Table{
		Range:     refRange,
		Name:      cfg.TableName,
		StyleName: "TableStyleLight9",
	}); err != nil {
		return err
	}
	return nil
}
