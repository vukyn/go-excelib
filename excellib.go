package excelib

import (
	"fmt"
	"reflect"

	"github.com/vukyn/go-excelib/config"
	"github.com/xuri/excelize/v2"
)

func ExportExcel(objs interface{}, cfg *config.ExportConfig) error {

	// validate
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
	validateConfig(cfg)

	// init value
	tbConfig := &config.TableConfig{}
	tbConfig.NumRows = values.Len()
	tbConfig.NumFields = values.Index(0).NumField() - 1
	tbConfig.ResetTableConfig()

	// init file
	f := excelize.NewFile()
	index, err := f.NewSheet(cfg.SheetName)
	if err != nil {
		return err
	}
	f.SetActiveSheet(index)

	if err := setMetadata(f, cfg, tbConfig); err != nil {
		return err
	}

	if err := setHeader(f, cfg, tbConfig, values); err != nil {
		return err
	}

	if err := setBody(f, cfg, tbConfig, values); err != nil {
		return err
	}

	if err := setFooter(f, cfg, tbConfig, values); err != nil {
		return err
	}

	if err := setStyle(f, cfg, tbConfig); err != nil {
		return err
	}

	if err := setTable(f, cfg, tbConfig); err != nil {
		return err
	}

	if err := f.SaveAs(generateExportPath(cfg.FileName)); err != nil {
		return err
	}

	return nil
}
