package excelib

import (
	"fmt"
	"reflect"
	"strings"
	"time"

	"github.com/vukyn/go-excelib/config"
	"github.com/xuri/excelize/v2"
)

type Excelib struct {
	cfg  *config.ExportConfig
	File *excelize.File
	// Stream *excelize.StreamWriter
}

func Init(cfg *config.ExportConfig) *Excelib {
	validateConfig(cfg)
	return &Excelib{
		cfg: cfg,
	}
}

func (e *Excelib) Process(objs interface{}) error {

	// validate
	values := reflect.ValueOf(objs)
	if values.Kind() != reflect.Slice {
		return fmt.Errorf("ExportExcel: objs must be a slice")
	}
	if values.Len() == 0 {
		return fmt.Errorf("ExportExcel: objs must be not empty")
	}
	if values.Len() > config.MAX_ROW {
		return fmt.Errorf("ExportExcel: objs must be less than %v records", config.MAX_ROW)
	}
	if values.Index(0).Kind() != reflect.Struct {
		return fmt.Errorf("ExportExcel: objs must be a slice of struct")
	}

	// init value
	tbConfig := &config.TableConfig{}
	tbConfig.NumRows = values.Len()
	tbConfig.NumFields = values.Index(0).NumField() - 1
	tbConfig.ResetTableConfig()

	// init file
	f := excelize.NewFile()
	index, err := f.NewSheet(e.cfg.SheetName)
	if err != nil {
		return err
	}
	f.SetActiveSheet(index)

	if err := setHeader(f, e.cfg, tbConfig, values); err != nil {
		return err
	}

	if e.cfg.HasDescription {
		if err := setDescription(f, e.cfg, tbConfig, values); err != nil {
			return err
		}
	}

	if err := setBody(f, e.cfg, tbConfig, values); err != nil {
		return err
	}

	if e.cfg.HasFooter {
		if err := setFooter(f, e.cfg, tbConfig, values); err != nil {
			return err
		}
	}

	if err := setStyle(f, e.cfg, tbConfig); err != nil {
		return err
	}

	if err := setTable(f, e.cfg, tbConfig); err != nil {
		return err
	}

	if err := setMetadata(f, e.cfg, tbConfig); err != nil {
		return err
	}

	e.File = f

	return nil
}

func (e *Excelib) ExportToFile(filePath, fileName string) error {
	pathExport := filePath + "/{time}_{file}.xlsx"
	pathExport = strings.ReplaceAll(pathExport, "{time}", time.Now().Format("2006_01_02_15_04_05"))
	pathExport = strings.ReplaceAll(pathExport, "{file}", fileName)
	if err := e.File.SaveAs(pathExport); err != nil {
		return err
	}
	return nil
}
