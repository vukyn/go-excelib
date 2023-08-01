package excelib

import (
	"fmt"
	"reflect"

	"github.com/vukyn/go-excelib/config"
	"github.com/vukyn/go-excelib/utils"
	"github.com/xuri/excelize/v2"
)

type Excelib interface {
	SetFileName(path, name string) string
	Process(sheetName string, objs interface{}) error
	ProcessStream(sheetName string, objs interface{}) error
	Save() error
}

type excelib struct {
	cfg    *config.ExportConfig
	File   *excelize.File
	Stream *excelize.StreamWriter
}

func New(cfg *config.ExportConfig) Excelib {
	validateConfig(cfg)
	return &excelib{
		cfg: cfg,
	}
}

func (e *excelib) SetFileName(path, name string) string {
	return e.cfg.SetFileName(path, name)
}

func (e *excelib) Process(sheetName string, objs interface{}) error {

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
	if e.File == nil {
		e.File = excelize.NewFile()
		e.File.SetSheetName("Sheet1", sheetName)
		e.File.SetActiveSheet(0)
	} else {
		index, err := e.File.NewSheet(sheetName)
		if err != nil {
			return err
		}
		e.File.SetActiveSheet(index)
	}

	e.recalculateConfig(tbConfig, values)

	if err := e.setMetadata(tbConfig); err != nil {
		return err
	}
	
	if err := e.setHeader(tbConfig, values); err != nil {
		return err
	}

	if e.cfg.HasDescription {
		if err := e.setDescription(tbConfig, values); err != nil {
			return err
		}
	}

	if err := e.setBody(tbConfig, values); err != nil {
		return err
	}

	if e.cfg.HasFooter {
		if err := e.setFooter(tbConfig, values); err != nil {
			return err
		}
	}

	if err := e.setStyle(tbConfig); err != nil {
		return err
	}

	if err := e.setTable(tbConfig); err != nil {
		return err
	}

	return nil
}

// ProcessStream writing data on a new existing empty worksheet with large amounts of data.
func (e *excelib) ProcessStream(sheetName string, objs interface{}) error {

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

	// init stream
	if e.File == nil {
		e.File = excelize.NewFile()
		e.File.SetSheetName("Sheet1", sheetName)
		stream, err := e.File.NewStreamWriter(sheetName)
		if err != nil {
			return err
		}
		e.Stream = stream
	}
	if e.Stream == nil {
		stream, err := e.File.NewStreamWriter(sheetName)
		if err != nil {
			return err
		}
		e.Stream = stream
	}

	e.recalculateConfig(tbConfig, values)

	if err := e.setStreamMetadata(tbConfig); err != nil {
		return err
	}

	if err := e.setStreamHeader(tbConfig, values); err != nil {
		return err
	}

	if e.cfg.HasDescription {
		if err := e.setStreamDescription(tbConfig, values); err != nil {
			return err
		}
	}

	if err := e.setStreamBody(tbConfig, values); err != nil {
		return err
	}

	if e.cfg.HasFooter {
		if err := e.setStreamFooter(tbConfig, values); err != nil {
			return err
		}
	}

	if err := e.setStreamTable(tbConfig); err != nil {
		return err
	}

	if err := e.Stream.Flush(); err != nil {
		return err
	}

	return nil
}

func (e *excelib) Save() error {
	if err := utils.CreateFilePath(e.cfg.GetFileName()); err != nil {
		return err
	}
	if err := e.File.SaveAs(e.cfg.GetFileName()); err != nil {
		return err
	}
	return nil
}
