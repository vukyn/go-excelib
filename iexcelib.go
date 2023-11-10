package excelib

import (
	"github.com/vukyn/go-excelib/config"
	"github.com/xuri/excelize/v2"
)

type Excelib interface {
	SetFileName(path, name string) string
	Process(sheetName string, objs []interface{}) error
	// ProcessStream writing data on a new existing empty worksheet with large amounts of data.
	ProcessStream(sheetName string, objs []interface{}) error
	SaveAs() error

	// SetMetadata provides a function to set document properties.
	//
	// For example:
	//
	//	err := e.SetMetadata(&config.Metadata{
	//	    Creator:        "Go Excelib",
	//	    Description:    "This file created by Go Excelib",
	//	    Subject:        "Test Subject",
	//	    Title:          "Test Title",
	//	    Language:       "en-US",
	//	    Version:        "1.0.0",
	//	})
	SetMetadata(metadata *config.Metadata) error

	StartStream(sheetName string) error
	AppendStream(objs []interface{}) error
	StopStream() error
}

type excelib struct {
	hasInit        bool
	hasHeader      bool
	hasMetadata    bool
	hasFooter      bool
	hasDescription bool
	headers        []interface{}
	footers        []interface{}
	descriptions   []interface{}
	exportCfg      *config.ExportConfig
	tbCfg          *config.TableConfig
	file           *excelize.File
	stream         *excelize.StreamWriter
}

func Init(cfg *config.ExportConfig) Excelib {
	validateConfig(cfg)
	e := &excelib{
		exportCfg: cfg,
		tbCfg:     &config.TableConfig{},
	}
	e.file = excelize.NewFile()
	e.file.SetActiveSheet(0)
	return e
}
