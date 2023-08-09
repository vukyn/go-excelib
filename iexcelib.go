package excelib

import (
	"github.com/vukyn/go-excelib/config"
	"github.com/xuri/excelize/v2"
)

type Excelib interface {
	SetFileName(path, name string) string
	Process(sheetName string, objs interface{}) error
	ProcessStream(sheetName string, objs interface{}) error
	Save() error

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
}

type excelib struct {
	cfg    *config.ExportConfig
	File   *excelize.File
	Stream *excelize.StreamWriter
}

func Init(cfg *config.ExportConfig) Excelib {
	validateConfig(cfg)
	e := &excelib{cfg: cfg}
	e.File = excelize.NewFile()
	e.File.SetActiveSheet(0)
	return e
}
