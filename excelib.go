package excelib

import (
	"github.com/vukyn/go-excelib/config"
	"github.com/vukyn/go-excelib/utils"
	"github.com/jinzhu/copier"
	"github.com/xuri/excelize/v2"
)

func (e *excelib) SetFileName(path, name string) string {
	return e.exportCfg.SetFileName(path, name)
}

func (e *excelib) SetMetadata(metadata *config.Metadata) error {
	docProps := &excelize.DocProperties{}
	copier.Copy(docProps, metadata)
	if err := e.file.SetDocProps(docProps); err != nil {
		return err
	}
	return nil
}

func (e *excelib) Process(sheetName string, objs []interface{}) error {

	// validate
	values, err := validateObj(objs)
	if err != nil {
		return err
	}

	// init value
	e.tbCfg.NumRows = len(objs)
	e.tbCfg.NumFields = values.NumField()
	e.recalculateConfig(values)

	// init sheet
	if e.file.GetSheetName(0) == "Sheet1" {
		e.file.SetSheetName("Sheet1", sheetName)
		e.file.SetActiveSheet(0)
	} else {
		index, err := e.file.NewSheet(sheetName)
		if err != nil {
			return err
		}
		e.file.SetActiveSheet(index)
	}

	if err := e.setMetadata(); err != nil {
		return err
	}

	if err := e.setHeader(values); err != nil {
		return err
	}

	if e.exportCfg.HasDescription {
		if err := e.setDescription(values); err != nil {
			return err
		}
	}

	if err := e.setBody(objs); err != nil {
		return err
	}

	if e.exportCfg.HasFooter {
		if err := e.setFooter(values); err != nil {
			return err
		}
	}

	if err := e.setStyle(); err != nil {
		return err
	}

	if err := e.setTable(); err != nil {
		return err
	}

	return nil
}

func (e *excelib) ProcessStream(sheetName string, objs []interface{}) error {

	// validate
	values, err := validateObj(objs)
	if err != nil {
		return err
	}

	// init value
	e.tbCfg.NumRows = values.Len()
	e.tbCfg.NumFields = values.Index(0).NumField()
	e.recalculateConfig(values)

	// init sheet
	if e.file.GetSheetName(0) == "Sheet1" {
		e.file.SetSheetName("Sheet1", sheetName)
		stream, err := e.file.NewStreamWriter(sheetName)
		if err != nil {
			return err
		}
		e.stream = stream
	} else {
		index, err := e.file.NewSheet(sheetName)
		if err != nil {
			return err
		}
		e.file.SetActiveSheet(index)
		stream, err := e.file.NewStreamWriter(sheetName)
		if err != nil {
			return err
		}
		e.stream = stream
	}

	if err := e.setStreamMetadata(); err != nil {
		return err
	}

	e.setStreamHeader(values)
	if err := e.writeStreamHeader(); err != nil {
		return err
	}

	if e.exportCfg.HasDescription {
		e.setStreamDescription(values)
		if err := e.writeStreamDescription(); err != nil {
			return err
		}
	}

	if err := e.setStreamBody(objs); err != nil {
		return err
	}

	if e.exportCfg.HasFooter {
		e.setStreamFooter(values)
		if err := e.writeStreamFooter(); err != nil {
			return err
		}
	}

	if err := e.setStreamTable(); err != nil {
		return err
	}

	if err := e.stream.Flush(); err != nil {
		return err
	}

	return nil
}

func (e *excelib) SaveAs() error {
	e.file.SetActiveSheet(0)
	if err := utils.CreateFilePath(e.exportCfg.GetFileName()); err != nil {
		return err
	}
	if err := e.file.SaveAs(e.exportCfg.GetFileName()); err != nil {
		return err
	}
	return nil
}

func (e *excelib) StartStream(sheetName string) error {
	if e.file.GetSheetName(0) == "Sheet1" {
		e.file.SetSheetName("Sheet1", sheetName)
		stream, err := e.file.NewStreamWriter(sheetName)
		if err != nil {
			return err
		}
		e.stream = stream
	} else {
		index, err := e.file.NewSheet(sheetName)
		if err != nil {
			return err
		}
		e.file.SetActiveSheet(index)
		stream, err := e.file.NewStreamWriter(sheetName)
		if err != nil {
			return err
		}
		e.stream = stream
	}
	if err := e.SaveAs(); err != nil {
		return err
	}
	return nil
}

func (e *excelib) AppendStream(objs []interface{}) error {

	// validate
	values, err := validateObj(objs)
	if err != nil {
		return err
	}

	// init value
	if !e.hasInit {
		e.tbCfg.NumFields = values.NumField()
		e.recalculateConfig(values)
		e.hasInit = true
	}

	if !e.hasMetadata {
		if err := e.setStreamMetadata(); err != nil {
			return err
		}
		e.hasMetadata = true
	}
	if !e.hasHeader {
		e.setStreamHeader(values)
		if err := e.writeStreamHeader(); err != nil {
			return err
		}
		e.hasHeader = true
	}
	if !e.hasDescription && e.exportCfg.HasDescription {
		e.setStreamDescription(values)
		if err := e.writeStreamDescription(); err != nil {
			return err
		}
		e.hasDescription = true
	}
	if !e.hasFooter && e.exportCfg.HasFooter {
		e.setStreamFooter(values)
		e.hasFooter = true
	}

	if err := e.setStreamBody(objs); err != nil {
		return err
	}
	if err := e.file.Save(); err != nil {
		return err
	}
	return nil
}

func (e *excelib) StopStream() error {
	if e.hasFooter && e.exportCfg.HasFooter {
		if err := e.writeStreamFooter(); err != nil {
			return err
		}
	}
	if err := e.setStreamTable(); err != nil {
		return err
	}
	if err := e.stream.Flush(); err != nil {
		return err
	}
	if err := e.file.Save(); err != nil {
		return err
	}
	return nil
}
