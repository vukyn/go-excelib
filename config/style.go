package config

import "github.com/xuri/excelize/v2"

func BoldCenter(f *excelize.File) int {
	styleBoldCenter, _ := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{
			Bold: true,
		},
		Alignment: &excelize.Alignment{
			Horizontal: "center",
			Vertical:   "center",
		},
	})
	return styleBoldCenter
}
