package config

import "github.com/xuri/excelize/v2"

func Center(f *excelize.File) int {
	styleCenter, _ := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{
			Horizontal: "center",
			Vertical:   "center",
		},
	})
	return styleCenter
}


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
