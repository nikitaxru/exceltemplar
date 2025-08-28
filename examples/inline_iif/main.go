package main

import (
	"log"
	"path/filepath"

	"github.com/xuri/excelize/v2"

	exceltemplar "github.com/nikitaxru/exceltemplar"
)

// Example: inline conditional expressions using iif(cond, then, else)
func main() {
	tmpDir := "."
	templatePath := filepath.Join(tmpDir, "inline_iif.xlsx")
	outputPath := filepath.Join(tmpDir, "inline_iif_output.xlsx")

	f := excelize.NewFile()
	sheet := "Sheet1"
	_ = f.SetCellValue(sheet, "A1", "{{= iif($.flag, 'ON', 'OFF') }}")
	_ = f.SetCellValue(sheet, "B1", "{{= iif(len($.arr) > 1, 'MANY', 'ONE') }}")
	_ = f.SetCellValue(sheet, "C1", "{{= iif($.missing, 'YES', 'NO') }}")
	_ = f.SetCellValue(sheet, "D1", "static")
	if err := f.SaveAs(templatePath); err != nil {
		log.Fatalf("save template: %v", err)
	}

	json := `{"flag": true, "arr": [1,2]}`
	if err := exceltemplar.WriteResultsWithTemplate(templatePath, outputPath, []string{json}); err != nil {
		log.Fatalf("render: %v", err)
	}

	log.Printf("written %s", outputPath)
}
