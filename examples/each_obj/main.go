package main

import (
	"log"
	"path/filepath"

	"github.com/xuri/excelize/v2"

	exceltemplar "github.com/nikitaxru/exceltemplar"
)

// Example: iterate over object (map) using {{#each-obj}}
func main() {
	wd := "."
	templatePath := filepath.Join(wd, "each_obj_template.xlsx")
	outputPath := filepath.Join(wd, "each_obj_output.xlsx")

	f := excelize.NewFile()
	sheet := "Sheet1"
	_ = f.SetCellValue(sheet, "A1", "Key")
	_ = f.SetCellValue(sheet, "B1", "Value")
	_ = f.SetCellValue(sheet, "A2", "{{#each-obj $.meta as $k $v}}")
	_ = f.SetCellValue(sheet, "A3", "{{= $k}}")
	_ = f.SetCellValue(sheet, "B3", "{{= $v}}")
	_ = f.SetCellValue(sheet, "A4", "{{/each}}")
	if err := f.SaveAs(templatePath); err != nil {
		log.Fatalf("save template: %v", err)
	}

	json := `{"meta": {"version": "1.0", "owner": "dept"}}`
	if err := exceltemplar.WriteResultsWithTemplate(templatePath, outputPath, []string{json}); err != nil {
		log.Fatalf("render: %v", err)
	}
	log.Printf("written %s", outputPath)
}
