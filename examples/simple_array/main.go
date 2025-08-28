package main

import (
	"log"
	"path/filepath"

	"github.com/xuri/excelize/v2"

	exceltemplar "github.com/nikitaxru/exceltemplar"
)

// Example: simple array with {{#each}}
func main() {
	tmpDir := "."
	templatePath := filepath.Join(tmpDir, "simple_array_template.xlsx")
	outputPath := filepath.Join(tmpDir, "simple_array_output.xlsx")

	f := excelize.NewFile()
	sheet := "Sheet1"
	_ = f.SetCellValue(sheet, "A1", "Defect table")
	_ = f.SetCellValue(sheet, "A2", "{{#each $.defect_table as $d}}")
	_ = f.SetCellValue(sheet, "A3", "{{= $d.unit_code}}")
	_ = f.SetCellValue(sheet, "B3", "{{= $d.unit_name}}")
	_ = f.SetCellValue(sheet, "C3", "{{= $d.defect_code}}")
	_ = f.SetCellValue(sheet, "D3", "{{= $d.defect_name}}")
	_ = f.SetCellValue(sheet, "A4", "{{/each}}")
	if err := f.SaveAs(templatePath); err != nil {
		log.Fatalf("save template: %v", err)
	}

	json := `{
        "defect_table": [
            {"unit_code":"U001","unit_name":"Unit 1","defect_code":"D001","defect_name":"Defect 1"},
            {"unit_code":"U002","unit_name":"Unit 2","defect_code":"D002","defect_name":"Defect 2"}
        ]
    }`

	if err := exceltemplar.WriteResultsWithTemplate(templatePath, outputPath, []string{json}); err != nil {
		log.Fatalf("render: %v", err)
	}

	log.Printf("written %s", outputPath)
}
