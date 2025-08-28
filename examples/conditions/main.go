package main

import (
	"log"
	"path/filepath"

	"github.com/xuri/excelize/v2"

	exceltemplar "github.com/nikitaxru/exceltemplar"
)

// Example: conditions with exists(), len(), and plain if/else blocks
func main() {
	wd := "."
	templatePath := filepath.Join(wd, "conditions_template.xlsx")
	outputPath := filepath.Join(wd, "conditions_output.xlsx")

	f := excelize.NewFile()
	sheet := "Sheet1"
	_ = f.SetCellValue(sheet, "A1", "Team: {{= $.team.name}}")
	_ = f.SetCellValue(sheet, "A2", "{{#if len($.team.members) > 0}}")
	_ = f.SetCellValue(sheet, "A3", "Members present")
	_ = f.SetCellValue(sheet, "A4", "{{else}}")
	_ = f.SetCellValue(sheet, "A5", "No members")
	_ = f.SetCellValue(sheet, "A6", "{{/if}}")
	_ = f.SetCellValue(sheet, "A8", "{{= iif(exists($.meta.owner), $.meta.owner, 'unknown') }}")
	if err := f.SaveAs(templatePath); err != nil {
		log.Fatalf("save template: %v", err)
	}

	json := `{"team": {"name": "Core", "members": []}, "meta": {"version": "1.0"}}`
	if err := exceltemplar.WriteResultsWithTemplate(templatePath, outputPath, []string{json}); err != nil {
		log.Fatalf("render: %v", err)
	}
	log.Printf("written %s", outputPath)
}
