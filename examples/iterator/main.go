package main

import (
	"log"
	"path/filepath"

	"github.com/xuri/excelize/v2"

	exceltemplar "github.com/nikitaxru/exceltemplar"
)

// Example: render a list of tasks using {{#each}}
func main() {
	tmpDir := "."
	templatePath := filepath.Join(tmpDir, "iterator_template.xlsx")
	outputPath := filepath.Join(tmpDir, "iterator_output.xlsx")

	f := excelize.NewFile()
	sheet := "Sheet1"
	_ = f.SetCellValue(sheet, "A1", "Task list")
	_ = f.SetCellValue(sheet, "A2", "{{#each $.tasks as $t i=$i}}")
	_ = f.SetCellValue(sheet, "A3", "{{= $t.name}}")
	_ = f.SetCellValue(sheet, "B3", "{{= $t.priority}}")
	_ = f.SetCellValue(sheet, "C3", "{{= $t.status}}")
	_ = f.SetCellValue(sheet, "A4", "{{/each}}")
	if err := f.SaveAs(templatePath); err != nil {
		log.Fatalf("save template: %v", err)
	}

	json := `{
        "tasks": [
            {"name": "Implement feature X", "priority": "high", "status": "in-progress"},
            {"name": "Fix bug Y", "priority": "medium", "status": "pending"}
        ]
    }`

	if err := exceltemplar.WriteResultsWithTemplate(templatePath, outputPath, []string{json}); err != nil {
		log.Fatalf("render: %v", err)
	}

	log.Printf("written %s", outputPath)
}
