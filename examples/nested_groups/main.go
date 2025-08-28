package main

import (
	"log"
	"path/filepath"

	"github.com/xuri/excelize/v2"

	exceltemplar "github.com/nikitaxru/exceltemplar"
)

// Example: nested groups (projects -> milestones -> tasks) with indices
func main() {
	wd := "."
	templatePath := filepath.Join(wd, "nested_groups_template.xlsx")
	outputPath := filepath.Join(wd, "nested_groups_output.xlsx")

	f := excelize.NewFile()
	sheet := "Sheet1"
	_ = f.SetCellValue(sheet, "A1", "{{#each $.projects as $p}}")
	_ = f.SetCellValue(sheet, "A2", "Project: {{= $p.name}} (milestones: {{= len($p.milestones)}})")
	_ = f.SetCellValue(sheet, "A3", "{{#if len($p.milestones) > 0}}")
	_ = f.SetCellValue(sheet, "A4", "{{#each $p.milestones as $m}}")
	_ = f.SetCellValue(sheet, "A5", "Milestone: {{= $m.name}} (tasks: {{= len($m.tasks)}})")
	_ = f.SetCellValue(sheet, "A6", "{{#each $m.tasks as $t i=$ti}}")
	_ = f.SetCellValue(sheet, "A7", "Task {{= $ti+1}}: {{= $t.t}}")
	_ = f.SetCellValue(sheet, "A8", "{{/each}}")
	_ = f.SetCellValue(sheet, "A9", "{{/each}}")
	_ = f.SetCellValue(sheet, "A10", "{{else}}")
	_ = f.SetCellValue(sheet, "A11", "No milestones")
	_ = f.SetCellValue(sheet, "A12", "{{/if}}")
	_ = f.SetCellValue(sheet, "A13", "{{/each}}")
	if err := f.SaveAs(templatePath); err != nil {
		log.Fatalf("save template: %v", err)
	}

	json := `{
        "projects": [
          {"name": "P1", "milestones": [
            {"name": "M1", "tasks": [{"t": "T1"}, {"t": "T2"}]},
            {"name": "M2", "tasks": [{"t": "T3"}]}
          ]}
        ]
      }`

	if err := exceltemplar.WriteResultsWithTemplate(templatePath, outputPath, []string{json}); err != nil {
		log.Fatalf("render: %v", err)
	}
	log.Printf("written %s", outputPath)
}
