package main

import (
	"log"
	"path/filepath"

	"github.com/xuri/excelize/v2"

	exceltemplar "github.com/nikitaxru/exceltemplar"
)

// Example: join(array, sep, field) inside a block
func main() {
	wd := "."
	templatePath := filepath.Join(wd, "join_fields_template.xlsx")
	outputPath := filepath.Join(wd, "join_fields_output.xlsx")

	f := excelize.NewFile()
	sheet := "Sheet1"
	_ = f.SetCellValue(sheet, "A1", "{{#each $.users as $u}}")
	_ = f.SetCellValue(sheet, "A2", "{{= $u.name}}: {{= join($u.phones, \", \", \"n\")}}")
	_ = f.SetCellValue(sheet, "A3", "{{/each}}")
	if err := f.SaveAs(templatePath); err != nil {
		log.Fatalf("save template: %v", err)
	}

	json := `{
        "users": [
            {"name": "A", "phones": [{"n": "111"}, {"n": "222"}]},
            {"name": "B", "phones": [{"n": "333"}]}
        ]
    }`

	if err := exceltemplar.WriteResultsWithTemplate(templatePath, outputPath, []string{json}); err != nil {
		log.Fatalf("render: %v", err)
	}
	log.Printf("written %s", outputPath)
}
