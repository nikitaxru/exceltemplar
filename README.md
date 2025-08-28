<div align="center">
  <img src="logo.png" alt="Logo" width="200" />
</div>

# ExcelTemplar

ExcelTemplar is a lightweight Excel templating engine for Go. It lets you design `.xlsx` templates in Excel using simple directives and render them with JSON data.

- Minimal, dependency-light core
- Works on top of `excelize`
- Explicit template syntax with `{{ }}` blocks

Repository: `https://github.com/nikitaxru/exceltemplar`

## Installation

```bash
go get github.com/nikitaxru/exceltemplar
```

## Quick Start

```go
package main

import (
    "log"

    excel "github.com/nikitaxru/exceltemplar"
)

func main() {
    // Path to an .xlsx file containing template directives
    templatePath := "./template.xlsx"
    outputPath := "./output.xlsx"

    // One or more JSON strings providing data for rendering
    data := []string{
        `{"tasks":[{"name":"Implement feature X","priority":"high"}]}`,
    }

    if err := excel.WriteResultsWithTemplate(templatePath, outputPath, data); err != nil {
        log.Fatalf("render failed: %v", err)
    }
}
```

## Template Syntax (in Excel cells)

- **Expression**: `{{= expr}}`
- **Each (list)**: `{{#each $.items as $it i=$i}} ... {{/each}}`
- **Each (object)**: `{{#each-obj $.dict as $k $v}} ... {{/each-obj}}`
- **If/Else**: `{{#if expr}} ... {{else}} ... {{/if}}`
- Built-ins: `len()`, `exists()`, `join()`

Examples (place in cells):
- `{{= $.title }}`
- `{{#if len($.items) > 0}} ... {{/if}}`
- Inside each: `{{= $it.name}}`

## API

- `LoadTemplate(path string) (*Template, error)`
- `(*Template).Render(outputs []string) error` — render with one or more JSON strings
- `(*Template).Save(destPath string) error`
- Convenience: `WriteResultsWithTemplate(templatePath, destPath string, outputs []string) error`

Utility:
- `NormalizeForExcel(jsonStrings []string) []string` — normalizes JSON for predictable rendering

## Full documentation

Complete documentation is available in the `docs/` folder:

- **English**: [Excel Template Guide](docs/excel_template_guide_en.md) - Complete guide with examples and API reference
- **Русский**: [Руководство по шаблонизатору Excel](docs/excel_template_guide_ru.md) - Полное руководство с примерами и справочником API

Both versions contain:
- Template syntax reference
- Path anchors and context explanation
- Practical examples with JSON data
- Excel template best practices
- Programmatic API usage
- Behavior with missing data

## Testing

```bash
go test ./...
```

## License

MIT © 2025 Nikita Samoylov
