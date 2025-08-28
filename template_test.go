package exceltemplar_test

import (
	"fmt"
	"path/filepath"
	"strings"
	"testing"

	"github.com/stretchr/testify/suite"
	"github.com/xuri/excelize/v2"

	"github.com/nikitaxru/exceltemplar"
)

// TemplateSuite — сьют тестов движка шаблонов Excel
type TemplateSuite struct {
	suite.Suite
}

// SetupSubTest — при необходимости переинициализируйте состояние для каждого subtest
func (s *TemplateSuite) SetupSubTest() {}

// Runner
func TestTemplateSuite(t *testing.T) {
	suite.Run(t, new(TemplateSuite))
}

// TestIterator — проверяет рендер each-блока со списком задач и удаление управляющих строк
func (s *TemplateSuite) TestIterator() {
	tmpDir := s.T().TempDir()
	tmpTemplate := filepath.Join(tmpDir, "iterator_template.xlsx")

	f := excelize.NewFile()
	sheet := "Sheet1"

	// Строка 1: заголовок
	_ = f.SetCellValue(sheet, "A1", "Список задач")
	// Строка 2: открывающий итератор
	_ = f.SetCellValue(sheet, "A2", "{{#each $.tasks as $t i=$i}}")
	// Строка 3: шаблон для каждой задачи
	_ = f.SetCellValue(sheet, "A3", "{{= $t.name}}")
	_ = f.SetCellValue(sheet, "B3", "{{= $t.priority}}")
	_ = f.SetCellValue(sheet, "C3", "{{= $t.status}}")
	// Строка 4: закрываем итератор
	_ = f.SetCellValue(sheet, "A4", "{{/each}}")

	s.Require().NoError(f.SaveAs(tmpTemplate), "save template")

	// JSON с данными
	json := `{
        "tasks": [
            {"name": "Implement feature X", "priority": "high", "status": "in-progress"},
            {"name": "Fix bug Y", "priority": "medium", "status": "pending"}
        ]
    }`

	// Применяем шаблон
	tmpOutput := filepath.Join(tmpDir, "iterator_output.xlsx")
	s.Require().NoError(exceltemplar.WriteResultsWithTemplate(tmpTemplate, tmpOutput, []string{json}), "render")

	// Проверяем результат
	result, err := excelize.OpenFile(tmpOutput)
	s.Require().NoError(err, "open result")

	// Проверяем заголовок
	if v, _ := result.GetCellValue(sheet, "A1"); true {
		s.Assert().Equal("Список задач", v, "A1")
	}

	// Первая задача должна начинаться со строки 2 (строка 1 - заголовок)
	if v, _ := result.GetCellValue(sheet, "A2"); true {
		s.Assert().Equal("Implement feature X", v, "A2")
	}
	if v, _ := result.GetCellValue(sheet, "B2"); true {
		s.Assert().Equal("high", v, "B2")
	}
	if v, _ := result.GetCellValue(sheet, "C2"); true {
		s.Assert().Equal("in-progress", v, "C2")
	}

	// Вторая задача
	if v, _ := result.GetCellValue(sheet, "A3"); true {
		s.Assert().Equal("Fix bug Y", v, "A3")
	}
	if v, _ := result.GetCellValue(sheet, "B3"); true {
		s.Assert().Equal("medium", v, "B3")
	}
	if v, _ := result.GetCellValue(sheet, "C3"); true {
		s.Assert().Equal("pending", v, "C3")
	}

	// Должно быть 3 строки: заголовок + 2 задачи
	rows, err := result.GetRows(sheet)
	s.Require().NoError(err, "get rows")
	s.Assert().Equal(3, len(rows), "rows count")
}

// TestIteratorDebug — минимальный итератор без строгих проверок содержимого
func (s *TemplateSuite) TestIteratorDebug() {
	tmpDir := s.T().TempDir()
	tmpTemplate := filepath.Join(tmpDir, "debug_template.xlsx")

	f := excelize.NewFile()
	sheet := "Sheet1"

	// Простой случай: один итератор, одна строка внутри (новый синтаксис)
	f.SetCellValue(sheet, "A1", "{{#each $.items as $it}}")
	f.SetCellValue(sheet, "A2", "{{= $it.name}}")
	f.SetCellValue(sheet, "A3", "{{/each}}")

	s.Require().NoError(f.SaveAs(tmpTemplate), "save template")

	// Пропускаем диагностический вывод структуры AST (устаревшие методы удалены)

	// JSON с данными
	json := `{"items": [{"name": "Item 1"}, {"name": "Item 2"}]}`
	tmpOutput := filepath.Join(tmpDir, "debug_output.xlsx")
	s.Require().NoError(exceltemplar.WriteResultsWithTemplate(tmpTemplate, tmpOutput, []string{json}), "render")

	// Проверяем результат
	result, err := excelize.OpenFile(tmpOutput)
	s.Require().NoError(err, "open result")

	rows, _ := result.GetRows(sheet)
	s.T().Logf("Output rows count: %d", len(rows))
	for i, row := range rows {
		s.T().Logf("Row %d: %v", i, row)
	}
}

// TestBoolExpr — проверяет ветвление if/else и функции len/exists
func (s *TemplateSuite) TestBoolExpr() {
	tmpDir := s.T().TempDir()
	tmpTemplate := filepath.Join(tmpDir, "bool_expr.xlsx")

	f := excelize.NewFile()
	sheet := "Sheet1"

	// Условие с and/or/not и функциями len/exists
	_ = f.SetCellValue(sheet, "A1", "Header")
	_ = f.SetCellValue(sheet, "A2", "{{#if len($.arr) > 1 and exists($.flag)}}")
	_ = f.SetCellValue(sheet, "A3", "OK")
	_ = f.SetCellValue(sheet, "A4", "{{else}}")
	_ = f.SetCellValue(sheet, "A5", "NO")
	_ = f.SetCellValue(sheet, "A6", "{{/if}}")

	s.Require().NoError(f.SaveAs(tmpTemplate), "save template")

	json := `{"arr":[1,2],"flag":true}`
	tmpOutput := filepath.Join(tmpDir, "bool_expr_output.xlsx")
	s.Require().NoError(exceltemplar.WriteResultsWithTemplate(tmpTemplate, tmpOutput, []string{json}), "render")

	result, err := excelize.OpenFile(tmpOutput)
	s.Require().NoError(err, "open result")
	// Учитывая удаление строк-маркеров ({{#if}}, {{else}}, {{/if}}), строки сдвигаются.
	// Проверяем содержимое по факту: должен присутствовать "OK" и отсутствовать "NO".
	okCount := 0
	noCount := 0
	rows, _ := result.GetRows(sheet)
	for _, r := range rows {
		for _, c := range r {
			v := strings.TrimSpace(c)
			if v == "OK" {
				okCount++
			}
			if v == "NO" {
				noCount++
			}
		}
	}
	s.Require().GreaterOrEqual(okCount, 1, "expected OK >= 1")
	s.Require().Equal(0, noCount, "expected NO == 0")
}

// TestScalarPlaceholder — проверяет подстановку скалярного значения без any each/if
func (s *TemplateSuite) TestScalarPlaceholder() {
	tmpDir := s.T().TempDir()
	tmpTemplate := filepath.Join(tmpDir, "scalar_placeholder.xlsx")

	f := excelize.NewFile()
	sheet := "Sheet1"

	// Статические подписи и скалярные плейсхолдеры
	_ = f.SetCellValue(sheet, "A1", "Имя пользователя")
	_ = f.SetCellValue(sheet, "B1", "{{= $.user.name}}")
	_ = f.SetCellValue(sheet, "A2", "Приветствие")
	_ = f.SetCellValue(sheet, "B2", "Hello, {{= $.user.name}}!")
	_ = f.SetCellValue(sheet, "A3", "Отсутствующее поле (должно быть пусто)")
	_ = f.SetCellValue(sheet, "B3", "{{= $.user.age}}")

	s.Require().NoError(f.SaveAs(tmpTemplate), "save template")

	// JSON без массивов — только вложенный объект
	json := `{"user": {"name": "nikita"}}`
	tmpOutput := filepath.Join(tmpDir, "scalar_placeholder_output.xlsx")
	s.Require().NoError(exceltemplar.WriteResultsWithTemplate(tmpTemplate, tmpOutput, []string{json}), "render")

	// Проверяем результат
	result, err := excelize.OpenFile(tmpOutput)
	s.Require().NoError(err, "open result")

	if v, _ := result.GetCellValue(sheet, "B1"); true {
		s.Assert().Equal("nikita", strings.TrimSpace(v), "B1 should equal user.name")
	}
	if v, _ := result.GetCellValue(sheet, "B2"); true {
		s.Assert().Equal("Hello, nikita!", strings.TrimSpace(v), "B2 should render inline text + expr")
	}
	if v, _ := result.GetCellValue(sheet, "B3"); true {
		s.Assert().Equal("", strings.TrimSpace(v), "B3 should be empty when value is absent")
	}
}

// TestInlineIIF — проверяет inline-условия через iif(cond, then, else) в ячейках одной строки
func (s *TemplateSuite) TestInlineIIF() {
	tmpDir := s.T().TempDir()
	tmpTemplate := filepath.Join(tmpDir, "inline_iif.xlsx")

	f := excelize.NewFile()
	sheet := "Sheet1"

	// Одна строка с несколькими независимыми условиями в разных ячейках
	_ = f.SetCellValue(sheet, "A1", "{{= iif($.flag, 'ON', 'OFF') }}")
	_ = f.SetCellValue(sheet, "B1", "{{= iif(len($.arr) > 1, 'MANY', 'ONE') }}")
	_ = f.SetCellValue(sheet, "C1", "{{= iif($.missing, 'YES', 'NO') }}")
	_ = f.SetCellValue(sheet, "D1", "static")

	s.Require().NoError(f.SaveAs(tmpTemplate), "save template")

	json := `{"flag": true, "arr": [1,2]}`
	tmpOutput := filepath.Join(tmpDir, "inline_iif_output.xlsx")
	s.Require().NoError(exceltemplar.WriteResultsWithTemplate(tmpTemplate, tmpOutput, []string{json}), "render")

	result, err := excelize.OpenFile(tmpOutput)
	s.Require().NoError(err, "open result")

	if v, _ := result.GetCellValue(sheet, "A1"); true {
		s.Assert().Equal("ON", strings.TrimSpace(v), "A1")
	}
	if v, _ := result.GetCellValue(sheet, "B1"); true {
		s.Assert().Equal("MANY", strings.TrimSpace(v), "B1")
	}
	if v, _ := result.GetCellValue(sheet, "C1"); true {
		s.Assert().Equal("NO", strings.TrimSpace(v), "C1 when missing is falsy")
	}
	if v, _ := result.GetCellValue(sheet, "D1"); true {
		s.Assert().Equal("static", strings.TrimSpace(v), "D1 static remains")
	}
}

// TestSimpleArray — проверяет простой each-блок с плоским массивом
func (s *TemplateSuite) TestSimpleArray() {
	tmpDir := s.T().TempDir()
	tmpTemplate := filepath.Join(tmpDir, "simple_array_template.xlsx")

	f := excelize.NewFile()
	sheet := "Sheet1"

	// Создаём шаблон с простым массивом (новый синтаксис)
	f.SetCellValue(sheet, "A1", "Таблица дефектов")
	f.SetCellValue(sheet, "A2", "{{#each $.defect_table as $d}}")
	f.SetCellValue(sheet, "A3", "{{= $d.unit_code}}")
	f.SetCellValue(sheet, "B3", "{{= $d.unit_name}}")
	f.SetCellValue(sheet, "C3", "{{= $d.defect_code}}")
	f.SetCellValue(sheet, "D3", "{{= $d.defect_name}}")
	f.SetCellValue(sheet, "A4", "{{/each}}")

	// Сохраняем шаблон
	s.Require().NoError(f.SaveAs(tmpTemplate), "save template")

	// JSON с данными
	json := `{
		"defect_table": [
			{
				"unit_code": "U001",
				"unit_name": "Узел 1",
				"defect_code": "D001",
				"defect_name": "Дефект 1"
			},
			{
				"unit_code": "U002",
				"unit_name": "Узел 2",
				"defect_code": "D002",
				"defect_name": "Дефект 2"
			},
			{
				"unit_code": "U003",
				"unit_name": "Узел 3",
				"defect_code": "D003",
				"defect_name": "Дефект 3"
			}
		]
	}`

	// Применяем шаблон
	tmpOutput := filepath.Join(tmpDir, "simple_array_output.xlsx")
	s.Require().NoError(exceltemplar.WriteResultsWithTemplate(tmpTemplate, tmpOutput, []string{json}), "render")

	// Проверяем результат
	result, err := excelize.OpenFile(tmpOutput)
	s.Require().NoError(err, "open result")

	// Проверяем заголовок
	if v, _ := result.GetCellValue(sheet, "A1"); true {
		s.Assert().Equal("Таблица дефектов", v, "A1")
	}

	// Проверяем первую строку данных (после заголовка)
	if v, _ := result.GetCellValue(sheet, "A2"); true {
		s.Assert().Equal("U001", v, "A2")
	}
	if v, _ := result.GetCellValue(sheet, "B2"); true {
		s.Assert().Equal("Узел 1", v, "B2")
	}
	if v, _ := result.GetCellValue(sheet, "C2"); true {
		s.Assert().Equal("D001", v, "C2")
	}
	if v, _ := result.GetCellValue(sheet, "D2"); true {
		s.Assert().Equal("Дефект 1", v, "D2")
	}

	// Проверяем вторую строку данных
	if v, _ := result.GetCellValue(sheet, "A3"); true {
		s.Assert().Equal("U002", v, "A3")
	}

	// Проверяем третью строку данных
	if v, _ := result.GetCellValue(sheet, "A4"); true {
		s.Assert().Equal("U003", v, "A4")
	}
	if v, _ := result.GetCellValue(sheet, "D4"); true {
		s.Assert().Equal("Дефект 3", v, "D4")
	}
}

// TestPathBuilding — проверяет вложенные each-блоки и построение путей
func (s *TemplateSuite) TestPathBuilding() {
	tmpDir := s.T().TempDir()
	tmpTemplate := filepath.Join(tmpDir, "path_template.xlsx")

	f := excelize.NewFile()
	sheet := "Sheet1"

	// Создаём шаблон (новый синтаксис)
	f.SetCellValue(sheet, "A1", "{{#each $.groups as $g}}")
	f.SetCellValue(sheet, "A2", "{{= $g.title}}")
	f.SetCellValue(sheet, "A3", "{{#each $g.items as $it}}")
	f.SetCellValue(sheet, "A4", "{{= $it.name}}")
	f.SetCellValue(sheet, "A5", "{{/each}}")
	f.SetCellValue(sheet, "A6", "{{/each}}")

	s.Require().NoError(f.SaveAs(tmpTemplate), "save template")

	// Пропускаем диагностический вывод путей (устаревшие методы удалены)

	// JSON с данными
	json := `{
		"groups": [
			{
				"title": "Group A",
				"items": [
					{"name": "Item A1"},
					{"name": "Item A2"}
				]
			}
		]
	}`

	// Применяем шаблон
	tmpOutput := filepath.Join(tmpDir, "path_output.xlsx")
	s.Require().NoError(exceltemplar.WriteResultsWithTemplate(tmpTemplate, tmpOutput, []string{json}), "render")

	// Проверяем результат
	result, err := excelize.OpenFile(tmpOutput)
	s.Require().NoError(err, "open result")

	rows, err := result.GetRows(sheet)
	s.Require().NoError(err, "get rows")
	s.T().Logf("Output rows: %d", len(rows))
	for i, row := range rows {
		s.T().Logf("Row %d: %v", i, row)
	}
}

// TestDeepNested — проверяет глубокую вложенность each-блоков, join и наличие ключевых строк
func (s *TemplateSuite) TestDeepNested() {
	tmpDir := s.T().TempDir()
	tmpTemplate := filepath.Join(tmpDir, "deep_nested_template.xlsx")

	f := excelize.NewFile()
	sheet := "Sheet1"

	// Структура (новый синтаксис)
	_ = f.SetCellValue(sheet, "A1", "Complex Nested Rendering")
	_ = f.SetCellValue(sheet, "A2", "{{#each $.projects as $p}}")
	_ = f.SetCellValue(sheet, "A3", "Project: {{= $p.name}} | Owners: {{= join($p.owners, ', ')}}")
	_ = f.SetCellValue(sheet, "A4", "{{#each $p.teams as $t}}")
	_ = f.SetCellValue(sheet, "A5", "Team: {{= $t.team_name}}")
	_ = f.SetCellValue(sheet, "A6", "{{#each $t.members as $m}}")
	_ = f.SetCellValue(sheet, "A7", "- {{= $m.first_name}} {{= $m.last_name}} ({{= join($m.skills, ', ')}})")
	_ = f.SetCellValue(sheet, "A8", "{{/each}}")
	_ = f.SetCellValue(sheet, "A9", "{{/each}}")
	_ = f.SetCellValue(sheet, "A10", "{{#each $p.tasks as $task}}")
	_ = f.SetCellValue(sheet, "A11", "Task: {{= $task.title}} [{{= join($task.labels, ', ')}}]")
	_ = f.SetCellValue(sheet, "A12", "{{/each}}")
	_ = f.SetCellValue(sheet, "A13", "{{/each}}")

	s.Require().NoError(f.SaveAs(tmpTemplate), "save template")

	// Глубокий JSON
	json := `{
      "projects": [
        {
          "name": "Apollo",
          "owners": ["Alice", "Bob"],
          "teams": [
            {
              "team_name": "Core",
              "members": [
                {"first_name": "John", "last_name": "Doe", "skills": ["Go", "Excel"]},
                {"first_name": "Jane", "last_name": "Roe", "skills": ["Python"]}
              ]
            },
            {
              "team_name": "Infra",
              "members": [
                {"first_name": "Max", "last_name": "Payne", "skills": ["K8s", "Terraform", "Go"]}
              ]
            }
          ],
          "tasks": [
            {"title": "Boot", "labels": ["init", "setup"]},
            {"title": "Scale", "labels": ["k8s"]}
          ]
        },
        {
          "name": "Zeus",
          "owners": ["Carol"],
          "teams": [
            {
              "team_name": "AI",
              "members": [
                {"first_name": "Eva", "last_name": "Lee", "skills": ["ML", "Go"]}
              ]
            }
          ],
          "tasks": [
            {"title": "Train", "labels": ["ml", "gpu"]}
          ]
        }
      ]
    }`

	tmpOutput := filepath.Join(tmpDir, "deep_nested_output.xlsx")
	s.Require().NoError(exceltemplar.WriteResultsWithTemplate(tmpTemplate, tmpOutput, []string{json}), "render")

	// Проверяем результат
	result, err := excelize.OpenFile(tmpOutput)
	s.Require().NoError(err, "open result")

	rows, err := result.GetRows(sheet)
	s.Require().NoError(err, "get rows")

	var all []string
	for _, row := range rows {
		for _, v := range row {
			s := strings.TrimSpace(v)
			if s != "" {
				all = append(all, s)
			}
		}
	}

	mustContain := []string{
		"Project: Apollo | Owners: Alice, Bob",
		"Team: Core",
		"Team: Infra",
		"- John Doe (Go, Excel)",
		"- Jane Roe (Python)",
		"- Max Payne (K8s, Terraform, Go)",
		"Task: Boot [init, setup]",
		"Task: Scale [k8s]",
		"Project: Zeus | Owners: Carol",
		"Team: AI",
		"- Eva Lee (ML, Go)",
		"Task: Train [ml, gpu]",
	}

	for _, needle := range mustContain {
		found := false
		for _, s2 := range all {
			if s2 == needle {
				found = true
				break
			}
		}
		if !found {
			DumpSheetRangeAsMarkdown(s.T(), result, sheet, 1, 40, 1, 3)
			s.Require().Failf("not found", "expected to find %q in output. Sample: %v", needle, all)
		}
	}
}

// DumpSheetRangeAsMarkdown печатает диапазон листа в формате Markdown-таблицы.
// Первая строка диапазона трактуется как заголовок.
// Индексы строк/колонок 1-based (как в Excel).
func DumpSheetRangeAsMarkdown(t *testing.T, f *excelize.File, sheet string, startRow, endRow, startCol, endCol int) {
	if t == nil || f == nil {
		return
	}
	if startRow <= 0 || endRow < startRow || startCol <= 0 || endCol < startCol {
		t.Logf("DumpSheetRangeAsMarkdown: invalid range: rows %d-%d, cols %d-%d", startRow, endRow, startCol, endCol)
		return
	}

	// Читаем значения
	data := make([][]string, 0, endRow-startRow+1)
	for r := startRow; r <= endRow; r++ {
		rowVals := make([]string, 0, endCol-startCol+1)
		for c := startCol; c <= endCol; c++ {
			addr, _ := excelize.CoordinatesToCellName(c, r)
			val, _ := f.GetCellValue(sheet, addr)
			// Уберём переводы строк для компактного вывода
			val = strings.ReplaceAll(val, "\n", " ")
			val = strings.TrimSpace(val)
			rowVals = append(rowVals, val)
		}
		data = append(data, rowVals)
	}

	if len(data) == 0 {
		t.Log("(empty range)")
		return
	}

	header := data[0]
	body := data[1:]

	// Формируем Markdown
	var b strings.Builder
	// Шапка
	b.WriteString("|")
	for i := range header {
		b.WriteString(" ")
		if header[i] == "" {
			b.WriteString(" ")
		} else {
			b.WriteString(header[i])
		}
		b.WriteString(" |")
	}
	b.WriteString("\n|")
	for range header {
		b.WriteString("---|")
	}
	b.WriteString("\n")
	// Тело
	for _, row := range body {
		b.WriteString("|")
		for i := range row {
			cell := row[i]
			if cell == "" {
				cell = " "
			}
			b.WriteString(" ")
			b.WriteString(cell)
			b.WriteString(" |")
		}
		b.WriteString("\n")
	}

	t.Log("\n" + b.String())
}

// TestEachObjIteration — проверяет {{#each-obj}} по объекту (map) с сортировкой ключей
func (s *TemplateSuite) TestEachObjIteration() {
	tmpDir := s.T().TempDir()
	tmpTemplate := filepath.Join(tmpDir, "each_obj_template.xlsx")

	f := excelize.NewFile()
	sheet := "Sheet1"

	_ = f.SetCellValue(sheet, "A1", "Key")
	_ = f.SetCellValue(sheet, "B1", "Value")
	_ = f.SetCellValue(sheet, "A2", "{{#each-obj $.meta as $k $v}}")
	_ = f.SetCellValue(sheet, "A3", "{{= $k}}")
	_ = f.SetCellValue(sheet, "B3", "{{= $v}}")
	_ = f.SetCellValue(sheet, "A4", "{{/each-obj}}")

	s.Require().NoError(f.SaveAs(tmpTemplate), "save template")

	// Нарочно перемешанные ключи — движок должен отсортировать их
	json := `{"meta": {"version": "1.0", "owner": "dept"}}`
	tmpOutput := filepath.Join(tmpDir, "each_obj_output.xlsx")
	s.Require().NoError(exceltemplar.WriteResultsWithTemplate(tmpTemplate, tmpOutput, []string{json}), "render")

	res, err := excelize.OpenFile(tmpOutput)
	s.Require().NoError(err, "open result")

	// Ожидаем порядок по ключам: owner, version
	if v, _ := res.GetCellValue(sheet, "A2"); true {
		s.Assert().Equal("owner", strings.TrimSpace(v), "A2 key")
	}
	if v, _ := res.GetCellValue(sheet, "B2"); true {
		s.Assert().Equal("dept", strings.TrimSpace(v), "B2 val")
	}
	if v, _ := res.GetCellValue(sheet, "A3"); true {
		s.Assert().Equal("version", strings.TrimSpace(v), "A3 key")
	}
	if v, _ := res.GetCellValue(sheet, "B3"); true {
		s.Assert().Equal("1.0", strings.TrimSpace(v), "B3 val")
	}
}

// TestJoinWithField — проверяет join(array, sep, fieldPath)
func (s *TemplateSuite) TestJoinWithField() {
	tmpDir := s.T().TempDir()
	tmpTemplate := filepath.Join(tmpDir, "join_field_template.xlsx")

	f := excelize.NewFile()
	sheet := "Sheet1"

	_ = f.SetCellValue(sheet, "A1", "Users")
	_ = f.SetCellValue(sheet, "A2", "{{#each $.users as $u}}")
	_ = f.SetCellValue(sheet, "A3", "{{= $u.name}}: {{= join($u.phones, ', ', 'n')}}")
	_ = f.SetCellValue(sheet, "A4", "{{/each}}")

	s.Require().NoError(f.SaveAs(tmpTemplate), "save template")

	json := `{
        "users": [
            {"name": "A", "phones": [{"n": "111"}, {"n": "222"}]},
            {"name": "B", "phones": [{"n": "333"}]}
        ]
    }`
	tmpOutput := filepath.Join(tmpDir, "join_field_output.xlsx")
	s.Require().NoError(exceltemplar.WriteResultsWithTemplate(tmpTemplate, tmpOutput, []string{json}), "render")

	res, err := excelize.OpenFile(tmpOutput)
	s.Require().NoError(err, "open result")

	if v, _ := res.GetCellValue(sheet, "A2"); true {
		s.Assert().Equal("A: 111, 222", strings.TrimSpace(v), "A2")
	}
	if v, _ := res.GetCellValue(sheet, "A3"); true {
		s.Assert().Equal("B: 333", strings.TrimSpace(v), "A3")
	}
}

// TestDynamicIndexInPath — проверяет использование динамического индекса [$ri]
func (s *TemplateSuite) TestDynamicIndexInPath() {
	tmpDir := s.T().TempDir()
	tmpTemplate := filepath.Join(tmpDir, "dynamic_index_template.xlsx")

	f := excelize.NewFile()
	sheet := "Sheet1"

	_ = f.SetCellValue(sheet, "A1", "{{#each $.groups as $g i=$gi}}")
	_ = f.SetCellValue(sheet, "A2", "Group {{= $gi+1}}: {{= $g.title}}")
	_ = f.SetCellValue(sheet, "A3", "{{#each $g.rows as $r i=$ri}}")
	_ = f.SetCellValue(sheet, "A4", "Row {{= $ri+1}}: {{= $g.rows[$ri]}}")
	_ = f.SetCellValue(sheet, "A5", "{{/each}}")
	_ = f.SetCellValue(sheet, "A6", "{{/each}}")

	s.Require().NoError(f.SaveAs(tmpTemplate), "save template")

	json := `{
        "groups": [
            {"title": "G1", "rows": ["A", "B"]},
            {"title": "G2", "rows": ["X"]}
        ]
    }`
	tmpOutput := filepath.Join(tmpDir, "dynamic_index_output.xlsx")
	s.Require().NoError(exceltemplar.WriteResultsWithTemplate(tmpTemplate, tmpOutput, []string{json}), "render")

	res, err := excelize.OpenFile(tmpOutput)
	s.Require().NoError(err, "open result")

	// Соберём все непустые значения колонки A
	rows, _ := res.GetRows(sheet)
	var vals []string
	for _, r := range rows {
		if len(r) > 0 {
			v := strings.TrimSpace(r[0])
			if v != "" {
				vals = append(vals, v)
			}
		}
	}
	// Проверяем присутствие строк с динамическим индексом
	must := []string{"Group 1: G1", "Row 1: A", "Row 2: B", "Group 2: G2", "Row 1: X"}
	for _, m := range must {
		found := false
		for _, v := range vals {
			if v == m {
				found = true
				break
			}
		}
		if !found {
			s.Assert().Failf("missing", "expected to find %q in output: %v", m, vals)
		}
	}
}

// TestRootAnchorUsage — проверяет использование якоря $root во вложенных блоках
func (s *TemplateSuite) TestRootAnchorUsage() {
	tmpDir := s.T().TempDir()
	tmpTemplate := filepath.Join(tmpDir, "root_anchor_template.xlsx")

	f := excelize.NewFile()
	sheet := "Sheet1"

	_ = f.SetCellValue(sheet, "A1", "{{#each $.users as $u}}")
	_ = f.SetCellValue(sheet, "A2", "{{= $u.name}} ({{= $root.report.date}})")
	_ = f.SetCellValue(sheet, "A3", "{{/each}}")

	s.Require().NoError(f.SaveAs(tmpTemplate), "save template")

	json := `{"users":[{"name":"A"},{"name":"B"}], "report": {"date": "2025-08-06"}}`
	tmpOutput := filepath.Join(tmpDir, "root_anchor_output.xlsx")
	s.Require().NoError(exceltemplar.WriteResultsWithTemplate(tmpTemplate, tmpOutput, []string{json}), "render")

	res, err := excelize.OpenFile(tmpOutput)
	s.Require().NoError(err, "open result")

	if v, _ := res.GetCellValue(sheet, "A1"); true {
		s.Assert().Equal("A (2025-08-06)", strings.TrimSpace(v), "A1")
	}
	if v, _ := res.GetCellValue(sheet, "A2"); true {
		s.Assert().Equal("B (2025-08-06)", strings.TrimSpace(v), "A2")
	}
}

// TestMultipleRootsResolution — проверяет, что значение берется из первого корня, где оно найдено
func (s *TemplateSuite) TestMultipleRootsResolution() {
	tmpDir := s.T().TempDir()
	tmpTemplate := filepath.Join(tmpDir, "multi_roots_template.xlsx")

	f := excelize.NewFile()
	sheet := "Sheet1"

	_ = f.SetCellValue(sheet, "A1", "{{= $.value}}")
	s.Require().NoError(f.SaveAs(tmpTemplate), "save template")

	json1 := `{"x": 1}`
	json2 := `{"value": "OK"}`
	tmpOutput := filepath.Join(tmpDir, "multi_roots_output.xlsx")
	s.Require().NoError(exceltemplar.WriteResultsWithTemplate(tmpTemplate, tmpOutput, []string{json1, json2}), "render")

	res, err := excelize.OpenFile(tmpOutput)
	s.Require().NoError(err, "open result")
	if v, _ := res.GetCellValue(sheet, "A1"); true {
		s.Assert().Equal("OK", strings.TrimSpace(v))
	}
}

// TestHorizontalMergesDuplicated — проверяет дублирование горизонтальных merge на каждой сгенерированной строке
func (s *TemplateSuite) TestHorizontalMergesDuplicated() {
	tmpDir := s.T().TempDir()
	tmpTemplate := filepath.Join(tmpDir, "merges_template.xlsx")

	f := excelize.NewFile()
	sheet := "Sheet1"

	_ = f.SetCellValue(sheet, "A1", "Header")
	_ = f.SetCellValue(sheet, "A2", "{{#each $.items as $it}}")
	_ = f.SetCellValue(sheet, "A3", "{{= $it.name}}")
	_ = f.SetCellValue(sheet, "B3", " ")
	_ = f.SetCellValue(sheet, "A4", "{{/each}}")
	// Горизонтальное объединение в строке-шаблоне A3:B3
	_ = f.MergeCell(sheet, "A3", "B3")

	s.Require().NoError(f.SaveAs(tmpTemplate), "save template")

	json := `{"items": [{"name":"X"}, {"name":"Y"}]}`
	tmpOutput := filepath.Join(tmpDir, "merges_output.xlsx")
	s.Require().NoError(exceltemplar.WriteResultsWithTemplate(tmpTemplate, tmpOutput, []string{json}), "render")

	res, err := excelize.OpenFile(tmpOutput)
	s.Require().NoError(err, "open result")

	merges, _ := res.GetMergeCells(sheet)
	// Ищем объединения A2:B2 и A3:B3 (две сгенерированные строки)
	want := map[int]bool{2: false, 3: false}
	for _, m := range merges {
		sAxis := m.GetStartAxis()
		eAxis := m.GetEndAxis()
		sc, sr, _ := excelize.SplitCellName(sAxis)
		ec, er, _ := excelize.SplitCellName(eAxis)
		if sc == "A" && ec == "B" && sr == er {
			if _, ok := want[sr]; ok {
				want[sr] = true
			}
		}
	}
	for r, ok := range want {
		if !ok {
			s.Assert().Failf("merge", "expected horizontal merge A:B at row %d", r)
		}
	}
}

// TestScalarInsertCollectionError — проверяет ошибку при попытке вставить коллекцию через {{= expr}}
func (s *TemplateSuite) TestScalarInsertCollectionError() {
	tmpDir := s.T().TempDir()
	tmpTemplate := filepath.Join(tmpDir, "scalar_collection_err.xlsx")

	f := excelize.NewFile()
	sheet := "Sheet1"
	_ = f.SetCellValue(sheet, "A1", "{{= $.arr}}")
	s.Require().NoError(f.SaveAs(tmpTemplate), "save template")

	json := `{"arr": [1,2,3]}`
	tmpOutput := filepath.Join(tmpDir, "scalar_collection_err_output.xlsx")
	err := exceltemplar.WriteResultsWithTemplate(tmpTemplate, tmpOutput, []string{json})
	s.Require().Error(err, "expected error when inserting collection into scalar placeholder")
}

// TestAbsentArrayEach — проверяет поведение each при отсутствии массива (данных нет)
func (s *TemplateSuite) TestAbsentArrayEach() {
	tmpDir := s.T().TempDir()
	tmpTemplate := filepath.Join(tmpDir, "absent_array_template.xlsx")

	f := excelize.NewFile()
	sheet := "Sheet1"
	_ = f.SetCellValue(sheet, "A1", "Header")
	_ = f.SetCellValue(sheet, "A2", "{{#each $.items as $it}}")
	_ = f.SetCellValue(sheet, "A3", "{{= $it.name}}")
	_ = f.SetCellValue(sheet, "A4", "{{/each}}")
	s.Require().NoError(f.SaveAs(tmpTemplate), "save template")

	// Массив отсутствует целиком
	json := `{}`
	tmpOutput := filepath.Join(tmpDir, "absent_array_output.xlsx")
	s.Require().NoError(exceltemplar.WriteResultsWithTemplate(tmpTemplate, tmpOutput, []string{json}), "render")

	res, err := excelize.OpenFile(tmpOutput)
	s.Require().NoError(err, "open result")
	rows, _ := res.GetRows(sheet)
	// Не должно остаться плейсхолдеров и управляющих строк
	if len(rows) > 0 {
		for r := 1; r <= 20 && r <= len(rows); r++ {
			for _, v := range rows[r-1] {
				if strings.Contains(v, "{{=") || strings.Contains(v, "{{#") || strings.Contains(v, "{{/") || strings.Contains(v, "{{else}}") {
					s.Require().Failf("placeholder", "placeholder left after absent array render at row %d: %q", r, v)
				}
			}
		}
	}
}

// TestMultiSheetDataOnOne — проверяет, что данные на одном листе не влияют на другой
func (s *TemplateSuite) TestMultiSheetDataOnOne() {
	tmpDir := s.T().TempDir()
	tmpTemplate := filepath.Join(tmpDir, "multi_sheet_template.xlsx")

	f := excelize.NewFile()
	sheet1 := "Sheet1"
	sheet2 := "Sheet2"
	_ = f.SetCellValue(sheet1, "A1", "T1")
	_ = f.SetCellValue(sheet1, "B1", "{{= $.a }}")
	_, _ = f.NewSheet(sheet2)
	_ = f.SetCellValue(sheet2, "A1", "T2")
	_ = f.SetCellValue(sheet2, "B1", "{{= $.b }}")

	s.Require().NoError(f.SaveAs(tmpTemplate), "save template")

	json := `{"a":"X"}`
	tmpOutput := filepath.Join(tmpDir, "multi_sheet_output.xlsx")
	s.Require().NoError(exceltemplar.WriteResultsWithTemplate(tmpTemplate, tmpOutput, []string{json}), "render")

	res, err := excelize.OpenFile(tmpOutput)
	s.Require().NoError(err, "open result")

	if v, _ := res.GetCellValue(sheet1, "B1"); true {
		s.Assert().Equal("X", strings.TrimSpace(v), "Sheet1 B1")
	}
	if v, _ := res.GetCellValue(sheet2, "B1"); true {
		s.Assert().Equal("", strings.TrimSpace(v), "Sheet2 B1 should be empty")
	}
}

// TestTwoTablesStacked — проверяет, что две таблицы одна под другой не затирают данные друг друга
func (s *TemplateSuite) TestTwoTablesStacked() {
	tmpDir := s.T().TempDir()
	tmpTemplate := filepath.Join(tmpDir, "two_tables_template.xlsx")

	f := excelize.NewFile()
	sheet := "Sheet1"

	// Заголовок первой таблицы (без маркеров)
	_ = f.SetCellValue(sheet, "A1", "Таблица A")
	// Открывающий итератор первой таблицы
	_ = f.SetCellValue(sheet, "A2", "{{#each $.tableA as $row i=$i}}")
	// Строка шаблона для первой таблицы (колонки A, B одинаковые для обеих таблиц)
	_ = f.SetCellValue(sheet, "A3", "{{= $i+1}}")
	_ = f.SetCellValue(sheet, "B3", "{{= $row.name}}")
	// Закрывающий итератор первой таблицы
	_ = f.SetCellValue(sheet, "A4", "{{/each}}")

	// Немного отступа между таблицами, затем заголовок второй
	_ = f.SetCellValue(sheet, "A6", "Таблица B")
	// Открывающий итератор второй таблицы
	_ = f.SetCellValue(sheet, "A7", "{{#each $.tableB as $row i=$i}}")
	// Строка шаблона для второй таблицы (те же столбцы A, B)
	_ = f.SetCellValue(sheet, "A8", "{{= $i+1}}")
	_ = f.SetCellValue(sheet, "B8", "{{= $row.name}}")
	// Закрывающий итератор второй таблицы
	_ = f.SetCellValue(sheet, "A9", "{{/each}}")

	s.Require().NoError(f.SaveAs(tmpTemplate), "save template")

	// Готовим JSON: по 10 записей в каждой таблице (A-1..A-10 и B-1..B-10)
	jsonPayload := fmt.Sprintf(`{
        "tableA": [%s],
        "tableB": [%s]
    }`,
		strings.Join(func() []string {
			out := make([]string, 10)
			for i := 1; i <= 10; i++ {
				out[i-1] = fmt.Sprintf(`{"id": %d, "name": "A-%d"}`, i, i)
			}
			return out
		}(), ","),
		strings.Join(func() []string {
			out := make([]string, 10)
			for i := 1; i <= 10; i++ {
				out[i-1] = fmt.Sprintf(`{"id": %d, "name": "B-%d"}`, i, i)
			}
			return out
		}(), ","),
	)

	// Рендерим
	tmpOutput := filepath.Join(tmpDir, "two_tables_output.xlsx")
	s.Require().NoError(exceltemplar.WriteResultsWithTemplate(tmpTemplate, tmpOutput, []string{jsonPayload}), "render")

	// Открываем результат
	result, err := excelize.OpenFile(tmpOutput)
	s.Require().NoError(err, "open result")

	// Сканируем колонку B (имена) сверху вниз и собираем позиции строк
	posA := make(map[string]int) // name->row
	posB := make(map[string]int)
	for r := 1; r <= 300; r++ {
		v, _ := result.GetCellValue(sheet, fmt.Sprintf("B%d", r))
		v = strings.TrimSpace(v)
		if strings.HasPrefix(v, "A-") {
			posA[v] = r
		}
		if strings.HasPrefix(v, "B-") {
			posB[v] = r
		}
	}

	// Проверяем, что каждая таблица имеет ровно 10 записей
	s.Require().Equal(10, len(posA), "rows in table A")
	s.Require().Equal(10, len(posB), "rows in table B")

	// Проверяем, что все элементы присутствуют и не затёрты
	for i := 1; i <= 10; i++ {
		a := fmt.Sprintf("A-%d", i)
		b := fmt.Sprintf("B-%d", i)
		if _, ok := posA[a]; !ok {
			s.Require().Failf("missing", "не найдена запись %q из первой таблицы", a)
		}
		if _, ok := posB[b]; !ok {
			s.Require().Failf("missing", "не найдена запись %q из второй таблицы", b)
		}
	}

	// Ключевая проверка: последняя строка первой таблицы ДОЛЖНА быть выше первой строки второй таблицы
	maxRowA := 0
	minRowB := 1<<31 - 1
	for i := 1; i <= 10; i++ {
		if r := posA[fmt.Sprintf("A-%d", i)]; r > maxRowA {
			maxRowA = r
		}
		if r := posB[fmt.Sprintf("B-%d", i)]; r < minRowB {
			minRowB = r
		}
	}
	if !(maxRowA < minRowB) {
		DumpSheetRangeAsMarkdown(s.T(), result, sheet, maxRowA-3, minRowB+3, 1, 2)
		s.Require().Failf("order", "ожидалось, что записи второй таблицы будут ниже первой (maxA=%d, minB=%d)", maxRowA, minRowB)
	}

	// Дополнительная проверка: шапка таблицы B не затёрта и присутствует ВЫШЕ её данных
	headerRowB := -1
	for r := 1; r <= 300; r++ {
		a, _ := result.GetCellValue(sheet, fmt.Sprintf("A%d", r))
		if strings.TrimSpace(a) == "Таблица B" {
			headerRowB = r
			break
		}
	}
	if headerRowB == -1 {
		DumpSheetRangeAsMarkdown(s.T(), result, sheet, maxRowA-5, maxRowA+15, 1, 2)
		s.Require().Fail("header missing", "шапка второй таблицы не найдена — вероятно, затёрта рендером")
	}
	if !(headerRowB < minRowB) {
		DumpSheetRangeAsMarkdown(s.T(), result, sheet, headerRowB-3, minRowB+3, 1, 2)
		s.Require().Failf("header order", "шапка второй таблицы должна быть выше её данных (headerB=%d, firstDataB=%d)", headerRowB, minRowB)
	}
}

// TestEmptyBlock — проверяет удаление пустого each-блока и отсутствие плейсхолдеров в результате
func (s *TemplateSuite) TestEmptyBlock() {
	tmpDir := s.T().TempDir()
	tmpTemplate := filepath.Join(tmpDir, "empty_block_template.xlsx")

	f := excelize.NewFile()
	sheet := "Sheet1"
	// Управляющие маркеры и строка-шаблон между ними
	_ = f.SetCellValue(sheet, "A1", "Header")
	_ = f.SetCellValue(sheet, "A2", "{{#each $.items as $it}}")
	_ = f.SetCellValue(sheet, "A3", "{{= $it.name}}")
	_ = f.SetCellValue(sheet, "A4", "{{/each}}")
	s.Require().NoError(f.SaveAs(tmpTemplate), "save template")

	// Данные: пустой массив, блок должен быть удалён
	json := `{"items": []}`
	tmpOutput := filepath.Join(tmpDir, "empty_block_output.xlsx")
	s.Require().NoError(exceltemplar.WriteResultsWithTemplate(tmpTemplate, tmpOutput, []string{json}), "render")

	// Проверяем, что строки 2..4 удалены, а плейсхолдеров нет
	result, err := excelize.OpenFile(tmpOutput)
	s.Require().NoError(err, "open result")
	rows, _ := result.GetRows(sheet)
	if len(rows) > 0 {
		for r := 1; r <= 20 && r <= len(rows); r++ {
			for _, v := range rows[r-1] {
				if strings.Contains(v, "{{=") || strings.Contains(v, "{{#") || strings.Contains(v, "{{/") || strings.Contains(v, "{{else}}") {
					s.Require().Failf("placeholder", "placeholder left after empty render at row %d: %q", r, v)
				}
			}
		}
	}
}

// TestCustomTechOpsTable — проверяет сложную таблицу технологических операций
func (s *TemplateSuite) TestCustomTechOpsTable() {
	tmpDir := s.T().TempDir()
	tmpTemplate := filepath.Join(tmpDir, "custom_tech_ops_template.xlsx")

	f := excelize.NewFile()
	sheet := "Sheet1"

	// Заголовок таблицы
	_ = f.SetCellValue(sheet, "B35", "Технологическая операция")

	// Шапка колонок
	_ = f.SetCellValue(sheet, "B36", "Код операции")
	_ = f.SetCellValue(sheet, "C36", "Код следующей операции")
	_ = f.SetCellValue(sheet, "D36", "Наименование")
	_ = f.SetCellValue(sheet, "E36", "Обоснование")
	_ = f.SetCellValue(sheet, "F36", "Объем")
	_ = f.SetCellValue(sheet, "G36", "Единица")

	// Номера колонок
	_ = f.SetCellValue(sheet, "B37", "1")
	_ = f.SetCellValue(sheet, "C37", "2")
	_ = f.SetCellValue(sheet, "D37", "3")
	_ = f.SetCellValue(sheet, "E37", "4")
	_ = f.SetCellValue(sheet, "F37", "5")
	_ = f.SetCellValue(sheet, "G37", "6")

	// Обязательность полей
	_ = f.SetCellValue(sheet, "B38", "-")
	_ = f.SetCellValue(sheet, "C38", "-")
	_ = f.SetCellValue(sheet, "D38", "+")
	_ = f.SetCellValue(sheet, "E38", "-")
	_ = f.SetCellValue(sheet, "F38", "+/-")
	_ = f.SetCellValue(sheet, "G38", "+/-")

	// Подготовительные операции
	_ = f.SetCellValue(sheet, "B39", "Подготовительные операции")
	_ = f.SetCellValue(sheet, "B40", "{{#each $.tech_operation.preparation as $op}}")
	_ = f.SetCellValue(sheet, "B41", "{{= $op.operation_code}}")
	_ = f.SetCellValue(sheet, "C41", "{{= $op.next_operation_code}}")
	_ = f.SetCellValue(sheet, "D41", "{{= $op.name}}")
	_ = f.SetCellValue(sheet, "E41", "{{= $op.justification}}")
	_ = f.SetCellValue(sheet, "F41", "{{= $op.volume}}")
	_ = f.SetCellValue(sheet, "G41", "{{= $op.unit}}")
	_ = f.SetCellValue(sheet, "B42", "{{/each}}")

	// Основные операции (группы)
	_ = f.SetCellValue(sheet, "B43", "Основные операции")
	_ = f.SetCellValue(sheet, "B44", "{{#each $.tech_operation.main as $grp}}")
	_ = f.SetCellValue(sheet, "B45", "{{= $grp.group_title}}")
	_ = f.SetCellValue(sheet, "B46", "{{#each $grp.operations as $op}}")
	_ = f.SetCellValue(sheet, "B47", "{{= $op.operation_code}}")
	_ = f.SetCellValue(sheet, "C47", "{{= $op.next_operation_code}}")
	_ = f.SetCellValue(sheet, "D47", "{{= $op.name}}")
	_ = f.SetCellValue(sheet, "E47", "{{= $op.justification}}")
	_ = f.SetCellValue(sheet, "F47", "{{= $op.volume}}")
	_ = f.SetCellValue(sheet, "G47", "{{= $op.unit}}")
	_ = f.SetCellValue(sheet, "B48", "{{/each}}")
	_ = f.SetCellValue(sheet, "B49", "{{/each}}")

	// Заключительные операции
	_ = f.SetCellValue(sheet, "B50", "Заключительные операции")
	_ = f.SetCellValue(sheet, "B51", "{{#each $.tech_operation.final as $op}}")
	_ = f.SetCellValue(sheet, "B52", "{{= $op.operation_code}}")
	_ = f.SetCellValue(sheet, "C52", "{{= $op.next_operation_code}}")
	_ = f.SetCellValue(sheet, "D52", "{{= $op.name}}")
	_ = f.SetCellValue(sheet, "E52", "{{= $op.justification}}")
	_ = f.SetCellValue(sheet, "F52", "{{= $op.volume}}")
	_ = f.SetCellValue(sheet, "G52", "{{= $op.unit}}")
	_ = f.SetCellValue(sheet, "B53", "{{/each}}")

	s.Require().NoError(f.SaveAs(tmpTemplate), "save template")

	// Данные для заполнения (выдуманные, согласованы с документацией и плейсхолдерами)
	json := `{
        "tech_operation": {
            "preparation": [
                {
                    "operation_code": "00-01-0-0-1",
                    "next_operation_code": "00-01-0-0-2",
                    "name": "Получение наряд-допуска",
                    "justification": "БЦЗ-010601-0101",
                    "volume": 1,
                    "unit": "1 задвижка"
                },
                {
                    "operation_code": "00-02-0-0-1",
                    "next_operation_code": "00-02-0-0-2",
                    "name": "Подготовка рабочего места",
                    "justification": "БЦЗ-010601-0102",
                    "volume": 1,
                    "unit": "1 задвижка"
                }
            ],
            "main": [
                {
                    "group_title": "Демонтаж задвижки",
                    "operations": [
                        {
                            "operation_code": "09-02-0-0",
                            "next_operation_code": "00-03-0-0",
                            "name": "Снятие привода",
                            "justification": "БЦЗ-010601-0103",
                            "volume": 1,
                            "unit": "1 задвижка"
                        },
                        {
                            "operation_code": "09-03-0-0",
                            "next_operation_code": "00-04-0-0",
                            "name": "Отсоединение трубопроводов",
                            "justification": "БЦЗ-010601-0104",
                            "volume": 2,
                            "unit": "соединение"
                        }
                    ]
                },
                {
                    "group_title": "Полная разборка задвижки",
                    "operations": [
                        {
                            "operation_code": "00-04-0-0-1",
                            "next_operation_code": "00-04-0-0-2",
                            "name": "Отвинтить винты (гайки), крепящие привод поз.10",
                            "justification": "БЦЗ-010601-0101",
                            "volume": 1,
                            "unit": "1 задвижка"
                        }
                    ]
                }
            ],
            "final": [
                {
                    "operation_code": "00-20-0-0-1",
                    "next_operation_code": null,
                    "name": "Приемка выполненных работ",
                    "justification": "БЦЗ-010601-0199",
                    "volume": 1,
                    "unit": "1 задвижка"
                }
            ]
        }
    }`

	// Рендерим
	tmpOutput := filepath.Join(tmpDir, "custom_tech_ops_output.xlsx")
	s.Require().NoError(exceltemplar.WriteResultsWithTemplate(tmpTemplate, tmpOutput, []string{json}), "render")

	// Проверяем результат
	result, err := excelize.OpenFile(tmpOutput)
	s.Require().NoError(err, "open result")

	// Ищем якорь по заголовку таблицы
	anchor := -1
	for rr := 25; rr <= 80; rr++ {
		bAddr, _ := excelize.CoordinatesToCellName(2, rr)
		v, _ := result.GetCellValue(sheet, bAddr)
		if strings.TrimSpace(v) == "Технологическая операция" {
			anchor = rr
			break
		}
	}
	s.Require().NotEqual(-1, anchor, "anchor not found in B25..B80")

	// Проверяем заголовки разделов
	mustHaveTitles := []string{"Подготовительные операции", "Основные операции", "Заключительные операции"}
	seen := map[string]bool{}
	for r := anchor - 10; r <= anchor+220; r++ {
		b, _ := result.GetCellValue(sheet, fmt.Sprintf("B%d", r))
		s := strings.TrimSpace(b)
		if s != "" {
			seen[s] = true
		}
	}
	for _, title := range mustHaveTitles {
		if !seen[title] {
			s.Assert().Failf("title", "section title %q not found", title)
		}
	}

	// Ожидаемые операции по колонкам B..G
	type expOp struct {
		nextCode      string
		name          string
		justification string
		volume        string
		unit          string
	}
	expected := map[string]expOp{
		"00-01-0-0-1": {nextCode: "00-01-0-0-2", name: "Получение наряд-допуска", justification: "БЦЗ-010601-0101", volume: "1", unit: "1 задвижка"},
		"00-02-0-0-1": {nextCode: "00-02-0-0-2", name: "Подготовка рабочего места", justification: "БЦЗ-010601-0102", volume: "1", unit: "1 задвижка"},
		"09-02-0-0":   {nextCode: "00-03-0-0", name: "Снятие привода", justification: "БЦЗ-010601-0103", volume: "1", unit: "1 задвижка"},
		"09-03-0-0":   {nextCode: "00-04-0-0", name: "Отсоединение трубопроводов", justification: "БЦЗ-010601-0104", volume: "2", unit: "соединение"},
		"00-04-0-0-1": {nextCode: "00-04-0-0-2", name: "Отвинтить винты (гайки), крепящие привод поз.10", justification: "БЦЗ-010601-0101", volume: "1", unit: "1 задвижка"},
		"00-20-0-0-1": {nextCode: "", name: "Приемка выполненных работ", justification: "БЦЗ-010601-0199", volume: "1", unit: "1 задвижка"},
	}

	validated := map[string]bool{}
	for r := anchor; r <= anchor+240; r++ {
		b, _ := result.GetCellValue(sheet, fmt.Sprintf("B%d", r))
		code := strings.TrimSpace(b)
		exp, ok := expected[code]
		if !ok {
			continue
		}
		c, _ := result.GetCellValue(sheet, fmt.Sprintf("C%d", r))
		d, _ := result.GetCellValue(sheet, fmt.Sprintf("D%d", r))
		e, _ := result.GetCellValue(sheet, fmt.Sprintf("E%d", r))
		ff, _ := result.GetCellValue(sheet, fmt.Sprintf("F%d", r))
		g, _ := result.GetCellValue(sheet, fmt.Sprintf("G%d", r))
		s.Assert().Equal(exp.nextCode, strings.TrimSpace(c), "row %d next_operation_code", r)
		s.Assert().Equal(exp.name, strings.TrimSpace(d), "row %d name", r)
		s.Assert().Equal(exp.justification, strings.TrimSpace(e), "row %d justification", r)
		s.Assert().Equal(exp.volume, strings.TrimSpace(ff), "row %d volume", r)
		s.Assert().Equal(exp.unit, strings.TrimSpace(g), "row %d unit", r)
		validated[code] = true
	}
	for code := range expected {
		if !validated[code] {
			s.Assert().Failf("missing", "operation_code %q row not found for full validation", code)
		}
	}

	// Названия групп присутствуют
	for _, title := range []string{"Демонтаж задвижки", "Полная разборка задвижки"} {
		found := false
		for r := anchor; r <= anchor+220; r++ {
			b, _ := result.GetCellValue(sheet, fmt.Sprintf("B%d", r))
			if strings.TrimSpace(b) == title {
				found = true
				break
			}
		}
		if !found {
			s.Assert().Failf("group", "group title %q not found", title)
		}
	}

	// Контроль: управляющие маркеры удалены
	for r := 38; r <= 60; r++ {
		for c := 2; c <= 7; c++ {
			addr, _ := excelize.CoordinatesToCellName(c, r)
			v, _ := result.GetCellValue(sheet, addr)
			if strings.Contains(v, "{{") || strings.Contains(v, "}}") {
				s.Assert().Failf("marker", "control marker left in %s: %q", addr, v)
			}
		}
	}

}

// TestSafetyRequirementsPlaceholders_Table — проверяет inline iif/exists/join в B6:E6
//
// B6: {{= iif(exists($.safety_requirements[0].occupational), join($.safety_requirements[0].occupational, "\n\n"), "") }}
// C6: {{= iif(exists($.safety_requirements[0].industrial),   join($.safety_requirements[0].industrial,   "\n\n"), "") }}
// D6: {{= iif(exists($.safety_requirements[0].fire),         join($.safety_requirements[0].fire,         "\n\n"), "") }}
// E6: {{= iif(exists($.safety_requirements[0].environmental), join($.safety_requirements[0].environmental, "\n\n"), "") }}
func (s *TemplateSuite) TestSafetyRequirementsPlaceholders_Table() {
	tmpDir := s.T().TempDir()
	tmpTemplate := filepath.Join(tmpDir, "safety_placeholders_template.xlsx")

	f := excelize.NewFile()
	sheet := "Sheet1"

	// Заполним A1..A5 статикой, чтобы индексы строк совпадали с реальными (особенность GetRows).
	_ = f.SetCellValue(sheet, "A1", "pad1")
	_ = f.SetCellValue(sheet, "A2", "pad2")
	_ = f.SetCellValue(sheet, "A3", "pad3")
	_ = f.SetCellValue(sheet, "A4", "pad4")
	_ = f.SetCellValue(sheet, "A5", "pad5")
	// Плейсхолдеры в B6..E6 (строка 6). Добавляем A6 как статический якорь.
	_ = f.SetCellValue(sheet, "A6", "Якорь")
	_ = f.SetCellValue(sheet, "B6", "{{= iif(exists($.safety_requirements[0].occupational), join($.safety_requirements[0].occupational, '\n\n'), '') }}")
	_ = f.SetCellValue(sheet, "C6", "{{= iif(exists($.safety_requirements[0].industrial),   join($.safety_requirements[0].industrial,   '\n\n'), '') }}")
	_ = f.SetCellValue(sheet, "D6", "{{= iif(exists($.safety_requirements[0].fire),         join($.safety_requirements[0].fire,         '\n\n'), '') }}")
	_ = f.SetCellValue(sheet, "E6", "{{= iif(exists($.safety_requirements[0].environmental), join($.safety_requirements[0].environmental, '\n\n'), '') }}")

	s.Require().NoError(f.SaveAs(tmpTemplate), "save template")

	type exp struct{ b, c, d, e string }
	cases := []struct {
		name string
		json string
		want exp
	}{
		{
			name: "all_present_two_each",
			json: `{
                "safety_requirements": [
                    {
                        "occupational": ["Охрана труда 1", "Охрана труда 2"],
                        "industrial":   ["Промышленная 1", "Промышленная 2"],
                        "fire":         ["Пожарная 1"],
                        "environmental": ["Эко 1", "Эко 2", "Эко 3"]
                    }
                ]
            }`,
			want: exp{
				b: "Охрана труда 1\n\nОхрана труда 2",
				c: "Промышленная 1\n\nПромышленная 2",
				d: "Пожарная 1",
				e: "Эко 1\n\nЭко 2\n\nЭко 3",
			},
		},
		{
			name: "partial_only_industrial",
			json: `{
                "safety_requirements": [
                    { "industrial": ["IND-1"] }
                ]
            }`,
			want: exp{b: "", c: "IND-1", d: "", e: ""},
		},
		{
			name: "present_but_empty_arrays",
			json: `{
                "safety_requirements": [
                    { "occupational": [], "industrial": [], "fire": [], "environmental": [] }
                ]
            }`,
			want: exp{b: "", c: "", d: "", e: ""},
		},
		{
			name: "no_field",
			json: `{}`,
			want: exp{b: "", c: "", d: "", e: ""},
		},
		{
			name: "empty_list",
			json: `{"safety_requirements": []}`,
			want: exp{b: "", c: "", d: "", e: ""},
		},
	}

	for _, tc := range cases {
		tc := tc
		s.Run(tc.name, func() {
			tmpOutput := filepath.Join(tmpDir, fmt.Sprintf("safety_%s.xlsx", strings.ReplaceAll(tc.name, "/", "_")))
			s.Require().NoError(exceltemplar.WriteResultsWithTemplate(tmpTemplate, tmpOutput, []string{tc.json}), "render")

			res, err := excelize.OpenFile(tmpOutput)
			s.Require().NoError(err, "open result")

			get := func(addr string) string {
				v, _ := res.GetCellValue(sheet, addr)
				return strings.TrimSpace(v)
			}

			s.Assert().Equal(tc.want.b, get("B6"), "B6 (occupational)")
			s.Assert().Equal(tc.want.c, get("C6"), "C6 (industrial)")
			s.Assert().Equal(tc.want.d, get("D6"), "D6 (fire)")
			s.Assert().Equal(tc.want.e, get("E6"), "E6 (environmental)")
		})
	}
}
