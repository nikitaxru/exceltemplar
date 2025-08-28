## Excel Template Engine Documentation (pkg/excel)

The new template engine supports any level of JSON nesting and explicit block syntax. It renders data in memory and then applies it to Excel sheets, correctly copying styles and horizontal merges from template rows.

### Supported Syntax

- Value insertion: `{{= expr}}`
  - Examples: `{{= .name}}`, `{{= $.company.title}}`, `{{= $i+1}}`
- Inline condition: `iif(cond, then, else)` inside `{{= ...}}`
  - Examples: `{{= iif($.flag, 'Yes', 'No') }}`, `{{= iif(len($.arr)>1, 'MANY', 'ONE') }}`
- Array loop: `{{#each path as $item i=$i}} ... {{/each}}`
  - `path`: absolute (`$.departments`) or relative (`.employees`)
  - `$item`: name of the current element variable (default `$`)
  - `i=$i`: name of the index variable (optional)
- Object (map) iteration: `{{#each-obj path as $k $v}} ... {{/each-obj}}`
- Conditions: `{{#if expr}} ... {{else}} ... {{/if}}`
  - Expression examples: `exists(.field)`, `len(.arr) > 0`, `.status == "ok"`
- Functions in expressions:
  - `len(x)` — length of array/string/object
  - `exists(x)` — check for value existence at path
  - `join(arrayPath, sep, [fieldPath])` — array concatenation, optionally by field

- Indexed access:
  - `path[index]` — index can be a number or expression/variable from block context: `[$i]`, `[$k]`, `[$var]`.
  - Examples: `{{= $.list[$i].name}}`, `{{= .rows[$ri]}}`

Rule: inside `{{= ...}}` the result must be a scalar. If the path leads to an array/object — use `{{#each}}` or `join()`.

Context:
- `.` — current element
- `$` or `$root` — JSON root
- Variables from loops: e.g., `$item`, `$i`

### Path Anchors: $, ., variables (with examples)

- What are anchors:
  - `$.path` — absolute path from root (doesn't depend on current context).
  - `.path` — relative path from current element (context is set by the nearest `each`/`each-obj`).
  - `$var.path` — path from explicitly named variable declared in `as $var` of outer loop.

- When to choose what:
  - Use `$.…` when you need to reliably access data "from root" from any depth (cross-references).
  - Use `.…` immediately after opening a block if you're accessing fields of the current element.
  - Use `$var.…` in nested blocks to unambiguously reference an element from outer level and not depend on `.` changes inside inner loops.

- Equivalence inside `each` block:
  - Inside `{{#each $.arr as $row}}` expressions `{{#each .items}}` and `{{#each $row.items}}` are equivalent until you change context with another nested `each`.

#### Example A: nested blocks with explicit variable

JSON:

```json
{
  "material_resources": {
    "sections": [
      {"title": "S1", "entries": [{"name": "A"}, {"name": "B"}]},
      {"title": "S2", "entries": []}
    ]
  }
}
```

Template (fragment):

| A |
|---|
| {{#each $.material_resources.sections as $sec i=$si}} |
| Section {{= $si+1}}: {{= $sec.title}} |
| {{#each $sec.entries as $e i=$ei}} |
| - {{= $e.name}} |
| {{/each}} |
| {{/each}} |

Alternative: replace `{{#each $sec.entries …}}` with `{{#each .entries …}}`, since inside outer `each` the current element (`.`) equals `$sec`.

#### Example B: cross-reference to root from nested block

JSON:

```json
{
  "users": [{"name":"A"}, {"name":"B"}],
  "report": {"date": "2025-08-06"}
}
```

Template (fragment):

| A |
|---|
| {{#each $.users as $u}} |
| {{= $u.name}} ({{= $.report.date}}) |
| {{/each}} |

Result: `A (2025-08-06)`, `B (2025-08-06)` — date is read via absolute path `$.report.date` regardless of current `each`.

### Programmatic API

```go
// Base path: pkg/excel

// High-level function (ready solution):
func WriteResultsWithTemplate(templatePath, destPath string, outputs []string) error

// Low-level control:
tmpl, _ := excel.LoadTemplate(templatePath) // build AST from sheets
_ = tmpl.Render(outputs)                   // render in memory (outputs — JSON strings)
_ = tmpl.Save(destPath)                    // save result
```

`outputs` — slice of JSON strings (arrays/objects). During rendering, the engine searches for needed paths in each of the passed roots; the first found one is used.

---

### Examples

#### 1) Simple substitutions

JSON:
```json
{
  "title": "Report",
  "date": "2025-08-06",
  "author": {"name": "Ivan"}
}
```

Template (Excel sheet as table):

| A                 | B                 |
|-------------------|-------------------|
| Header            | {{= $.title}}     |
| Date              | {{= $.date}}      |
| Author            | {{= $.author.name}} |

Result:

| A       | B         |
|---------|-----------|
| Header  | Report    |
| Date    | 2025-08-06|
| Author  | Ivan      |

##### Scalar substitution from nested object (without loops)

JSON:
```json
{
  "user": { "name": "nikita" }
}
```

Template:

| A                | B                       |
|------------------|-------------------------|
| User Name        | {{= $.user.name}}       |
| Greeting         | Hello, {{= $.user.name}}!|
| Age              | {{= $.user.age}}        |

Result:

| A                | B                |
|------------------|------------------|
| User Name        | nikita           |
| Greeting         | Hello, nikita!   |
| Age              |                  |

Explanation:
- `{{= $.user.name}}` takes value by absolute path from JSON root.
- If path is missing (`$.user.age`), an empty string is written to the cell.

##### Inline conditions in cells of one row

JSON:
```json
{
  "flag": true,
  "arr": [1,2]
}
```

Template:

| A                               | B                                        | C                                  |
|---------------------------------|------------------------------------------|------------------------------------|
| {{= iif($.flag, 'ON','OFF') }}  | {{= iif(len($.arr) > 1, 'MANY','ONE') }} | {{= iif($.missing, 'YES','NO') }}  |

Result:

| A   | B     | C  |
|-----|-------|----|
| ON  | MANY  | NO |

Notes:
- `iif(cond, then, else)` is evaluated inline and allows having independent conditions in different cells of one row.
- `cond` — any boolean expression supported by the engine (`len`, `exists`, comparisons, and/or/not).

---

#### 2) Array row iteration (user table)

JSON:
```json
{
  "users": [
    {"name": "A", "age": 30},
    {"name": "B", "age": 25}
  ]
}
```

Template:

| A                                  | B                |
|------------------------------------|------------------|
| Name                               | Age              |
| {{#each $.users as $u i=$i}}       |                  |
| {{= $u.name}}                      | {{= $u.age}}     |
| {{/each}}                          |                  |

Result:

| A   | B   |
|-----|-----|
| Name| Age |
| A   | 30  |
| B   | 25  |

---

#### 3) Nested groups (projects → milestones → tasks)

JSON:
```json
{
  "projects": [
    {
      "name": "P1",
      "milestones": [
        {"name": "M1", "tasks": [{"t": "T1"}, {"t": "T2"}]},
        {"name": "M2", "tasks": [{"t": "T3"}]}
      ]
    }
  ]
}
```

Template:

| A                                                   |
|-----------------------------------------------------|
| {{#each $.projects as $p}}                           |
| Project: {{= $p.name}} (milestones: {{= len($p.milestones)}}) |
| {{#if len($p.milestones) > 0}}                       |
| {{#each $p.milestones as $m}}                        |
| Milestone: {{= $m.name}} (tasks: {{= len($m.tasks)}})     |
| {{#each $m.tasks as $t i=$ti}}                       |
| Task {{= $ti+1}}: {{= $t.t}}                       |
| {{/each}}                                            |
| {{/each}}                                            |
| {{else}}                                             |
| No milestones                                        |
| {{/if}}                                              |
| {{/each}}                                            |

Result:

| A                                            |
|----------------------------------------------|
| Project: P1 (milestones: 2)                  |
| Milestone: M1 (tasks: 2)                     |
| Task 1: T1                                   |
| Task 2: T2                                   |
| Milestone: M2 (tasks: 1)                     |
| Task 1: T3                                   |

---

#### 4) join with object field

JSON:
```json
{
  "users": [
    {"name": "A", "phones": [{"n": "111"}, {"n": "222"}]},
    {"name": "B", "phones": [{"n": "333"}]}
  ]
}
```

Template:

| A                                      |
|----------------------------------------|
| {{#each $.users as $u}}                |
| {{= $u.name}}: {{= join($u.phones, ", ", "n")}} |
| {{/each}}                              |

Result:

| A           |
|-------------|
| A: 111, 222 |
| B: 333      |

---

#### 5) Object (map) iteration

JSON:
```json
{
  "meta": {"version": "1.0", "owner": "dept"}
}
```

Template:

| A                                        | B        |
|------------------------------------------|----------|
| Key                                       | Value    |
| {{#each-obj $.meta as $k $v}}            |          |
| {{= $k}}                                 | {{= $v}} |
| {{/each-obj}}                            |          |

Result:

| A        | B        |
|----------|----------|
| owner    | dept     |
| version  | 1.0      |

---

#### 6) Conditions and empty arrays

JSON:
```json
{
  "team": { "name": "Core", "members": [] }
}
```

Template:

| A                                                       |
|---------------------------------------------------------|
| Team: {{= $.team.name}}                                 |
| {{#if len($.team.members) > 0}}                         |
| Members exist                                            |
| {{else}}                                                |
| No members                                               |
| {{/if}}                                                 |

Result:

| A              |
|----------------|
| Team: Core     |
| No members     |

---

### Behavior when data is missing

- **`{{= expr}}` insertions**: if there's no value at the path, an empty string is written to the cell (cell remains empty).
- **`{{#each}}` on empty/missing array**: block body is not rendered; template row with expressions is removed; rows containing only control markers (`{{#each}}`, `{{/each}}`) are removed.
- **`{{#if}}`**: when expression evaluates to false, `{{else}}` branch is rendered (if specified); rows with markers are removed.
- **Static strings** (without markers): preserved unchanged.
- **Unified rules for all sheets**: engine goes through each sheet and applies the same logic: no data → no placeholders in result; service rows are removed.

#### Examples for scenarios without data

1) Single insertion without value

Template:

| A    | B                    |
|------|----------------------|
| Name | {{= $.user.name }}   |

Data: `{"user":{}}`

Result:

| A    | B    |
|------|------|
| Name |      |

2) Empty each block

Template:

| A                                |
|----------------------------------|
| List                             |
| {{#each $.items as $it}}         |
| {{= $it.name}}                   |
| {{/each}}                        |

Data: `{"items":[]}`

Result:

| A        |
|----------|
| List     |

3) Two tables, data only for one

Template:

| A                                | B                 |
|----------------------------------|-------------------|
| Table A                          |                   |
| {{#each $.tableA as $row i=$i}}  |                   |
| {{= $i+1}}                       | {{= $row.name}}   |
| {{/each}}                        |                   |
| Table B                          |                   |
| {{#each $.tableB as $row i=$i}}  |                   |
| {{= $i+1}}                       | {{= $row.name}}   |
| {{/each}}                        |                   |

Data: `{"tableA":[{"name":"A-1"},{"name":"A-2"}]}`

Result:

| A          | B    |
|------------|------|
| Table A    |      |
| 1          | A-1  |
| 2          | A-2  |
| Table B    |      |

4) Multiple sheets — data only on one

Sheet `Sheet1` (template):

| A    | B                  |
|------|--------------------|
| T1   | {{= $.a }}         |

Sheet `Sheet2` (template):

| A    | B                  |
|------|--------------------|
| T2   | {{= $.b }}         |

Data: `{"a":"X"}`

Sheet1 (result):

| A    | B  |
|------|----|
| T1   | X  |

Sheet2 (result):

| A    | B  |
|------|----|
| T2   |    |

### Excel Template Recommendations

- Place block markers (`{{#each}}`, `{{/each}}`, `{{#if}}`, `{{/if}}`) entirely in separate cells of a row. The entire sequence of rows between `{{#each}}` and `{{/each}}` is repeated as a group.
- Rows containing only control markers (`{{#each}}`, `{{/each}}`, `{{#if}}`, `{{/if}}`, `{{else}}`) are automatically removed from the final sheet.
- Use relative paths from current element (`.field`) inside nested blocks.
- For concatenating lists, use `join()` instead of direct array insertion.
- Excel formulas and absolute addresses are not rewritten; use relative references or named ranges when possible.
- Horizontal merges found in the template row will be duplicated on each generated row.

#### Multiple tables on one sheet (vertically)

When multiple tables are placed on one sheet one below the other (e.g., two `each` blocks on the same columns separated by a header), the engine renders them independently and safely:

- Rendering for each table is "anchored" on its template row — this is the row where there are expressions `{{= ...}}`.
- For each array element, the engine inserts a new row immediately BEFORE the template row, copies styles and horizontal merges from that row, and writes values.
- After insertion, the original template row is removed. Any STATIC rows between tables (headers, empty spacing, formulas) are preserved and NOT overwritten.
- Rows containing ONLY control markers (`{{#each}}`, `{{/each}}`, `{{#if}}`, `{{/if}}`, `{{else}}`) are removed.

Requirements and tips:
- Place the second table header in a separate static row between the first table's `{{/each}}` and second table's `{{#each}}` blocks.
- Don't mix headers with control markers in one row — such rows will be removed as marker rows.
- Vertical cell merges are not copied automatically; horizontal merges from the template row are duplicated for each generated row.

Template example (sheet fragment):

```
A1: Table A
A2: {{#each $.tableA as $row i=$i}}
A3: {{= $i+1}}   | B3: {{= $row.name}}
A4: {{/each}}

A6: Table B
A7: {{#each $.tableB as $row i=$i}}
A8: {{= $i+1}}   | B8: {{= $row.name}}
A9: {{/each}}
```

JSON:

```json
{
  "tableA": [{"name": "A-1"}, {"name": "A-2"}, ..., {"name": "A-10"}],
  "tableB": [{"name": "B-1"}, {"name": "B-2"}, ..., {"name": "B-10"}]
}
```

Result: rows `A-1..A-10` go immediately under "Table A"; header "Table B" stays in place, and data `B-1..B-10` is rendered BELOW it. Thus, the first table doesn't overwrite the second regardless of row count.

Correct/incorrect:

```
// ✅ Correct: header between blocks — separate static row
... {{/each}}
Table B
{{#each $.tableB as $row}}
...
{{/each}}

// ❌ Incorrect: header in one row with marker (row will be removed)
Table B {{#each $.tableB as $row}}
...
{{/each}}
```

### Example with dynamic index

JSON:
```json
{
  "groups": [
    { "title": "G1", "rows": ["A", "B"] },
    { "title": "G2", "rows": ["X"] }
  ]
}
```

Template:

| A                                         |
|-------------------------------------------|
| {{#each $.groups as $g i=$gi}}            |
| Header {{= $gi+1}}: {{= $g.title}}        |
| {{#each $g.rows as $r i=$ri}}             |
| Element {{= $ri+1}}: {{= $g.rows[$ri]}}   |
| {{/each}}                                 |
| {{/each}}                                 |

Result:

| A                          |
|----------------------------|
| Header 1: G1               |
| Element 1: A               |
| Element 2: B               |
| Header 2: G2               |
| Element 1: X               |

### Code Example

```go
package main

import (
    "log"
    excel "github.com/your/module/pkg/excel"
)

func main() {
    templatePath := "templates/tech_map.xlsx"
    destPath := "results/out.xlsx"
    jsonData := `{"users":[{"name":"A","age":30},{"name":"B","age":25}]}`

    if err := excel.WriteResultsWithTemplate(templatePath, destPath, []string{jsonData}); err != nil {
        log.Fatalf("failed: %v", err)
    }
}
```
