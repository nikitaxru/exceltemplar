package exceltemplar

import (
	"encoding/json"
	"errors"
	"fmt"
	"regexp"
	"sort"
	"strconv"
	"strings"

	expro "github.com/expr-lang/expr"
	"github.com/xuri/excelize/v2"
)

// Новый движок шаблонов для Excel с явным синтаксисом {{...}}.
// Поддержка:
// - {{= expr}}
// - {{#each path as $item i=$i}} ... {{/each}}
// - {{#each-obj path as $k $v}} ... {{/each-obj}}
// - {{#if expr}} ... {{else}} ... {{/if}}
// - функции: len(), exists(), join()
// Внешний API сохранён: LoadTemplate, Render, Save.

// -----------------------------
// AST
// -----------------------------

type node interface{}

type rowNode struct {
	sheet string
	row   int // 1-based
	cells []cellTpl
}

type eachNode struct {
	path     string
	itemVar  string
	indexVar string
	children []node
}

type eachObjNode struct {
	path     string
	keyVar   string
	valVar   string
	children []node
}

type ifNode struct {
	expr      string
	thenNodes []node
	elseNodes []node
}

type cellTokenKind int

const (
	tokenText cellTokenKind = iota
	tokenExpr
)

type cellToken struct {
	kind cellTokenKind
	text string
	expr string
}

type cellTpl struct {
	col    int
	raw    string
	tokens []cellToken
}

type sheetTemplate struct {
	name    string
	nodes   []node
	minRow  int
	maxRow  int
	rowTpls map[int]rowTpl
}

type Template struct {
	f      *excelize.File
	sheets map[string]*sheetTemplate
}

// rowTpl описывает свойства шаблонной строки (стили, исходные значения, горизонтальные слияния)
type rowTpl struct {
	styles  map[int]int
	rawVals map[int]string
	merges  []struct {
		startCol int
		endCol   int
	}
}

// -----------------------------
// Парсер листа
// -----------------------------

var (
	rxCtrlEach       = regexp.MustCompile(`^\{\{#each\s+(.+?)\}\}$`)
	rxCtrlEachObj    = regexp.MustCompile(`^\{\{#each-obj\s+(.+?)\}\}$`)
	rxCtrlEndEach    = regexp.MustCompile(`^\{\{\/each\}\}$`)
	rxCtrlEndEachObj = regexp.MustCompile(`^\{\{\/each-obj\}\}$`)
	rxCtrlIf         = regexp.MustCompile(`^\{\{#if\s+(.+?)\}\}$`)
	rxCtrlElse       = regexp.MustCompile(`^\{\{else\}\}$`)
	rxCtrlEndIf      = regexp.MustCompile(`^\{\{\/if\}\}$`)
	// Разрешаем любые символы внутри выражения (включая потенциальные скрытые/служебные символы Excel)
	rxExpr = regexp.MustCompile(`\{\{=\s*([\s\S]+?)\s*\}\}`)
)

// LoadTemplate строит AST для каждого листа
func LoadTemplate(path string) (*Template, error) {
	f, err := excelize.OpenFile(path)
	if err != nil {
		return nil, err
	}
	t := &Template{f: f, sheets: map[string]*sheetTemplate{}}
	for _, sheet := range f.GetSheetList() {
		st, err := parseSheet(f, sheet)
		if err != nil {
			return nil, fmt.Errorf("парсинг листа %s: %w", sheet, err)
		}
		t.sheets[sheet] = st
	}
	return t, nil
}

func parseSheet(f *excelize.File, sheet string) (*sheetTemplate, error) {
	rows, err := f.GetRows(sheet)
	if err != nil {
		return nil, err
	}
	var nodes []node
	type stackItem struct {
		kind   string // each | each-obj | if
		en     *eachNode
		eo     *eachObjNode
		in     *ifNode
		target *[]node
	}
	var stack []stackItem
	var minRow, maxRow int

	appendNode := func(n node) {
		if len(stack) == 0 {
			nodes = append(nodes, n)
		} else {
			*stack[len(stack)-1].target = append(*stack[len(stack)-1].target, n)
		}
	}

	for rIdx, row := range rows {
		rowNum := rIdx + 1
		// Контрольные маркеры
		ctrl := false
		for _, cell := range row {
			trimmed := strings.TrimSpace(cell)
			if trimmed == "" {
				continue
			}
			if m := rxCtrlEach.FindStringSubmatch(trimmed); len(m) == 2 {
				path, itemVar, indexVar := parseEachHeader(m[1])
				en := &eachNode{path: path, itemVar: itemVar, indexVar: indexVar}
				en.children = []node{}
				stack = append(stack, stackItem{kind: "each", en: en, target: &en.children})
				ctrl = true
				break
			}
			if m := rxCtrlEachObj.FindStringSubmatch(trimmed); len(m) == 2 {
				path, kVar, vVar := parseEachObjHeader(m[1])
				eo := &eachObjNode{path: path, keyVar: kVar, valVar: vVar}
				eo.children = []node{}
				stack = append(stack, stackItem{kind: "each-obj", eo: eo, target: &eo.children})
				ctrl = true
				break
			}
			if m := rxCtrlIf.FindStringSubmatch(trimmed); len(m) == 2 {
				in := &ifNode{expr: m[1]}
				in.thenNodes = []node{}
				stack = append(stack, stackItem{kind: "if", in: in, target: &in.thenNodes})
				ctrl = true
				break
			}
			if rxCtrlElse.MatchString(trimmed) {
				if len(stack) == 0 || stack[len(stack)-1].kind != "if" {
					return nil, fmt.Errorf("некорректный else на строке %d", rowNum)
				}
				it := &stack[len(stack)-1]
				it.target = &it.in.elseNodes
				ctrl = true
				break
			}
			if rxCtrlEndEach.MatchString(trimmed) || rxCtrlEndEachObj.MatchString(trimmed) {
				if len(stack) == 0 || (stack[len(stack)-1].kind != "each" && stack[len(stack)-1].kind != "each-obj") {
					return nil, fmt.Errorf("некорректный /each на строке %d", rowNum)
				}
				top := stack[len(stack)-1]
				stack = stack[:len(stack)-1]
				if top.kind == "each" {
					appendNode(top.en)
				} else {
					appendNode(top.eo)
				}
				ctrl = true
				break
			}
			if rxCtrlEndIf.MatchString(trimmed) {
				if len(stack) == 0 || stack[len(stack)-1].kind != "if" {
					return nil, fmt.Errorf("некорректный /if на строке %d", rowNum)
				}
				top := stack[len(stack)-1]
				stack = stack[:len(stack)-1]
				appendNode(top.in)
				ctrl = true
				break
			}
		}
		if ctrl {
			if minRow == 0 || rowNum < minRow {
				minRow = rowNum
			}
			if rowNum > maxRow {
				maxRow = rowNum
			}
			continue
		}

		// Обычная строка: ищем {{= expr}}
		var cells []cellTpl
		has := false
		for cIdx, cell := range row {
			toks := parseCellTokens(cell)
			if len(toks) > 0 {
				has = true
				cells = append(cells, cellTpl{col: cIdx + 1, raw: cell, tokens: toks})
			}
		}
		// Если строка внутри блока (each/if), учитываем даже статические строки,
		// чтобы их можно было повторять вместе с данными
		if has || len(stack) > 0 {
			rn := &rowNode{sheet: sheet, row: rowNum, cells: cells}
			appendNode(rn)
			if minRow == 0 || rowNum < minRow {
				minRow = rowNum
			}
			if rowNum > maxRow {
				maxRow = rowNum
			}
		}
	}
	if len(stack) != 0 {
		return nil, errors.New("несбалансированные блоки each/if")
	}
	// Собираем шаблонные строки (rowTpls) для устойчивого копирования стилей/значений/merge
	rowTpls := make(map[int]rowTpl)
	tplRowsSet := make(map[int]struct{})
	var collectTplRows func([]node)
	collectTplRows = func(ns []node) {
		for _, n := range ns {
			switch nn := n.(type) {
			case *rowNode:
				tplRowsSet[nn.row] = struct{}{}
			case *eachNode:
				collectTplRows(nn.children)
			case *eachObjNode:
				collectTplRows(nn.children)
			case *ifNode:
				collectTplRows(nn.thenNodes)
				collectTplRows(nn.elseNodes)
			}
		}
	}
	collectTplRows(nodes)

	const maxCols = 100
	for tplRow := range tplRowsSet {
		rt := rowTpl{styles: make(map[int]int), rawVals: make(map[int]string)}
		for col := 1; col <= maxCols; col++ {
			addr, _ := excelize.CoordinatesToCellName(col, tplRow)
			v, _ := f.GetCellValue(sheet, addr)
			if v != "" {
				rt.rawVals[col] = v
			}
			if sid, err := f.GetCellStyle(sheet, addr); err == nil && sid != 0 {
				rt.styles[col] = sid
			}
		}
		merges, _ := f.GetMergeCells(sheet)
		for _, m := range merges {
			s := m.GetStartAxis()
			e := m.GetEndAxis()
			sc, sr, _ := excelize.SplitCellName(s)
			ec, er, _ := excelize.SplitCellName(e)
			if sr == tplRow && er == tplRow {
				scn, _ := excelize.ColumnNameToNumber(sc)
				ecn, _ := excelize.ColumnNameToNumber(ec)
				rt.merges = append(rt.merges, struct{ startCol, endCol int }{startCol: scn, endCol: ecn})
			}
		}
		rowTpls[tplRow] = rt
	}

	return &sheetTemplate{name: sheet, nodes: nodes, minRow: minRow, maxRow: maxRow, rowTpls: rowTpls}, nil
}

func parseCellTokens(s string) []cellToken {
	ms := rxExpr.FindAllStringSubmatchIndex(s, -1)
	if len(ms) == 0 {
		return nil
	}
	// Быстрый guard: если вся ячейка — это один единственный expr без окружения текста,
	// мы считаем, что строка содержит выражение и должна рендериться.
	// Это уже покрывается логикой ниже, оставляем как комментарий для ясности.
	var toks []cellToken
	last := 0
	for _, m := range ms {
		start, end := m[0], m[1]
		es, ee := m[2], m[3]
		if start > last {
			toks = append(toks, cellToken{kind: tokenText, text: s[last:start]})
		}
		toks = append(toks, cellToken{kind: tokenExpr, expr: strings.TrimSpace(s[es:ee])})
		last = end
	}
	if last < len(s) {
		toks = append(toks, cellToken{kind: tokenText, text: s[last:]})
	}
	return toks
}

func parseEachHeader(src string) (path, itemVar, indexVar string) {
	// path [as $item] [i=$i]
	itemVar = "$"
	parts := strings.Fields(src)
	if len(parts) == 0 {
		return "", "$", ""
	}
	// path до ключевых слов
	var pparts []string
	i := 0
	for i < len(parts) {
		p := parts[i]
		if p == "as" || strings.HasPrefix(p, "i=") {
			break
		}
		pparts = append(pparts, p)
		i++
	}
	path = strings.Join(pparts, " ")
	for i < len(parts) {
		p := parts[i]
		if p == "as" && i+1 < len(parts) {
			itemVar = parts[i+1]
			i += 2
			continue
		}
		if strings.HasPrefix(p, "i=") {
			indexVar = strings.TrimPrefix(p, "i=")
			i++
			continue
		}
		i++
	}
	return
}

func parseEachObjHeader(src string) (path, keyVar, valVar string) {
	parts := strings.Fields(src)
	if len(parts) == 0 {
		return "", "$k", "$v"
	}
	path = parts[0]
	keyVar, valVar = "$k", "$v"
	if len(parts) >= 3 && parts[1] == "as" {
		keyVar = parts[2]
		if len(parts) >= 4 {
			valVar = parts[3]
		}
	}
	return
}

// -----------------------------
// Eval и резолвер путей
// -----------------------------

type renderRow struct {
	sheet  string
	tplRow int
	values map[int]string
}

type evalContext struct {
	current interface{}
	parent  *evalContext
	root    []interface{}
	vars    map[string]interface{}
}

func resolvePath(ctx *evalContext, path string) (interface{}, bool) {
	path = strings.TrimSpace(path)
	if path == "." {
		return ctx.current, ctx.current != nil
	}
	// Явный абсолютный путь от корня через $root
	if strings.HasPrefix(path, "$root") {
		var rest string
		if path == "$root" {
			rest = ""
		} else {
			_, rest = splitFirst(path, ".")
		}
		for _, r := range ctx.root {
			if rest == "" {
				if r != nil {
					return r, true
				}
				continue
			}
			if v, ok := drillWithCtx(ctx, r, rest); ok {
				return v, true
			}
		}
		return nil, false
	}
	// Явный абсолютный путь от корня: $.something
	if strings.HasPrefix(path, "$.") {
		_, rest := splitFirst(path, ".")
		for _, r := range ctx.root {
			if v, ok := drillWithCtx(ctx, r, rest); ok {
				return v, true
			}
		}
		return nil, false
	}
	if strings.HasPrefix(path, "$") && path != "$root" && path != "$" {
		name, rest := splitFirst(path[1:], ".")
		v, ok := ctx.vars["$"+name]
		if !ok {
			return nil, false
		}
		return drillWithCtx(ctx, v, rest)
	}
	if path == "$" || strings.HasPrefix(path, "$root") {
		var rest string
		if path == "$" {
			rest = ""
		} else {
			_, rest = splitFirst(path, ".")
		}
		for _, r := range ctx.root {
			if rest == "" {
				if r != nil {
					return r, true
				}
				continue
			}
			if v, ok := drillWithCtx(ctx, r, rest); ok {
				return v, true
			}
		}
		return nil, false
	}
	if strings.HasPrefix(path, ".") {
		_, rest := splitFirst(path, ".")
		if rest == "" {
			return ctx.current, ctx.current != nil
		}
		return drillWithCtx(ctx, ctx.current, rest)
	}
	// абсолютный путь от корня
	for _, r := range ctx.root {
		if v, ok := drillWithCtx(ctx, r, path); ok {
			return v, true
		}
	}
	return nil, false
}

func splitFirst(s, sep string) (string, string) {
	if s == "" {
		return "", ""
	}
	if idx := strings.Index(s, sep); idx >= 0 {
		return s[:idx], s[idx+len(sep):]
	}
	return s, ""
}

func drill(v interface{}, path string) (interface{}, bool) {
	if path == "" {
		return v, true
	}
	cur := v
	rest := path
	for rest != "" {
		seg, tail := nextSeg(rest)
		if strings.HasPrefix(seg, "[") {
			// индекс
			if arr, ok := cur.([]interface{}); ok {
				idxStr := strings.Trim(seg, "[]")
				i, err := strconv.Atoi(idxStr)
				if err != nil || i < 0 || i >= len(arr) {
					return nil, false
				}
				cur = arr[i]
			} else {
				return nil, false
			}
		} else {
			if m, ok := cur.(map[string]interface{}); ok {
				nv, ok := m[seg]
				if !ok {
					return nil, false
				}
				cur = nv
			} else {
				return nil, false
			}
		}
		rest = tail
	}
	return cur, true
}

// drillWithCtx поддерживает динамические индексы вида [$i] (значение берётся из контекста)
func drillWithCtx(ctx *evalContext, v interface{}, path string) (interface{}, bool) {
	if path == "" {
		return v, true
	}
	cur := v
	rest := path
	for rest != "" {
		seg, tail := nextSeg(rest)
		if strings.HasPrefix(seg, "[") {
			if arr, ok := cur.([]interface{}); ok {
				idxStr := strings.Trim(seg, "[]")
				// сначала попробуем как число
				if i, err := strconv.Atoi(idxStr); err == nil {
					if i < 0 || i >= len(arr) {
						return nil, false
					}
					cur = arr[i]
				} else {
					// как ссылка/переменная
					if iv, ok := resolvePath(ctx, idxStr); ok {
						switch vv := iv.(type) {
						case float64:
							i := int(vv)
							if i < 0 || i >= len(arr) {
								return nil, false
							}
							cur = arr[i]
						case string:
							if ii, err := strconv.Atoi(vv); err == nil {
								if ii < 0 || ii >= len(arr) {
									return nil, false
								}
								cur = arr[ii]
							} else {
								return nil, false
							}
						default:
							return nil, false
						}
					} else {
						return nil, false
					}
				}
			} else {
				return nil, false
			}
		} else {
			if m, ok := cur.(map[string]interface{}); ok {
				nv, ok := m[seg]
				if !ok {
					return nil, false
				}
				cur = nv
			} else {
				return nil, false
			}
		}
		rest = tail
	}
	return cur, true
}

func nextSeg(path string) (seg string, tail string) {
	if path == "" {
		return "", ""
	}
	if path[0] == '[' {
		if i := strings.Index(path, "]"); i >= 0 {
			seg = path[:i+1]
			if i+1 < len(path) && path[i+1] == '.' {
				tail = path[i+2:]
			} else {
				tail = path[i+1:]
			}
			return
		}
	}
	i := 0
	for i < len(path) && path[i] != '.' && path[i] != '[' {
		i++
	}
	seg = path[:i]
	if i < len(path) && path[i] == '.' {
		tail = path[i+1:]
	} else {
		tail = path[i:]
	}
	return
}

func toString(v interface{}) string {
	switch vv := v.(type) {
	case nil:
		return ""
	case string:
		return vv
	case float64:
		if vv == float64(int64(vv)) {
			return fmt.Sprintf("%d", int64(vv))
		}
		return fmt.Sprintf("%v", vv)
	case bool:
		if vv {
			return "true"
		}
		return "false"
	default:
		return fmt.Sprintf("%v", vv)
	}
}

func fnLen(ctx *evalContext, arg string) (float64, bool) {
	v, ok := resolvePath(ctx, arg)
	if !ok {
		return 0, false
	}
	switch vv := v.(type) {
	case []interface{}:
		return float64(len(vv)), true
	case string:
		return float64(len(vv)), true
	case map[string]interface{}:
		return float64(len(vv)), true
	default:
		return 0, false
	}
}

func fnExists(ctx *evalContext, arg string) bool {
	_, ok := resolvePath(ctx, arg)
	return ok
}

func fnJoin(ctx *evalContext, args []string) (string, error) {
	if len(args) < 2 {
		return "", fmt.Errorf("join: минимум 2 аргумента")
	}
	arrPath := strings.TrimSpace(args[0])
	sep := strings.Trim(strings.TrimSpace(args[1]), "\"'")
	var field string
	if len(args) >= 3 {
		field = strings.Trim(strings.TrimSpace(args[2]), "\"'")
	}
	v, ok := resolvePath(ctx, arrPath)
	if !ok {
		return "", nil
	}
	arr, ok := v.([]interface{})
	if !ok {
		return "", fmt.Errorf("join: не массив")
	}
	vals := make([]string, 0, len(arr))
	for _, it := range arr {
		var val interface{} = it
		if field != "" {
			if m, ok := it.(map[string]interface{}); ok {
				if f, ok := drill(m, field); ok {
					val = f
				} else {
					val = nil
				}
			} else {
				val = nil
			}
		}
		vals = append(vals, toString(val))
	}
	return strings.Join(vals, sep), nil
}

func evalScalar(ctx *evalContext, expr string) (interface{}, error) {
	expr = strings.TrimSpace(expr)
	// функции
	if strings.HasPrefix(expr, "iif(") && strings.HasSuffix(expr, ")") {
		inner := strings.TrimSuffix(strings.TrimPrefix(expr, "iif("), ")")
		args := splitArgs(inner)
		// Требуем минимум 2 аргумента: cond, then [, else]
		if len(args) < 2 {
			return "", fmt.Errorf("iif: минимум 2 аргумента (cond, then[, else])")
		}
		condStr := args[0]
		thenStr := args[1]
		var elseStr string
		if len(args) >= 3 {
			elseStr = args[2]
		}
		cond, err := evalBool(ctx, condStr)
		if err != nil {
			return "", fmt.Errorf("iif: ошибка условия: %w", err)
		}
		if cond {
			v, err := evalScalar(ctx, thenStr)
			if err != nil {
				return "", err
			}
			return v, nil
		}
		v, err := evalScalar(ctx, elseStr)
		if err != nil {
			return "", err
		}
		return v, nil
	}
	if strings.HasPrefix(expr, "join(") && strings.HasSuffix(expr, ")") {
		inner := strings.TrimSuffix(strings.TrimPrefix(expr, "join("), ")")
		return fnJoin(ctx, splitArgs(inner))
	}
	if strings.HasPrefix(expr, "len(") && strings.Contains(expr, ")") {
		arg := strings.TrimSuffix(strings.TrimPrefix(expr, "len("), ")")
		l, _ := fnLen(ctx, arg)
		return l, nil
	}
	if strings.HasPrefix(expr, "exists(") && strings.HasSuffix(expr, ")") {
		arg := strings.TrimSuffix(strings.TrimPrefix(expr, "exists("), ")")
		return fnExists(ctx, arg), nil
	}
	// арифметика вида $i+1
	if strings.Contains(expr, "+") {
		parts := strings.Split(expr, "+")
		if len(parts) == 2 {
			left := strings.TrimSpace(parts[0])
			right := strings.TrimSpace(parts[1])
			lv, _ := resolvePath(ctx, left)
			li := toFloat(lv)
			ri, _ := strconv.Atoi(right)
			return li + float64(ri), nil
		}
	}
	// литералы
	if (strings.HasPrefix(expr, "\"") && strings.HasSuffix(expr, "\"")) || (strings.HasPrefix(expr, "'") && strings.HasSuffix(expr, "'")) {
		return strings.Trim(expr, "\"'"), nil
	}
	if v, ok := resolvePath(ctx, expr); ok {
		switch v.(type) {
		case []interface{}, map[string]interface{}:
			return nil, fmt.Errorf("скалярная вставка получила коллекцию; используйте each/join")
		}
		return v, nil
	}
	if n, err := strconv.Atoi(expr); err == nil {
		return float64(n), nil
	}
	return "", nil
}

func toFloat(v interface{}) float64 {
	switch vv := v.(type) {
	case float64:
		return vv
	case int:
		return float64(vv)
	case string:
		n, _ := strconv.Atoi(vv)
		return float64(n)
	default:
		return 0
	}
}

func evalBool(ctx *evalContext, expr string) (bool, error) {
	// Трансформируем обращения вида $.a.b и $var.c[d] в path("...") для expr-lang
	expr = transformBoolExprPaths(expr)
	// Окружение с функциями, используемое как на этапе компиляции, так и выполнения
	env := map[string]interface{}{
		"len": func(x interface{}) float64 {
			switch v := x.(type) {
			case []interface{}:
				return float64(len(v))
			case string:
				return float64(len(v))
			case map[string]interface{}:
				return float64(len(v))
			default:
				return 0
			}
		},
		"exists": func(x interface{}) bool {
			switch v := x.(type) {
			case string:
				_, ok := resolvePath(ctx, v)
				return ok
			default:
				return truthy(v)
			}
		},
		// Доступ к значениям по пути
		"path": func(p string) interface{} {
			if v, ok := resolvePath(ctx, p); ok {
				return v
			}
			return nil
		},
		// Сокращение: $(".a.b") / $("$.x")
		"$": func(p string) interface{} {
			if v, ok := resolvePath(ctx, p); ok {
				return v
			}
			return nil
		},
	}
	// Полный парсинг булевых выражений через expr-lang с тем же env
	program, err := expro.Compile(expr, expro.Env(env))
	if err != nil {
		return false, err
	}
	out, err := expro.Run(program, env)
	if err != nil {
		return false, err
	}
	if b, ok := out.(bool); ok {
		return b, nil
	}
	return truthy(out), nil
}

// transformBoolExprPaths преобразует обращения вида $.a.b или $var.c[d] в path("...")
// для корректной работы в expr-lang (который не знает нотацию $.). Обрабатывает только
// участки вне кавычек и не затрагивает уже существующие вызовы path("...").
func transformBoolExprPaths(src string) string {
	if !strings.Contains(src, "$") {
		return src
	}
	var out strings.Builder
	inQuote := byte(0)
	i := 0
	for i < len(src) {
		ch := src[i]
		if inQuote == 0 && (ch == '\'' || ch == '"') {
			inQuote = ch
			out.WriteByte(ch)
			i++
			continue
		}
		if inQuote != 0 {
			out.WriteByte(ch)
			if ch == inQuote {
				inQuote = 0
			}
			i++
			continue
		}
		if ch == '$' {
			start := i
			j := i + 1
			// допускаем либо точку (для $.), либо имя переменной
			if j < len(src) && src[j] == '.' {
				j++
			} else {
				for j < len(src) && ((src[j] >= 'a' && src[j] <= 'z') || (src[j] >= 'A' && src[j] <= 'Z') || (src[j] == '_')) {
					j++
				}
				if j < len(src) && src[j] == '.' {
					j++
				}
			}
			// продолжаем до первого символа, не относящегося к пути
			k := j
			for k < len(src) {
				c := src[k]
				if (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') || (c >= '0' && c <= '9') || c == '_' || c == '.' || c == '[' || c == ']' || c == '$' {
					k++
					continue
				}
				break
			}
			if k > start+1 {
				path := src[start:k]
				out.WriteString("path(\"")
				out.WriteString(path)
				out.WriteString("\")")
				i = k
				continue
			}
		}
		out.WriteByte(ch)
		i++
	}
	return out.String()
}

// legacy helpers compare/toStringMaybe удалены как неиспользуемые

func truthy(v interface{}) bool {
	switch vv := v.(type) {
	case nil:
		return false
	case bool:
		return vv
	case string:
		return vv != ""
	case []interface{}:
		return len(vv) > 0
	case map[string]interface{}:
		return len(vv) > 0
	case float64:
		return vv != 0
	default:
		return true
	}
}

func splitArgs(s string) []string {
	var args []string
	var b strings.Builder
	quote := byte(0)
	depth := 0 // глубина круглых скобок для вложенных вызовов (join(...), iif(...))
	for i := 0; i < len(s); i++ {
		ch := s[i]
		if quote == 0 {
			// Управление кавычками
			if ch == '\'' || ch == '"' {
				quote = ch
				b.WriteByte(ch)
				continue
			}
			// Учёт вложенных круглых скобок
			if ch == '(' {
				depth++
				b.WriteByte(ch)
				continue
			}
			if ch == ')' {
				if depth > 0 {
					depth--
				}
				b.WriteByte(ch)
				continue
			}
			// Разделяем по запятым только на верхнем уровне
			if ch == ',' && depth == 0 {
				args = append(args, strings.TrimSpace(b.String()))
				b.Reset()
				continue
			}
			b.WriteByte(ch)
			continue
		}
		// Внутри кавычек: копируем до закрытия
		b.WriteByte(ch)
		if ch == quote {
			quote = 0
		}
	}
	if b.Len() > 0 {
		args = append(args, strings.TrimSpace(b.String()))
	}
	return args
}

// -----------------------------
// Рендер в память и применение к Excel
// -----------------------------

func (t *Template) Render(outputs []string) error {
	// парсим outputs в корневые объекты
	var roots []interface{}
	for _, s := range outputs {
		s = sanitizeJSONBlock(s)
		if strings.TrimSpace(s) == "" {
			continue
		}
		var v interface{}
		if err := json.Unmarshal([]byte(s), &v); err == nil {
			roots = append(roots, v)
		}
	}
	for _, st := range t.sheets {
		rendered, err := renderSheet(st, &evalContext{current: nil, parent: nil, root: roots, vars: map[string]interface{}{}})
		if err != nil {
			return fmt.Errorf("лист %s: %w", st.name, err)
		}
		if err := t.applyRendered(st, rendered); err != nil {
			return fmt.Errorf("лист %s: %w", st.name, err)
		}
	}
	return nil
}

func renderSheet(st *sheetTemplate, ctx *evalContext) ([]renderRow, error) {
	var out []renderRow
	var walk func([]node, *evalContext) error
	walk = func(nodes []node, ctx *evalContext) error {
		for _, n := range nodes {
			switch nn := n.(type) {
			case *rowNode:
				vals := map[int]string{}
				for _, c := range nn.cells {
					var sb strings.Builder
					for _, tk := range c.tokens {
						if tk.kind == tokenText {
							sb.WriteString(tk.text)
							continue
						}
						v, err := evalScalar(ctx, tk.expr)
						if err != nil {
							return err
						}
						sb.WriteString(toString(v))
					}
					vals[c.col] = sb.String()
				}
				out = append(out, renderRow{sheet: nn.sheet, tplRow: nn.row, values: vals})
			case *eachNode:
				v, ok := resolvePath(ctx, nn.path)
				if !ok {
					continue
				}
				arr, ok := v.([]interface{})
				if !ok {
					continue
				}
				for i, item := range arr {
					nctx := &evalContext{current: item, parent: ctx, root: ctx.root, vars: map[string]interface{}{}}
					for k, v := range ctx.vars {
						nctx.vars[k] = v
					}
					if nn.itemVar != "" {
						nctx.vars[nn.itemVar] = item
					}
					if nn.indexVar != "" {
						nctx.vars[nn.indexVar] = float64(i)
					}
					if err := walk(nn.children, nctx); err != nil {
						return err
					}
				}
			case *eachObjNode:
				v, ok := resolvePath(ctx, nn.path)
				if !ok {
					continue
				}
				m, ok := v.(map[string]interface{})
				if !ok {
					continue
				}
				keys := make([]string, 0, len(m))
				for k := range m {
					keys = append(keys, k)
				}
				sort.Strings(keys)
				for _, k := range keys {
					val := m[k]
					nctx := &evalContext{current: val, parent: ctx, root: ctx.root, vars: map[string]interface{}{}}
					for kk, vv := range ctx.vars {
						nctx.vars[kk] = vv
					}
					if nn.keyVar != "" {
						nctx.vars[nn.keyVar] = k
					}
					if nn.valVar != "" {
						nctx.vars[nn.valVar] = val
					}
					if err := walk(nn.children, nctx); err != nil {
						return err
					}
				}
			case *ifNode:
				cond, err := evalBool(ctx, nn.expr)
				if err != nil {
					return err
				}
				if cond {
					if err := walk(nn.thenNodes, ctx); err != nil {
						return err
					}
				} else {
					if err := walk(nn.elseNodes, ctx); err != nil {
						return err
					}
				}
			}
		}
		return nil
	}
	if err := walk(st.nodes, ctx); err != nil {
		return nil, err
	}
	return out, nil
}

func (t *Template) applyRendered(st *sheetTemplate, rows []renderRow) error {
	sheet := st.name
	if st.minRow == 0 && st.maxRow == 0 {
		return nil
	}
	// Поддерживаем актуальные позиции исходных шаблонных строк
	tplPos := make(map[int]int)
	for tplRow := range st.rowTpls {
		tplPos[tplRow] = tplRow
	}
	// Глобальный барьер: запрещает вставку выше уже вставленных данных, чтобы сохранять порядок rows
	barrier := st.minRow

	// Вставляем строки в порядке rows, вычисляя позицию как max(barrier, текущая позиция шаблонной строки)
	for _, rr := range rows {
		rt := st.rowTpls[rr.tplRow]
		curTpl := tplPos[rr.tplRow]
		insertAt := curTpl
		if barrier > insertAt {
			insertAt = barrier
		}
		if err := t.f.InsertRows(sheet, insertAt, 1); err != nil {
			return err
		}
		// Заполняем вставленную строку
		dstRow := insertAt
		// Стили из образца
		for col, sid := range rt.styles {
			addr, _ := excelize.CoordinatesToCellName(col, dstRow)
			if err := t.f.SetCellStyle(sheet, addr, addr, sid); err != nil {
				return err
			}
		}
		// Статические значения (без выражений) из образца
		for col, rawv := range rt.rawVals {
			addr, _ := excelize.CoordinatesToCellName(col, dstRow)
			if rxExpr.MatchString(rawv) {
				if err := t.f.SetCellValue(sheet, addr, ""); err != nil {
					return err
				}
			} else {
				if err := t.f.SetCellValue(sheet, addr, rawv); err != nil {
					return err
				}
			}
		}
		// Рендеренные значения поверх
		for col, val := range rr.values {
			addr, _ := excelize.CoordinatesToCellName(col, dstRow)
			if err := t.f.SetCellValue(sheet, addr, val); err != nil {
				return err
			}
		}
		// Горизонтальные слияния
		for _, mg := range rt.merges {
			c1, _ := excelize.CoordinatesToCellName(mg.startCol, dstRow)
			c2, _ := excelize.CoordinatesToCellName(mg.endCol, dstRow)
			_ = t.f.MergeCell(sheet, c1, c2)
		}

		// Обновляем барьер и позиции всех шаблонных строк на листе (всё ниже insertAt сдвигается на +1)
		barrier = insertAt + 1
		for tr, pos := range tplPos {
			if pos >= insertAt {
				tplPos[tr] = pos + 1
			}
		}
	}

	// Удаляем исходные шаблонные строки (которые теперь смещены согласно tplPos)
	// Удаляем снизу вверх
	var toDelete []int
	for _, pos := range tplPos {
		toDelete = append(toDelete, pos)
	}
	sort.Sort(sort.Reverse(sort.IntSlice(toDelete)))
	for _, r := range toDelete {
		if err := t.f.RemoveRow(sheet, r); err != nil {
			return err
		}
	}

	// Удаляем строки, содержащие только управляющие маркеры ({{#each}}, {{/each}}, {{#if}}, {{/if}}, {{else}})
	if err := removeControlMarkerRows(t.f, sheet); err != nil {
		return err
	}
	return nil
}

// removeControlMarkerRows удаляет строки, которые содержат только управляющие маркеры шаблона
// (каждая непустая ячейка строки полностью совпадает с одним из маркеров).
func removeControlMarkerRows(f *excelize.File, sheet string) error {
	rows, err := f.GetRows(sheet)
	if err != nil {
		return err
	}
	var toDelete []int
	isCtrl := func(s string) bool {
		s = strings.TrimSpace(s)
		if s == "" {
			return false
		}
		return rxCtrlEach.MatchString(s) || rxCtrlEndEach.MatchString(s) || rxCtrlEachObj.MatchString(s) || rxCtrlEndEachObj.MatchString(s) || rxCtrlIf.MatchString(s) || rxCtrlEndIf.MatchString(s) || rxCtrlElse.MatchString(s)
	}
	for i, row := range rows {
		hasCtrl := false
		removable := true
		for _, cell := range row {
			c := strings.TrimSpace(cell)
			if c == "" {
				continue
			}
			if isCtrl(c) {
				hasCtrl = true
				continue
			}
			// есть содержимое, не являющееся маркером → строку нельзя удалять
			removable = false
			break
		}
		if hasCtrl && removable {
			toDelete = append(toDelete, i+1) // 1-based
		}
	}
	// удаляем снизу вверх
	for i := len(toDelete) - 1; i >= 0; i-- {
		_ = f.RemoveRow(sheet, toDelete[i])
	}
	return nil
}

// Save сохраняет файл
func (t *Template) Save(destPath string) error { return t.f.SaveAs(destPath) }
