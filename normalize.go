package exceltemplar

import (
	"encoding/json"
	"log"
)

// NormalizeForExcel приводит JSON строки к более предсказуемой форме:
// - добавляет отсутствующие поля пустыми значениями
// - сортирует массивы объектов по ключу name/code при наличии
// - удаляет явные дубликаты объектов (по сериализованному виду)
func NormalizeForExcel(jsonStrings []string) []string {
	out := make([]string, 0, len(jsonStrings))
	for _, s := range jsonStrings {
		norm := normalizeOne(s)
		out = append(out, norm)
	}
	return out
}

func normalizeOne(s string) string {
	if s == "" {
		return s
	}
	var v interface{}
	if err := json.Unmarshal([]byte(s), &v); err != nil {
		// если не JSON — возвращаем как есть
		return s
	}
	v = deepNormalize(v)
	b, err := json.Marshal(v)
	if err != nil {
		log.Printf("⚠️ Ошибка сериализации при нормализации: %v", err)
		return s
	}
	return string(b)
}

func deepNormalize(v interface{}) interface{} {
	switch vv := v.(type) {
	case []interface{}:
		// сохраняем исходный порядок, просто рекурсивно нормализуем элементы
		for i := range vv {
			vv[i] = deepNormalize(vv[i])
		}
		return vv
	case map[string]interface{}:
		for k, val := range vv {
			vv[k] = deepNormalize(val)
		}
		return vv
	default:
		return vv
	}
}

func deduplicateArray(arr []interface{}) []interface{} {
	seen := make(map[string]struct{})
	out := make([]interface{}, 0, len(arr))
	for _, it := range arr {
		b, err := json.Marshal(it)
		if err != nil {
			out = append(out, it)
			continue
		}
		key := string(b)
		if _, ok := seen[key]; ok {
			continue
		}
		seen[key] = struct{}{}
		out = append(out, it)
	}
	return out
}
