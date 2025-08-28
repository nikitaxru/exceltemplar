package exceltemplar

import (
	"encoding/json"
	"fmt"
	"log"
	"regexp"
	"strings"
	"time"
)

// sanitizeJSONBlock извлекает JSON, обёрнутый в тройные кавычки ``` ... ```.
// Если таких кавычек нет, либо структура неверная, возвращает исходную строку.
var fenceRx = regexp.MustCompile("(?s)```[a-zA-Z]*\\n(.*?)```")

func sanitizeJSONBlock(s string) string {
	if !strings.Contains(s, "```") {
		return s
	}
	m := fenceRx.FindStringSubmatch(s)
	if len(m) >= 2 {
		return strings.TrimSpace(m[1])
	}
	return s
}

// valToCell нормализует значение перед записью в Excel.
func valToCell(v interface{}) interface{} {
	if v == nil {
		return ""
	}
	switch vv := v.(type) {
	case string:
		return vv
	case []interface{}:
		allStr := true
		strs := make([]string, len(vv))
		for i, it := range vv {
			if s, ok := it.(string); ok {
				strs[i] = s
			} else {
				allStr = false
				break
			}
		}
		if allStr {
			return strings.Join(strs, ", ")
		}
		b, _ := json.Marshal(vv)
		return string(b)
	default:
		return fmt.Sprintf("%v", vv)
	}
}

// WriteResultsWithTemplate оставляем без изменений
func WriteResultsWithTemplate(templatePath, destPath string, outputs []string) error {
	log.Printf("📊 Начинаем запись результатов в Excel...")
	log.Printf("📁 Шаблон: %s", templatePath)
	log.Printf("📄 Выходной файл: %s", destPath)
	log.Printf("📝 Количество этапов для записи: %d", len(outputs))

	startTime := time.Now()

	// Логируем размеры данных каждого этапа
	for i, output := range outputs {
		log.Printf("📊 Этап %d: %d символов", i+1, len(output))
	}

	log.Printf("🔄 Загрузка Excel шаблона...")
	tmpl, err := LoadTemplate(templatePath)
	if err != nil {
		log.Printf("❌ Ошибка загрузки шаблона: %v", err)
		return err
	}
	log.Printf("✅ Шаблон загружен успешно")

	// Нормализуем JSON перед рендером для устойчивости
	normalized := NormalizeForExcel(outputs)

	log.Printf("🔄 Рендеринг данных в шаблон...")
	if err := tmpl.Render(normalized); err != nil {
		log.Printf("❌ Ошибка рендеринга: %v", err)
		return err
	}
	log.Printf("✅ Рендеринг завершен")

	log.Printf("💾 Сохранение файла...")
	if err := tmpl.Save(destPath); err != nil {
		log.Printf("❌ Ошибка сохранения: %v", err)
		return err
	}

	duration := time.Since(startTime)
	log.Printf("✅ Excel файл создан за %v", duration)
	log.Printf("📄 Результат сохранен в: %s", destPath)

	return nil
}
