package exceltemplar

import "testing"

func TestResolvePath_BasicAndRelative(t *testing.T) {
	root := map[string]interface{}{
		"a": map[string]interface{}{
			"b": map[string]interface{}{"c": 42.0},
		},
		"arr": []interface{}{
			map[string]interface{}{"name": "zero"},
			map[string]interface{}{"name": "one"},
		},
	}
	ctx := &evalContext{current: root, parent: nil, root: []interface{}{root}, vars: map[string]interface{}{}}

	if v, ok := resolvePath(ctx, "$.a.b.c"); !ok || v.(float64) != 42.0 {
		t.Fatalf("resolve $.a.b.c => %v ok=%v", v, ok)
	}

	// относительный путь от current
	if v, ok := resolvePath(ctx, ".a.b.c"); !ok || v.(float64) != 42.0 {
		t.Fatalf("resolve .a.b.c => %v ok=%v", v, ok)
	}

	// индекс массива
	if v, ok := resolvePath(ctx, "$.arr[1].name"); !ok || v.(string) != "one" {
		t.Fatalf("resolve $.arr[1].name => %v ok=%v", v, ok)
	}
}

func TestResolvePath_Variables(t *testing.T) {
	root := map[string]interface{}{"x": 1.0}
	current := map[string]interface{}{"y": 2.0}
	ctx := &evalContext{current: current, parent: nil, root: []interface{}{root}, vars: map[string]interface{}{"$v": map[string]interface{}{"z": 3.0}}}

	if v, ok := resolvePath(ctx, "$v.z"); !ok || v.(float64) != 3.0 {
		t.Fatalf("resolve $v.z => %v ok=%v", v, ok)
	}
	if v, ok := resolvePath(ctx, ".y"); !ok || v.(float64) != 2.0 {
		t.Fatalf("resolve .y => %v ok=%v", v, ok)
	}
	if v, ok := resolvePath(ctx, "$.x"); !ok || v.(float64) != 1.0 {
		t.Fatalf("resolve $.x => %v ok=%v", v, ok)
	}
}

func TestNextSeg(t *testing.T) {
	// имя
	if seg, tail := nextSeg("foo.bar"); seg != "foo" || tail != "bar" {
		t.Fatalf("nextSeg name: seg=%q tail=%q", seg, tail)
	}
	// индекс
	if seg, tail := nextSeg("[10].rest"); seg != "[10]" || tail != "rest" {
		t.Fatalf("nextSeg index: seg=%q tail=%q", seg, tail)
	}
}
