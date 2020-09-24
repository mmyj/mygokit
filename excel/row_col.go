package excel

import (
	"log"
	"reflect"

	"github.com/fatih/structtag"
)

const colTag = "col"

type ColMate struct {
	v   interface{}
	val string
	tag string
}

func (k ColMate) Val() string {
	return k.val
}

func (k ColMate) Tag() string {
	return k.tag
}

type RowMate struct {
	structMap map[uintptr]*ColMate
	columnMap map[uintptr]*ColMate
}

func (m *RowMate) init(i interface{}) {
	rv := reflect.ValueOf(i).Elem()
	rt := reflect.TypeOf(i).Elem()

	for i := 0; i < rt.NumField(); i++ {
		fv := rv.Field(i)
		field := rt.Field(i)
		tags, err := structtag.Parse(string(field.Tag))
		if err != nil {
			continue
		}

		tag, err := tags.Get(colTag)
		if err != nil {
			continue
		}

		it := &ColMate{
			val: field.Name,
			tag: tag.Name,
			v:   fv,
		}

		switch field.Type.Kind() {
		case reflect.Struct:
			if fv.Addr().CanInterface() {
				m.structMap[fv.UnsafeAddr()] = it
				m.init(fv.Addr().Interface())
			}
		default:
			m.columnMap[fv.UnsafeAddr()] = it
		}
	}
}

func (m *RowMate) Column(i interface{}) ColMate {
	if reflect.TypeOf(i).Kind() != reflect.Ptr {
		log.Fatalln("NEED PTR")
	}
	v := reflect.ValueOf(i).Elem()
	addr := v.UnsafeAddr()
	retMap := m.columnMap
	if v.Kind() == reflect.Struct {
		retMap = m.structMap
	}
	if retMap[addr] == nil {
		log.Fatalln("NO FIELD")
	}
	return *retMap[addr]
}

func ReflectRomMate(i interface{}) *RowMate {
	m := new(RowMate)
	m.structMap = make(map[uintptr]*ColMate)
	m.columnMap = make(map[uintptr]*ColMate)

	if reflect.TypeOf(i).Elem().Kind() != reflect.Struct &&
		reflect.TypeOf(i).Kind() != reflect.Ptr {
		log.Fatalln("NEED A PTR TO STRUCT")
	}

	m.init(i)
	return m
}

func GetAllColumnName(i interface{}) (ret []string) {
	m := ReflectRomMate(i)
	for _, c := range m.columnMap {
		ret = append(ret, c.tag)
	}
	return
}

func GetAllColumnNameInterface(i interface{}) (ret []interface{}) {
	m := ReflectRomMate(i)
	for _, c := range m.columnMap {
		ret = append(ret, c.tag)
	}
	return
}

func GetAllColumnValue(i interface{}) (ret []interface{}) {
	m := ReflectRomMate(i)
	for _, c := range m.columnMap {
		ret = append(ret, c.v)
	}
	return
}
