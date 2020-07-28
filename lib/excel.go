package lib

import (
	"os"
	"strings"
	"sync"

	"github.com/360EntSecGroup-Skylar/excelize"
)

// デリミタ  ex) "## hoge ##"
const delimiter = "##"
const specialKeyword = "keys"

// GetBookByPath パスで指定した Excel ファイルが存在しない場合は新規作成、
// 存在する場合は読み込む
func GetBookByPath(path string) (*excelize.File, error) {
	if _, err := os.Stat(path); err != nil {
		// ファイルが存在しない場合、新規作成
		f := excelize.NewFile()
		err := f.SaveAs(path)
		return f, err
	}
	// すでにファイルが存在する場合は読込
	return excelize.OpenFile(path)
}

// GetSheetByName 指定した名前のシートの存在を確認
// 指定した名前のシートが存在しないかつ、 "create" が true だった場合新規作成
func GetSheetByName(f *excelize.File, name string, create bool) (existed bool) {
	sheetList := f.GetSheetList()
	if i := Contains(sheetList, name); i == -1 {
		if create {
			f.NewSheet(name)
		}
		return false
	}
	return true
}

// Contains スライスに要素が存在するか判定
// 存在する場合はそのインデックス、存在しない場合は −1 を返す
func Contains(s []string, e string) int {
	for i, v := range s {
		if v == e {
			return i
		}
	}
	return -1
}

// IsKeyword デリミタを持つキーワードかどうか判定
// デリミタを持っている場合、デリミタ部分を削除しキーワードのみを返す
func IsKeyword(word string) (string, bool) {
	// 前後の空白文字を削除
	w := strings.TrimSpace(word)
	if strings.HasPrefix(w, delimiter) && strings.HasSuffix(w, delimiter) {
		return strings.TrimSpace(w[2 : len(w)-2]), true
	}
	return word, false
}

// ReplaceKeyToVal キーをデータに置き換える
// (*) goroutine で実行する
func ReplaceKeyToVal(f *excelize.File, d *Data, wg *sync.WaitGroup, c chan bool) {
	defer func() {
		<-c
	}()

	fc := func(sheet string) {
		rows, _ := f.GetRows(sheet)
		for r, row := range rows {
			for c, val := range row {
				key, ok := IsKeyword(val)
				if ok {
					axis, _ := excelize.CoordinatesToCellName(c+1, r+1)
					v := d.GetValByKey(key)
					if v != nil {
						f.SetCellValue(sheet, axis, v)
						f.CalcCellValue(sheet, axis)
					}
				}
			}
		}
	}

	sheets := f.GetSheetList()
	for _, s := range sheets {
		fc(s)
	}
	wg.Done()
}

// GetKeyBaseCell "## keys ##" があるセルを基準としてデータ定義を
// 取得するための基準セルの座標を取得
func GetKeyBaseCell(f *excelize.File, sheet string) (string, error) {
	rows, err := f.GetRows(sheet)
	if err != nil {
		return "", err
	}
	for r, row := range rows {
		for c, v := range row {
			if key, _ := IsKeyword(v); key == specialKeyword {
				return excelize.CoordinatesToCellName(c+1, r+1)
			}
		}
	}
	return "", nil
}

func SaveWithA1(f *excelize.File) error {
	ss := f.GetSheetList()
	for _, s := range ss {
		f.GetCellValue(s, "A1")
	}
	return f.Save()
}
