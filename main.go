package main

import (
	"fmt"
	"log"
	"sync"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/guromityan/go-excellent/lib"
	"gopkg.in/alecthomas/kingpin.v2"
)

const limit = 10

var (
	// コマンドオプション
	dataPath     = kingpin.Flag("data", "データが定義された Excel のパス").Short('d').Required().ExistingFile()
	sheetName    = kingpin.Flag("sheet", "データが定義されたシート名").Short('s').String()
	templatePath = kingpin.Flag("template", "テンプレートの Excel のパス").Short('t').Required().ExistingFile()
	concurrency  = kingpin.Flag("concurrency", fmt.Sprintf("同時平行処理数, デフォルト: %v", limit)).Short('c').Int()
)

func main() {
	// オプションをパース
	kingpin.Parse()

	if *concurrency == 0 {
		*concurrency = limit
	}

	// データ定義ファイルを読込
	path := *dataPath
	f, err := lib.GetBookByPath(path)
	if err != nil {
		log.Fatalln(err)
	}

	// シート名の指定がない場合は、"Sheet1" を読込
	if *sheetName == "" {
		*sheetName = "Sheet1"
	}

	co, err := lib.GetKeyBaseCell(f, *sheetName)
	if co == "" || err != nil {
		log.Fatalln("\"## keys ##\" is missing.")
		log.Fatalln(err)
	}

	// "## keys ##" があるセルの座標を取得
	bcol, brow, err := excelize.CellNameToCoordinates(co)
	if err != nil {
		log.Fatalln(err)
	}

	rows, err := f.GetRows(*sheetName)
	if err != nil {
		log.Fatalln(err)
	}

	// "## keys ##" より左側を除いたデータスライス
	newRows := rows[brow-1:]

	// Data を生成
	ds := make([]*lib.Data, 0)
	for r, row := range newRows {
		if r == 0 {
			continue
		}
		d, err := lib.NewData(*templatePath, row[bcol-1])
		if err != nil {
			log.Fatalln(err)
		}
		ds = append(ds, d)
	}

	// Data にパラメータを追加
	for c := bcol; len(newRows[0]) > c; c++ {
		key := ""
		for r := 0; len(newRows) > r; r++ {
			v := newRows[r][c]
			if r == 0 {
				key = v
			} else {
				ds[r-1].AddData(key, v)
			}
		}
	}

	// goroutine の上限を決めるためのセマフォ
	c := make(chan bool, *concurrency)
	// goroutine で値の置き換えを行う
	wg := sync.WaitGroup{}
	for _, d := range ds {
		c <- true
		wg.Add(1)
		f, err := lib.GetBookByPath(d.Filename)
		defer lib.SaveWithA1(f)
		if err != nil {
			log.Fatalln(err)
		}
		go lib.ReplaceKeyToVal(f, d, &wg, c)
	}
	wg.Wait()
}
