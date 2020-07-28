package lib

import (
	"io"
	"os"
	"strconv"
)

// Data ファイル名とデータのマッピング
type Data struct {
	Filename string
	Params   map[string]string
}

// NewData Data の構造体を新規作成
func NewData(template, filename string) (*Data, error) {
	err := CopyFile(template, filename)
	data := Data{Filename: filename}
	return &data, err
}

// AddData データの追加
func (data *Data) AddData(key, value string) {
	if len(data.Params) == 0 {
		data.Params = map[string]string{key: value}
	} else {
		data.Params[key] = value
	}
}

// GetValByKey キーに対応するデータの値を取得
// int に変換できるものは、int に変換
func (data *Data) GetValByKey(key string) interface{} {
	v, ok := data.Params[key]
	// 存在しない key は nil を返す
	if !ok {
		return nil
	}

	// int に変換
	i, _ := strconv.Atoi(v)
	s := strconv.Itoa(i)
	if s == v {
		return i
	}

	// float64 に変換
	f, _ := strconv.ParseFloat(v, 64)
	s = strconv.FormatFloat(f, 'f', -1, 64)
	if s == v {
		return f
	}

	// string
	return v
}

// CopyFile ファイルのコピー
func CopyFile(src, dest string) error {
	s, err := os.Open(src)
	if err != nil {
		return err
	}
	defer s.Close()

	d, err := os.Create(dest)
	if err != nil {
		return err
	}
	defer d.Close()

	_, err = io.Copy(d, s)
	return err
}
