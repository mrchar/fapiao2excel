package main

import (
	"encoding/json"
	"errors"
	"fmt"
	"io/ioutil"
	"log"
	"os"
	"path/filepath"
	"strconv"
	"strings"

	"github.com/tealeg/xlsx/v3"
)

const (
	magicNumber = "01"
)

var fapiaoKindMap = map[string]string{
	"10": "增值税电子普通发票",
	"04": "增值税普通发票",
	"01": "增值税专用发票",
}

type fapiao struct {
	Kind         string  `json:"发票类型"` // 发票类型
	Code         string  `json:"发票代码"` //发票代码
	SerialNumber string  `json:"发票号码"` // 发票号码
	Amount       float64 `json:"总额"`   // 金额
	Date         string  `json:"时间"`   // 时间
	Checkcode    string  `json:"校验码"`  // 校验码
}

func parse(input string) (*fapiao, error) {
	var err error
	// 拆分字符串
	var seq = strings.Split(input, ",")
	if len(seq) < 7 {
		return nil, errors.New("这不是 发票二维码的内容，因为长度不足")
	}
	if seq[0] != magicNumber {
		return nil, errors.New("这不是发票二维码的内容")
	}

	var f = &fapiao{}

	// 获取发票类型
	var ok bool
	f.Kind, ok = fapiaoKindMap[seq[1]]
	if !ok {
		return nil, fmt.Errorf("找不到%s对应的发票类型", seq[1])
	}

	// 发票代码
	f.Code = seq[2]
	// 发票号码
	f.SerialNumber = seq[3]
	// 发票金额
	f.Amount, err = strconv.ParseFloat(seq[4], 64)
	if err != nil {
		return nil, err
	}
	// 开票日期
	// timestamp, err := strconv.ParseInt(seq[5], 10, 64)
	// if err != nil {
	// 	return nil, err
	// }
	// f.date = time.Unix(timestamp, 0)
	f.Date = seq[5]
	// 校验码
	f.Checkcode = seq[6]
	return f, nil
}

func export(name string, fapiaos ...*fapiao) error {
	wb := xlsx.NewFile()
	sh, err := wb.AddSheet("发票")
	if err != nil {
		return err
	}

	defer sh.Close()

	// 表头
	header := sh.AddRow()
	header.SetHeight(12)
	for _, content := range []string{"发票类型", "发票代码", "发票号码", "金额", "时间", "校验码"} {
		cell := header.AddCell()
		cell.SetString(content)
	}

	for _, fp := range fapiaos {
		row := sh.AddRow()
		row.SetHeight(12)
		row.AddCell().SetString(fp.Kind)
		row.AddCell().SetString(fp.Code)
		row.AddCell().SetString(fp.SerialNumber)
		row.AddCell().SetFloat(fp.Amount)
		row.AddCell().SetString(fp.Date)
		row.AddCell().SetString(fp.Checkcode)
	}

	return wb.Save(name)
}

func init() {
	log.SetFlags(log.Lshortfile)
	myStyle := xlsx.NewStyle()
	myStyle.Alignment.Horizontal = "right"
	myStyle.Fill.FgColor = "FFFFFF00"
	myStyle.Fill.PatternType = "solid"
	myStyle.Font.Name = "Georgia"
	myStyle.Font.Size = 11
	myStyle.Font.Bold = true
	myStyle.ApplyAlignment = true
	myStyle.ApplyFill = true
	myStyle.ApplyFont = true
}

func main() {
	fmt.Println("请使用扫码枪扫描二维码，结束扫描请按回车键：")
	var fapiaos = make([]*fapiao, 0)

	for {
		var input string
		if _, err := fmt.Scanln(&input); err != nil && err.Error() != "unexpected newline" {
			log.Fatal(err)
		}

		if input == "" {
			break
		}

		fp, err := parse(input)
		if err != nil {
			log.Fatal(err)
		}
		// 保存扫描后的发票
		if buf, err := json.Marshal(fp); err != nil {
			log.Println(err)
		} else {
			ioutil.WriteFile(filepath.Join("data", fp.Code), buf, os.FileMode(0644))
		}

		fapiaos = append(fapiaos, fp)
	}

	if err := export("./发票.xlsx", fapiaos...); err != nil {
		log.Fatal(err)
	}
}
