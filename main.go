package main

import (
	"errors"
	"fmt"
	"log"
	"strconv"
	"strings"
	"time"

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
	kind         string    // 发票类型
	code         string    //发票代码
	serialNumber string    // 发票号码
	amount       float64   // 金额
	date         time.Time // 时间
	checkcode    string    // 校验码
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
	f.kind, ok = fapiaoKindMap[seq[1]]
	if !ok {
		return nil, fmt.Errorf("找不到%s对应的发票类型", seq[1])
	}

	// 发票代码
	f.code = seq[2]
	// 发票号码
	f.serialNumber = seq[3]
	// 发票金额
	f.amount, err = strconv.ParseFloat(seq[4], 64)
	if err != nil {
		return nil, err
	}
	// 开票日期
	timestamp, err := strconv.ParseInt(seq[5], 10, 64)
	if err != nil {
		return nil, err
	}
	f.date = time.Unix(timestamp, 0)
	// 校验码
	f.checkcode = seq[6]
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
		row.AddCell().SetString(fp.kind)
		row.AddCell().SetString(fp.code)
		row.AddCell().SetString(fp.serialNumber)
		row.AddCell().SetFloat(fp.amount)
		row.AddCell().SetDate(fp.date)
		row.AddCell().SetString(fp.checkcode)
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
		fapiaos = append(fapiaos, fp)
	}

	if err := export("./发票.xlsx", fapiaos...); err != nil {
		log.Fatal(err)
	}
}
