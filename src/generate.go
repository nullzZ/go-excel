/*
@Author: nullzz
@Date: 2021/12/22 2:38 下午
@Version: 1.0
@DEC:
*/
package src

import (
	"errors"
	"fmt"
	"github.com/tealeg/xlsx"
	"io/ioutil"
	"path"
	"strings"
	"unicode"
)

const (
	LineNumber = 4 //每个sheet开始行数
)

var (
	ErrorSourcePath = errors.New("ReadExcel sourcePath nil")
	ErrorToPath     = errors.New("ReadExcel stoPath nil")
	ErrorMaxRow     = errors.New("ReadExcel row<4")
)

/**
 * 将excel中的前四列转化为struct
 * 第一列字段类型		如 int
 * 第二列字段名称
 * 第三列字段名		如 id
 * 第四列s,c,all 	s表示服务端使用 c表示客户端使用 all表示都使用
 */
type GenerateExcel struct {
	toPath     string //存储路径
	sourcePath string //源路径
}

func (g *GenerateExcel) ReadFile(sourcePath, toPath string) error {
	if sourcePath == "" {
		return ErrorSourcePath
	}
	if toPath == "" {
		return ErrorToPath
	}

	files, err := ioutil.ReadDir(sourcePath)
	if err != nil {
		return err
	}
	for _, file := range files {
		if path.Ext(file.Name()) != ".xlsx" {
			continue
		}
		f, err := xlsx.OpenFile(sourcePath + "\\" + file.Name())
		if err != nil {
			return err
		}
		for _, sheet := range f.Sheets {
			if g.isContinue(sheet) { //跳过sheet
				continue
			}

			if sheet.MaxRow < LineNumber {
				fmt.Errorf("ReadFile sheet.MaxRow < 4 sheet=%s", sheet.Name)
				return ErrorMaxRow
			}
			// 遍历列
			for i := 0; i < sheet.MaxCol; i++ {
				// 判断某一列的第一行是否为空
				c := sheet.Cell(0, i)
				if c.Value == "" {
					continue
				}
				cellData := make([]string, 0)
				// 遍历行
				for j := 0; j < LineNumber; j++ {
					c := sheet.Cell(j, i)
					cellData = append(cellData, c.Value)
				}
				//sheetData = append(sheetData, cellData)
				for _, cc := range cellData {
					fmt.Println("cellData=", cc)
				}
			}

		}
	}
	return nil
}

func (GenerateExcel) isContinue(sheet *xlsx.Sheet) bool {
	if strings.Index(sheet.Name, "Sheet") != -1 { //排除默认sheet
		return true
	}
	for _, v := range []rune(sheet.Name) { //排除汉子
		if unicode.Is(unicode.Han, v) {
			return true
		}
	}
	return false
}
