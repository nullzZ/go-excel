/*
@Author: nullzz
@Date: 2021/12/22 2:38 下午
@Version: 1.0
@DEC:
*/
package main

import (
	"flag"
	"go-excel/src"
)

func main() {
	sourcePath := flag.String("s", "", "source")
	toPath := flag.String("t", "", "source")
	src.GenerateExcel{}
}
