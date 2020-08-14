// Pixel2Excel project main.go
package main

import (
	"fmt"
	"image"
	"image/color"

	"errors"
	"image/jpeg"
	"image/png"
	"os"

	"strings"
	"time"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/tealeg/xlsx"
)

func SetXlsCellSize(picWidth, picHeight int, xlsxFilePath string) {

	xlsxFile, err := excelize.OpenFile(xlsxFilePath)
	if nil != err {
		panic(err)
	}

	for rowNo := 1; rowNo < (1 + picHeight); rowNo++ {
		xlsxFile.SetRowHeight("Sheet1", rowNo, 6)
	}

	for colNo := 1; colNo < (1 + picWidth); colNo++ {
		colName, err := excelize.ColumnNumberToName(colNo)
		if nil != err {
			panic(err)
		}
		xlsxFile.SetColWidth("Sheet1", colName, colName, 1)
	}

	if err := xlsxFile.Save(); err != nil {
		panic(err)
	}
}

func main() {
	xlsFilePath := ".\\test.xlsx"
	picFilePath := ".\\test.jpg"

	args := os.Args
	argc := len(args)

	switch argc {
	case 2, 3:
		if !strings.HasSuffix(args[1], ".jpg") && !strings.HasSuffix(args[1], ".png") {
			panic(errors.New("Invalid command format"))
		}
		picFilePath = args[1]

		if argc == 3 && !strings.HasSuffix(args[2], ".xlsx") {
			xlsFilePath = args[2]
		} else {
			xlsFilePath = picFilePath[:len(picFilePath)-4] + ".xlsx"
		}

	default:
		fmt.Println("Usage: Pixel2Excel xxx.jpg|xxx.png")
		return
	}
	fmt.Printf("pic = %s xlsx = %s\r\n", picFilePath, xlsFilePath)

	picFile, err := os.Open(picFilePath)
	if err != nil {
		panic(err)
	}

	defer picFile.Close()

	var img image.Image
	var picErr error
	if strings.HasSuffix(picFilePath, ".jpg") {
		img, picErr = jpeg.Decode(picFile)
	} else {
		img, picErr = png.Decode(picFile)
	}
	if picErr != nil {
		panic(err)
	}
	rect := img.Bounds()
	picWidth, picHeight := rect.Dx(), rect.Dy()

	fmt.Printf("picture width = %v height = %v\r\n", picWidth, picHeight)

	xlsxFile := xlsx.NewFile()

	t1 := time.Now()

	sheet, err := xlsxFile.AddSheet("Sheet1")
	if err != nil {
		panic(err)
	}

	//fgColorStr := "FF00FF00"
	for rowNo := 0; rowNo < (0 + picHeight); rowNo++ {
		for colNo := 0; colNo < (0 + picWidth); colNo++ {
			cell := sheet.Cell(rowNo, colNo)

			if nil != err {
				panic(err)
			}
			pixColor := img.At(colNo, rowNo)
			colorValue := color.NRGBAModel.Convert(pixColor).(color.NRGBA)
			fgColorStr := strings.ToUpper(fmt.Sprintf("%.2x%.2x%.2x", colorValue.R, colorValue.G, colorValue.B))

			cellStyle := cell.GetStyle()
			cellStyle.Fill.FgColor = fgColorStr
			cellStyle.Fill.PatternType = "solid"
			cellStyle.ApplyFill = true
			cell.SetStyle(cellStyle)
		}
	}

	fmt.Printf("Elapse pixel = %v\r\n", time.Since(t1).Seconds())
	t2 := time.Now()
	fmt.Println("Saving xlsx......")
	if err := xlsxFile.Save(xlsFilePath); err != nil {
		panic(err)
	}
	fmt.Printf("Elapse save = %v\r\n", time.Since(t2).Seconds())

	SetXlsCellSize(picWidth, picHeight, xlsFilePath)

	fmt.Println("Create excel pixel picture successfully.")
}
