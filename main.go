package main

import (
	"fmt"
	"log"
	"os"
	"path/filepath"

	"github.com/adnsv/go-xl/xl"
)

func main() {

	fmt.Printf("Generating XLSX file\n")

	wb := xl.NewWorkbook()
	wb.AppName = "My App"

	sheet, err := wb.AddSheet("sheet1")
	if err != nil {
		log.Fatal(err)
	}

	sheet.SetColumnWidth(3, 40)

	row := sheet.AddRow()
	row.AddCell().SetStr("col1")
	row.AddCell().SetStr("col2")
	row.AddCell().SetStr("col3")

	row = sheet.AddRow()
	row.AddCell().SetInt(1)
	row.AddCell().SetInt(2)
	row.AddCell().SetInt(3)

	row = sheet.AddRow()
	row.Height = 64
	{
		fn := "./testdata/image1.png"
		blob, err := os.ReadFile(fn)
		if err != nil {
			log.Fatal(err)
		}
		row.AddCell().SetPicture(&xl.PictureInfo{
			Extension: filepath.Ext(fn),
			Blob:      blob,
		})
	}
	{
		fn := "./testdata/image2.jpeg"
		blob, err := os.ReadFile(fn)
		if err != nil {
			log.Fatal(err)
		}
		row.AddCell().SetPicture(&xl.PictureInfo{
			Extension: filepath.Ext(fn),
			Blob:      blob,
		})
	}

	debugdir := "./testdata/dbg"
	fmt.Printf("Writing file parts into %s\n", debugdir)
	ds := xl.NewDirStorage("./testdata/dbg")
	dw := xl.NewWriter(ds)
	err = dw.Write(wb)
	if err != nil {
		log.Fatal(err)
	}

	outfn := "./testdata/dbg.xlsx"
	fmt.Printf("Writing %s\n", outfn)
	f, err := os.Create(outfn)
	if err != nil {
		log.Fatal(err)
	}
	defer f.Close()

	zs := xl.NewZipStorage(f)
	zw := xl.NewWriter(zs)
	defer zs.Close()
	zw.Write(wb)

	fmt.Printf("Mission Accomplished\n")
}
