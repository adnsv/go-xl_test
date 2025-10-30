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
	row.Height = 30
	cell11 := row.AddCell()
	cell11.XF.Alignment.Vertical = "top"
	cell11.SetStr("col1")
	cell12 := row.AddCell()
	cell12.XF.Alignment.Vertical = "center"
	cell12.SetStr("col2")
	cell13 := row.AddCell()
	cell13.XF.Alignment.Vertical = "bottom"
	cell13.SetStr("col3")

	row = sheet.AddRow()
	row.Height = 30
	cell21 := row.AddCell()
	cell21.XF.Alignment.Vertical = "center"
	cell21.SetInt(1)
	cell22 := row.AddCell()
	cell22.XF.Alignment.Vertical = "center"
	cell22.SetInt(2)
	cell23 := row.AddCell()
	cell23.XF.Alignment.Vertical = "center"
	cell23.SetInt(3)

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

	// Add a row to demonstrate merged cells
	row = sheet.AddRow()
	row.Height = 40
	mergedCell := row.AddCell()
	mergedCell.SetStr("This cell is merged across columns A-C")
	mergedCell.XF.Alignment.Horizontal = "center"
	mergedCell.XF.Alignment.Vertical = "center"
	row.AddCell() // Empty cell that will be part of the merge
	row.AddCell() // Empty cell that will be part of the merge

	// Merge cells using string reference
	err = sheet.Merge("A4:C4")
	if err != nil {
		log.Fatal(err)
	}

	// Add another row to demonstrate MergeRange
	row = sheet.AddRow()
	row.Height = 40
	mergedCell2 := row.AddCell()
	mergedCell2.SetStr("Merged using MergeRange")
	mergedCell2.XF.Alignment.Horizontal = "center"
	mergedCell2.XF.Alignment.Vertical = "center"
	row.AddCell() // Empty cell

	// Merge cells using coordinate-based API (columns 1-2, row 5)
	err = sheet.MergeRange(1, 5, 2, 5)
	if err != nil {
		log.Fatal(err)
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
