package xlsx

import (
	"log"
	"os"
	"testing"
)

func TestMain(m *testing.M) {
	// show file & location, date & time
	log.SetFlags(log.Ldate | log.Ltime | log.Lshortfile)

	exitVal := m.Run()

	os.Exit(exitVal)
}

func TestCreateWorkbook(t *testing.T) {
	// output to xlsx
	x := NewXlsx()

	// create the sheet and get rid of default sheet
	err := x.CreateSheet("test")
	if err != nil {
		log.Printf("%+v", err)
		return
	}
	err = x.RemoveSheet("Sheet1")
	if err != nil {
		log.Printf("%+v", err)
		return
	}

	// write out some data
	aas := []struct {
		Number int    `xls:"Number,{\"number_format\":1}"`
		Name   string `xls:"Name"`
	}{
		{1, "a"},
		{2, "b"},
		{3, "c"},
	}
	for _, aa := range aas {
		err = x.WriteRow("test", aa)
		if err != nil {
			log.Printf("%+v", err)
			return
		}
	}

	// save to file
	err = x.Close("test.xlsx")
	if err != nil {
		log.Printf("%+v", err)
		return
	}
}
