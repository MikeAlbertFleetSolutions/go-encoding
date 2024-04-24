package xlsx

import (
	"log"
	"os"
	"testing"
	"time"
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
	type subone struct {
		OneField string `xls:"OneField"`
	}
	type subtwo struct {
		TwoField string `xls:"TwoField"`
		SubOne   subone
	}
	aas := []struct {
		Number    int       `xls:"Number,{\"number_format\":2}"`
		Name      string    `xls:"Name"`
		TimeField time.Time `xls:"Requested Date,{\"number_format\":14}"`
		Sub       subtwo
	}{
		{1, "a", time.Now(), subtwo{TwoField: "aa", SubOne: subone{OneField: "aaa"}}},
		{2, "b", time.Now().Add(2 * time.Hour), subtwo{TwoField: "bb", SubOne: subone{OneField: "bbb"}}},
		{3, "c", time.Now().Add(3 * 24 * time.Hour), subtwo{TwoField: "cc", SubOne: subone{OneField: "ccc"}}},
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
