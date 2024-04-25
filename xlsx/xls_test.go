package xlsx

import (
	"log"
	"os"
	"testing"
	"time"
)

type CustomTime time.Time

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

	t1 := time.Now()
	t2 := time.Now().Add(2 * time.Hour)
	t3 := time.Now().Add(3 * 24 * time.Hour)
	ct1 := CustomTime(t1)
	ct2 := CustomTime(t2)
	ct3 := CustomTime(t3)

	aas := []struct {
		Number  int         `xls:"Number,{\"number_format\":2}"`
		Name    string      `xls:"Name"`
		Date    time.Time   `xls:"Date,{\"number_format\":14}"`
		PDate   *time.Time  `xls:"Pointer to Date,{\"number_format\":14}"`
		MyDate  CustomTime  `xls:"MyDate,{\"number_format\":14}"`
		PMyDate *CustomTime `xls:"Pointer to MyDate,{\"number_format\":14}"`
		Sub     subtwo
	}{
		{1, "a", t1, &t1, ct1, &ct1, subtwo{TwoField: "aa", SubOne: subone{OneField: "aaa"}}},
		{2, "b", t1, &t2, ct2, &ct2, subtwo{TwoField: "bb", SubOne: subone{OneField: "bbb"}}},
		{3, "c", t1, &t3, ct3, &ct3, subtwo{TwoField: "cc", SubOne: subone{OneField: "ccc"}}},
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
