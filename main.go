package main

import (
	"encoding/json"
	"flag"
	"fmt"
	"log"
	"sort"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

var (
	// intFlag  int
	pathFlag string

// boolFlag bool
)

func main() {
	//define the arguments for command line
	// flag.IntVar(&intFlag, "int", 1234, "help message")
	flag.StringVar(&pathFlag, "path", "default", "help message")
	// flag.BoolVar(&boolFlag, "bool", false, "help message")
	flag.Parse()

	// fmt.Println("intFlag value is: ", intFlag)
	// fmt.Println("strFlag value is: ", strFlag)
	// fmt.Println("boolFlag value is: ", boolFlag)

	//get the file
	// filePath := "C:/Users/abrah/rekap_fmis/rekap_wsl.xlsx"
	filePath := pathFlag
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		fmt.Println(err)
		return
	}
	defer func() {
		// Close the spreadsheet.
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	//get table data

	rows, err := f.GetRows("Sheet2")
	if err != nil {
		fmt.Println(err)
		return
	}

	//unmarshal excel json into rawData struct
	var unformatedData AllRawData

	resp, err := excelToJson(rows)
	if err != nil {
		log.Fatal(err)
	}
	err = json.Unmarshal(resp, &unformatedData.AllData)
	if err != nil {
		log.Fatal(err)
	}

	//call the struct method to fix the format before grouping the data by the dbh
	formated := unformatedData.FormatRow()
	// for i, v := range formated.AllData {
	// 	fmt.Printf("%v -->  %+v \n", i, v)
	// }

	//call the struct method to group the data by the dbh and count the sum and height average
	// grouped := formated.GroupByDHB()
	// for i, v := range grouped.AllData {
	// 	fmt.Printf("%v -->  %+v: \n", i, v)
	// }

	//init []convertedData struct to save data in excel
	var AllConvertedData []ConvertedData

	groupedData := formated.GroupBy()

	// Count total stem and the average of height
	for dbh, stem := range groupedData {
		mainStemCount := 0
		var totalHeight float64
		if dbh == 0 {
			for _, row := range groupedData[dbh] {
				AllConvertedData = append(AllConvertedData, ConvertedData{
					FeatID:  row.FeatID,
					OldComp: row.OldComp,
					Survey:  7,
					PlotNo:  row.PlotNo,
					Remark:  row.Remark,
				})
			}
		} else {
			allRemarks := make([]string, 0, len(stem))
			for j, p := range stem {
				if p.ThirdStem == 0 {
					mainStemCount++
				}
				totalHeight += p.Height
				if stem[j].Remark != "" {
					allRemarks = append(allRemarks, stem[j].Remark)
				}
			}
			secondStemCount := len(stem) - mainStemCount
			heightAvg := totalHeight / float64(len(stem))
			AllConvertedData = append(AllConvertedData, ConvertedData{
				FeatID:     stem[0].FeatID,
				OldComp:    stem[0].OldComp,
				Survey:     7,
				PlotNo:     stem[0].PlotNo,
				DBH:        dbh,
				MainStem:   mainStemCount,
				SecondStem: secondStemCount,
				HT1:        roundFloat(heightAvg, 1),
				Remark:     strings.Join(allRemarks, ", "),
			})
		}
		sort.Slice(AllConvertedData, func(i, j int) bool {
			return AllConvertedData[i].DBH < AllConvertedData[j].DBH
		})
	}

	//save file
	sheetName := "Sheet1"

	rows, err = f.GetRows(sheetName)
	if err != nil {
		log.Fatal(err)
	}

	// f.SetCellValue(sheetName, "A1", "Feat ID")
	// f.SetCellValue(sheetName, "B1", "Old Comp")
	// f.SetCellValue(sheetName, "C1", "Survey")
	// f.SetCellValue(sheetName, "D1", "Plot No")
	// f.SetCellValue(sheetName, "E1", "DBH")
	// f.SetCellValue(sheetName, "F1", "Main St")
	// f.SetCellValue(sheetName, "G1", "Sec St")
	// f.SetCellValue(sheetName, "H1", "Fallen St")
	// f.SetCellValue(sheetName, "I1", "Sick St")
	// f.SetCellValue(sheetName, "J1", "DR")
	// f.SetCellValue(sheetName, "K1", "DR Freq")
	// f.SetCellValue(sheetName, "L1", "Dead St")
	// f.SetCellValue(sheetName, "M1", "HT1")
	// f.SetCellValue(sheetName, "N1", "HT2")
	// f.SetCellValue(sheetName, "O1", "Planted_date")
	// f.SetCellValue(sheetName, "P1", "PHI_survey")
	// f.SetCellValue(sheetName, "Q1", "Plan Harvesting")
	// f.SetCellValue(sheetName, "R1", "Remark")

	// for i, v := range AllConvertedData {
	// 	if v.DBH == 0 {
	// 		AllConvertedData = append(AllConvertedData[:i], AllConvertedData[:i+1]...)
	// 		AllConvertedData = append(AllConvertedData, v)
	// 	}
	// }

	//sorting dbh that == 0 to be on the last rows
	indexCount := 0
	for _, v := range AllConvertedData {
		if v.DBH == 0 {
			indexCount++
		}
	}
	temp := AllConvertedData[indexCount:]
	AllConvertedData = append(temp, AllConvertedData[:indexCount]...)

	// for i, v := range AllConvertedData {
	// 	fmt.Printf("%+v --> %+v \n", i, v)
	// }

	for i, data := range AllConvertedData {
		row := len(rows) + i + 1 // Excel rows start from 1, so offset by 2
		if data.DBH > 0 {
			f.SetCellValue(sheetName, fmt.Sprintf("A%d", row), data.FeatID)
			f.SetCellValue(sheetName, fmt.Sprintf("B%d", row), data.OldComp)
			f.SetCellValue(sheetName, fmt.Sprintf("C%d", row), data.Survey)
			f.SetCellValue(sheetName, fmt.Sprintf("D%d", row), data.PlotNo)
			f.SetCellValue(sheetName, fmt.Sprintf("E%d", row), data.DBH)
			f.SetCellValue(sheetName, fmt.Sprintf("F%d", row), data.MainStem)
			f.SetCellValue(sheetName, fmt.Sprintf("G%d", row), data.SecondStem)
			f.SetCellValue(sheetName, fmt.Sprintf("M%d", row), data.HT1)
			f.SetCellValue(sheetName, fmt.Sprintf("R%d", row), data.Remark)
		} else {
			f.SetCellValue(sheetName, fmt.Sprintf("A%d", row), data.FeatID)
			f.SetCellValue(sheetName, fmt.Sprintf("B%d", row), data.OldComp)
			f.SetCellValue(sheetName, fmt.Sprintf("C%d", row), data.Survey)
			f.SetCellValue(sheetName, fmt.Sprintf("D%d", row), data.PlotNo)
			f.SetCellValue(sheetName, fmt.Sprintf("E%d", row), "-")
			f.SetCellValue(sheetName, fmt.Sprintf("F%d", row), "-")
			f.SetCellValue(sheetName, fmt.Sprintf("G%d", row), "-")
			f.SetCellValue(sheetName, fmt.Sprintf("R%d", row), data.Remark)
		}
	}

	f.SetActiveSheet(len(rows) + len(AllConvertedData))

	if err := f.SaveAs(filePath); err != nil {
		log.Fatal(err)
	}
	fmt.Printf("Excel berhasil di save di : %s ", filePath)
}

func excelToJson(rows [][]string) ([]byte, error) {
	headers := rows[0]
	slice := make([]map[string]interface{}, 0)
	for _, row := range rows[1:] {
		tempData := make(map[string]interface{}, 0)
		for j, cellValue := range row {
			tempData[headers[j]] = cellValue
			if headers[j] == "Main Stem" || headers[j] == "Second Stem" || headers[j] == "Third Stem" || headers[j] == "Fourth Stem" || headers[j] == "Height" || headers[j] == "Dead Stem" || headers[j] == "Second Height" || headers[j] == "Third Height" || headers[j] == "Fourth Height" {
				if cellValue == "" || cellValue == "-" {
					tempData[headers[j]] = 0.0
				} else {
					convVal, err := strconv.ParseFloat(cellValue, 64)
					if err != nil {
						fmt.Println(err)
						continue
					}
					tempData[headers[j]] = convVal
				}
			}
		}
		slice = append(slice, tempData)
	}

	jsonData, err := json.Marshal(slice)
	if err != nil {
		return nil, err
	}

	return jsonData, nil
}
