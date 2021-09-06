package main

import (
	"fmt"
	"github.com/tealeg/xlsx"
	"strconv"
	"strings"
)

type Row struct {
	Snumber     string //序号
	Title       string //title
	Content     string //content
	Keywords    string //keywords
	Description string //description

	SheetName string // 时间
}
type ora struct {
	corp  string //单位
	name  string //业务系统
	name2 string //进程名
	v1000 string //10点值
	v2000 string //20点值
	H     string // 最大值
	L     string // 最小值
	A     string // 平均值
	TIME  string // 时间
}

const DESC_SCALE = 0.7

/**对Excel文件数据做批量处理：
1、A列是序号
2、B列是标题
3、C列是内容
4、D列是关键词
5、E列是摘要
根据B列标题来合并数据
假设B2跟B3的相似度>70%，则两个合并；如果<70%，则继续向下匹配；两个单元格数据匹配上后，不再参与循环；
新Excel文档：
1、新A列，为：A2+A3
2、新B列，为：B2_B3
3、新C列，为：B2+C2+B3+C3；其中B2和B3用h2标签包装，即<h2>B2</h2>
4、新D列，为：D2
5、新E列，为：E2
6、所有未被匹配到的数据，单独导出一份Excel
*/
func main() {
	fmt.Printf("START")
	excelFileName := "D:\\merge\\采集数据.xlsx"
	rows := readXlsx(excelFileName)
	rowRow, ints := analyzeData(rows)
	fmt.Println(fmt.Sprintf("%v,%v", rowRow, ints))

	mergeRows, noMergeRows := getMergeRowsAndNoMergeRows(rows, ints, rowRow)
	writingXlsx2(mergeRows, "D:\\merge\\合并导出.xlsx")
	writingXlsx2(noMergeRows, "D:\\merge\\未合并导出.xlsx")

}

func getMergeRowsAndNoMergeRows(rows []Row, ints []int, rowRow map[int]int) ([]Row, []Row) {

	noMergeRows := []Row{}
	for i, row := range rows {

		if !exclude(ints, i) {
			noMergeRows = append(noMergeRows, row)
		}
	}
	mergeRows := []Row{{
		Snumber:     rows[0].Snumber,
		Title:       rows[0].Title,
		Content:     rows[0].Content,
		Keywords:    rows[0].Keywords,
		Description: rows[0].Description,
		SheetName:   rows[0].SheetName,
	}}
	for row2, row3 := range rowRow {
		mergeRow := Row{}
		row2 := rows[row2]
		row3 := rows[row3]
		mergeRow.Snumber = fmt.Sprintf("%v+%v", row2.Snumber, row3.Snumber)
		mergeRow.Title = fmt.Sprintf("%v_%v", row2.Title, row3.Title)
		mergeRow.Content = fmt.Sprintf("<h2>%v</h2>+%v+<h2>%v</h2>+%v", row2.Title, row2.Content, row3.Title, row3.Content)
		mergeRow.Keywords = row2.Keywords
		mergeRow.Description = row2.Description
		mergeRows = append(mergeRows, mergeRow)
	}
	return mergeRows, noMergeRows
}

func readXlsx(filename string) []Row {
	var listOra []Row
	xlFile, err := xlsx.OpenFile(filename)
	if err != nil {
		fmt.Printf("open failed: %s\n", err)
		panic("未找到文件")
	}
	for _, sheet := range xlFile.Sheets {

		//fmt.Printf("Sheet Name: %s\n", sheet.Name)
		tmpOra := Row{}
		// 获取标签页(时间)
		//tmpOra.TIME = sheet.Name
		for _, row := range sheet.Rows {

			var strs []string

			for _, cell := range row.Cells {
				text := cell.String()
				strs = append(strs, text)
			}
			// 获取标签页(时间)
			tmpOra.SheetName = sheet.Name
			tmpOra.Snumber = strs[0]
			tmpOra.Title = strs[1]
			tmpOra.Content = strs[2]
			tmpOra.Keywords = strs[3]
			tmpOra.Description = strs[4]
			listOra = append(listOra, tmpOra)
		}
	}
	return listOra
}

func writingXlsx2(rows []Row, fileName string) {
	var file *xlsx.File
	var sheet *xlsx.Sheet
	//var row *xlsx.Row
	var cell *xlsx.Cell
	var err error
	if len(rows) == 0 {
		return
	}
	firsRow := rows[0]
	file = xlsx.NewFile()
	sheet, err = file.AddSheet(firsRow.SheetName)
	if err != nil {
		fmt.Printf(err.Error())
	}
	//row = sheet.AddRow()
	//row.SetHeightCM(0.5)
	//
	//cell = row.AddCell()
	//cell.Value = firsRow.Snumber
	//cell = row.AddCell()
	//cell.Value = firsRow.Title
	//cell = row.AddCell()
	//cell.Value = firsRow.Content
	//cell = row.AddCell()
	//cell.Value = firsRow.Keywords
	//cell = row.AddCell()
	//cell.Value = firsRow.Description

	for _, r := range rows {
		var row *xlsx.Row
		row = sheet.AddRow()
		row.SetHeightCM(0.5)
		cell = row.AddCell()
		cell.Value = r.Snumber
		cell = row.AddCell()
		cell.Value = r.Title
		cell = row.AddCell()
		cell.Value = r.Content
		cell = row.AddCell()
		cell.Value = r.Keywords
		cell = row.AddCell()
		cell.Value = r.Description
	}
	err = file.Save(fileName)
	if err != nil {
		fmt.Printf(err.Error())
	}
}

/**对Excel文件数据做批量处理：
1、A列是序号
2、B列是标题
3、C列是内容
4、D列是关键词
5、E列是摘要
根据B列标题来合并数据
假设B2跟B3的相似度>70%，则两个合并；如果<70%，则继续向下匹配；两个单元格数据匹配上后，不再参与循环；
新Excel文档：
1、新A列，为：A2+A3
2、新B列，为：B2_B3
3、新C列，为：B2+C2+B3+C3；其中B2和B3用h2标签包装，即<h2>B2</h2>
4、新D列，为：D2
5、新E列，为：E2
6、所有未被匹配到的数据，单独导出一份Excel
假设b2有10个字，在b3找到其中7个，就匹配上
*/
func analyzeData(rows []Row) (map[int]int, []int) {
	rowRow := make(map[int]int)
	aready := []int{}
	if len(rows) == 0 {
		return map[int]int{}, []int{}
	}

	for i := 1; i < len(rows); i++ {
		if exclude(aready, i) {
			continue
		}
		row := rows[i]

		for j := i + 1; j < len(rows); j++ {
			if exclude(aready, i) {
				continue
			}
			if calculateScale(row.Title, rows[j].Title) {
				rowRow[i] = j
				aready = append(aready, i)
				aready = append(aready, j)
			}
		}
	}
	return rowRow, aready
}

//排除已经合并的元素
func exclude(aready []int, i int) bool {
	for _, r := range aready {
		if i == r {
			return true
		}
	}
	return false
}

func calculateScale(title1, title2 string) bool {
	s1 := strings.Split(title1, "")
	s2 := strings.Split(title2, "")
	if len(s2) == 0 {
		return false
	}
	equalStr := []string{}
	for _, r1 := range s1 {
		for _, r2 := range s2 {
			if r1 == r2 {
				equalStr = append(equalStr, r1)
				break
			}
		}
	}
	rate := (len(equalStr) * 10) / len(s1)
	if rate >= DESC_SCALE*10 {
		return true
	}
	return false
}

func writingXlsx(oraList []ora) {
	var file *xlsx.File
	var sheet *xlsx.Sheet
	var row *xlsx.Row
	var cell *xlsx.Cell
	var err error

	file = xlsx.NewFile()
	sheet, err = file.AddSheet("Sheet1")
	if err != nil {
		fmt.Printf(err.Error())
	}
	row = sheet.AddRow()
	row.SetHeightCM(0.5)
	cell = row.AddCell()
	cell.Value = "单位"
	cell = row.AddCell()
	cell.Value = "业务系统"
	cell = row.AddCell()
	cell.Value = "进程名"
	cell = row.AddCell()
	cell.Value = "V1000"
	cell = row.AddCell()
	cell.Value = "V2000"
	cell = row.AddCell()
	cell.Value = "H"
	cell = row.AddCell()
	cell.Value = "L"
	cell = row.AddCell()
	cell.Value = "A"
	cell = row.AddCell()
	cell.Value = "TIME"

	for _, i := range oraList {
		if i.corp == "单位" {
			continue
		}

		// 判断是否为-9999，是的变为0.0
		var row1 *xlsx.Row
		if i.v1000 == "-9999" {
			i.v1000 = "0.0"
		}
		if i.v2000 == "-9999" {
			i.v2000 = "0.0"
		}
		if i.H == "-9999" {
			i.H = "0.0"
		}
		if i.L == "-9999" {
			i.L = "0.0"
		}

		row1 = sheet.AddRow()
		row1.SetHeightCM(0.5)

		cell = row1.AddCell()
		cell.Value = i.corp
		cell = row1.AddCell()
		cell.Value = i.name
		cell = row1.AddCell()
		cell.Value = i.name2

		// 判断值是大于7200，大于变成红色
		v1, _ := strconv.ParseFloat(i.v1000, 64)
		if v1 > 7200 {
			cell = row1.AddCell()
			cell.Value = i.v1000
			cell.GetStyle().Font.Color = "00FF0000"
		} else {
			cell = row1.AddCell()
			cell.Value = i.v1000
		}

		//v2, _ := strconv.Atoi(i.v2000)
		v2, _ := strconv.ParseFloat(i.v2000, 64)
		if v2 > 7200 {
			cell = row1.AddCell()
			cell.Value = i.v2000
			cell.GetStyle().Font.Color = "00FF0000"
		} else {
			cell = row1.AddCell()
			cell.Value = i.v2000
		}

		//vH, _ := strconv.Atoi(i.H)
		vH, _ := strconv.ParseFloat(i.H, 64)
		if vH > 7200 {
			cell = row1.AddCell()
			cell.Value = i.H
			cell.GetStyle().Font.Color = "00FF0000"
		} else {
			cell = row1.AddCell()
			cell.Value = i.H
		}

		//vL, _ := strconv.Atoi(i.L)
		vL, _ := strconv.ParseFloat(i.L, 64)
		if vL > 7200 {
			cell = row1.AddCell()
			cell.Value = i.L
			cell.GetStyle().Font.Color = "00FF0000"
		} else {
			cell = row1.AddCell()
			cell.Value = i.L

		}

		//vA, _ := strconv.Atoi(i.A)
		vA, _ := strconv.ParseFloat(i.A, 64)
		if vA > 7200 {
			cell = row1.AddCell()
			cell.Value = i.A
			cell.GetStyle().Font.Color = "00FF0000"
		} else {
			cell = row1.AddCell()
			cell.Value = i.A
		}

		// 打印时间
		cell = row1.AddCell()
		cell.Value = i.TIME
	}

	err = file.Save("2019-_-_-2019-_-_Lag延时数据.xlsx")
	if err != nil {
		fmt.Printf(err.Error())
	}
}
