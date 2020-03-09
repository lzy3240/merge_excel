package main

import (
	"fmt"
	"io/ioutil"
	"os"
	"strconv"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize"
	config "github.com/Unknwon/goconfig"
	"github.com/lzy3240/mlog"
)

var log *mlog.Logger

// init log
func init() {
	log = mlog.Newlog("info", "./logs/", "Merge_excel_", "day", 1024)
}

//msg
func msg() {
	fmt.Println("******************")
	fmt.Println(" XX报表合并工具XX")
	fmt.Println(" Author : Lzy")
	fmt.Println(" Version :1.0.0")
	fmt.Println("******************")
	fmt.Println(" 按 Ctrl+C 退出")
	fmt.Println()
}

//checkErr 检查错误
func checkErr(err error) {
	if err != nil {
		log.Error("error:%v..", err)
	}
}

// rxlsx 读取excel，返回*[][]string切片指针，减少内存消耗
func rxlsx(fileName, sheetName string) (v *[][]string) {
	//sourceDate := make([][]string, 0)
	f, err := excelize.OpenFile(fileName)
	checkErr(err)

	rows, err := f.GetRows(sheetName)
	checkErr(err)
	/*for _, row := range rows {
		sourceDate = append(sourceDate, row)
		//fmt.Println(row)
		//fmt.Printf("%T\n", row)
		//for _, value := range row {
			//fmt.Printf("\t%s", value)
		//}
		//fmt.Println()
	}
	//fmt.Println(sourceDate)*/ //遍历二维切片
	return &rows
}

// wxlsx 写入目标表，传入*[][]string 切片指针
func wxlsx(filName string, sourceDate *[][]string) {
	var arrange = [...]string{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ"}
	_, err := os.Stat(filName + ".xlsx") //判断目标文件是否存在
	//文件不存在时，新建并写入
	if err != nil {
		log.Info("文件[%v]不存在，准备创建...", filName+".xlsx")
		xlsx := excelize.NewFile() //新建excel
		index := xlsx.NewSheet("Sheet1")
		for i, x := range *sourceDate {
			for j, y := range x {
				xlsx.SetCellValue("Sheet1", arrange[j]+strconv.Itoa(i+1), y) //按单元格写入值，按照列模板定位单元格
			} //end for
		} //end for
		xlsx.SetActiveSheet(index)            //激活sheet
		err := xlsx.SaveAs(filName + ".xlsx") //保存表
		if err != nil {
			log.Error("文件[%s]写入失败，失败原因：%s", filName+".xlsx", err)
		} else {
			log.Info("文件[%s]已创建，并写入成功...", filName+".xlsx")
		}
		//文件存在时，在已有文件内追加
	} else {
		xlsx, err := excelize.OpenFile(filName + ".xlsx") //打开已存在的表
		index := xlsx.GetSheetIndex("Sheet1")
		if err != nil {
			log.Error("打开文件[%s]异常，错误为：%s", filName+".xlsx", err)
		}
		rows, _ := xlsx.GetRows("Sheet1") //取得表内当前行，大量数据时资源消耗大
		//fmt.Println(rows, len(rows)) //34行
		for a := 1; a < len(*sourceDate); a++ { //a：去除原切片中的表头行
			for j, y := range (*sourceDate)[a] {
				xlsx.SetCellValue("Sheet1", arrange[j]+strconv.Itoa(a+len(rows)), y) //按单元格写入值，按照列模板定位单元格，
			} //endfor
		} //endfor
		xlsx.SetActiveSheet(index)           //激活sheet
		err = xlsx.SaveAs(filName + ".xlsx") //保存表
		if err != nil {
			log.Error("文件[%s]写入失败，失败原因：%s", filName+".xlsx", err)
		} else {
			log.Info("文件[%s]写入成功", filName+".xlsx")
		}
	}
}

//文件夹遍历，返回[]string切片
func listDir(folder string) (fslice []string) {
	files, _ := ioutil.ReadDir(folder)
	for _, file := range files {
		if !file.IsDir() {
			fslice = append(fslice, file.Name())
		}
	}
	return fslice
}

// func 主入口
func main() {
	msg()
	//配置读取
	cfg, err := config.LoadConfigFile("./config.ini")
	checkErr(err)
	values, err := cfg.GetValue("title", "name")
	checkErr(err)
	title := strings.Split(values, ",")
	//原表格目录
	dir, _ := os.Getwd()
	fileList := listDir(dir + "/excel")
	// 主循环
	for _, f := range fileList {
		log.Info("开始处理文件[%s]...", f)
		// title循环
		for _, g := range title {
			tmp := rxlsx(dir+"/excel/"+f, g)
			wxlsx(g, tmp)
		} //endfor
		log.Info("文件[%s]处理完成...", f)
	} //endfor
	fmt.Println()
	fmt.Println("[按[Ctrl+C]或[Enter]键结束...]")
	var s string
	fmt.Scanln(&s)
}
