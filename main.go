package main

import (
	"fmt"
	"os"
	"syscall"
	"unsafe"

	"github.com/xuri/excelize/v2"
)

// 将字符串转换为 *uint16
// func stringToUTF16Ptr(s string) *uint16 {
// 	ptr, _ := syscall.StringToUTF16Ptr(s)
// 	return ptr
// }
func IntPtr(n int) uintptr {
	return uintptr(n)
}
func StrPtr(s string) uintptr {
	return uintptr(unsafe.Pointer(syscall.StringToUTF16Ptr(s)))
}

func showMessageBox(tittle, text string) {
	user32dll, _ := syscall.LoadLibrary("user32.dll")
	user32 := syscall.NewLazyDLL("user32.dll")
	MessageBoxW := user32.NewProc("MessageBoxW")
	MessageBoxW.Call(IntPtr(0), StrPtr(text), StrPtr(tittle), IntPtr(0))
	defer syscall.FreeLibrary(user32dll)
}

// 显示消息框
// func showMessageBox(title, message string, flags uint32) {
// 	winapi.MessageBox(0, stringToUTF16Ptr(message), stringToUTF16Ptr(title), flags)
// }

func mergeExcels(outputFile string, inputFiles []string) error {
	// 创建输出的Excel文件
	outFile, err := excelize.OpenFile(inputFiles[0])
	if err != nil {
		return fmt.Errorf("无法打开文件 %s: %v", inputFiles[0], err)
	}

	// 获取第一个工作表
	sheetName := outFile.GetSheetName(0)

	// 获取表头
	rows, err := outFile.GetRows(sheetName)
	if err != nil {
		return fmt.Errorf("无法获取表头: %v", err)
	}
	currentRowCount := len(rows) + 1

	// 从第二个文件开始合并
	for _, inputFile := range inputFiles[1:] {
		// 打开当前Excel文件
		currFile, err := excelize.OpenFile(inputFile)
		if err != nil {
			return fmt.Errorf("无法打开文件 %s: %v", inputFile, err)
		}

		rows, err := currFile.Rows(sheetName)
		if err != nil {
			fmt.Println(err)
			return fmt.Errorf("无法获取行数据: %v", err)
		}

		// 从第二行开始添加数据（跳过表头）
		rows.Next()
		for rows.Next() {
			// 获取当前行的位置
			rowIndex := currentRowCount
			currentRowCount++

			// 使用 SetSheetRow 将一行数据赋值到目标位置
			cell := fmt.Sprintf("A%d", rowIndex) // 假设从 A 列开始
			row, _ := rows.Columns()
			err := outFile.SetSheetRow(sheetName, cell, &row)
			rowOpts := rows.GetRowOpts()
			outFile.SetRowStyle(sheetName, rowIndex, rowIndex, rowOpts.StyleID)
			if err != nil {
				return fmt.Errorf("无法设置行数据: %v", err)
			}
		}
	}

	// 保存合并后的文件
	if err := outFile.SaveAs(outputFile); err != nil {
		return fmt.Errorf("保存文件失败: %v", err)
	}

	return nil
}

func main() {
	args := os.Args[1:]
	if len(args) < 2 {
		showMessageBox("错误", "请拖拽至少两个Excel文件到该程序上执行")
		return
	}

	outputFile := "merged_output.xlsx"
	err := mergeExcels(outputFile, args)
	if err != nil {
		showMessageBox("错误", fmt.Sprintf("文件合并失败: %v", err))
		return
	}

	showMessageBox("成功", fmt.Sprintf("文件合并成功! 输出文件为: %s", outputFile))
}
