package main

import (
	"fmt"
	"log"
	"os"
	"sort"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

var (
	inputFile = "order.xlsx"
)

func init() {
	args := os.Args[1:]
	switch {
	case len(args) == 1:
		if strings.HasSuffix(args[0], "xlsx") {
			inputFile = args[0]
		}
	}
}

func main() {
	log.Println("1. read file: ", inputFile)
	f, err := excelize.OpenFile(inputFile)
	if err != nil {
		fmt.Println(err)
		return
	}
	// Get all the rows in the Sheet1.
	rows, err := f.GetRows("Sheet1")
	if err != nil {
		fmt.Println(err)
		return
	}
	log.Println("2. read sku")
	// sku id 对应的订单数
	skuOrderCountMap := make(map[string]int)
	// sku id 对应的订单集合
	skuOrderSetMap := make(map[string]map[string]int)
	// the set contains all order
	allOrderMap := make(map[string]int)
	for _, row := range rows {
		skuOrderCountMap[row[0]] += 1
		_, ok := skuOrderSetMap[row[0]]
		if ok {
			skuOrderSetMap[row[0]][row[1]] = 0
		} else {
			skuOrderSetMap[row[0]] = map[string]int{
				row[1]: 0,
			}
		}
		allOrderMap[row[1]] += 1
	}
	type skuOrderCount struct {
		skuID      string
		orderCount int
		ratio1     float64
		ratio2     float64
	}
	var skuOrderCounts []skuOrderCount
	for k, v := range skuOrderCountMap {
		skuOrderCounts = append(skuOrderCounts, skuOrderCount{
			skuID:      k,
			orderCount: v,
		})
	}
	// 根据每个 sku 对应的订单数倒序排列
	sort.SliceStable(skuOrderCounts, func(i, j int) bool {
		return skuOrderCounts[i].orderCount < skuOrderCounts[j].orderCount
	})

	// iterate sku list, from least order count to most order count
	log.Println("3. delete sku from sku list")
	totalOrder := len(allOrderMap)
	for i, cnt := range skuOrderCounts {
		for orderID := range skuOrderSetMap[cnt.skuID] {
			delete(allOrderMap, orderID)
		}
		skuOrderCounts[i].ratio1 = float64(len(allOrderMap)) / float64(totalOrder) * 100
		skuOrderCounts[i].ratio2 = float64(len(skuOrderSetMap[cnt.skuID])) / float64(totalOrder) * 100
	}

	// output
	log.Println("4. save result, total sku: ", len(skuOrderCounts), "total order:", totalOrder)
	nf := excelize.NewFile()
	// Create a new sheet.
	index := f.NewSheet("Sheet1")
	_ = nf.SetCellValue("Sheet1", "A1", "SKU ID")
	_ = nf.SetCellValue("Sheet1", "B1", "ORDER COUNT")
	_ = nf.SetCellValue("Sheet1", "C1", "RATIO 1")
	_ = nf.SetCellValue("Sheet1", "D1", "RATIO 2")
	// Set value of a cell.
	n := len(skuOrderCounts)
	for i := n - 1; i >= 0; i-- {
		x := n - i + 1
		_ = nf.SetCellValue("Sheet1", fmt.Sprintf("A%d", x), skuOrderCounts[i].skuID)
		_ = nf.SetCellValue("Sheet1", fmt.Sprintf("B%d", x), skuOrderCounts[i].orderCount)
		_ = nf.SetCellValue("Sheet1", fmt.Sprintf("C%d", x), skuOrderCounts[i].ratio1)
		_ = nf.SetCellValue("Sheet1", fmt.Sprintf("D%d", x), skuOrderCounts[i].ratio2)
	}
	// Set active sheet of the workbook.
	nf.SetActiveSheet(index)
	// Save spreadsheet by the given path.
	err = nf.SaveAs(fmt.Sprintf("Sku-%s.xlsx", time.Now().Format("2006-01-02 15-04-05")))
	if err != nil {
		fmt.Println(err)
	}
}
