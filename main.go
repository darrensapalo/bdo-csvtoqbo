package main

import (
	"fmt"
	"log"
	"math/big"
	"os"

	"github.com/darrensapalo/bdo-csvtoqbo/bdo"
	"github.com/tealeg/xlsx/v3"
)

// Create a new defaultStyle
var defaultStyle = xlsx.NewStyle()

func init() {
	// Set the font to Arial and size to 12
	defaultStyle.Font.Name = "Arial"
	defaultStyle.Font.Size = 11
}

func main() {
	// Open the xlsx file
	filePath := "raw/sample.xlsx"
	if _, err := os.Stat(filePath); os.IsNotExist(err) {
		log.Fatalf("File %s does not exist", filePath)
	}

	workbook, err := xlsx.OpenFile(filePath)
	if err != nil {
		log.Fatalf("Failed to open the xlsx file: %s", err)
	}

	if len(workbook.Sheets) != 1 {
		log.Fatalf("Expected exactly one worksheet; want 1, got %d", len(workbook.Sheets))
	}

	if err != nil {
		log.Fatalf("Failed to open the xlsx file: %s", err)
	}

	history := bdo.TransactionHistory{}

	err = history.LoadHeaders(workbook)
	if err != nil {
		log.Fatalf("Failed to load headers; %v", err)
	}

	err = history.LoadSections(workbook)
	if err != nil {
		log.Fatalf("Failed to load sections; %v", err)
	}

	log.Printf("Transaction history loaded; %+v\n", history)
	for i := 0; i < len(history.Transactions); i++ {
		fmt.Printf(" - %s\n", history.Transactions[i].String())
	}

	th := history

	// Create a new XLSX file
	newWorkbook := xlsx.NewFile()

	sheet, err := newWorkbook.AddSheet("Transactions")
	if err != nil {
		log.Fatalf("Failed to create sheet: %s", err)
	}

	// Add first sheet for headers
	headerSheet, err := newWorkbook.AddSheet("Account Information")
	if err != nil {
		log.Fatalf("Failed to create sheet: %s", err)
	}

	sheet.SetColWidth(1, 1, 16)
	sheet.SetColWidth(2, 2, 36)
	sheet.SetColWidth(3, 3, 64)
	sheet.SetColWidth(4, 7, 16)

	// Write transactions to the file
	writeTransactions(sheet, th.Transactions)

	// Write headers for the file
	writeHeaders(headerSheet, &th)

	// Save the file
	err = newWorkbook.Save("raw/output.xlsx")
	if err != nil {
		log.Fatalf("Failed to save file: %s", err)
	}

	log.Println("XLSX file created successfully!")

	return
}

// writeHeaders writes the main headers for the transaction history into the sheet.
func writeHeaders(sheet *xlsx.Sheet, th *bdo.TransactionHistory) {
	row := sheet.AddRow()
	c := row.AddCell()
	c.SetString("Corporation:")
	c.SetStyle(defaultStyle)
	c = row.AddCell()
	c.SetString(th.Corporation)
	c.SetStyle(defaultStyle)

	row = sheet.AddRow()
	c = row.AddCell()
	c.SetString("Requested Date:")
	c.SetStyle(defaultStyle)
	c = row.AddCell()
	c.SetString(th.RequestedDate)
	c.SetStyle(defaultStyle)

	row = sheet.AddRow()
	c = row.AddCell()
	c.SetString("Period Covered:")
	c.SetStyle(defaultStyle)
	c = row.AddCell()
	c.SetString(th.PeriodCovered)
	c.SetStyle(defaultStyle)

	row = sheet.AddRow()
	c = row.AddCell()
	c.SetString("Account Alias:")
	c.SetStyle(defaultStyle)
	c = row.AddCell()
	c.SetString(th.AccountAlias)
	c.SetStyle(defaultStyle)

	row = sheet.AddRow()
	c = row.AddCell()
	c.SetString("Account Number:")
	c.SetStyle(defaultStyle)
	c = row.AddCell()
	c.SetString(th.AccountNumber)
	c.SetStyle(defaultStyle)

	row = sheet.AddRow()
	c = row.AddCell()
	c.SetString("Currency:")
	c.SetStyle(defaultStyle)
	c = row.AddCell()
	c.SetString(th.Currency)
	c.SetStyle(defaultStyle)

	row = sheet.AddRow()
	c = row.AddCell()
	c.SetString("Account Name:")
	c.SetStyle(defaultStyle)
	c = row.AddCell()
	c.SetString(th.AccountName)
	c.SetStyle(defaultStyle)

	// Add a blank row before the transaction records
	sheet.AddRow()
}

// writeTransactions writes the transaction records into the sheet.
func writeTransactions(sheet *xlsx.Sheet, transactions []bdo.Transaction) {
	decimal := "#,##0.00"

	// Write the header row for transactions
	headerRow := sheet.AddRow()
	headerStyle := xlsx.NewStyle()
	headerStyle.Font.Bold = true
	headerStyle.Font.Name = "Arial"
	headerStyle.Font.Size = 11

	h := headerRow.AddCell()
	h.SetStyle(headerStyle)
	h.SetString("Posting Date")

	h = headerRow.AddCell()
	h.SetStyle(headerStyle)
	h.SetString("Branch")

	h = headerRow.AddCell()
	h.SetStyle(headerStyle)
	h.SetString("Description")

	h = headerRow.AddCell()
	h.SetStyle(headerStyle)
	h.SetString("Debit")

	h = headerRow.AddCell()
	h.SetStyle(headerStyle)
	h.SetString("Credit")

	h = headerRow.AddCell()
	h.SetStyle(headerStyle)
	h.SetString("Running Balance")

	h = headerRow.AddCell()
	h.SetStyle(headerStyle)
	h.SetString("Check Number")

	// Write each transaction row
	for _, t := range transactions {
		row := sheet.AddRow()
		cell := row.AddCell()
		cell.SetString(t.PostingDate)
		cell.SetStyle(defaultStyle)
		
		cell = row.AddCell()
		cell.SetString(t.Branch)
		cell.SetStyle(defaultStyle)
		
		cell = row.AddCell()
		cell.SetString(t.Description)
		cell.SetStyle(defaultStyle)

		d, _ := RoundUp(&t.Debit).Float64()
		cell = row.AddCell()
		cell.SetFloatWithFormat(d, "0.00")
		cell.NumFmt = decimal
		cell.SetStyle(defaultStyle)

		c, _ := RoundUp(&t.Credit).Float64()
		cell = row.AddCell()
		cell.SetFloatWithFormat(c, "0.00")
		cell.NumFmt = decimal
		cell.SetStyle(defaultStyle)

		rb, _ := RoundUp(&t.RunningBalance).Float64()
		cell = row.AddCell()
		cell.SetFloatWithFormat(rb, "0.00")
		cell.NumFmt = decimal
		cell.SetStyle(defaultStyle)

		if t.CheckNumber != "000000000" {
			cell = row.AddCell()
			cell.SetString(t.CheckNumber)
			cell.SetStyle(defaultStyle)
		}
	}
}

// RoundUp rounds up a big.Float to 2 decimal places.
func RoundUp(f *big.Float) *big.Float {
	scale := big.NewFloat(100) // Scaling factor for two decimal places

	// Multiply by 100 to shift decimal point
	fScaled := new(big.Float).Mul(f, scale)

	// Get the integer part (floor value)
	fInt, _ := fScaled.Int(nil)

	// Convert the integer back to a big.Float
	fIntFloat := new(big.Float).SetInt(fInt)

	// If fScaled is not equal to fIntFloat, we need to round up
	if fScaled.Cmp(fIntFloat) != 0 {
		fIntFloat.Add(fIntFloat, big.NewFloat(1)) // Round up
	}

	// Divide by 100 to scale it back to 2 decimal places
	result := new(big.Float).Quo(fIntFloat, scale)
	return result
}
