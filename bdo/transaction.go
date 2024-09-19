package bdo

import (
	"fmt"
	"io"
	"log"
	"math/big"
	"strings"

	"github.com/tealeg/xlsx/v3"
)

// TransactionHistory stores the parsed header data.
type TransactionHistory struct {
	Corporation        string
	RequestedDate 		 string
	PeriodCovered      string
	AccountAlias       string
	AccountNumber      string
	Currency           string
	AccountName        string
	Transactions			 []Transaction
}

// Transaction ...
type Transaction struct {
	PostingDate string
	Branch string
	Description string
	Debit big.Float
	Credit big.Float
	RunningBalance big.Float
	CheckNumber string
}

func (t Transaction) String() string {
	return fmt.Sprintf("posting date: %s; branch: %s; description: %s; debit: %s; credit: %s; running balance: %s; check number: %s", 
		t.PostingDate, t.Branch, t.Description, t.Debit.String(), t.Credit.String(), t.RunningBalance.String(), t.CheckNumber,
	)
}

// IsValid ...
func (t Transaction) IsValid() bool {
	return t.PostingDate != "" && t.Branch != "" && t.Description != "" && t.CheckNumber != ""
}

// GetCellAtSheet ... 
func GetCellAtSheet(sheet *xlsx.Sheet, rowIdx, colIdx int) (cell *xlsx.Cell, err error) {
	row, err := sheet.Row(rowIdx)
	if err != nil {
		return
	}

	cell = row.GetCell(colIdx)
	return
}

var bdoCoords map[string]coordinates = map[string]coordinates{
	"Requested Date": []int{0, 6, 4},
	"Corporation": []int{0, 8, 4},
	"Period Covered": []int{0, 9, 4},
	"Account Alias": []int{0, 10, 4},
	"Account Number": []int{0, 11, 4},
	"Currency": []int{0, 12, 4},
	"Account Name": []int{0, 13, 4},
	"StartData": []int{0, 15, 1},
}

// LoadHeaders ...
func (t *TransactionHistory) LoadHeaders(workbook *xlsx.File) error {
	src, err := workbook.ToSliceUnmerged()

	if err != nil {
		return err
	}

	t.RequestedDate = bdoCoords["Requested Date"].valueFrom(src)
	t.Corporation = bdoCoords["Corporation"].valueFrom(src)
	t.PeriodCovered = bdoCoords["Period Covered"].valueFrom(src)
	t.AccountAlias = bdoCoords["Account Alias"].valueFrom(src)
	t.AccountNumber = bdoCoords["Account Number"].valueFrom(src)
	t.Currency = bdoCoords["Currency"].valueFrom(src)
	t.AccountName = bdoCoords["Account Name"].valueFrom(src)
	return nil
}

func getTransactionLineItemCoords(row int, column string) coordinates {
	switch column {
	case "posting_date":
		return []int{0, row, 1}
	case "branch":
		return []int{0, row, 3}
	case "description":
		return []int{0, row, 6}
	case "debit":
		return []int{0, row, 7}
	case "credit":
		return []int{0, row, 8}
	case "running_balance":
		return []int{0, row, 9}
	case "check_number":
		return []int{0, row, 11}
	}
	return []int{0, row, 1}
}

// LoadSections ...
func (t *TransactionHistory) LoadSections(workbook *xlsx.File) error {
	src, err := workbook.ToSliceUnmerged()

	if err != nil {
		log.Printf("Error slicing unmerged: %+v\n", err)
		return err
	}

	startSection := bdoCoords["StartData"]
	startRow := startSection[1]

	for i := startRow; i < len(src[0]); i++ {
		debitCoords := getTransactionLineItemCoords(i, "debit")
		debitRaw := strings.ReplaceAll(debitCoords.valueFrom(src), ",", "")
		debit, _, err := big.ParseFloat(debitRaw, 10, 30, big.ToNearestEven)
		if err != nil { 
			if err == io.EOF {
				debit = big.NewFloat(0)
			} else {
				log.Printf("Error; failed to parse debit: %+v", err)
				continue
			}
		}
		
		
		creditCoords := getTransactionLineItemCoords(i, "credit")
		creditRaw := strings.ReplaceAll(creditCoords.valueFrom(src), ",", "")
		credit, _, err := big.ParseFloat(creditRaw, 10, 30, big.ToNearestEven)
		if err != nil { 
			if err == io.EOF {
				credit = big.NewFloat(0)
			} else {
				log.Printf("Error; failed to parse credit: %+v", err)
				continue
			}
		}
		
		runningBalanceCoords := getTransactionLineItemCoords(i, "running_balance")
		runningBalanceRaw := strings.ReplaceAll(runningBalanceCoords.valueFrom(src), ",", "")
		runningBalance, _, err := big.ParseFloat(runningBalanceRaw, 10, 30, big.ToNearestEven)
		if err != nil { 
			if err == io.EOF {
				runningBalance = big.NewFloat(0)
			} else {
				log.Printf("Error; failed to parse running balance: %+v", err)
				continue
			}
		}

		log.Printf("runningBalance = %+v\n", runningBalance.String())

		// Skip if running balance is not filled up with a non-zero number
		nonZeroRunningBalance := runningBalance.Cmp(big.NewFloat(0)) == 0
		postingDate := getTransactionLineItemCoords(i, "posting_date").valueFrom(src)
		dateIsSet := postingDate != ""
		log.Printf("Date is set = %+v; nonZeroRunningBalance = %+v\n", dateIsSet, nonZeroRunningBalance)
		if dateIsSet && nonZeroRunningBalance { continue }

		description := getTransactionLineItemCoords(i, "description").valueFrom(src)
		description = strings.ReplaceAll(description, "  ", " ")
		description = strings.ReplaceAll(description, "  ", " ")

		txn := Transaction{
			PostingDate: postingDate,
			Branch: getTransactionLineItemCoords(i, "branch").valueFrom(src),
			Description: description,
			Debit: *debit,
			Credit: *credit,
			RunningBalance: *runningBalance,
			CheckNumber: getTransactionLineItemCoords(i, "check_number").valueFrom(src),
		}

		if txn.IsValid() {
			t.Transactions = append(t.Transactions, txn)
		}
	}

	return nil
}