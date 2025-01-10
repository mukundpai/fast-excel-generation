package main

import (
	"database/sql"
	"fmt"
	"log"
	"time"

	_ "github.com/lib/pq"
	"github.com/xuri/excelize/v2"
)

func main() {
	startTime := time.Now()

	// Connect to the database
	connStr := "postgres://user:password@localhost:5432/test_db?sslmode=disable"
	db, err := sql.Open("postgres", connStr)
	if err != nil {
		log.Fatal(err)
	}
	defer db.Close()

	db.SetMaxOpenConns(25)
	db.SetMaxIdleConns(25)

	// Create a new Excel file
	f := excelize.NewFile()
	streamWriter, err := f.NewStreamWriter("Sheet1")
	if err != nil {
		log.Fatal(err)
	}

	// Use batch size for fetching data
	const batchSize = 10000
	query := `SELECT * FROM test_table LIMIT 1000000`

	rows, err := db.Query(query)
	if err != nil {
		log.Fatal(err)
	}
	defer rows.Close()

	// Get column names
	columns, err := rows.Columns()
	if err != nil {
		log.Fatal(err)
	}

	// Write headers
	headerRow := make([]interface{}, len(columns))
	for i, colName := range columns {
		headerRow[i] = colName
	}
	if err := streamWriter.SetRow("A1", headerRow); err != nil {
		log.Fatal(err)
	}

	// Create channels
	batchChan := make(chan [][]interface{}, 10)
	errorChan := make(chan error, 1)
	doneChan := make(chan bool)

	// Start single writer goroutine
	rowNum := 2
	go func() {
		for batch := range batchChan {
			for _, row := range batch {
				cell, _ := excelize.CoordinatesToCellName(1, rowNum)
				if err := streamWriter.SetRow(cell, row); err != nil {
					errorChan <- err
					return
				}
				rowNum++
			}
		}
		doneChan <- true
	}()

	// Process in batches
	currentBatch := make([][]interface{}, 0, batchSize)
	totalRows := 0
	values := make([]interface{}, len(columns))
	valuePtrs := make([]interface{}, len(columns))
	for i := range values {
		valuePtrs[i] = &values[i]
	}

	for rows.Next() {
		if err := rows.Scan(valuePtrs...); err != nil {
			log.Fatal(err)
		}

		rowValues := make([]interface{}, len(values))
		for i, v := range values {
			rowValues[i] = v
		}
		currentBatch = append(currentBatch, rowValues)
		totalRows++

		if len(currentBatch) >= batchSize {
			batchChan <- currentBatch
			currentBatch = make([][]interface{}, 0, batchSize)
		}
	}

	if len(currentBatch) > 0 {
		batchChan <- currentBatch
	}

	// Close channels and wait for completion
	close(batchChan)
	select {
	case err := <-errorChan:
		log.Fatal(err)
	case <-doneChan:
	}

	if err := streamWriter.Flush(); err != nil {
		log.Fatal(err)
	}

	if err := f.SaveAs("database_export.xlsx"); err != nil {
		log.Fatal(err)
	}

	totalElapsed := time.Since(startTime)
	fmt.Printf("\nProcess completed:\n")
	fmt.Printf("Total rows exported: %d\n", totalRows)
	fmt.Printf("Total execution time: %s\n", totalElapsed.Round(time.Second))
	fmt.Printf("Average rows per second: %.2f\n", float64(totalRows)/totalElapsed.Seconds())
}
