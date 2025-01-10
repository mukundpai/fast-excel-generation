# Excel Creation with Go

### Goal
The goal was to export 1 million rows from a PostgreSQL database to an Excel file efficiently. Through several iterations, I reduced the execution time to approximately 35 seconds.

## How to Run
  - Make sure you have a PostgreSQL database with the test_db schema

  - Install Go Lang
  - Make sure you have the correct   credentials in the main.go file
  - install the dependencies

```bash
go mod tidy
```
  -- run the program
```bash
go run main.go
```

### Performance Achievement
- Successfully exported 1 million rows to Excel in ~38 seconds
- Processing speed: ~26,316 rows per second
- Memory-efficient streaming approach

### Optimization Journey

1. **Initial Approach (Crashed)**
   - Basic row-by-row processing
   - Single-threaded execution
   - Individual cell writes
   ```go
   for rows.Next() {
       // Process one row at a time
       f.SetCellValue("Sheet1", cell, value)
   }
   ```

2. **First Optimization (76s)**
   - Added Excel streaming mode
   - Basic batch processing
   - Database connection pooling
   ```go
   streamWriter, err := f.NewStreamWriter("Sheet1")
   db.SetMaxOpenConns(25)
   db.SetMaxIdleConns(25)
   ```

3. **Final Optimization (38-44s)**
   ```go
   --Used Channels for concurrent processing
   --Used Goroutines for concurrent writing
   --Used Pool for memory management
   // Channel setup for concurrent processing
   batchChan := make(chan [][]interface{}, 5)
   errorChan := make(chan error, 1)
   doneChan := make(chan bool)

   // Writer goroutine
   go func() {
       for batch := range batchChan {
           for _, row := range batch {
               cell, _ := excelize.CoordinatesToCellName(1, rowNum)
               streamWriter.SetRow(cell, row)
               rowNum++
           }
       }
       doneChan <- true
   }()
   ```

### Implementation Methods

1. **Database Connection Setup**
   ```go
   // Initialize database connection with optimized settings
   db, err := sql.Open("postgres", connStr)
   db.SetMaxOpenConns(25)
   db.SetMaxIdleConns(25)
   ```

2. **Excel Stream Writer Setup**
   ```go
   // Create Excel file with streaming mode
   f := excelize.NewFile()
   streamWriter, err := f.NewStreamWriter("Sheet1")
   ```

3. **Batch Processing Implementation**
   ```go
   // Process rows in batches
   const batchSize = 5000
   currentBatch := make([][]interface{}, 0, batchSize)
   
   for rows.Next() {
       // Add row to current batch
       currentBatch = append(currentBatch, rowValues)
       
       if len(currentBatch) >= batchSize {
           batchChan <- currentBatch
           currentBatch = make([][]interface{}, 0, batchSize)
       }
   }
   ```

4. **Concurrent Writing with Goroutines**
   ```go
   // Writer goroutine for Excel operations
   go func() {
       for batch := range batchChan {
           for _, row := range batch {
               cell, _ := excelize.CoordinatesToCellName(1, rowNum)
               streamWriter.SetRow(cell, row)
               rowNum++
           }
       }
       doneChan <- true
   }()
   ```

5. **Channel Coordination**
   ```go
   // Channel setup
   batchChan := make(chan [][]interface{}, 5)
   errorChan := make(chan error, 1)
   doneChan := make(chan bool)

   // Wait for completion or error
   select {
   case err := <-errorChan:
       log.Fatal(err)
   case <-doneChan:
       // Writer completed successfully
   }
   ```

6. **Memory Management**
   ```go
   // Pre-allocate slices
   values := make([]interface{}, len(columns))
   valuePtrs := make([]interface{}, len(columns))
   for i := range values {
       valuePtrs[i] = &values[i]
   }
   ```

### Key Performance Features

1. **Concurrent Processing**
   - Separate goroutine for Excel writing
   - Main thread focuses on database reading
   - Channel-based communication

2. **Batch Processing**
   - Optimal batch size of 10000 rows
   - Pre-allocated slices
   - Reduced memory allocations

3. **Database Optimization**
   - Connection pool settings
   - Efficient query execution

4. **Excel Streaming**
   - Memory-efficient writing
   - Continuous data flow

## Why Go Over Python?

1. **Performance**
   - Go's concurrent processing with goroutines
   - Efficient memory management
   - Faster execution speed
   - Better handling of large datasets

2. **Memory Efficiency**
   - Streaming data processing
   - Lower memory footprint
   - Garbage collection optimization

3. **Concurrency**
   - Built-in concurrency with goroutines
   - Channel-based communication
   - Easy-to-implement parallel processing

4. **Type Safety**
   - Compile-time type checking
   - Reduced runtime errors
   - Better code reliability

## Libraries Used
- `database/sql` - Database operations
- `github.com/lib/pq` - PostgreSQL driver
- `github.com/xuri/excelize/v2` - Excel operations

## Future Improvements
1. Add support for multiple sheets
2. Add data formatting options
3. Support for different data types