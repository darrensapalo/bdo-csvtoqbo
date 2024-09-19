## Banco de Oro - Business Online Banking XLS to Quickbooks-compatible CSV

This tool transforms an XLS file downloadable from BDO's business online 
banking tool into a quickbooks-compatible CSV.

Note: no raw data in the `raw/` folder is provided. You must provide your own
data.

## Capabilities

✅ Renders the excel data into a better data format that is easier to work with.
Renders the excel data into a quickbooks-compliant transaction file.

## Usage

_Note: Banco de Oro does not save files as `xlsx` and instead saves files using `xls`.
You will need to convert the file from `xls` to `xlsx` via "File > Save as" on Microsoft Excel._

Currently runs with `go run main.go` and loads data from `raw/sample.xlsx` to 
renders output into`output.xlsx`.


## Roadmap

1. Implement configurable inputs and outputs (Go viper)
2. Define scope of basic unit tests
3. Provide sample raw data with synthetic data/fake information
4. Quickbooks-compliant transaction file structure