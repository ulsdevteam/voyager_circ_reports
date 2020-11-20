# Voyager Circ Reports

This console application pulls monthly circulation data from a Voyager Oracle database, and outputs it either to the console or to an Excel spreadsheet.

## Usage

This program requires the .Net Core runtime, version 3.1 or greater.

To run a report for March of 2020, you would use this syntax:

    dotnet run -- 3 2020 -l <library location code> -c <oracle connection string> -o report.xlsx

If the library location code is omitted, it will retrieve data for all locations. If the output file path is omitted, it will print the data to the console.

## License

Copyright University of Pittsburgh.

Freely licensed for reuse under the MIT License.