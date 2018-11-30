# XLSX2MSSQL
Reads in an XLSX spreadsheet, evaluates all the columns for datatypes, creates an MSSQL table, and imports all of the data into the SQL database.

Uses two libraries: NDesk.Options.0.2.1 for handling command line arguments and EPPlus.4.5.2.1 for reading excel data.

The connection string to the SQL Database can either be set in the app.config file or specified using the -c parameter.

Usage: XLSX2MSSQL [OPTIONS]+
Import a XLSX file into a MSSQL Database.

Options:
  -d, --drop                 Delete the table and then recreate
  -c, --connectionstring=VALUE
                             The connection string to connect to SQL with. If
                               not specified, the app.config file will be used.
  -f, --filepath=VALUE       The excel spreadsheet filepath
  -r, --rename=VALUE         rename a column in the format of oldcolumnnam-
                               e,newcolumnname
  -t, --tablename=VALUE      The SQL tablename to copy the excel file to
  -w, --worksheet=VALUE      The Excel worksheet name to read in
  -h, --help                 Show this message and exit


Example:

XLSX2MSSQL.exe -f "\\server\folder1\\MasterFile.xlsx" -t MasterTable -d

This will read in the MasterFile.xlsx file, drop the table called MasterTable in SQL since -d is specified, recreate the SQL table using create statements, and then import all of the rows.
