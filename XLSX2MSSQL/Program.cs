using System;
using System.Data;
using System.IO;
using OfficeOpenXml;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Configuration;
using NDesk.Options;

namespace XLS2MSSQL
{
    class Program
    {

        static void Main(string[] args)
        {
            //Parse the command line options and configure the application settings.
            ApplicationSettingsClass application_settings = new ApplicationSettingsClass(args);

            //Read in the excel spreadsheet to a datatable value
            DataTable excel_datatable = read_xlsx_file_into_datatable(application_settings);

            //Open the SQL database connection
            application_settings.sql_connection = find_connection_string_and_initialize_sql_connection(application_settings.connection_string);

            //If the drop table parameter is specified, then delete(drop) the table first.
            if (application_settings.drop_table)
            {
                drop_table_if_exists(application_settings);
            }

            //Create the table if it does not exist. If the table already exists then add any new columns.
            Create_Table_If_Not_Exists_Or_Add_Columns_If_Needed(application_settings, excel_datatable);

            //Insert all of the rows from the datatable into the SQL table
            insert_datatable_into_mssql(application_settings, excel_datatable);

            verify_sql_table_row_counts(application_settings, excel_datatable.Rows.Count);
            //Close the SQL connection
            close_mssql_connection_if_open(application_settings.sql_connection);

            //Exit without any issues
            //Console.ReadLine();
            Environment.Exit(0);
        }


        public class ApplicationSettingsClass
        {
            //initialize settings
            public bool show_help = false;
            public bool drop_table = false;
            public string file_path = "";
            public string database_table_name = "";
            public string worksheet_name = "";
            public string connection_string = "";
            public List<string> rename_column_string_list = new List<string>();
            public SqlConnection sql_connection;
            public OptionSet option_set_object;

            public ApplicationSettingsClass(string[] input_args)
            {
                this.read_command_line_arguments_and_configure_settings(input_args);
            }

            //Set the command line values
            public void read_command_line_arguments_and_configure_settings(string[] args)
            {

                //Initialize the command line arguement values and values they'll be stored in
                this.option_set_object = new OptionSet() {
                { "d|drop",  "Delete the table and then recreate",
                    v => this.drop_table = v != null },
                { "c|connectionstring=",
                    "The connection string to connect to SQL with. If not specified, the app.config file will be used.",
                    v => this.connection_string = v },
                { "f|filepath=",
                    "The excel spreadsheet filepath",
                    v => this.file_path = v },
                { "r|rename=", "rename a column in the format of oldcolumnname,newcolumnname",
                    v => this.rename_column_string_list.Add (v) },
                { "t|tablename=",
                    "The SQL tablename to copy the excel file to",
                    v => this.database_table_name = v },
                { "w|worksheet=",
                    "The Excel worksheet name to read in",
                    v => this.worksheet_name = v },
                { "h|help",  "Show this message and exit",
                    v => this.show_help = v != null },
                };

                //Parse the command line arguments.
                List<string> extra;
                try
                {
                    extra = this.option_set_object.Parse(args);
                }
                catch (OptionException e)
                {
                    Console.Write("XLSX2MSSQL: ");
                    Console.WriteLine(e.Message);
                    Console.WriteLine("Try `XLSX2MSSQL --help' for more information.");
                    Environment.Exit(4);
                }

                //show help if there are issues with the command line arguments
                if (this.show_help || args.Length < 1 || this.file_path == "")
                {
                    this.ShowHelp();
                }


                //If no tablename is specified then set it to the excel spreadsheet name without the extension
                if (this.database_table_name == "")
                {
                    this.database_table_name = Path.GetFileNameWithoutExtension(this.file_path);
                }

                //If no connection string is specified then use the value in the app.config file
                if (connection_string == "")
                {
                    try
                    {
                        connection_string = ConfigurationManager.ConnectionStrings["MyConnectionStringName"].ConnectionString;
                    }
                    catch
                    {
                        Console.WriteLine("The connection string was not set and could not find the value in the app.config file. Please use the -c parameter or app.config file");
                        Environment.Exit(8);
                    }
                }

            }

            //Show help message
            public void ShowHelp()
            {
                Console.WriteLine("Usage: XLSX2MSSQL [OPTIONS]+");
                Console.WriteLine("Import a XLSX file into a MSSQL Database.");
                Console.WriteLine();
                Console.WriteLine("Options:");
                this.option_set_object.WriteOptionDescriptions(Console.Out);
                Environment.Exit(5);
            }
        }







        static bool verify_sql_table_row_counts(ApplicationSettingsClass application_settings, int datatable_row_count)
        {
            bool return_value = false;
            //Verify how many rows exist
            try
            {

                string sqlcmd = "SELECT COUNT(*) FROM [" + application_settings.database_table_name + "]";
                SqlCommand cmd = new SqlCommand(sqlcmd, application_settings.sql_connection);
                Int32 verified_row_count = (Int32)cmd.ExecuteScalar();

                if (verified_row_count == datatable_row_count)
                {
                    Console.WriteLine("Confirmed table " + application_settings.database_table_name + " now has " + verified_row_count.ToString() + " rows.");
                    return_value = true;
                }
                else
                {
                    Console.WriteLine("Table Counts do not match. " + application_settings.database_table_name + " now has " + verified_row_count.ToString() + " rows but the spreadsheet read in " + datatable_row_count.ToString() + " rows.");
                }

            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            return return_value;
        }

        static SqlConnection find_connection_string_and_initialize_sql_connection(string connection_string)
        {



            //Open the SQL database connection
            SqlConnection myConnection = new SqlConnection(connection_string);
            open_mssql_connection_if_closed(myConnection);

            return myConnection;
        }


        //Read in the excel spreadsheet to a datatable value
        static DataTable read_xlsx_file_into_datatable(ApplicationSettingsClass input_CommandLineArgument_object)
        {

            //Check if the file exists, if it does not then exit and throw an error.
            if (!File.Exists(input_CommandLineArgument_object.file_path))
            {
                Console.WriteLine("Error, the file path could not be found at: " + input_CommandLineArgument_object.file_path);
                Environment.Exit(8);
            }


            //Opening an existing Excel file worksheet
            ExcelWorksheet workSheet = open_excel_worksheet(input_CommandLineArgument_object.file_path, input_CommandLineArgument_object.worksheet_name);

            DataTable tbl = new DataTable(input_CommandLineArgument_object.worksheet_name);

            add_all_columns_to_datatable(tbl, workSheet);

            //Rename any columns if there were parameters to do so
            rename_column_headers_if_specified(tbl, input_CommandLineArgument_object.rename_column_string_list);


            //Loop through every row, then loop through every column to get each cell value
            add_rows_to_datatable(tbl, workSheet);
            
            Console.WriteLine("Read in " + tbl.Rows.Count.ToString() + " rows from " + input_CommandLineArgument_object.file_path);

            return tbl;
        }

        //Rename any columns if there were parameters to do so
        static void rename_column_headers_if_specified(DataTable tbl, List<String> rename_column_string_list)
        {
            foreach (string current_rename_string in rename_column_string_list)
            {
                string[] rename_column_array = current_rename_string.Split(',');
                if (rename_column_array.Length > 1)
                {
                    string old_column_name = rename_column_array[0];
                    string new_column_name = rename_column_array[1];

                    if (tbl.Columns.Contains(old_column_name))
                    {
                        tbl.Columns[old_column_name].ColumnName = new_column_name;
                    }

                }
            }

        }




    static ExcelWorksheet open_excel_worksheet(string file_path, string worksheet_name)
    {
        //Opening an existing Excel file
        FileInfo fi = new FileInfo(file_path);
        var package = new ExcelPackage(fi);

        ExcelWorksheet workSheet;

        //If the worksheet name is specified then use that, otherwise default to the first sheet.
        if (worksheet_name != "")
        {
            workSheet = package.Workbook.Worksheets[worksheet_name];
        }
        else
        {
            workSheet = package.Workbook.Worksheets[1];
        }

        return workSheet;

    }


    static void add_all_columns_to_datatable(DataTable input_data_table, ExcelWorksheet input_workSheet, bool hasHeader = true)
    {

        DataTable tbl = new DataTable("dtImage");

        int startRow = hasHeader ? 2 : 1;

        //Loop through every column in order to add all of the columns to the datatable that will be returned
        for (int column_index_number = 1; column_index_number <= input_workSheet.Dimension.End.Column; column_index_number++)
        {

            //We determine if the field should be a string, integer, or float
            //This requires looping through all the cells in the column to see if they meet the requirements
            //Default to string if the fields include non-numeric data
            bool column_is_float = check_if_column_is_float(input_workSheet, startRow, column_index_number);
            bool column_is_integer = false;
            if (!column_is_float)
            {
                column_is_integer = check_if_column_is_integer(input_workSheet, startRow, column_index_number);

            }


            var firstRowCell = input_workSheet.Cells[1, column_index_number, 1, column_index_number];
            string column_header_name = hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column);
            column_header_name = column_header_name.Trim();

            Type current_column_data_type;
            if (column_is_float)
            {
                current_column_data_type = typeof(double);
            }
            else if (column_is_integer)
            {
                current_column_data_type = typeof(int);
            }
            else
            {
                current_column_data_type = typeof(string);
            }


            if (!input_data_table.Columns.Contains(column_header_name))
            {
                input_data_table.Columns.Add(column_header_name, current_column_data_type);
            }
            else
            {
                input_data_table.Columns.Add(column_header_name + "2", current_column_data_type);
            }


        }

    }


        static void add_rows_to_datatable(DataTable tbl, ExcelWorksheet workSheet, bool hasHeader = true)
        {

            int startRow = hasHeader ? 2 : 1;   
            
            for (int rowNum = startRow; rowNum <= workSheet.Dimension.End.Row; rowNum++)
            {
                var wsRow = workSheet.Cells[rowNum, 1, rowNum, workSheet.Dimension.End.Column];
                DataRow row = tbl.Rows.Add();
                foreach (var cell in wsRow)
                {

                    if (cell.Value == null)
                    {
                        row[cell.Start.Column - 1] = DBNull.Value;
                    }
                    else
                    {
                        Type cell_data_type = tbl.Columns[cell.Start.Column - 1].DataType;

                        if (cell_data_type == typeof(System.String))
                        {
                            row[cell.Start.Column - 1] = cell.Value;
                        }
                        else if (cell_data_type == typeof(System.Int32))
                        {
                            int write_int;
                            Int32.TryParse(cell.Value.ToString(), out write_int);
                            row[cell.Start.Column - 1] = write_int;
                        }
                        else if (cell_data_type == typeof(System.Double))
                        {
                            Double write_double;
                            Double.TryParse(cell.Value.ToString(), out write_double);
                            row[cell.Start.Column - 1] = write_double;
                        }
                    }
                }
            }

        }



        //Loop through cell in the column and decide if all the values are float. They need to have a period in the value and numeric values in every field.
        static bool check_if_column_is_float(ExcelWorksheet workSheet, int start_row, int column_index_number)
        {
            bool found_non_blank_value = false;
            bool period_found = false;

            for (int rowNum = start_row; rowNum <= workSheet.Dimension.End.Row; rowNum++)
            {
            object current_cell_value = workSheet.Cells[rowNum, column_index_number].Value;
                string cell = "";
                if (current_cell_value != null)
                {
                    cell = current_cell_value.ToString();
                    if (cell.Contains("."))
                    {
                        period_found = true;
                    }
                }
                
                if ((cell.Trim() != "") && (!value_is_float(cell)))
                {
                    return false;
                }
                else if (cell.Trim() != "")
                {
                    found_non_blank_value = true;
                }
            }

            if(!period_found)
            {
                return false;
            }

            if (!found_non_blank_value)
            {
                return false;
            }



            return true;
        }

        //Loop through cell in the column and decide if all the values are integers. Every field needs to be an integer.
        static bool check_if_column_is_integer(ExcelWorksheet workSheet, int start_row, int column_index_number)
        {
            bool found_non_blank_value = false;

            for (int rowNum = start_row; rowNum <= workSheet.Dimension.End.Row; rowNum++)
            {
                object current_cell_value = workSheet.Cells[rowNum, column_index_number].Value;

                string cell = "";

                if(current_cell_value != null)
                {
                    cell = current_cell_value.ToString();
                }


                if (cell.Trim() != "" && !value_is_integer(cell))
                {
                    return false;
                }
                else if (cell.Trim() != "")
                {
                    found_non_blank_value = true;
                }


            }

            if(!found_non_blank_value)
            {
                return false;
            }

            return true;

        }

        //Test if the value can be converted to float
        static bool value_is_float(string input_string)
        {
            bool return_value_is_float = float.TryParse(input_string, out float result);
            return return_value_is_float;
        }

        //Test if the value can be converted to integer
        static bool value_is_integer(string input_string)
        {
            bool return_value_is_integer = Int32.TryParse(input_string, out int result);
            return return_value_is_integer;
        }



        //Test if the MSSQL table exists
        static bool check_if_mssql_table_exists(SqlConnection mySqlConnection, string Tablename)
        {
            try
            {
                SqlCommand command = new SqlCommand("select 1 from [" + Tablename + "] where 1 = 0", mySqlConnection);
                command.ExecuteNonQuery();

            }
            catch
            {
                return false;
            }

            return true;
        }


        //Drop the table if it exists
        static void drop_table_if_exists(ApplicationSettingsClass application_settings)
        {

            open_mssql_connection_if_closed(application_settings.sql_connection);

            bool table_exists = check_if_mssql_table_exists(application_settings.sql_connection, application_settings.database_table_name);

            if (table_exists)
            {
                string sql_cmd = "DROP TABLE [" + application_settings.database_table_name + "];";
                Console.WriteLine("Dropping table using command: " + sql_cmd);
                SqlCommand command = new SqlCommand(sql_cmd, application_settings.sql_connection);
                command.ExecuteNonQuery();
            }
        }

        //Check if the MSSQL connection is open. If it is not, then open it
        static void open_mssql_connection_if_closed(SqlConnection mySqlConnection)
        {
            if (mySqlConnection != null && mySqlConnection.State == ConnectionState.Closed)
            {
                Console.WriteLine("Opening SQL Connection: " + mySqlConnection.ConnectionString);
                try
                {
                    mySqlConnection.Open();
                    
                    Console.WriteLine("Successfully connected to server " + mySqlConnection.DataSource.ToString() + " on database: " + mySqlConnection.Database.ToString());
                }
                catch (Exception e)
                {
                    Console.WriteLine("Error connecting to the database. Exiting.");
                    Console.WriteLine(e.Message);
                    Environment.Exit(6);
                }
            }
        }

        //Close the MSSQL connection if it is active
        static void close_mssql_connection_if_open(SqlConnection mySqlConnection)
        {
            if (mySqlConnection != null && mySqlConnection.State == ConnectionState.Open)
            {
                Console.WriteLine("Closing SQL Connection.");
                mySqlConnection.Close();
            }
        }

        //This function translates the C# Datatable types into the MSSQL column type to create the SQL table or add columns
        //Default to nvarchar if it can't find a match
        static string convert_csharp_object_type_to_sql_column_name(DataColumn data_column_object)
        {
            switch (data_column_object.DataType.ToString())
            {
                case "System.UInt16":
                    return "int";
                case "System.Int32":
                    return "int";
                case "System.UInt32":
                    return " int";
                case "System.Int64":
                    return "bigint";
                case "System.Int16":
                    return "smallint";
                case "System.Byte":
                    return "tinyint";
                case "System.Single":
                    return "float";
                case "System.SByte":
                    return "smallint";
                case "System.Double":
                    return "float";
                case "System.Decimal":
                    return "decimal";
                case "System.DateTime":
                    return "datetime";
                default:
                    return $" nvarchar({(data_column_object.MaxLength == -1 ? "max" : data_column_object.MaxLength.ToString())})";
            }


        }

        //Gets a list of all the columns from a SQL database table
        static List<string> get_mssql_table_column_names(SqlConnection myConnection, string table_name)
        {
            open_mssql_connection_if_closed(myConnection);

            string[] restrictionsColumns = new string[4];

            restrictionsColumns[2] = table_name;

            List<string> return_column_header_list = new List<string>();

            DataTable schemaColumns = myConnection.GetSchema("Columns", restrictionsColumns);
            foreach (DataRow rowColumn in schemaColumns.Rows)
            {
                string ColumnName = rowColumn[3].ToString();
                return_column_header_list.Add(ColumnName);
            }

            return return_column_header_list;
        }

        //Create the table if it does not exist. If the table already exists then add any new columns.
        static void Create_Table_If_Not_Exists_Or_Add_Columns_If_Needed(ApplicationSettingsClass application_settings, DataTable table)
        {
            open_mssql_connection_if_closed(application_settings.sql_connection);

            bool table_exists = check_if_mssql_table_exists(application_settings.sql_connection, application_settings.database_table_name);

            if (!table_exists)
            {
                string sqlsc = "CREATE TABLE [" + application_settings.database_table_name + "] (";
                foreach (DataColumn current_data_column in table.Columns)
                {
                    sqlsc += "\n [" + current_data_column.ColumnName + "] ";

                    sqlsc += " " + convert_csharp_object_type_to_sql_column_name(current_data_column) + " ";

                    if (current_data_column.AutoIncrement)
                        sqlsc += " IDENTITY(" + current_data_column.AutoIncrementSeed + "," + current_data_column.AutoIncrementStep + ") ";
                    if (!current_data_column.AllowDBNull)
                        sqlsc += " NOT NULL ";
                    sqlsc += ",";

                }
                sqlsc = sqlsc.Substring(0, sqlsc.Length - 1) + "\n)";

                Console.WriteLine("Creating Table with command: " + sqlsc);
                SqlCommand command = new SqlCommand(sqlsc, application_settings.sql_connection);
                command.ExecuteNonQuery();
            }
            else
            {
                List<String> database_column_header_list = get_mssql_table_column_names(application_settings.sql_connection, application_settings.database_table_name);

                foreach (DataColumn current_column in table.Columns)
                {
                    if (!database_column_header_list.Contains(current_column.ColumnName))
                    {
                        string data_type_string = convert_csharp_object_type_to_sql_column_name(current_column);
                        string sqlsc = "alter table [" + application_settings.database_table_name + "] add [" + current_column.ColumnName + "] " + data_type_string;
                        if (current_column.AutoIncrement)
                            sqlsc += " IDENTITY(" + current_column.AutoIncrementSeed + "," + current_column.AutoIncrementStep + ") ";
                        if (!current_column.AllowDBNull)
                            sqlsc += " NOT NULL ";
                        Console.WriteLine("Creating Column with command: " + sqlsc);
                        SqlCommand cmd = new SqlCommand(sqlsc, application_settings.sql_connection);
                        cmd.ExecuteNonQuery();

                    }
                }
            }
        }

        //Insert all of the rows from the datatable into the SQL table
        static void insert_datatable_into_mssql(ApplicationSettingsClass application_settings, DataTable table)
        {

            open_mssql_connection_if_closed(application_settings.sql_connection);
            SqlBulkCopy bulkcopy = new SqlBulkCopy(application_settings.sql_connection)
            {
                DestinationTableName = "[" + application_settings.database_table_name + "]",
                BatchSize = 20000,
                BulkCopyTimeout = 1500
            };

            foreach (DataColumn current_column in table.Columns)
            {
                bulkcopy.ColumnMappings.Add(current_column.ColumnName, current_column.ColumnName);
            }

            Console.WriteLine("Inserting " + table.Rows.Count + " rows into table " + application_settings.database_table_name);

            //Table has already been created, insert the data using bulk copy
            try
            {
                bulkcopy.WriteToServer(table);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }

        }








    }
}









