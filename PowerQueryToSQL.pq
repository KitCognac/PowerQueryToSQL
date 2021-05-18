let
    PowerQueryToSQLFunction =
        (//
        PQ_DATASOURCE as table, //
        UPLOAD_SQL_SERVER as text, // 
        UPLOAD_SQL_DATABASE as text, //
        UPLOAD_SQL_TABLE as text, //
        UPLOAD_SQL_BATCH_ROWS as number, // 
        UPLOAD_SQL_BATCH_BEFORE as any//
        ) as table =>
            let
                // Set source as a PQ Source
                COMBINE_COL_VALUE = PQ_DATASOURCE,
                // Log Start Time
                StartTime = DateTime.LocalNow(),
                // COMBINE ALL TABLES AND GETS ORIGINAL COLUMNS NAME
                TBL_COL_NAMES = Table.ColumnNames(COMBINE_COL_VALUE),
                // Prepare Type TXT to convert
                TBL_COLS_TYPE_TXT =
                    List.Transform(
                        TBL_COL_NAMES,
                        each
                            {
                                _,
                                type text
                            }
                    ),
                // GROUP DATA into Batch for Large Table
                CREATE_BATCH_ID =
                    Table.AddIndexColumn(
                        COMBINE_COL_VALUE,
                        "BATCH_ID",
                        1,
                        1,
                        Int64.Type
                    ),
                // BASED on BatchRows Settings 
                CREATE_BATCH_GROUP =
                    Table.AddColumn(
                        CREATE_BATCH_ID,
                        "BATCH",
                        each Number.RoundUp([BATCH_ID] / UPLOAD_SQL_BATCH_ROWS)
                    ),
                // Group Rows into Table Object
                CREATE_BATCH_TABLE =
                    Table.Group(
                        CREATE_BATCH_GROUP,
                        {"BATCH"},
                        {
                            {
                                "ROWS",
                                each Table.RowCount(_),
                                Int64.Type
                            },
                            {
                                "TBL",
                                each _
                            }
                        }
                    ),
                // CONCAT ALL COLS TO ONE COLS
                // Using Locale en-GB to have Date Format in 'dd/MM/yyyy'
                // In INSERT_STATEMENT I use SET DATEFORMAT DMY so Locale en-GB is necessary
                // Use Before Batch SQL Script to setup DDL correctly becasue 
                //      all String set to Unicode, all conversion will rely on implicit SQL Server. 
                // Concat using custom delimiter @@,N@@ to clean Single Quote first then convert it back.
                CONCAT_COLS =
                    Table.AddColumn(
                        CREATE_BATCH_TABLE,
                        "CONCAT_COLS",
                        each
                            Table.CombineColumns(
                                Table.TransformColumnTypes(
                                    [TBL],
                                    TBL_COLS_TYPE_TXT,
                                    "en-GB"
                                ),
                                TBL_COL_NAMES,
                                Combiner.CombineTextByDelimiter(
                                    "@@,N@@",
                                    QuoteStyle.None
                                ),
                                "SQL_CONCAT_COLS"
                            )
                    ),
                // Clean single quote to double quote for SQL Syntax Correction
                CLEAN_SINGLE_QUOTE =
                    Table.AddColumn(
                        CONCAT_COLS,
                        "CLEAN_SINGLE_QUOTE",
                        each
                            let
                                __COL = [CONCAT_COLS]
                            in
                                Table.ReplaceValue(
                                    __COL,
                                    "'",
                                    "''",
                                    Replacer.ReplaceText,
                                    {"SQL_CONCAT_COLS"}
                                )
                    ),
                // Convert back @@,N@@ to ',N'
                CONVERT_DELIMITER =
                    Table.AddColumn(
                        CLEAN_SINGLE_QUOTE,
                        "CONVERT_DELIMITER",
                        each
                            let
                                __COL = [CLEAN_SINGLE_QUOTE]
                            in
                                Table.ReplaceValue(
                                    __COL,
                                    "@@,N@@",
                                    "',N'",
                                    Replacer.ReplaceText,
                                    {"SQL_CONCAT_COLS"}
                                )
                    ),
                // ADD SELECT SYNTAX AND QUOTE
                ADD_SELECT_SYNTAX =
                    Table.AddColumn(
                        CONVERT_DELIMITER,
                        "ADD_SELECT",
                        each
                            let
                                __COL = [CONVERT_DELIMITER]
                            in
                                Table.AddColumn(
                                    __COL,
                                    "ADD_SELECT",
                                    each " SELECT N'" & [SQL_CONCAT_COLS] & "'"
                                )
                    ),
                // Convert all N'' to NULL value
                CLEAN_NULL_VALUE =
                    Table.AddColumn(
                        ADD_SELECT_SYNTAX,
                        "CLEAN_NULL",
                        each
                            let
                                __COL = [ADD_SELECT]
                            in
                                Table.ReplaceValue(
                                    __COL,
                                    "N''",
                                    "NULL",
                                    Replacer.ReplaceText,
                                    {"ADD_SELECT"}
                                )
                    ),
                // Combine all Rows into one SQL Query
                INSERT_STATEMENT =
                    Table.AddColumn(
                        CLEAN_NULL_VALUE,
                        "SQL_QUERY",
                        each
                            let
                                __TBL = [CLEAN_NULL],
                                __BATCH_NO = [BATCH]
                            in
                                (
                                    if __BATCH_NO = 1 then
                                        UPLOAD_SQL_BATCH_BEFORE
                                    else
                                        ""
                                )
                                & " SET DATEFORMAT DMY "
                                & " DECLARE @START_TIME DATETIME = CURRENT_TIMESTAMP "
                                & " ; INSERT INTO "
                                & UPLOAD_SQL_TABLE
                                & " WITH (TABLOCK) "
                                & Text.Combine(
                                    __TBL[ADD_SELECT],
                                    " UNION ALL "
                                )
                                & " GO "
                                & " SELECT  DATEDIFF( MILLISECOND, @START_TIME, CURRENT_TIMESTAMP ) AS UPLOAD_TIME ;"
                    ),
                // Filter out unnecessary columns
                FILTER_STATEMENT_COLS =
                    Table.SelectColumns(
                        INSERT_STATEMENT,
                        {
                            "BATCH",
                            "ROWS",
                            "SQL_QUERY"
                        }
                    ),
                // Set up SQL Connections
                SQL_SOURCE =
                    Sql.Database(
                        UPLOAD_SQL_SERVER,
                        UPLOAD_SQL_DATABASE
                    ),
                // Exec SQL Batch using Value.NativeQuery
                EXEC_BATCHS =
                    Table.AddColumn(
                        FILTER_STATEMENT_COLS,
                        "EXEC_TIME",
                        each
                            Value.NativeQuery(
                                SQL_SOURCE,
                                [SQL_QUERY]
                            )
                                [UPLOAD_TIME]{0}
                    ),
                // Filter out unnecessary columns
                IMPORT_RESULT =
                    Table.SelectColumns(
                        EXEC_BATCHS,
                        {
                            "BATCH",
                            "ROWS",
                            "EXEC_TIME"
                        }
                    ),
                // Group all rows to make sure all batch sent to server. 
                // This step trying to prevent unexpected import partial data when user hover over Queries in Queries&Connection Task Pane.
                TIME_RESULT =
                    Table.Group(
                        IMPORT_RESULT,
                        {},
                        {
                            {
                                "EXEC_TIME",
                                each List.Sum([EXEC_TIME]),
                                type number
                            },
                            {
                                "ROWS_AFFECTED",
                                each List.Sum([ROWS]),
                                type number
                            }
                        }
                    ),
                PQ_EXEC_TIME =
                    Table.AddColumn(
                        TIME_RESULT,
                        "PQ_EXEC_TIME",
                        each Duration.ToText(DateTime.LocalNow() - StartTime)
                    )
            in
                PQ_EXEC_TIME,
    // Add documentation
    PowerQueryToSQLType =
        type function (//
        PQ_DATASOURCE as
            (
                type table
                meta
                [
                    Documentation.FieldCaption = "Source Table",
                    Documentation.FieldDescription = "The table to upload",
                    Documentation.SampleValues = {
                        "Queries Name"
                    }
                ]
            ), //
        UPLOAD_SQL_SERVER as
            (
                type text
                meta
                [
                    Documentation.FieldCaption = "Target SQL Server Address",
                    Documentation.FieldDescription = "Ex: localhost; 192.168.1.1; ...",
                    Documentation.SampleValues = {
                        "localhost"
                    }
                ]
            ), //
        UPLOAD_SQL_DATABASE as
            (
                type text
                meta
                [
                    Documentation.FieldCaption = "Target SQL Database Name",
                    Documentation.FieldDescription = "Ex: AdventureWorksDW2020; TailspinToys2020-US ...",
                    Documentation.SampleValues = {
                        "AdventureWorksDW2020"
                    }
                ]
            ), //
        UPLOAD_SQL_TABLE as
            (
                type text
                meta
                [
                    Documentation.FieldCaption = "Target SQL Schema and Table Name",
                    Documentation.FieldDescription = "Ex: dbo.Date; dbo.Sales ...",
                    Documentation.SampleValues = {
                        "dbo.Sales"
                    }
                ]
            ), //
        UPLOAD_SQL_BATCH_ROWS as
            (
                type number
                meta
                [
                    Documentation.FieldCaption = "Batch Rows for Split Large Table to Upload.",
                    Documentation.FieldDescription = "Ex: 100, 1000, 10000, ... Optimal is 1000",
                    Documentation.SampleValues = {
                        1000
                    }
                ]
            ), //
        UPLOAD_SQL_BATCH_BEFORE as
            (
                type text
                meta
                [
                    Documentation.FieldCaption = "SQL Snippet to be done before begin upload. ",
                    Documentation.FieldDescription = "Ex: TRUNCATE TABLE dbo.Dates . If error add space or ; in the end",
                    Documentation.SampleValues = {
                        " TRUNCATE TABLE dbo.Sales "
                    },
                    Formatting.IsMultiLine = true,
                    Formatting.IsCode = true
                ]
            )) as table
        meta
        [
            Documentation.Name = "PowerQueryToSQL",
            Documentation.Description = "Automate EXEC DML&DDL Before Upload PQ Object to SQL Server",
            Documentation.LongDescription =
                "
            For ADVANCED USER only. This function will not replace SSIS or any kind of insert bulk, but it will: 
            <br> - Ultilize settings per table in excel workbook through Queries&Connection Task Pane. 
            <br> - Easy maintain DML&DDL before upload
            <br> - Upload directly while Excel Workbook is opened.
            <br>Limitation:
            <br> - Unexpected import Data when user hover queries in Queries&Connection Task Pane
            <br> - Slow compared to insert Bulk
            ",
            Documentation.Category = "Table",
            Documentation.Author = "Kit Cognac",
            Documentation.Examples = {
                [
                    Description = "Replace values on all columns",
                    Code =
                        "PowerQueryToSQL(
                    SOURCE_TABLE, 
                    ""localhost"",
                    ""AdventureWorksDW2020"",
                    ""dbo.Sales"",
                    1000,
                    "" TRUNCATE TABLE dbo.Sales ""
                    )",
                    Result = "#table({""EXEC_TIME"", ""ROWS_AFFECTED"",""PQ_EXEC_TIME""}, {{""8000ms"",5000,""00:00:01.2672214""}})"
                ]
            }
        ]
in
    Value.ReplaceType(
        PowerQueryToSQLFunction,
        PowerQueryToSQLType
    )