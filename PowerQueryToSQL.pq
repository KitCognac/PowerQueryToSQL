let
    PowerQueryToSQLFunction =
        (//
        PQ_DATASOURCE as table, //
        UPLOAD_SQL_SERVER as text, // 
        UPLOAD_SQL_DATABASE as text, //
        UPLOAD_SQL_TABLE as text, //
        UPLOAD_SQL_BATCH_ROWS as number, // 
        UPLOAD_SQL_IS_START as text, //   
        UPLOAD_SQL_DDL as text, //
        UPLOAD_SQL_DML as text//
        ) as table =>
            let
                // Set source as a PQ Source
                COMBINE_COL_VALUE = PQ_DATASOURCE,
                // SET a temp table name
                PQ_GUID =
                    // "tempdb."
                    UPLOAD_SQL_TABLE
                    & Text.Replace(Text.NewGuid(), "-", ""),
                // Log Start Time
                START_TIME = DateTime.LocalNow(),
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
                                    each
                                        " SELECT N'"
                                        & [SQL_CONCAT_COLS]
                                        & "' #(lf) "
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
                // Temp Data Table Before Creat a Huge SQL Files
                TMP_CLEAN_NULL_VALUE = Table.Buffer(CLEAN_NULL_VALUE),
                // Combine all Rows into one SQL Query
                INSERT_STATEMENT =
                    Table.AddColumn(
                        TMP_CLEAN_NULL_VALUE,
                        "SQL_QUERY",
                        each
                            let
                                __TBL = [CLEAN_NULL],
                                __BATCH_NO = [BATCH]
                            in
                                if UPLOAD_SQL_IS_START = "YES" then
                                    (
                                        " SET DATEFORMAT DMY ; #(lf)"
                                        & " DECLARE @START_TIME DATETIME = CURRENT_TIMESTAMP ; #(lf)"
                                        & " INSERT INTO "
                                        & PQ_GUID
                                        & "#(lf)"
                                        & Text.Combine(
                                            __TBL[ADD_SELECT],
                                            " UNION ALL #(lf)"
                                        )
                                        & "#(lf)"
                                        & " SELECT DATEDIFF(MILLISECOND,@START_TIME,CURRENT_TIMESTAMP) AS UPLOAD_TIME ;#(lf)"
                                    )
                                else
                                    " SELECT 0 AS UPLOAD_TIME ;#(lf)"
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
                CACHE_SQL_QUERIES = Table.Buffer(FILTER_STATEMENT_COLS),
                // Buffer Table so that excel wont send another batchs
                // Using List.Generate to ensure Batch is sent by orders
                // ACCU_EXEC_TIME make sure previous query is done
                EXEC_BATCHS =
                    let
                        // TEMP Final Batch Number
                        __FINAL_BATCH_NO = List.Max(CACHE_SQL_QUERIES[BATCH]),
                        __SQL_CONNECTION =
                            Sql.Database(
                                UPLOAD_SQL_SERVER,
                                UPLOAD_SQL_DATABASE
                            ),
                        // Clean-up any leftover temptable because of disconnection
                        // Make sure only lastest queries running
                        __INIT =
                            "#(lf)
                        DECLARE @LOOP_DROP_TBL VARCHAR (255) ;
                        DECLARE LOOP_DROP_TBL CURSOR FOR
                            SELECT
                                'DROP TABLE IF EXISTS ' + TABLE_SCHEMA + '.' + TABLE_NAME + ''
                            FROM
                                INFORMATION_SCHEMA.TABLES
                            WHERE
                                (TABLE_SCHEMA + '.' + TABLE_NAME) LIKE '"
                            & UPLOAD_SQL_TABLE
                            & "________________________________' ;

                        OPEN LOOP_DROP_TBL ;
                        WHILE 1 = 1
                            BEGIN
                                FETCH LOOP_DROP_TBL
                                INTO
                                    @LOOP_DROP_TBL ;
                                IF @@fetch_status != 0
                                    BREAK ;
                                EXEC (@LOOP_DROP_TBL) ;
                            END ;
                        CLOSE LOOP_DROP_TBL ;
                        DEALLOCATE LOOP_DROP_TBL ;
                        #(lf)",
                        __DDL_EXEC =
                            " DECLARE @START_TIME DATETIME = CURRENT_TIMESTAMP ; #(lf)"
                            & UPLOAD_SQL_DDL
                            & "#(lf)"
                            & " SELECT DATEDIFF(MILLISECOND,@START_TIME,CURRENT_TIMESTAMP) AS UPLOAD_TIME ;#(lf)"
                            & " SELECT * INTO "
                            & PQ_GUID
                            & " FROM "
                            & UPLOAD_SQL_TABLE
                            & " WHERE 1 = 0 ;#(lf)",
                        __DML_EXEC =
                            "#(lf)"
                            & " BEGIN TRAN #(lf)"
                            & UPLOAD_SQL_DML
                            & "#(lf)"
                            & " INSERT INTO "
                            & UPLOAD_SQL_TABLE
                            & " WITH(TABLOCK) #(lf)"
                            & " SELECT * FROM "
                            & PQ_GUID
                            & "#(lf)"
                            & " DROP TABLE IF EXISTS "
                            & PQ_GUID
                            & "#(lf)"
                            & " COMMIT TRAN #(lf)"
                    in
                        Table.Buffer(
                            Table.FromRecords(
                                List.Generate(
                                    () =>
                                        [
                                            __BATCH = 0,
                                            __ROWS = 0,
                                            __QUERY =
                                                if UPLOAD_SQL_IS_START = "YES" then
                                                    __INIT & __DDL_EXEC
                                                else
                                                    " SELECT 0 AS UPLOAD_TIME ;#(lf)",
                                            __EXEC =
                                                Value.NativeQuery(
                                                    __SQL_CONNECTION,
                                                    __QUERY,
                                                    null,
                                                    [EnableFolding = false]
                                                )
                                        // __ACCU_EXEC_TIME = __EXEC[UPLOAD_TIME]{0}
                                        ],
                                    each [__BATCH] <= __FINAL_BATCH_NO,
                                    each
                                        [
                                            __BATCH = [__BATCH] + 1,
                                            __TBL = CACHE_SQL_QUERIES{[BATCH = __BATCH]},
                                            // __ROWS = __TBL[ROWS]{0},
                                            __ROWS = CACHE_SQL_QUERIES{[BATCH = __BATCH]}[ROWS],
                                            __QUERY =
                                                __TBL[SQL_QUERY]
                                                & (
                                                    if __BATCH = __FINAL_BATCH_NO then
                                                        if UPLOAD_SQL_IS_START = "YES" then
                                                            __DML_EXEC
                                                        else
                                                            " SELECT 0 AS UPLOAD_TIME ;#(lf)"
                                                    else
                                                        ""
                                                ),
                                            __EXEC =
                                                Value.NativeQuery(
                                                    __SQL_CONNECTION,
                                                    __QUERY,
                                                    null,
                                                    [EnableFolding = false]
                                                )
                                        // __ACCU_EXEC_TIME =
                                        //     __EXEC[UPLOAD_TIME]{0}
                                        //     + [__ACCU_EXEC_TIME]
                                        ],
                                    each
                                        [
                                            BATCH = [__BATCH],
                                            ROWS = [__ROWS],
                                            SQL_QUERY = [__QUERY],
                                            EXEC_TIME = [__EXEC][UPLOAD_TIME]{0}
                                        // ACCU_EXEC_TIME = [__ACCU_EXEC_TIME]
                                        ]
                                )
                            )
                        ),
                GROUP_TIMING =
                    Table.Group(
                        EXEC_BATCHS,
                        {},
                        {
                            {
                                "ROWS",
                                each List.Sum([ROWS]),
                                type number
                            },
                            {
                                "SQL_EXEC_TIME",
                                each
                                    Duration.ToText(
                                        Duration.From(
                                            List.Sum([EXEC_TIME])
                                            / (1000 * 60 * 60 * 24)
                                        )
                                    ),
                                type any
                            }
                        }
                    ),
                PQ_EXEC_TIME =
                    Table.AddColumn(
                        GROUP_TIMING,
                        "PQ_EXEC_TIME",
                        each
                            let
                                __END_TIME = DateTime.LocalNow()
                            in
                                Duration.ToText(__END_TIME - START_TIME)
                    ),
                ADD_TARGET_NAME =
                    Table.AddColumn(
                        PQ_EXEC_TIME,
                        "TARGET",
                        each UPLOAD_SQL_TABLE
                    ),
                ADD_IS_START_STATUS =
                    Table.AddColumn(
                        ADD_TARGET_NAME,
                        "IS_START",
                        each UPLOAD_SQL_IS_START
                    )
            in
                ADD_IS_START_STATUS,
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
        UPLOAD_SQL_IS_START as
            (
                type text
                meta
                [
                    Documentation.FieldCaption = "Functions Start",
                    Documentation.FieldDescription = "Ex: YES, NO",
                    Documentation.AllowedValues = {
                        "NO",
                        "YES"
                    },
                    Documentation.SampleValues = {
                        "YES"
                    }
                ]
            ), //
        UPLOAD_SQL_DDL as
            (
                type text
                meta
                [
                    Documentation.FieldCaption = "DDL Statements. ",
                    Documentation.FieldDescription = "DDL Statemnents to be executed",
                    Documentation.SampleValues = {
                        " DROP TABLE... CREATE TABLE... "
                    },
                    Formatting.IsMultiLine = true,
                    Formatting.IsCode = true
                ]
            ), //
        UPLOAD_SQL_BATCH_DML as
            (
                type text
                meta
                [
                    Documentation.FieldCaption = "DML Statements ",
                    Documentation.FieldDescription = "DML Statemnents to be executed",
                    Documentation.SampleValues = {
                        " TRUNCATE TABLE... DELETE FROM... "
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
                    "" TRUNCATE TABLE dbo.Sales "",
                    ""NO"" //Function only runs with YES value. Create Parameter link to this or set manually.
                    )",
                    Result = "Single Rows Table. Update Later."
                ]
            }
        ]
in
    Value.ReplaceType(
        PowerQueryToSQLFunction,
        PowerQueryToSQLType
    )