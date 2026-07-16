// Power Query (M) — paste in Power BI Desktop:
// Home → Transform data → New Source → Blank Query → Advanced Editor → paste this
// Then update the file path on the Source line.

let
    // 1) Load Excel
    Source = Excel.Workbook(
        File.Contents("C:/Users/raviranjan/Downloads/data-18.xlsx"),
        null,
        true
    ),
    Export_Sheet = Source{[Item="Export", Kind="Sheet"]}[Data],

    // 2) Promote first row as headers (dates across columns)
    Promoted = Table.PromoteHeaders(Export_Sheet, [PromoteAllScalars=true]),

    // 3) Keep only metric rows (drop FormattedFullDate / Settlement Type / filter note rows if present)
    // First column is the metric name — rename it
    Cols = Table.ColumnNames(Promoted),
    MetricCol = Cols{0},
    RenamedMetric = Table.RenameColumns(Promoted, {{MetricCol, "Metric"}}),

    // 4) Unpivot all date columns → long format
    Unpivoted = Table.UnpivotOtherColumns(RenamedMetric, {"Metric"}, "DateRaw", "Amount"),

    // 5) Clean types
    Typed = Table.TransformColumnTypes(
        Unpivoted,
        {{"Metric", type text}, {"Amount", type number}}
    ),

    // 6) Parse Date; keep "Total" out of daily checks
    AddDateLabel = Table.AddColumn(
        Typed,
        "DateLabel",
        each Text.From([DateRaw]),
        type text
    ),
    AddIsTotal = Table.AddColumn(
        AddDateLabel,
        "IsTotal",
        each Text.Upper(Text.Trim([DateLabel])) = "TOTAL",
        type logical
    ),
    AddDate = Table.AddColumn(
        AddIsTotal,
        "Date",
        each
            if [IsTotal] then null
            else
                try Date.From([DateRaw])
                otherwise
                    try Date.From(DateTime.From([DateRaw]))
                    otherwise null,
        type date
    ),

    // 7) Daily only
    DailyOnly = Table.SelectRows(AddDate, each [IsTotal] = false and [Date] <> null),

    // 8) Pivot metrics to columns per day
    Pivoted = Table.Pivot(
        Table.SelectColumns(DailyOnly, {"Date", "Metric", "Amount"}),
        List.Distinct(DailyOnly[Metric]),
        "Metric",
        "Amount",
        List.Sum
    ),

    // 9) Fill blanks with 0 for numeric metrics used in the check
    MetricNames = {
        "Starting Balance",
        "Collected",
        "Collections",
        "Refunds",
        "Chargeback/Chargeback Repost",
        "Returns",
        "Deduction",
        "Reimbursement",
        "Less: Payments made at Club",
        "Deposit",
        "Ending Balance"
    },
    EnsureCols = List.Accumulate(
        MetricNames,
        Pivoted,
        (state, col) =>
            if Table.HasColumns(state, {col}) then state
            else Table.AddColumn(state, col, each null, type number)
    ),
    ReplaceNulls = Table.ReplaceValue(
        EnsureCols,
        null,
        0,
        Replacer.ReplaceValue,
        MetricNames
    ),

    // 10) Balance check columns
    WithCalc = Table.AddColumn(
        ReplaceNulls,
        "Calculated Ending",
        each
            [Starting Balance]
            + [Collected]
            + [Collections]
            + [Refunds]
            + [#"Chargeback/Chargeback Repost"]
            + [Returns]
            + [Deduction]
            + [Reimbursement]
            + [#"Less: Payments made at Club"]
            + [Deposit],
        type number
    ),
    WithDiff = Table.AddColumn(
        WithCalc,
        "Difference",
        each [Calculated Ending] - [Ending Balance],
        type number
    ),
    WithStatus = Table.AddColumn(
        WithDiff,
        "Status",
        each if Number.Abs([Difference]) < 0.02 then "Match" else "Mismatch",
        type text
    ),
    WithAbsDiff = Table.AddColumn(
        WithStatus,
        "Abs Difference",
        each Number.Abs([Difference]),
        type number
    )
in
    WithAbsDiff
