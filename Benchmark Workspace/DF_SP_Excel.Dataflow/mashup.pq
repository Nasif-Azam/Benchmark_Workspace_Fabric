[StagingDefinition = [Kind = "FastCopy"]]
section Section1;
[DataDestinations = {[Definition = [Kind = "Reference", QueryName = "Financial Budget Data_DataDestination", IsNewTarget = true], Settings = [Kind = "Automatic", TypeSettings = [Kind = "Table"]]]}]
shared #"Financial Budget Data" = let
    Source = SharePoint.Files("https://anzir.sharepoint.com/sites/SharePointDataExcel", [ApiVersion = 15]),
    #"Filtered Hidden Files1" = Table.SelectRows(Source, each [Attributes]?[Hidden]? <> true),
    #"Invoke Custom Function1" = Table.AddColumn(#"Filtered Hidden Files1", "Transform File", each #"Transform file"([Content])),
    #"Renamed Columns1" = Table.RenameColumns(#"Invoke Custom Function1", {"Name", "Source.Name"}),
    #"Removed Other Columns1" = Table.SelectColumns(#"Renamed Columns1", {"Source.Name", "Transform File"}),
    #"Expanded Table Column1" = Table.ExpandTableColumn(#"Removed Other Columns1", "Transform File", Table.ColumnNames(#"Transform file"(#"Sample file"))),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded Table Column1",{{"Source.Name", type text}, {"Expense Group", type text}, {"GL", type text}, {"GL Name", type text}, {"Month", type text}, {"Amount", Int64.Type}})
in
    #"Changed Type";
shared #"Sample file" = let
  Source = SharePoint.Files("https://anzir.sharepoint.com/sites/SharePointDataExcel", [ApiVersion = 15]),
  #"Filtered hidden files" = Table.SelectRows(Source, each [Attributes]?[Hidden]? <> true),
  Navigation = #"Filtered hidden files"{[Name = "IT Budget Template FY23.xlsx", #"Folder Path" = "https://anzir.sharepoint.com/sites/SharePointDataExcel/Shared Documents/Financial Data Excel/"]}[Content]
in
  Navigation;
shared Parameter = let
  Parameter = #"Sample file" meta [IsParameterQuery = true, IsParameterQueryRequired = false, Type = type binary, BinaryIdentifier = #"Sample file"]
in
  Parameter;
shared #"Transform Sample file" = // let
//   Source = Excel.Workbook(Parameter, null, true),
//   Navigation = Source{[Item = "Information Technology", Kind = "Sheet"]}[Data]
// in
//   Navigation

let
    Source = Excel.Workbook(Parameter, null, true),
    #"Filtered Rows" = Table.SelectRows(Source, each ([Kind] = "Sheet") and ([Hidden] = false)),
    #"Expanded Data" = Table.ExpandTableColumn(#"Filtered Rows", "Data", {"Column1", "Column2", "Column3", "Column4", "Column5", "Column6", "Column7", "Column8", "Column9", "Column10", "Column11", "Column12", "Column13", "Column14", "Column15", "Column16", "Column17", "Column18", "Column19", "Column20", "Column21", "Column22", "Column23", "Column24", "Column25", "Column26", "Column27", "Column28", "Column29", "Column30", "Column31", "Column32", "Column33", "Column34", "Column35", "Column36", "Column37", "Column38", "Column39", "Column40", "Column41", "Column42", "Column43", "Column44", "Column45", "Column46"}, {"Column1", "Column2", "Column3", "Column4", "Column5", "Column6", "Column7", "Column8", "Column9", "Column10", "Column11", "Column12", "Column13", "Column14", "Column15", "Column16", "Column17", "Column18", "Column19", "Column20", "Column21", "Column22", "Column23", "Column24", "Column25", "Column26", "Column27", "Column28", "Column29", "Column30", "Column31", "Column32", "Column33", "Column34", "Column35", "Column36", "Column37", "Column38", "Column39", "Column40", "Column41", "Column42", "Column43", "Column44", "Column45", "Column46"}),
    #"Removed Columns" = Table.RemoveColumns(#"Expanded Data",{"Name"}),
    #"Removed Top Rows" = Table.Skip(#"Removed Columns",2),
    #"Removed Other Columns" = Table.SelectColumns(#"Removed Top Rows",{"Column1", "Column2", "Column3", "Column18", "Column19", "Column20", "Column21", "Column22", "Column23", "Column24", "Column25", "Column26", "Column27", "Column28", "Column29"}),
    #"Promoted Headers" = Table.PromoteHeaders(#"Removed Other Columns", [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Expense Group", type text}, {"GL", type any}, {"GL Name", type text}, {"Jan", Int64.Type}, {"Feb", Int64.Type}, {"Mar", Int64.Type}, {"Apr", Int64.Type}, {"May", Int64.Type}, {"Jun", Int64.Type}, {"Jul", Int64.Type}, {"Aug", Int64.Type}, {"Sep", Int64.Type}, {"Oct", Int64.Type}, {"Nov", Int64.Type}, {"Dec", Int64.Type}}),
    #"Filled Down" = Table.FillDown(#"Changed Type",{"Expense Group"}),
    #"Filtered Rows1" = Table.SelectRows(#"Filled Down", each ([GL] <> null)),
    #"Unpivoted Other Columns" = Table.UnpivotOtherColumns(#"Filtered Rows1", {"Expense Group", "GL", "GL Name"}, "Attribute", "Value"),
    #"Renamed Columns" = Table.RenameColumns(#"Unpivoted Other Columns",{{"Attribute", "Month"}, {"Value", "Amount"}}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Renamed Columns",{{"Expense Group", type text}, {"GL", type text}, {"GL Name", type text}, {"Month", type text}, {"Amount", Int64.Type}})
in
    #"Changed Type1";
shared #"Transform file" = let
    Source = (Parameter as binary) => let
    Source = Excel.Workbook(Parameter, null, true),
    #"Filtered Rows" = Table.SelectRows(Source, each ([Kind] = "Sheet") and ([Hidden] = false)),
    #"Expanded Data" = Table.ExpandTableColumn(#"Filtered Rows", "Data", {"Column1", "Column2", "Column3", "Column4", "Column5", "Column6", "Column7", "Column8", "Column9", "Column10", "Column11", "Column12", "Column13", "Column14", "Column15", "Column16", "Column17", "Column18", "Column19", "Column20", "Column21", "Column22", "Column23", "Column24", "Column25", "Column26", "Column27", "Column28", "Column29", "Column30", "Column31", "Column32", "Column33", "Column34", "Column35", "Column36", "Column37", "Column38", "Column39", "Column40", "Column41", "Column42", "Column43", "Column44", "Column45", "Column46"}, {"Column1", "Column2", "Column3", "Column4", "Column5", "Column6", "Column7", "Column8", "Column9", "Column10", "Column11", "Column12", "Column13", "Column14", "Column15", "Column16", "Column17", "Column18", "Column19", "Column20", "Column21", "Column22", "Column23", "Column24", "Column25", "Column26", "Column27", "Column28", "Column29", "Column30", "Column31", "Column32", "Column33", "Column34", "Column35", "Column36", "Column37", "Column38", "Column39", "Column40", "Column41", "Column42", "Column43", "Column44", "Column45", "Column46"}),
    #"Removed Columns" = Table.RemoveColumns(#"Expanded Data",{"Name"}),
    #"Removed Top Rows" = Table.Skip(#"Removed Columns",2),
    #"Removed Other Columns" = Table.SelectColumns(#"Removed Top Rows",{"Column1", "Column2", "Column3", "Column18", "Column19", "Column20", "Column21", "Column22", "Column23", "Column24", "Column25", "Column26", "Column27", "Column28", "Column29"}),
    #"Promoted Headers" = Table.PromoteHeaders(#"Removed Other Columns", [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Expense Group", type text}, {"GL", type any}, {"GL Name", type text}, {"Jan", Int64.Type}, {"Feb", Int64.Type}, {"Mar", Int64.Type}, {"Apr", Int64.Type}, {"May", Int64.Type}, {"Jun", Int64.Type}, {"Jul", Int64.Type}, {"Aug", Int64.Type}, {"Sep", Int64.Type}, {"Oct", Int64.Type}, {"Nov", Int64.Type}, {"Dec", Int64.Type}}),
    #"Filled Down" = Table.FillDown(#"Changed Type",{"Expense Group"}),
    #"Filtered Rows1" = Table.SelectRows(#"Filled Down", each ([GL] <> null)),
    #"Unpivoted Other Columns" = Table.UnpivotOtherColumns(#"Filtered Rows1", {"Expense Group", "GL", "GL Name"}, "Attribute", "Value"),
    #"Renamed Columns" = Table.RenameColumns(#"Unpivoted Other Columns",{{"Attribute", "Month"}, {"Value", "Amount"}}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Renamed Columns",{{"Expense Group", type text}, {"GL", type text}, {"GL Name", type text}, {"Month", type text}, {"Amount", Int64.Type}})
in
    #"Changed Type1"
in
    Source;
shared #"Financial Budget Data_DataDestination" = let
  Pattern = Lakehouse.Contents([HierarchicalNavigation = null, CreateNavigationProperties = false, EnableFolding = false]),
  Navigation_1 = Pattern{[workspaceId = "c38149e8-f089-4b53-8c51-b2db1915323e"]}[Data],
  Navigation_2 = Navigation_1{[lakehouseId = "56769383-991d-41aa-bdd4-c09889526219"]}[Data],
  TableNavigation = Navigation_2{[Id = "Financial Budget Data", ItemKind = "Table"]}?[Data]?
in
  TableNavigation;
