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
shared #"IT Budget (2022)" = let
    Source = SharePoint.Files("https://anzir.sharepoint.com/sites/SharePointDataExcel", [ApiVersion = 15]),
    Navigation = Source{[Name = "IT Budget Template FY23.xlsx", #"Folder Path" = "https://anzir.sharepoint.com/sites/SharePointDataExcel/Shared Documents/Financial Data Excel/"]}[Content],
    #"Imported Excel workbook" = Excel.Workbook(Navigation, null, true),
    #"Filtered Rows" = Table.SelectRows(#"Imported Excel workbook", each ([Kind] = "Sheet") and ([Hidden] = false)),
    #"Expanded Data" = Table.ExpandTableColumn(#"Filtered Rows", "Data", {"Column1", "Column2", "Column3", "Column4", "Column5", "Column6", "Column7", "Column8", "Column9", "Column10", "Column11", "Column12", "Column13", "Column14", "Column15", "Column16", "Column17", "Column18", "Column19", "Column20", "Column21", "Column22", "Column23", "Column24", "Column25", "Column26", "Column27", "Column28", "Column29", "Column30", "Column31", "Column32", "Column33", "Column34", "Column35", "Column36", "Column37", "Column38", "Column39", "Column40", "Column41", "Column42", "Column43", "Column44", "Column45", "Column46"}, {"Column1", "Column2", "Column3", "Column4", "Column5", "Column6", "Column7", "Column8", "Column9", "Column10", "Column11", "Column12", "Column13", "Column14", "Column15", "Column16", "Column17", "Column18", "Column19", "Column20", "Column21", "Column22", "Column23", "Column24", "Column25", "Column26", "Column27", "Column28", "Column29", "Column30", "Column31", "Column32", "Column33", "Column34", "Column35", "Column36", "Column37", "Column38", "Column39", "Column40", "Column41", "Column42", "Column43", "Column44", "Column45", "Column46"}),
    #"Removed Columns" = Table.RemoveColumns(#"Expanded Data",{"Name"}),
    #"Removed Other Columns1" = Table.SelectColumns(#"Removed Columns",{"Column1", "Column2", "Column3", "Column4", "Column5", "Column6", "Column7", "Column8", "Column9", "Column10", "Column11", "Column12", "Column13", "Column14", "Column15"}),
    #"Removed Top Rows" = Table.Skip(#"Removed Other Columns1",2),
    #"Promoted Headers" = Table.PromoteHeaders(#"Removed Top Rows", [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Expense Group", type text}, {"GL", type any}, {"GL Name", type text}, {"Jan", Int64.Type}, {"Feb", Int64.Type}, {"Mar", Int64.Type}, {"Apr", Int64.Type}, {"May", Int64.Type}, {"Jun", Int64.Type}, {"Jul", Int64.Type}, {"Aug", Int64.Type}, {"Sep", Int64.Type}, {"Oct", Int64.Type}, {"Nov", Int64.Type}, {"Dec", Int64.Type}}),
    #"Filled Down" = Table.FillDown(#"Changed Type",{"Expense Group"}),
    #"Filtered Rows1" = Table.SelectRows(#"Filled Down", each ([GL] <> null)),
    #"Unpivoted Other Columns" = Table.UnpivotOtherColumns(#"Filtered Rows1", {"Expense Group", "GL", "GL Name"}, "Attribute", "Value"),
    #"Renamed Columns" = Table.RenameColumns(#"Unpivoted Other Columns",{{"Attribute", "Month"}, {"Value", "Amount"}}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Renamed Columns",{{"Expense Group", type text}, {"GL", type text}, {"GL Name", type text}, {"Month", type text}, {"Amount", Int64.Type}}),
    #"Added Custom" = Table.AddColumn(#"Changed Type1", "FY Year", each 2022),
    #"Changed Type2" = Table.TransformColumnTypes(#"Added Custom",{{"FY Year", Int64.Type}}),
    #"Reordered Columns" = Table.ReorderColumns(#"Changed Type2",{"Expense Group", "GL", "GL Name", "Month", "FY Year", "Amount"})
in
    #"Reordered Columns";
shared #"IT Budget (2023)" = let
    Source = SharePoint.Files("https://anzir.sharepoint.com/sites/SharePointDataExcel", [ApiVersion = 15]),
    Navigation = Source{[Name = "IT Budget Template FY23.xlsx", #"Folder Path" = "https://anzir.sharepoint.com/sites/SharePointDataExcel/Shared Documents/Financial Data Excel/"]}[Content],
    #"Imported Excel workbook" = Excel.Workbook(Navigation, null, true),
    #"Filtered Rows" = Table.SelectRows(#"Imported Excel workbook", each ([Kind] = "Sheet") and ([Hidden] = false)),
    #"Expanded Data" = Table.ExpandTableColumn(#"Filtered Rows", "Data", {"Column1", "Column2", "Column3", "Column4", "Column5", "Column6", "Column7", "Column8", "Column9", "Column10", "Column11", "Column12", "Column13", "Column14", "Column15", "Column16", "Column17", "Column18", "Column19", "Column20", "Column21", "Column22", "Column23", "Column24", "Column25", "Column26", "Column27", "Column28", "Column29", "Column30", "Column31", "Column32", "Column33", "Column34", "Column35", "Column36", "Column37", "Column38", "Column39", "Column40", "Column41", "Column42", "Column43", "Column44", "Column45", "Column46"}, {"Column1", "Column2", "Column3", "Column4", "Column5", "Column6", "Column7", "Column8", "Column9", "Column10", "Column11", "Column12", "Column13", "Column14", "Column15", "Column16", "Column17", "Column18", "Column19", "Column20", "Column21", "Column22", "Column23", "Column24", "Column25", "Column26", "Column27", "Column28", "Column29", "Column30", "Column31", "Column32", "Column33", "Column34", "Column35", "Column36", "Column37", "Column38", "Column39", "Column40", "Column41", "Column42", "Column43", "Column44", "Column45", "Column46"}),
    #"Removed Columns" = Table.RemoveColumns(#"Expanded Data",{"Name"}),
    #"Removed Other Columns" = Table.SelectColumns(#"Removed Columns",{"Column1", "Column2", "Column3", "Column18", "Column19", "Column20", "Column21", "Column22", "Column23", "Column24", "Column25", "Column26", "Column27", "Column28", "Column29"}),
    #"Removed Top Rows" = Table.Skip(#"Removed Other Columns",2),
    #"Promoted Headers" = Table.PromoteHeaders(#"Removed Top Rows", [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Expense Group", type text}, {"GL", type any}, {"GL Name", type text}, {"Jan", Int64.Type}, {"Feb", Int64.Type}, {"Mar", Int64.Type}, {"Apr", Int64.Type}, {"May", Int64.Type}, {"Jun", Int64.Type}, {"Jul", Int64.Type}, {"Aug", Int64.Type}, {"Sep", Int64.Type}, {"Oct", Int64.Type}, {"Nov", Int64.Type}, {"Dec", Int64.Type}}),
    #"Filled Down" = Table.FillDown(#"Changed Type",{"Expense Group"}),
    #"Filtered Rows1" = Table.SelectRows(#"Filled Down", each ([GL] <> null)),
    #"Unpivoted Other Columns" = Table.UnpivotOtherColumns(#"Filtered Rows1", {"Expense Group", "GL", "GL Name"}, "Attribute", "Value"),
    #"Renamed Columns" = Table.RenameColumns(#"Unpivoted Other Columns",{{"Attribute", "Month"}, {"Value", "Amount"}}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Renamed Columns",{{"Expense Group", type text}, {"GL", type text}, {"GL Name", type text}, {"Month", type text}, {"Amount", Int64.Type}}),
    #"Added Custom" = Table.AddColumn(#"Changed Type1", "FY Year", each 2023),
    #"Changed Type2" = Table.TransformColumnTypes(#"Added Custom",{{"FY Year", Int64.Type}}),
    #"Reordered Columns" = Table.ReorderColumns(#"Changed Type2",{"Expense Group", "GL", "GL Name", "Month", "FY Year", "Amount"})
in
    #"Reordered Columns";
shared #"IT Budget (Delta)" = let
    Source = SharePoint.Files("https://anzir.sharepoint.com/sites/SharePointDataExcel", [ApiVersion = 15]),
    Navigation = Source{[Name = "IT Budget Template FY23.xlsx", #"Folder Path" = "https://anzir.sharepoint.com/sites/SharePointDataExcel/Shared Documents/Financial Data Excel/"]}[Content],
    #"Imported Excel workbook" = Excel.Workbook(Navigation, null, true),
    #"Filtered Rows" = Table.SelectRows(#"Imported Excel workbook", each ([Kind] = "Sheet") and ([Hidden] = false)),
    #"Expanded Data" = Table.ExpandTableColumn(#"Filtered Rows", "Data", {"Column1", "Column2", "Column3", "Column4", "Column5", "Column6", "Column7", "Column8", "Column9", "Column10", "Column11", "Column12", "Column13", "Column14", "Column15", "Column16", "Column17", "Column18", "Column19", "Column20", "Column21", "Column22", "Column23", "Column24", "Column25", "Column26", "Column27", "Column28", "Column29", "Column30", "Column31", "Column32", "Column33", "Column34", "Column35", "Column36", "Column37", "Column38", "Column39", "Column40", "Column41", "Column42", "Column43", "Column44", "Column45", "Column46"}, {"Column1", "Column2", "Column3", "Column4", "Column5", "Column6", "Column7", "Column8", "Column9", "Column10", "Column11", "Column12", "Column13", "Column14", "Column15", "Column16", "Column17", "Column18", "Column19", "Column20", "Column21", "Column22", "Column23", "Column24", "Column25", "Column26", "Column27", "Column28", "Column29", "Column30", "Column31", "Column32", "Column33", "Column34", "Column35", "Column36", "Column37", "Column38", "Column39", "Column40", "Column41", "Column42", "Column43", "Column44", "Column45", "Column46"}),
    #"Removed Columns" = Table.RemoveColumns(#"Expanded Data",{"Name"}),
    #"Removed Other Columns1" = Table.SelectColumns(#"Removed Columns",{"Column1", "Column2", "Column3", "Column32", "Column33", "Column34", "Column35", "Column36", "Column37", "Column38", "Column39", "Column40", "Column41", "Column42", "Column43"}),
    #"Removed Top Rows" = Table.Skip(#"Removed Other Columns1",2),
    #"Promoted Headers" = Table.PromoteHeaders(#"Removed Top Rows", [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Expense Group", type text}, {"GL", type any}, {"GL Name", type text}, {"Jan", Int64.Type}, {"Feb", Int64.Type}, {"Mar", Int64.Type}, {"Apr", Int64.Type}, {"May", Int64.Type}, {"Jun", Int64.Type}, {"Jul", Int64.Type}, {"Aug", Int64.Type}, {"Sep", Int64.Type}, {"Oct", Int64.Type}, {"Nov", Int64.Type}, {"Dec", Int64.Type}}),
    #"Filled Down" = Table.FillDown(#"Changed Type",{"Expense Group"}),
    #"Filtered Rows1" = Table.SelectRows(#"Filled Down", each ([GL] <> null)),
    #"Unpivoted Other Columns" = Table.UnpivotOtherColumns(#"Filtered Rows1", {"Expense Group", "GL", "GL Name"}, "Attribute", "Value"),
    #"Renamed Columns" = Table.RenameColumns(#"Unpivoted Other Columns",{{"Attribute", "Month"}, {"Value", "Amount"}}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Renamed Columns",{{"Expense Group", type text}, {"GL", type text}, {"GL Name", type text}, {"Month", type text}, {"Amount", Int64.Type}}),
    #"Added Custom" = Table.AddColumn(#"Changed Type1", "FY Year", each "2022-2023"),
    #"Changed Type2" = Table.TransformColumnTypes(#"Added Custom",{{"FY Year", type text}})
in
    #"Changed Type2";
shared #"HR Budget (2022)" = let
    Source = SharePoint.Files("https://anzir.sharepoint.com/sites/SharePointDataExcel", [ApiVersion = 15]),
    Navigation = Source{[Name = "HR Budget Template FY23.xlsx", #"Folder Path" = "https://anzir.sharepoint.com/sites/SharePointDataExcel/Shared Documents/Financial Data Excel/"]}[Content],
    #"Imported Excel workbook" = Excel.Workbook(Navigation, null, true),
    #"Filtered Rows" = Table.SelectRows(#"Imported Excel workbook", each ([Kind] = "Sheet") and ([Hidden] = false)),
    #"Expanded Data" = Table.ExpandTableColumn(#"Filtered Rows", "Data", {"Column1", "Column2", "Column3", "Column4", "Column5", "Column6", "Column7", "Column8", "Column9", "Column10", "Column11", "Column12", "Column13", "Column14", "Column15", "Column16", "Column17", "Column18", "Column19", "Column20", "Column21", "Column22", "Column23", "Column24", "Column25", "Column26", "Column27", "Column28", "Column29", "Column30", "Column31", "Column32", "Column33", "Column34", "Column35", "Column36", "Column37", "Column38", "Column39", "Column40", "Column41", "Column42", "Column43", "Column44", "Column45", "Column46"}, {"Column1", "Column2", "Column3", "Column4", "Column5", "Column6", "Column7", "Column8", "Column9", "Column10", "Column11", "Column12", "Column13", "Column14", "Column15", "Column16", "Column17", "Column18", "Column19", "Column20", "Column21", "Column22", "Column23", "Column24", "Column25", "Column26", "Column27", "Column28", "Column29", "Column30", "Column31", "Column32", "Column33", "Column34", "Column35", "Column36", "Column37", "Column38", "Column39", "Column40", "Column41", "Column42", "Column43", "Column44", "Column45", "Column46"}),
    #"Removed Columns" = Table.RemoveColumns(#"Expanded Data",{"Name"}),
    #"Removed Other Columns1" = Table.SelectColumns(#"Removed Columns",{"Column1", "Column2", "Column3", "Column4", "Column5", "Column6", "Column7", "Column8", "Column9", "Column10", "Column11", "Column12", "Column13", "Column14", "Column15"}),
    #"Removed Top Rows" = Table.Skip(#"Removed Other Columns1",2),
    #"Promoted Headers" = Table.PromoteHeaders(#"Removed Top Rows", [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Expense Group", type text}, {"GL", type any}, {"GL Name", type text}, {"Jan", Int64.Type}, {"Feb", Int64.Type}, {"Mar", Int64.Type}, {"Apr", Int64.Type}, {"May", Int64.Type}, {"Jun", Int64.Type}, {"Jul", Int64.Type}, {"Aug", Int64.Type}, {"Sep", Int64.Type}, {"Oct", Int64.Type}, {"Nov", Int64.Type}, {"Dec", Int64.Type}}),
    #"Filled Down" = Table.FillDown(#"Changed Type",{"Expense Group"}),
    #"Filtered Rows1" = Table.SelectRows(#"Filled Down", each ([GL] <> null)),
    #"Unpivoted Other Columns" = Table.UnpivotOtherColumns(#"Filtered Rows1", {"Expense Group", "GL", "GL Name"}, "Attribute", "Value"),
    #"Renamed Columns" = Table.RenameColumns(#"Unpivoted Other Columns",{{"Attribute", "Month"}, {"Value", "Amount"}}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Renamed Columns",{{"Expense Group", type text}, {"GL", type text}, {"GL Name", type text}, {"Month", type text}, {"Amount", Int64.Type}}),
    #"Added Custom" = Table.AddColumn(#"Changed Type1", "FY Year", each 2022),
    #"Changed Type2" = Table.TransformColumnTypes(#"Added Custom",{{"FY Year", Int64.Type}}),
    #"Reordered Columns" = Table.ReorderColumns(#"Changed Type2",{"Expense Group", "GL", "GL Name", "Month", "FY Year", "Amount"})
in
    #"Reordered Columns";
shared #"HR Budget (2023)" = let
    Source = SharePoint.Files("https://anzir.sharepoint.com/sites/SharePointDataExcel", [ApiVersion = 15]),
    Navigation = Source{[Name = "HR Budget Template FY23.xlsx", #"Folder Path" = "https://anzir.sharepoint.com/sites/SharePointDataExcel/Shared Documents/Financial Data Excel/"]}[Content],
    #"Imported Excel workbook" = Excel.Workbook(Navigation, null, true),
    #"Filtered Rows" = Table.SelectRows(#"Imported Excel workbook", each ([Kind] = "Sheet") and ([Hidden] = false)),
    #"Expanded Data" = Table.ExpandTableColumn(#"Filtered Rows", "Data", {"Column1", "Column2", "Column3", "Column4", "Column5", "Column6", "Column7", "Column8", "Column9", "Column10", "Column11", "Column12", "Column13", "Column14", "Column15", "Column16", "Column17", "Column18", "Column19", "Column20", "Column21", "Column22", "Column23", "Column24", "Column25", "Column26", "Column27", "Column28", "Column29", "Column30", "Column31", "Column32", "Column33", "Column34", "Column35", "Column36", "Column37", "Column38", "Column39", "Column40", "Column41", "Column42", "Column43", "Column44", "Column45", "Column46"}, {"Column1", "Column2", "Column3", "Column4", "Column5", "Column6", "Column7", "Column8", "Column9", "Column10", "Column11", "Column12", "Column13", "Column14", "Column15", "Column16", "Column17", "Column18", "Column19", "Column20", "Column21", "Column22", "Column23", "Column24", "Column25", "Column26", "Column27", "Column28", "Column29", "Column30", "Column31", "Column32", "Column33", "Column34", "Column35", "Column36", "Column37", "Column38", "Column39", "Column40", "Column41", "Column42", "Column43", "Column44", "Column45", "Column46"}),
    #"Removed Columns" = Table.RemoveColumns(#"Expanded Data",{"Name"}),
    #"Removed Other Columns" = Table.SelectColumns(#"Removed Columns",{"Column1", "Column2", "Column3", "Column18", "Column19", "Column20", "Column21", "Column22", "Column23", "Column24", "Column25", "Column26", "Column27", "Column28", "Column29"}),
    #"Removed Top Rows" = Table.Skip(#"Removed Other Columns",2),
    #"Promoted Headers" = Table.PromoteHeaders(#"Removed Top Rows", [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Expense Group", type text}, {"GL", type any}, {"GL Name", type text}, {"Jan", Int64.Type}, {"Feb", Int64.Type}, {"Mar", Int64.Type}, {"Apr", Int64.Type}, {"May", Int64.Type}, {"Jun", Int64.Type}, {"Jul", Int64.Type}, {"Aug", Int64.Type}, {"Sep", Int64.Type}, {"Oct", Int64.Type}, {"Nov", Int64.Type}, {"Dec", Int64.Type}}),
    #"Filled Down" = Table.FillDown(#"Changed Type",{"Expense Group"}),
    #"Filtered Rows1" = Table.SelectRows(#"Filled Down", each ([GL] <> null)),
    #"Unpivoted Other Columns" = Table.UnpivotOtherColumns(#"Filtered Rows1", {"Expense Group", "GL", "GL Name"}, "Attribute", "Value"),
    #"Renamed Columns" = Table.RenameColumns(#"Unpivoted Other Columns",{{"Attribute", "Month"}, {"Value", "Amount"}}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Renamed Columns",{{"Expense Group", type text}, {"GL", type text}, {"GL Name", type text}, {"Month", type text}, {"Amount", Int64.Type}}),
    #"Added Custom" = Table.AddColumn(#"Changed Type1", "FY Year", each 2023),
    #"Changed Type2" = Table.TransformColumnTypes(#"Added Custom",{{"FY Year", Int64.Type}}),
    #"Reordered Columns" = Table.ReorderColumns(#"Changed Type2",{"Expense Group", "GL", "GL Name", "Month", "FY Year", "Amount"})
in
    #"Reordered Columns";
shared #"HR Budget (Delta)" = let
    Source = SharePoint.Files("https://anzir.sharepoint.com/sites/SharePointDataExcel", [ApiVersion = 15]),
    Navigation = Source{[Name = "HR Budget Template FY23.xlsx", #"Folder Path" = "https://anzir.sharepoint.com/sites/SharePointDataExcel/Shared Documents/Financial Data Excel/"]}[Content],
    #"Imported Excel workbook" = Excel.Workbook(Navigation, null, true),
    #"Filtered Rows" = Table.SelectRows(#"Imported Excel workbook", each ([Kind] = "Sheet") and ([Hidden] = false)),
    #"Expanded Data" = Table.ExpandTableColumn(#"Filtered Rows", "Data", {"Column1", "Column2", "Column3", "Column4", "Column5", "Column6", "Column7", "Column8", "Column9", "Column10", "Column11", "Column12", "Column13", "Column14", "Column15", "Column16", "Column17", "Column18", "Column19", "Column20", "Column21", "Column22", "Column23", "Column24", "Column25", "Column26", "Column27", "Column28", "Column29", "Column30", "Column31", "Column32", "Column33", "Column34", "Column35", "Column36", "Column37", "Column38", "Column39", "Column40", "Column41", "Column42", "Column43", "Column44", "Column45", "Column46"}, {"Column1", "Column2", "Column3", "Column4", "Column5", "Column6", "Column7", "Column8", "Column9", "Column10", "Column11", "Column12", "Column13", "Column14", "Column15", "Column16", "Column17", "Column18", "Column19", "Column20", "Column21", "Column22", "Column23", "Column24", "Column25", "Column26", "Column27", "Column28", "Column29", "Column30", "Column31", "Column32", "Column33", "Column34", "Column35", "Column36", "Column37", "Column38", "Column39", "Column40", "Column41", "Column42", "Column43", "Column44", "Column45", "Column46"}),
    #"Removed Columns" = Table.RemoveColumns(#"Expanded Data",{"Name"}),
    #"Removed Other Columns1" = Table.SelectColumns(#"Removed Columns",{"Column1", "Column2", "Column3", "Column32", "Column33", "Column34", "Column35", "Column36", "Column37", "Column38", "Column39", "Column40", "Column41", "Column42", "Column43"}),
    #"Removed Top Rows" = Table.Skip(#"Removed Other Columns1",2),
    #"Promoted Headers" = Table.PromoteHeaders(#"Removed Top Rows", [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Expense Group", type text}, {"GL", type any}, {"GL Name", type text}, {"Jan", Int64.Type}, {"Feb", Int64.Type}, {"Mar", Int64.Type}, {"Apr", Int64.Type}, {"May", Int64.Type}, {"Jun", Int64.Type}, {"Jul", Int64.Type}, {"Aug", Int64.Type}, {"Sep", Int64.Type}, {"Oct", Int64.Type}, {"Nov", Int64.Type}, {"Dec", Int64.Type}}),
    #"Filled Down" = Table.FillDown(#"Changed Type",{"Expense Group"}),
    #"Filtered Rows1" = Table.SelectRows(#"Filled Down", each ([GL] <> null)),
    #"Unpivoted Other Columns" = Table.UnpivotOtherColumns(#"Filtered Rows1", {"Expense Group", "GL", "GL Name"}, "Attribute", "Value"),
    #"Renamed Columns" = Table.RenameColumns(#"Unpivoted Other Columns",{{"Attribute", "Month"}, {"Value", "Amount"}}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Renamed Columns",{{"Expense Group", type text}, {"GL", type text}, {"GL Name", type text}, {"Month", type text}, {"Amount", Int64.Type}}),
    #"Added Custom" = Table.AddColumn(#"Changed Type1", "FY Year", each "2022-2023"),
    #"Changed Type2" = Table.TransformColumnTypes(#"Added Custom",{{"FY Year", type text}})
in
    #"Changed Type2";
shared #"HR Budget Informations" = let
    Source = Table.NestedJoin(#"HR Budget (2022)", {"GL"}, #"HR Budget (2023)", {"GL"}, "HR Budget (2023)", JoinKind.LeftOuter),
    #"Expanded Transform Sample File (2023)" = Table.ExpandTableColumn(Source, "HR Budget (2023)", {"Amount", "FY Year", "Month"}, {"Amount.1", "FY Year.1", "Month.1"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Expanded Transform Sample File (2023)",{"Expense Group", "GL", "GL Name", "Month", "FY Year", "Amount", "Month.1", "FY Year.1", "Amount.1"}),
    #"Merged Columns" = Table.CombineColumns(Table.TransformColumnTypes(#"Reordered Columns", {{"FY Year", type text}}, "en-US"),{"Month", "FY Year"},Combiner.CombineTextByDelimiter("-", QuoteStyle.None),"Month-Year 2022"),
    #"Merged Columns1" = Table.CombineColumns(Table.TransformColumnTypes(#"Merged Columns", {{"FY Year.1", type text}}, "en-US"),{"Month.1", "FY Year.1"},Combiner.CombineTextByDelimiter("-", QuoteStyle.None),"Month-Year 2023"),
    #"Renamed Columns1" = Table.RenameColumns(#"Merged Columns1",{{"Amount", "Amount of 2022"}, {"Amount.1", "Amount of 2023"}}),
    #"Added Custom" = Table.AddColumn(#"Renamed Columns1", "Total Ammount", each [Amount of 2022]+[Amount of 2023]),
    #"Changed Type" = Table.TransformColumnTypes(#"Added Custom",{{"Total Ammount", type number}}),
    #"Merged Queries" = Table.NestedJoin(#"Changed Type", {"GL"}, #"HR Budget (Delta)", {"GL"}, "HR Budget (Delta)", JoinKind.LeftOuter),
    #"Expanded Transform Sample File (Delta 22-23)" = Table.ExpandTableColumn(#"Merged Queries", "HR Budget (Delta)", {"Amount"}, {"Amount"}),
    #"Renamed Columns" = Table.RenameColumns(#"Expanded Transform Sample File (Delta 22-23)",{{"Amount", "Total Dela Ammount"}})
in
    #"Renamed Columns";
shared #"Accounting Budget (2022)" = let
    Source = SharePoint.Files("https://anzir.sharepoint.com/sites/SharePointDataExcel", [ApiVersion = 15]),
    Navigation = Source{[Name = "Accounting Budget Template FY23.xlsx", #"Folder Path" = "https://anzir.sharepoint.com/sites/SharePointDataExcel/Shared Documents/Financial Data Excel/"]}[Content],
    #"Imported Excel workbook" = Excel.Workbook(Navigation, null, true),
    #"Filtered Rows" = Table.SelectRows(#"Imported Excel workbook", each ([Kind] = "Sheet") and ([Hidden] = false)),
    #"Expanded Data" = Table.ExpandTableColumn(#"Filtered Rows", "Data", {"Column1", "Column2", "Column3", "Column4", "Column5", "Column6", "Column7", "Column8", "Column9", "Column10", "Column11", "Column12", "Column13", "Column14", "Column15", "Column16", "Column17", "Column18", "Column19", "Column20", "Column21", "Column22", "Column23", "Column24", "Column25", "Column26", "Column27", "Column28", "Column29", "Column30", "Column31", "Column32", "Column33", "Column34", "Column35", "Column36", "Column37", "Column38", "Column39", "Column40", "Column41", "Column42", "Column43", "Column44", "Column45", "Column46"}, {"Column1", "Column2", "Column3", "Column4", "Column5", "Column6", "Column7", "Column8", "Column9", "Column10", "Column11", "Column12", "Column13", "Column14", "Column15", "Column16", "Column17", "Column18", "Column19", "Column20", "Column21", "Column22", "Column23", "Column24", "Column25", "Column26", "Column27", "Column28", "Column29", "Column30", "Column31", "Column32", "Column33", "Column34", "Column35", "Column36", "Column37", "Column38", "Column39", "Column40", "Column41", "Column42", "Column43", "Column44", "Column45", "Column46"}),
    #"Removed Columns" = Table.RemoveColumns(#"Expanded Data",{"Name"}),
    #"Removed Other Columns1" = Table.SelectColumns(#"Removed Columns",{"Column1", "Column2", "Column3", "Column4", "Column5", "Column6", "Column7", "Column8", "Column9", "Column10", "Column11", "Column12", "Column13", "Column14", "Column15"}),
    #"Removed Top Rows" = Table.Skip(#"Removed Other Columns1",2),
    #"Promoted Headers" = Table.PromoteHeaders(#"Removed Top Rows", [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Expense Group", type text}, {"GL", type any}, {"GL Name", type text}, {"Jan", Int64.Type}, {"Feb", Int64.Type}, {"Mar", Int64.Type}, {"Apr", Int64.Type}, {"May", Int64.Type}, {"Jun", Int64.Type}, {"Jul", Int64.Type}, {"Aug", Int64.Type}, {"Sep", Int64.Type}, {"Oct", Int64.Type}, {"Nov", Int64.Type}, {"Dec", Int64.Type}}),
    #"Filled Down" = Table.FillDown(#"Changed Type",{"Expense Group"}),
    #"Filtered Rows1" = Table.SelectRows(#"Filled Down", each ([GL] <> null)),
    #"Unpivoted Other Columns" = Table.UnpivotOtherColumns(#"Filtered Rows1", {"Expense Group", "GL", "GL Name"}, "Attribute", "Value"),
    #"Renamed Columns" = Table.RenameColumns(#"Unpivoted Other Columns",{{"Attribute", "Month"}, {"Value", "Amount"}}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Renamed Columns",{{"Expense Group", type text}, {"GL", type text}, {"GL Name", type text}, {"Month", type text}, {"Amount", Int64.Type}}),
    #"Added Custom" = Table.AddColumn(#"Changed Type1", "FY Year", each 2022),
    #"Changed Type2" = Table.TransformColumnTypes(#"Added Custom",{{"FY Year", Int64.Type}}),
    #"Reordered Columns" = Table.ReorderColumns(#"Changed Type2",{"Expense Group", "GL", "GL Name", "Month", "FY Year", "Amount"})
in
    #"Reordered Columns";
shared #"Accounting Budget (2023)" = let
    Source = SharePoint.Files("https://anzir.sharepoint.com/sites/SharePointDataExcel", [ApiVersion = 15]),
    Navigation = Source{[Name = "Accounting Budget Template FY23.xlsx", #"Folder Path" = "https://anzir.sharepoint.com/sites/SharePointDataExcel/Shared Documents/Financial Data Excel/"]}[Content],
    #"Imported Excel workbook" = Excel.Workbook(Navigation, null, true),
    #"Filtered Rows" = Table.SelectRows(#"Imported Excel workbook", each ([Kind] = "Sheet") and ([Hidden] = false)),
    #"Expanded Data" = Table.ExpandTableColumn(#"Filtered Rows", "Data", {"Column1", "Column2", "Column3", "Column4", "Column5", "Column6", "Column7", "Column8", "Column9", "Column10", "Column11", "Column12", "Column13", "Column14", "Column15", "Column16", "Column17", "Column18", "Column19", "Column20", "Column21", "Column22", "Column23", "Column24", "Column25", "Column26", "Column27", "Column28", "Column29", "Column30", "Column31", "Column32", "Column33", "Column34", "Column35", "Column36", "Column37", "Column38", "Column39", "Column40", "Column41", "Column42", "Column43", "Column44", "Column45", "Column46"}, {"Column1", "Column2", "Column3", "Column4", "Column5", "Column6", "Column7", "Column8", "Column9", "Column10", "Column11", "Column12", "Column13", "Column14", "Column15", "Column16", "Column17", "Column18", "Column19", "Column20", "Column21", "Column22", "Column23", "Column24", "Column25", "Column26", "Column27", "Column28", "Column29", "Column30", "Column31", "Column32", "Column33", "Column34", "Column35", "Column36", "Column37", "Column38", "Column39", "Column40", "Column41", "Column42", "Column43", "Column44", "Column45", "Column46"}),
    #"Removed Columns" = Table.RemoveColumns(#"Expanded Data",{"Name"}),
    #"Removed Other Columns" = Table.SelectColumns(#"Removed Columns",{"Column1", "Column2", "Column3", "Column18", "Column19", "Column20", "Column21", "Column22", "Column23", "Column24", "Column25", "Column26", "Column27", "Column28", "Column29"}),
    #"Removed Top Rows" = Table.Skip(#"Removed Other Columns",2),
    #"Promoted Headers" = Table.PromoteHeaders(#"Removed Top Rows", [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Expense Group", type text}, {"GL", type any}, {"GL Name", type text}, {"Jan", Int64.Type}, {"Feb", Int64.Type}, {"Mar", Int64.Type}, {"Apr", Int64.Type}, {"May", Int64.Type}, {"Jun", Int64.Type}, {"Jul", Int64.Type}, {"Aug", Int64.Type}, {"Sep", Int64.Type}, {"Oct", Int64.Type}, {"Nov", Int64.Type}, {"Dec", Int64.Type}}),
    #"Filled Down" = Table.FillDown(#"Changed Type",{"Expense Group"}),
    #"Filtered Rows1" = Table.SelectRows(#"Filled Down", each ([GL] <> null)),
    #"Unpivoted Other Columns" = Table.UnpivotOtherColumns(#"Filtered Rows1", {"Expense Group", "GL", "GL Name"}, "Attribute", "Value"),
    #"Renamed Columns" = Table.RenameColumns(#"Unpivoted Other Columns",{{"Attribute", "Month"}, {"Value", "Amount"}}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Renamed Columns",{{"Expense Group", type text}, {"GL", type text}, {"GL Name", type text}, {"Month", type text}, {"Amount", Int64.Type}}),
    #"Added Custom" = Table.AddColumn(#"Changed Type1", "FY Year", each 2023),
    #"Changed Type2" = Table.TransformColumnTypes(#"Added Custom",{{"FY Year", Int64.Type}}),
    #"Reordered Columns" = Table.ReorderColumns(#"Changed Type2",{"Expense Group", "GL", "GL Name", "Month", "FY Year", "Amount"})
in
    #"Reordered Columns";
shared #"Accounting Budget (Delta)" = let
    Source = SharePoint.Files("https://anzir.sharepoint.com/sites/SharePointDataExcel", [ApiVersion = 15]),
    Navigation = Source{[Name = "Accounting Budget Template FY23.xlsx", #"Folder Path" = "https://anzir.sharepoint.com/sites/SharePointDataExcel/Shared Documents/Financial Data Excel/"]}[Content],
    #"Imported Excel workbook" = Excel.Workbook(Navigation, null, true),
    #"Filtered Rows" = Table.SelectRows(#"Imported Excel workbook", each ([Kind] = "Sheet") and ([Hidden] = false)),
    #"Expanded Data" = Table.ExpandTableColumn(#"Filtered Rows", "Data", {"Column1", "Column2", "Column3", "Column4", "Column5", "Column6", "Column7", "Column8", "Column9", "Column10", "Column11", "Column12", "Column13", "Column14", "Column15", "Column16", "Column17", "Column18", "Column19", "Column20", "Column21", "Column22", "Column23", "Column24", "Column25", "Column26", "Column27", "Column28", "Column29", "Column30", "Column31", "Column32", "Column33", "Column34", "Column35", "Column36", "Column37", "Column38", "Column39", "Column40", "Column41", "Column42", "Column43", "Column44", "Column45", "Column46"}, {"Column1", "Column2", "Column3", "Column4", "Column5", "Column6", "Column7", "Column8", "Column9", "Column10", "Column11", "Column12", "Column13", "Column14", "Column15", "Column16", "Column17", "Column18", "Column19", "Column20", "Column21", "Column22", "Column23", "Column24", "Column25", "Column26", "Column27", "Column28", "Column29", "Column30", "Column31", "Column32", "Column33", "Column34", "Column35", "Column36", "Column37", "Column38", "Column39", "Column40", "Column41", "Column42", "Column43", "Column44", "Column45", "Column46"}),
    #"Removed Columns" = Table.RemoveColumns(#"Expanded Data",{"Name"}),
    #"Removed Other Columns1" = Table.SelectColumns(#"Removed Columns",{"Column1", "Column2", "Column3", "Column32", "Column33", "Column34", "Column35", "Column36", "Column37", "Column38", "Column39", "Column40", "Column41", "Column42", "Column43"}),
    #"Removed Top Rows" = Table.Skip(#"Removed Other Columns1",2),
    #"Promoted Headers" = Table.PromoteHeaders(#"Removed Top Rows", [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Expense Group", type text}, {"GL", type any}, {"GL Name", type text}, {"Jan", Int64.Type}, {"Feb", Int64.Type}, {"Mar", Int64.Type}, {"Apr", Int64.Type}, {"May", Int64.Type}, {"Jun", Int64.Type}, {"Jul", Int64.Type}, {"Aug", Int64.Type}, {"Sep", Int64.Type}, {"Oct", Int64.Type}, {"Nov", Int64.Type}, {"Dec", Int64.Type}}),
    #"Filled Down" = Table.FillDown(#"Changed Type",{"Expense Group"}),
    #"Filtered Rows1" = Table.SelectRows(#"Filled Down", each ([GL] <> null)),
    #"Unpivoted Other Columns" = Table.UnpivotOtherColumns(#"Filtered Rows1", {"Expense Group", "GL", "GL Name"}, "Attribute", "Value"),
    #"Renamed Columns" = Table.RenameColumns(#"Unpivoted Other Columns",{{"Attribute", "Month"}, {"Value", "Amount"}}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Renamed Columns",{{"Expense Group", type text}, {"GL", type text}, {"GL Name", type text}, {"Month", type text}, {"Amount", Int64.Type}}),
    #"Added Custom" = Table.AddColumn(#"Changed Type1", "FY Year", each "2022-2023"),
    #"Changed Type2" = Table.TransformColumnTypes(#"Added Custom",{{"FY Year", type text}})
in
    #"Changed Type2";
shared #"Accounting Budget Informations" = let
    Source = Table.NestedJoin(#"Accounting Budget (2022)", {"GL"}, #"Accounting Budget (2023)", {"GL"}, "Accounting Budget (2023)", JoinKind.LeftOuter),
    #"Expanded Transform Sample File (2023)" = Table.ExpandTableColumn(Source, "Accounting Budget (2023)", {"Amount", "FY Year", "Month"}, {"Amount.1", "FY Year.1", "Month.1"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Expanded Transform Sample File (2023)",{"Expense Group", "GL", "GL Name", "Month", "FY Year", "Amount", "Month.1", "FY Year.1", "Amount.1"}),
    #"Merged Columns" = Table.CombineColumns(Table.TransformColumnTypes(#"Reordered Columns", {{"FY Year", type text}}, "en-US"),{"Month", "FY Year"},Combiner.CombineTextByDelimiter("-", QuoteStyle.None),"Month-Year 2022"),
    #"Merged Columns1" = Table.CombineColumns(Table.TransformColumnTypes(#"Merged Columns", {{"FY Year.1", type text}}, "en-US"),{"Month.1", "FY Year.1"},Combiner.CombineTextByDelimiter("-", QuoteStyle.None),"Month-Year 2023"),
    #"Renamed Columns1" = Table.RenameColumns(#"Merged Columns1",{{"Amount", "Amount of 2022"}, {"Amount.1", "Amount of 2023"}}),
    #"Added Custom" = Table.AddColumn(#"Renamed Columns1", "Total Ammount", each [Amount of 2022]+[Amount of 2023]),
    #"Changed Type" = Table.TransformColumnTypes(#"Added Custom",{{"Total Ammount", type number}}),
    #"Merged Queries" = Table.NestedJoin(#"Changed Type", {"GL"}, #"Accounting Budget (Delta)", {"GL"}, "Accounting Budget (Delta)", JoinKind.LeftOuter),
    #"Expanded Transform Sample File (Delta 22-23)" = Table.ExpandTableColumn(#"Merged Queries", "Accounting Budget (Delta)", {"Amount"}, {"Amount"}),
    #"Renamed Columns" = Table.RenameColumns(#"Expanded Transform Sample File (Delta 22-23)",{{"Amount", "Total Dela Ammount"}})
in
    #"Renamed Columns";
shared #"IT Budget Informations" = let
    Source = Table.NestedJoin(#"IT Budget (2022)", {"GL"}, #"IT Budget (2023)", {"GL"}, "IT Budget (2023)", JoinKind.LeftOuter),
    #"Expanded Transform Sample File (2023)" = Table.ExpandTableColumn(Source, "IT Budget (2023)", {"Amount", "FY Year", "Month"}, {"Amount.1", "FY Year.1", "Month.1"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Expanded Transform Sample File (2023)",{"Expense Group", "GL", "GL Name", "Month", "FY Year", "Amount", "Month.1", "FY Year.1", "Amount.1"}),
    #"Merged Columns" = Table.CombineColumns(Table.TransformColumnTypes(#"Reordered Columns", {{"FY Year", type text}}, "en-US"),{"Month", "FY Year"},Combiner.CombineTextByDelimiter("-", QuoteStyle.None),"Month-Year 2022"),
    #"Merged Columns1" = Table.CombineColumns(Table.TransformColumnTypes(#"Merged Columns", {{"FY Year.1", type text}}, "en-US"),{"Month.1", "FY Year.1"},Combiner.CombineTextByDelimiter("-", QuoteStyle.None),"Month-Year 2023"),
    #"Renamed Columns1" = Table.RenameColumns(#"Merged Columns1",{{"Amount", "Amount of 2022"}, {"Amount.1", "Amount of 2023"}}),
    #"Added Custom" = Table.AddColumn(#"Renamed Columns1", "Total Ammount", each [Amount of 2022]+[Amount of 2023]),
    #"Changed Type" = Table.TransformColumnTypes(#"Added Custom",{{"Total Ammount", type number}}),
    #"Merged Queries" = Table.NestedJoin(#"Changed Type", {"GL"}, #"IT Budget (Delta)", {"GL"}, "IT Budget (Delta)", JoinKind.LeftOuter),
    #"Expanded Transform Sample File (Delta 22-23)" = Table.ExpandTableColumn(#"Merged Queries", "IT Budget (Delta)", {"Amount"}, {"Amount"}),
    #"Renamed Columns" = Table.RenameColumns(#"Expanded Transform Sample File (Delta 22-23)",{{"Amount", "Total Dela Ammount"}})
in
    #"Renamed Columns";
shared #"Financial Budget Informations" = let
  Source = Table.Combine({#"IT Budget Informations", #"Accounting Budget Informations", #"HR Budget Informations"})
in
  Source;
shared #"Financial Budget Data_DataDestination" = let
  Pattern = Lakehouse.Contents([HierarchicalNavigation = null, CreateNavigationProperties = false, EnableFolding = false]),
  Navigation_1 = Pattern{[workspaceId = "c38149e8-f089-4b53-8c51-b2db1915323e"]}[Data],
  Navigation_2 = Navigation_1{[lakehouseId = "56769383-991d-41aa-bdd4-c09889526219"]}[Data],
  TableNavigation = Navigation_2{[Id = "Financial Budget Data", ItemKind = "Table"]}?[Data]?
in
  TableNavigation;
