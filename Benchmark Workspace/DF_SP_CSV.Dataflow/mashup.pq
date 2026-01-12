[StagingDefinition = [Kind = "FastCopy"]]
section Section1;
[DataDestinations = {[Definition = [Kind = "Reference", QueryName = "CSV Financial Data_DataDestination", IsNewTarget = true], Settings = [Kind = "Automatic", TypeSettings = [Kind = "Table"]]]}]
shared #"CSV Financial Data" = let
  Source = SharePoint.Files("https://anzir.sharepoint.com/sites/SharePointData", [ApiVersion = 15]),
  #"Filtered hidden files" = Table.SelectRows(Source, each [Attributes]?[Hidden]? <> true),
  #"Invoke custom function" = Table.AddColumn(#"Filtered hidden files", "Transform file", each #"Transform file"([Content])),
  #"Renamed columns" = Table.RenameColumns(#"Invoke custom function", {{"Name", "Source.Name"}}),
  #"Removed other columns" = Table.SelectColumns(#"Renamed columns", {"Source.Name", "Transform file"}),
  #"Expanded table column" = Table.ExpandTableColumn(#"Removed other columns", "Transform file", Table.ColumnNames(#"Transform file"(#"Sample file"))),
  #"Changed column type 1" = Table.TransformColumnTypes(#"Expanded table column", {{"DateKey", type date}, {"OrganizationKey", Int64.Type}, {"DepartmentKey", Int64.Type}, {" AccountKey ", Int64.Type}, {"Amount_LocalCCY", type number}}),
  #"Removed columns" = Table.RemoveColumns(#"Changed column type 1", {"Source.Name"}),
  #"Renamed columns 1" = Table.RenameColumns(#"Removed columns", {{" AccountKey ", "AccountKey"}, {"Amount_LocalCCY", "Amount Local CCY"}}),
  AddedYearColumn = Table.AddColumn(#"Renamed columns 1", "Year", each Date.Year([DateKey]), Int64.Type),
  AddedMonthColumn = Table.AddColumn(#"AddedYearColumn", "Month", each Date.Month([DateKey]), Int64.Type),
  AddedCustomColumn = Table.AddColumn(#"AddedMonthColumn", "Amount BDT", each [Amount Local CCY] * 120),
  #"Changed column type 2" = Table.TransformColumnTypes(AddedCustomColumn, {{"Amount BDT", type number}}),
      AddedCategoryColumn = Table.AddColumn(#"Changed column type 2", "Amount Category USD", each if [Amount Local CCY] > 1000 then "High" else if [Amount Local CCY] >= 501 then "Medium" else "Low"),
  #"Changed column type 3" = Table.TransformColumnTypes(AddedCategoryColumn, {{"Amount Category USD", type text}, {"DateKey", type date}, {"OrganizationKey", type text}})
in
  #"Changed column type 3";
shared #"Sample file" = let
  Source = Source,
  #"Filtered hidden files" = Table.SelectRows(Source, each [Attributes]?[Hidden]? <> true),
  Navigation = #"Filtered hidden files"{0}[Content]
in
  Navigation;
shared Parameter = let
  Parameter = #"Sample file" meta [IsParameterQuery = true, IsParameterQueryRequired = false, Type = type binary, BinaryIdentifier = #"Sample file"]
in
  Parameter;
shared #"Transform Sample file" = let
  Source = Csv.Document(Parameter, [Delimiter = ",", Columns = 5, QuoteStyle = QuoteStyle.None]),
  #"Removed top rows" = Table.Skip(Source, 4),
  #"Promoted headers" = Table.PromoteHeaders(#"Removed top rows", [PromoteAllScalars = true]),
  #"Changed column type" = Table.TransformColumnTypes(#"Promoted headers", {{"DateKey", type date}, {"OrganizationKey", Int64.Type}, {"DepartmentKey", Int64.Type}, {" AccountKey ", Int64.Type}, {"Amount_LocalCCY", type number}})
in
  #"Changed column type";
[FunctionQueryBinding = "{""exemplarFormulaName"":""Transform Sample file""}"]
shared #"Transform file" = (Parameter as binary) => let
  Source = Csv.Document(Parameter, [Delimiter = ",", Columns = 5, QuoteStyle = QuoteStyle.None]),
  #"Removed top rows" = Table.Skip(Source, 4),
  #"Promoted headers" = Table.PromoteHeaders(#"Removed top rows", [PromoteAllScalars = true]),
  #"Changed column type" = Table.TransformColumnTypes(#"Promoted headers", {{"DateKey", type date}, {"OrganizationKey", Int64.Type}, {"DepartmentKey", Int64.Type}, {" AccountKey ", Int64.Type}, {"Amount_LocalCCY", type number}})
in
  #"Changed column type";
shared Source = let
  Source = SharePoint.Files("https://anzir.sharepoint.com/sites/SharePointData", [ApiVersion = 15])
in
  Source;
shared #"CSV Financial Data_DataDestination" = let
  Pattern = Lakehouse.Contents([HierarchicalNavigation = null, CreateNavigationProperties = false, EnableFolding = false]),
  Navigation_1 = Pattern{[workspaceId = "c38149e8-f089-4b53-8c51-b2db1915323e"]}[Data],
  Navigation_2 = Navigation_1{[lakehouseId = "49fb1398-9cf3-429b-8179-d491430024a1"]}[Data],
  TableNavigation = Navigation_2{[Id = "CSV Financial Data", ItemKind = "Table"]}?[Data]?
in
  TableNavigation;
