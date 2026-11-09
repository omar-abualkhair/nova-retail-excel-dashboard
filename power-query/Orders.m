shared Orders = let
    Source = Excel.Workbook(File.Contents("C:\Data\1- Raw Data - Business Data Set.xlsx"), null, true),
    Orders_Sheet = Source{[Item="Orders",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(Orders_Sheet, [PromoteAllScalars=true]),
    // Raw "Ship Date" is a number such as 1101102016 = month 11, "0", day 11, "0", year 2016.
    // Rebuild it as a real date; when a split is ambiguous, keep the date 0-10 days after the order.
    #"Fixed Ship Date" = Table.AddColumn(#"Promoted Headers", "ship Date", each
        let
            t     = Text.From([Ship Date]),
            y     = Number.From(Text.End(t, 4)),
            core  = Text.Start(t, Text.Length(t) - 5),
            od    = Date.From([Order Date]),
            opts  = List.RemoveNulls(List.Transform({1 .. Text.Length(core) - 2}, (i) =>
                        if Text.At(core, i) = "0" and Text.At(core, i + 1) <> "0"
                        then try #date(y, Number.From(Text.Start(core, i)), Number.From(Text.Range(core, i + 1))) otherwise null
                        else null)),
            valid = List.Select(opts, (d) => Duration.Days(d - od) >= 0 and Duration.Days(d - od) <= 10)
        in
            List.First(valid), type date),
    #"Merged Columns2" = Table.RemoveColumns(#"Fixed Ship Date", {"Ship Date"}),
    #"Split Column by Delimiter" = Table.SplitColumn(#"Merged Columns2", "Location", Splitter.SplitTextByDelimiter(",", QuoteStyle.Csv), {"Location.1", "Location.2"}),
    #"Changed Type5" = Table.TransformColumnTypes(#"Split Column by Delimiter",{{"Location.1", type text}, {"Location.2", type text}}),
    #"Renamed Columns3" = Table.RenameColumns(#"Changed Type5",{{"Location.1", "Country"}, {"Location.2", "City"}}),
    #"Changed Type6" = Table.TransformColumnTypes(#"Renamed Columns3",{{"ship Date", type date}, {"Order Date", type date}, {"Discount", Percentage.Type}}),
    #"Added Custom" = Table.AddColumn(#"Changed Type6", "DiscountValue", each [Sales]*-[Discount]),
    #"Reordered Columns" = Table.ReorderColumns(#"Added Custom",{"Row ID", "Order ID", "Order Date", "ship Date", "Ship Mode", "Customer ID", "Customer Name", "Segment", "Country", "City", "State", "Postal Code", "Region", "Product ID", "Category", "Sub-Category", "Product Name", "Sales", "Quantity", "Discount", "DiscountValue", "Profit"}),
    #"Added Custom1" = Table.AddColumn(#"Reordered Columns", "COGS", each [Sales]-[Profit]),
    #"Added Custom2" = Table.AddColumn(#"Added Custom1", "COGS2", each -([Sales]+[DiscountValue]-[Profit])),
    #"Removed Columns1" = Table.RemoveColumns(#"Added Custom2",{"COGS"}),
    #"Reordered Columns1" = Table.ReorderColumns(#"Removed Columns1",{"Row ID", "Order ID", "Order Date", "ship Date", "Ship Mode", "Customer ID", "Customer Name", "Segment", "Country", "City", "State", "Postal Code", "Region", "Product ID", "Category", "Sub-Category", "Product Name", "Sales", "Quantity", "Discount", "DiscountValue", "COGS2", "Profit"}),
    #"Added Index" = Table.AddIndexColumn(#"Reordered Columns1", "Index", 1002, 4, Int64.Type),
    #"Renamed Columns4" = Table.RenameColumns(#"Added Index",{{"Index", "Serial Number"}}),
    #"Changed Type7" = Table.TransformColumnTypes(#"Renamed Columns4",{{"Serial Number", Int64.Type}}),
    #"Added Custom3" = Table.AddColumn(#"Changed Type7", "Custom", each [ship Date]-[Order Date]),
    #"Changed Type8" = Table.TransformColumnTypes(#"Added Custom3",{{"Custom", Int64.Type}}),
    #"Renamed Columns5" = Table.RenameColumns(#"Changed Type8",{{"Custom", "Delivery Date"}})
in
    #"Renamed Columns5";

