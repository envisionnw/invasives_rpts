Operation =1
Option =0
Where ="(((MSysObjects.Name) Not Like \"~*\") AND ((tsys_Link_Tables.Link_table) Is Null"
    ") AND ((MSysObjects.Type) In (4,6)))"
Begin InputTables
    Name ="MSysObjects"
    Name ="tsys_Link_Tables"
End
Begin OutputColumns
    Alias ="CurrTable"
    Expression ="MSysObjects.Name"
    Alias ="CurrDb"
    Expression ="IIf([Type]=4,fxnParseConnectionStr([Connect]),fxnParseFileName([Database]))"
    Alias ="CurrServer"
    Expression ="IIf([Type]=4,fxnParseConnectionStr([Connect],'SERVER='))"
    Alias ="CurrPath"
    Expression ="IIf([Type]=6,[Database])"
    Alias ="ODBC"
    Expression ="IIf([Type]=4,True,False)"
End
Begin Joins
    LeftTable ="MSysObjects"
    RightTable ="tsys_Link_Tables"
    Expression ="MSysObjects.Name = tsys_Link_Tables.Link_table"
    Flag =2
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbText "Description" ="Linked tables in MSysObjects that do not have records in tsys_Link_Tables (other"
    " than recently deleted objects that start with '~')"
dbBinary "GUID" = Begin
    0xdd820664635c7845b82eb1fc58c489f5
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="CurrDb"
        dbInteger "ColumnWidth" ="2430"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xeb851a7d67e9a941ab53874d3c698686
        End
    End
    Begin
        dbText "Name" ="CurrServer"
        dbInteger "ColumnWidth" ="1365"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x40ee735909cee44aa1fb9d3ef769d151
        End
    End
    Begin
        dbText "Name" ="CurrPath"
        dbInteger "ColumnWidth" ="9285"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x8c5db371a3123d47b34dc02779c203f7
        End
    End
    Begin
        dbText "Name" ="CurrTable"
        dbInteger "ColumnWidth" ="2505"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf1b3658e90131447b11e6450059163de
        End
    End
    Begin
        dbText "Name" ="ODBC"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2c9463880069f0408761f431a5181c72
        End
    End
End
Begin
    State =0
    Left =18
    Top =40
    Right =1130
    Bottom =352
    Left =-1
    Top =-1
    Right =1074
    Bottom =123
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =113
        Top =0
        Name ="MSysObjects"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =113
        Top =0
        Name ="tsys_Link_Tables"
        Name =""
    End
End
