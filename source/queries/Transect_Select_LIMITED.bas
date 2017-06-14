dbMemo "SQL" ="SELECT (l.Plot_ID  & \"_\" & t.Transect) AS ID, l.Unit_Code, Year([Start_Date]) "
    "AS Visit_Year, l.Plot_ID AS Route, t.Transect_ID, t.Transect, l.Area, t.E_Coord,"
    " t.N_Coord, q.ID AS Quadrat_ID, q.Quadrat, esp.Position_m, esp.ColName, q.IsSamp"
    "led, q.NoExotics\015\012FROM ((tbl_Locations AS l LEFT JOIN EventSamplePosition "
    "AS esp ON esp.Location_ID = l.Location_ID) LEFT JOIN Transect AS t ON t.Event_ID"
    " = esp.Event_ID) LEFT JOIN Quadrat AS q ON q.Transect_ID = t.Transect_ID\015\012"
    "WHERE esp.Quadrat = q.Quadrat\015\012ORDER BY l.Unit_Code, Year(esp.Start_Date),"
    " l.Plot_ID, t.Transect, q.Quadrat;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x66ba1ac1c0cbbc4ca73e455c1b9742db
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xbc20df4c02531849b8bc4b4a1b70784a
        End
        dbInteger "ColumnWidth" ="2925"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="l.Unit_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x67163259f30dd94fb569cc9488ee89f9
        End
    End
    Begin
        dbText "Name" ="Visit_Year"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x050230f5f358d1408ed61ee2bab61c1e
        End
    End
    Begin
        dbText "Name" ="Route"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd95736d428ffa34a93907d2a605b30cb
        End
    End
    Begin
        dbText "Name" ="t.Transect_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xfc4678161f6c444c9bd7e0af927fc997
        End
    End
    Begin
        dbText "Name" ="t.Transect"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x8dfa5afbd1c57b4381a009ba5d25e899
        End
    End
    Begin
        dbText "Name" ="l.Area"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc0a57d1ff11c6e499348ea5e098e6217
        End
    End
    Begin
        dbText "Name" ="t.E_Coord"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xee7c2c7a64a65942bf639d65b4455ba1
        End
    End
    Begin
        dbText "Name" ="t.N_Coord"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf6f7e8d8ef63a84a95ac6d3970c52305
        End
    End
    Begin
        dbText "Name" ="esp.Position_m"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf6c238ddbfa72745b5f91d76adbbfaf5
        End
    End
    Begin
        dbText "Name" ="esp.ColName"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x79bcc44819ced54dbd3ca2a5e04707bf
        End
    End
    Begin
        dbText "Name" ="q.IsSampled"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x991d1aaa7393cc45b1b28f1cbed448d7
        End
    End
    Begin
        dbText "Name" ="q.NoExotics"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc05090cd5336474f9d67f62eb0fd5c4b
        End
    End
    Begin
        dbText "Name" ="q.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="q.Quadrat"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Quadrat_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4cc02076de1cb04488ec84b9c3428e84
        End
    End
End
