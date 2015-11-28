dbMemo "SQL" ="SELECT Min(TgtYear), \"'\" & [TempVars]![Park] & \"'\"\015\012FROM qry_Annual_Co"
    "mplete_Tgt_Species_Lists\015\012WHERE Park = \"'\" & [TempVars]![Park] & \"'\";\015"
    "\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x415883e31135324a91d05d71ec3f1322
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="Expr1000"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1001"
        dbLong "AggregateType" ="-1"
    End
End
