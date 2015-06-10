Operation =1
Option =2
Begin InputTables
    Name ="qryAnnualCompleteTgtSpeciesLists"
End
Begin OutputColumns
    Expression ="qryAnnualCompleteTgtSpeciesLists.Species_Name"
    Expression ="qryAnnualCompleteTgtSpeciesLists.Family"
    Expression ="qryAnnualCompleteTgtSpeciesLists.utah_species"
    Expression ="qryAnnualCompleteTgtSpeciesLists.Co_Species"
    Expression ="qryAnnualCompleteTgtSpeciesLists.Wy_Species"
    Expression ="qryAnnualCompleteTgtSpeciesLists.Master_Plant_Code_FK"
    Expression ="qryAnnualCompleteTgtSpeciesLists.Master_Common_Name"
    Alias ="ARCHPriority"
    Expression ="IIf([Park]=\"ARCH\",[PriorityTarget],\"X\")"
    Alias ="BLCAPriority"
    Expression ="IIf([Park]=\"BLCA\",[PriorityTarget],\"X\")"
    Alias ="BRCAPriority"
    Expression ="IIf([Park]=\"BRCA\",[PriorityTarget],\"X\")"
    Alias ="CANYPriority"
    Expression ="IIf([Park]=\"CANY\",[PriorityTarget],\"X\")"
    Alias ="CAREPriority"
    Expression ="IIf([Park]=\"CARE\",[PriorityTarget],\"X\")"
    Alias ="CEBRPriority"
    Expression ="IIf([Park]=\"CEBR\",[PriorityTarget],\"X\")"
    Alias ="COLMPriority"
    Expression ="IIf([Park]=\"COLM\",[PriorityTarget],\"X\")"
    Alias ="CUREPriority"
    Expression ="IIf([Park]=\"CURE\",[PriorityTarget],\"X\")"
    Alias ="DINOPriority"
    Expression ="IIf([Park]=\"DINO\",[PriorityTarget],\"X\")"
    Alias ="FOBUPriority"
    Expression ="IIf([Park]=\"FOBU\",[PriorityTarget],\"X\")"
    Alias ="GOSPPriority"
    Expression ="IIf([Park]=\"GOSP\",[PriorityTarget],\"X\")"
    Alias ="HOVEPriority"
    Expression ="IIf([Park]=\"HOVE\",[PriorityTarget],\"X\")"
    Alias ="NABRPriority"
    Expression ="IIf([Park]=\"NABR\",[PriorityTarget],\"X\")"
    Alias ="PISPPriority"
    Expression ="IIf([Park]=\"PISP\",[PriorityTarget],\"X\")"
    Alias ="TICAPriority"
    Expression ="IIf([Park]=\"TICA\",[PriorityTarget],\"X\")"
    Alias ="ZIONPriority"
    Expression ="IIf([Park]=\"ZION\",[PriorityTarget],\"X\")"
End
Begin Groups
    Expression ="qryAnnualCompleteTgtSpeciesLists.Species_Name"
    GroupLevel =0
    Expression ="qryAnnualCompleteTgtSpeciesLists.Family"
    GroupLevel =0
    Expression ="qryAnnualCompleteTgtSpeciesLists.utah_species"
    GroupLevel =0
    Expression ="qryAnnualCompleteTgtSpeciesLists.Co_Species"
    GroupLevel =0
    Expression ="qryAnnualCompleteTgtSpeciesLists.Wy_Species"
    GroupLevel =0
    Expression ="qryAnnualCompleteTgtSpeciesLists.Master_Plant_Code_FK"
    GroupLevel =0
    Expression ="qryAnnualCompleteTgtSpeciesLists.Master_Common_Name"
    GroupLevel =0
    Expression ="IIf([Park]=\"ARCH\",[PriorityTarget],\"X\")"
    GroupLevel =0
    Expression ="IIf([Park]=\"BLCA\",[PriorityTarget],\"X\")"
    GroupLevel =0
    Expression ="IIf([Park]=\"BRCA\",[PriorityTarget],\"X\")"
    GroupLevel =0
    Expression ="IIf([Park]=\"CANY\",[PriorityTarget],\"X\")"
    GroupLevel =0
    Expression ="IIf([Park]=\"CARE\",[PriorityTarget],\"X\")"
    GroupLevel =0
    Expression ="IIf([Park]=\"CEBR\",[PriorityTarget],\"X\")"
    GroupLevel =0
    Expression ="IIf([Park]=\"COLM\",[PriorityTarget],\"X\")"
    GroupLevel =0
    Expression ="IIf([Park]=\"CURE\",[PriorityTarget],\"X\")"
    GroupLevel =0
    Expression ="IIf([Park]=\"DINO\",[PriorityTarget],\"X\")"
    GroupLevel =0
    Expression ="IIf([Park]=\"FOBU\",[PriorityTarget],\"X\")"
    GroupLevel =0
    Expression ="IIf([Park]=\"GOSP\",[PriorityTarget],\"X\")"
    GroupLevel =0
    Expression ="IIf([Park]=\"HOVE\",[PriorityTarget],\"X\")"
    GroupLevel =0
    Expression ="IIf([Park]=\"NABR\",[PriorityTarget],\"X\")"
    GroupLevel =0
    Expression ="IIf([Park]=\"PISP\",[PriorityTarget],\"X\")"
    GroupLevel =0
    Expression ="IIf([Park]=\"TICA\",[PriorityTarget],\"X\")"
    GroupLevel =0
    Expression ="IIf([Park]=\"ZION\",[PriorityTarget],\"X\")"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbBinary "GUID" = Begin
    0x5188a5b072348a4faa3426c65339009b
End
dbText "Description" ="Complete species list for all parks for a given year\015\012 (Target List Tool u"
    "pdate)"
Begin
    Begin
        dbText "Name" ="BRCAPriority"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xfd90e35782ff044d8e08b3b61bc1249f
        End
    End
    Begin
        dbText "Name" ="ARCHPriority"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xcbda7f14e3649a47bb090ed7d03d040e
        End
    End
    Begin
        dbText "Name" ="BLCAPriority"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6bb7d012df55e449bf89864305897e82
        End
    End
    Begin
        dbText "Name" ="CANYPriority"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa85f539a57375441b9c097ec3e73d0d3
        End
    End
    Begin
        dbText "Name" ="CAREPriority"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x26486d3361215c478cfc02d959d0f4a9
        End
    End
    Begin
        dbText "Name" ="ZIONPriority"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xcfe856f9eee23b4cb3d930251c9d30da
        End
    End
    Begin
        dbText "Name" ="CEBRPriority"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2573fa24ca02ff4a9cf3dc826168d85f
        End
    End
    Begin
        dbText "Name" ="COLMPriority"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x427f1d0d2ff32b49879cc19ab8cf1f10
        End
    End
    Begin
        dbText "Name" ="CUREPriority"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9354e28229f2594f8a87398ff24f2972
        End
    End
    Begin
        dbText "Name" ="DINOPriority"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe2db09460d968541b57985c94b0cc553
        End
    End
    Begin
        dbText "Name" ="FOBUPriority"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x8935791877fc5046b233504a736c9cb2
        End
    End
    Begin
        dbText "Name" ="GOSPPriority"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xda5457d4b557904b95346a7ef8ab76e3
        End
    End
    Begin
        dbText "Name" ="HOVEPriority"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xacfe4315d151b14e890da7231d83ec8f
        End
    End
    Begin
        dbText "Name" ="NABRPriority"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9113a64ada808b45a5b10adb1fad018c
        End
    End
    Begin
        dbText "Name" ="PISPPriority"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x640fa4aefd39064c80d10e3aca3b1c68
        End
    End
    Begin
        dbText "Name" ="TICAPriority"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc42d0bd436cad54e8a6ccb3ab3be9115
        End
    End
    Begin
        dbText "Name" ="qryAnnualCompleteTgtSpeciesLists.Species_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryAnnualCompleteTgtSpeciesLists.Family"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryAnnualCompleteTgtSpeciesLists.Master_Plant_Code_FK"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryAnnualCompleteTgtSpeciesLists.Master_Common_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryAnnualCompleteTgtSpeciesLists.utah_species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryAnnualCompleteTgtSpeciesLists.Co_Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryAnnualCompleteTgtSpeciesLists.Wy_Species"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =1699
    Bottom =805
    Left =-1
    Top =-1
    Right =1661
    Bottom =304
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =60
        Top =17
        Right =316
        Bottom =287
        Top =0
        Name ="qryAnnualCompleteTgtSpeciesLists"
        Name =""
    End
End
