dbMemo "SQL" ="SELECT tsca.Unit_Code, tsca.Visit_Year, tsca.Plot_ID AS Route, tsca.Transect, ts"
    "ca.Area, tsca.E_Coord, tsca.N_Coord, tsca.Species, tsca.Master_Common_Name, tsca"
    ".IsDead, tsca.Q1_0_5m, tsca.Q2_4_5m, tsca.Q3_9_5m, tsca.Q1_3m, tsca.Q2_8m, tsca."
    "Q3_13m, tsca.Q1, tsca.Q2, tsca.Q3, tsca.QuadratsSampled, tsca.TotalCover, tsca.A"
    "verageCover, ABS(tsca.Q1_0_5m - tsca.AverageCover) AS Dev_Q1_0_5m, ABS(tsca.Q2_4"
    "_5m - tsca.AverageCover) AS Dev_Q2_4_5m, ABS(tsca.Q3_9_5m - tsca.AverageCover) A"
    "S Dev_Q3_9_5m, ABS(tsca.Q1_3m - tsca.AverageCover) AS Dev_Q1_3m, ABS(tsca.Q2_8m "
    "- tsca.AverageCover) AS Dev_Q2_8m, ABS(tsca.Q3_13m - tsca.AverageCover) AS Dev_Q"
    "3_13m, ABS(tsca.Q1 - tsca.AverageCover) AS Dev_Q1, ABS(tsca.Q2 - tsca.AverageCov"
    "er) AS Dev_Q2, ABS(tsca.Q3 - tsca.AverageCover) AS Dev_Q3\015\012FROM Transect_S"
    "elect_Crosstab_with_AverageCover AS tsca\015\012ORDER BY tsca.Unit_Code, tsca.Vi"
    "sit_Year, tsca.Plot_ID, tsca.Transect, tsca.Species;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x83a9162518ee3d46b54b39a9f48484a4
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="tsca.Unit_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x91f8912ac957b442af0dd24e87f44c20
        End
    End
    Begin
        dbText "Name" ="tsca.Visit_Year"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x694aa81ae2b6b24d97b30749f754e12b
        End
    End
    Begin
        dbText "Name" ="Route"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x210b64207548524c98f4d632690ee320
        End
    End
    Begin
        dbText "Name" ="tsca.Transect"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe3e1595fae6b1b45b57ede650de56cad
        End
    End
    Begin
        dbText "Name" ="tsca.Area"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x23278cfb2646854c90d58db47a328614
        End
    End
    Begin
        dbText "Name" ="tsca.E_Coord"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xba58554909f7f946848374d4e8fd2a12
        End
    End
    Begin
        dbText "Name" ="tsca.N_Coord"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe6e09afeed2da0488f4bf4456f95dfb9
        End
    End
    Begin
        dbText "Name" ="tsca.Species"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6588c673fe760e488b6fe114b46bdd43
        End
    End
    Begin
        dbText "Name" ="tsca.Master_Common_Name"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x68e1687a3d60f44497031d5aa2d172bf
        End
    End
    Begin
        dbText "Name" ="tsca.IsDead"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x23fb276699db7f42ab1f8f3783359951
        End
    End
    Begin
        dbText "Name" ="tsca.Q1_0_5m"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf09599890c24bd438bd4b9f96faf65c0
        End
    End
    Begin
        dbText "Name" ="tsca.Q2_4_5m"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb1e9f238b10b924b80baff59586373af
        End
    End
    Begin
        dbText "Name" ="tsca.Q3_9_5m"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3362138cb4418c4da3d1bea5756c8a8d
        End
    End
    Begin
        dbText "Name" ="tsca.Q1_3m"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x46033ff7a1cc6b4a9684e5422f631ea9
        End
    End
    Begin
        dbText "Name" ="tsca.Q2_8m"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc508e72368b6e54da89e96d2c6bb4c27
        End
    End
    Begin
        dbText "Name" ="tsca.Q3_13m"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x85441c08b3f5864dadaaa26c00c9244d
        End
    End
    Begin
        dbText "Name" ="tsca.Q1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf2fdd82988e37344aeb52dfc58ca1169
        End
    End
    Begin
        dbText "Name" ="tsca.Q2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x0cf510c31aa320448a96fe7a1b211f58
        End
    End
    Begin
        dbText "Name" ="tsca.Q3"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9018bf01d9260a4c9d5767c96a38f866
        End
    End
    Begin
        dbText "Name" ="tsca.QuadratsSampled"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x53383ea32b68304084711906756732b6
        End
    End
    Begin
        dbText "Name" ="tsca.TotalCover"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x1bcd748019c8ff45afe14e4f72ba2d5d
        End
    End
    Begin
        dbText "Name" ="tsca.AverageCover"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x798ff4957c772d499df98cfd328e31f7
        End
    End
    Begin
        dbText "Name" ="Dev_Q1_0_5m"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd7262b9cc5432d47a147774c21059f0c
        End
    End
    Begin
        dbText "Name" ="Dev_Q2_4_5m"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xca1bab649045d54ab29fdc7475498254
        End
    End
    Begin
        dbText "Name" ="Dev_Q3_9_5m"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x8cec766be9a75a4e8ed8a0973f10fd5f
        End
    End
    Begin
        dbText "Name" ="Dev_Q1_3m"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3f8560066c65f240aa20fc55feae9cd4
        End
    End
    Begin
        dbText "Name" ="Dev_Q2_8m"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf9dab0cd9cdeb74998fdaa6cb6a59fbf
        End
    End
    Begin
        dbText "Name" ="Dev_Q3_13m"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x1703c21e5d6b624f9fba9746e77b905e
        End
    End
    Begin
        dbText "Name" ="Dev_Q1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x51d60d78db0ca04ba55a188b00fe477d
        End
    End
    Begin
        dbText "Name" ="Dev_Q2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x70fc80baf20ad449ac5698d73e6c5bcc
        End
    End
    Begin
        dbText "Name" ="Dev_Q3"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xad8133595b2d644aa81d01ac186b2d8c
        End
    End
End
