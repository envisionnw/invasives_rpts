dbMemo "SQL" ="TRANSFORM Min(sc.TransectsDetected) AS TransectsDetected\015\012SELECT sc.Unit_C"
    "ode, sc.Visit_Year, sc.Species, sc.Master_Common_Name, sc.[Alive?]\015\012FROM R"
    "oute_SpeciesCover AS sc\015\012WHERE sc.Unit_Code = 'ZION' AND sc.Visit_Year = 2"
    "015\015\012GROUP BY sc.Unit_Code, sc.Visit_Year, sc.Species, sc.Master_Common_Na"
    "me, sc.[Alive?]\015\012PIVOT sc.ColRouteTransects;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xbe0713d75d68ba45a31865fc473ebac8
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="Cundick Ridge Road (4) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xcc420982f202894ca9d7f91e484649e3
        End
    End
    Begin
        dbText "Name" ="Eagle Nest Point Road (2) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x31221555b6ca3243a259fad84f843037
        End
    End
    Begin
        dbText "Name" ="Fenceline Route (30) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x101b66ab48440f4c8ced9604dc03b5bb
        End
    End
    Begin
        dbText "Name" ="Fruita Canyon (4) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd5b75dad60699d4088cf34874698794b
        End
    End
    Begin
        dbText "Name" ="Fruita Canyon (5) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xdd396ba150b2ff44b8be7be3092b6526
        End
    End
    Begin
        dbText "Name" ="Gold Star Canyon (3) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x712f470fd72a644c9499e8c8ff247a10
        End
    End
    Begin
        dbText "Name" ="Green River, Left 1 (3) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x291ddc72aeb0d0458c30fb1c468fdb73
        End
    End
    Begin
        dbText "Name" ="Green River, Left 3 (57) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x771c96c12f5aa44eb69d3bb800c04e6b
        End
    End
    Begin
        dbText "Name" ="Green River, Right 2 (6) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa1a7911926b5a84db22e66e017cf246a
        End
    End
    Begin
        dbText "Name" ="Green River, Right 4 (36) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa226ce08365e9a4599c7c8b47e725e2b
        End
    End
    Begin
        dbText "Name" ="Archeology Trail (1) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x18a362d9aeb7184f900f6038fa05e02d
        End
        dbInteger "ColumnWidth" ="2640"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Blue Mesa Reservoir 1 (240) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7bf947c8f361e0408260123c03266b46
        End
    End
    Begin
        dbText "Name" ="Chicken Creek Nature Trail (5) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x617d1de8fd252541a5ab128482e36569
        End
    End
    Begin
        dbText "Name" ="Columbus Canyon (4) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xdd0caafce39b844b84c78d7bbfc2de62
        End
    End
    Begin
        dbText "Name" ="Corral Service Road (2) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x164ccec86da6814b900da788f28e5dc4
        End
    End
    Begin
        dbText "Name" ="East Glade Park Road (17) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x46dafa9280c1ca4391b929f074bb86d3
        End
    End
    Begin
        dbText "Name" ="East Mesa Trail (14) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe1f3c1d26fb4e143aa3f371256b18abc
        End
    End
    Begin
        dbText "Name" ="East Rim Trail (34) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x54c61312512b6d48b6ae4506f6b72faa
        End
    End
    Begin
        dbText "Name" ="Lower Kolob Terrace Road (26) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x48d1b63ac94318449944c42cfe890d28
        End
    End
    Begin
        dbText "Name" ="Main Park Road (23) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xba07b882fdf67f4fb94e32aed81f9419
        End
    End
    Begin
        dbText "Name" ="Millet Canyon (2) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc269c38564be4e44b881c9ea1d66dfdc
        End
    End
    Begin
        dbText "Name" ="No Thoroughfare Trail (2) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9ccbf7e7d8cadb4ba8a213afbd12de2e
        End
    End
    Begin
        dbText "Name" ="No Thoroughfare Trail (9) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd5ecd41899153c41b15666a40d47748a
        End
    End
    Begin
        dbText "Name" ="Northgate Peaks Trail (3) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb690868545bfdc4f8844328defb9ed02
        End
    End
    Begin
        dbText "Name" ="Oak Creek (15) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x632232fb7d04a3479ee73fd71dffd077
        End
        dbInteger "ColumnWidth" ="1785"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Poulsons Road (4) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd068c4e08d5d354d8b61ebbc5636e102
        End
    End
    Begin
        dbText "Name" ="Poulsons Road (5) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x1fd341b296ee5b4a995ad233b8063a0d
        End
    End
    Begin
        dbText "Name" ="Red Canyon (8) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x841546bf193ec54bb45fefbbb3e8a330
        End
    End
    Begin
        dbText "Name" ="Residence Service Road (1) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xdfe1cbe31f40014f9a3fdaf4225108b2
        End
    End
    Begin
        dbText "Name" ="Gunnison - Cooper Neversink, SN (7) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x19f47d6e01955549ab77ca7513d1a063
        End
    End
    Begin
        dbText "Name" ="Gunnison River 3 (2) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x0e2a89880e37fe40ab786762ac2af015
        End
    End
    Begin
        dbText "Name" ="Hop Valley Trail (4) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc36f1dca29170e47b33b3ebb6a14cb4f
        End
    End
    Begin
        dbText "Name" ="Hydro 1 (1) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x021afcaa7e35374481cd5b5d636a32d2
        End
    End
    Begin
        dbText "Name" ="Hydro 82 (3) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x0099fa579f21704f879f3799db1fa64d
        End
    End
    Begin
        dbText "Name" ="Hydro 95 (4) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe854eb5aaaa00047a614d326aafdbb60
        End
    End
    Begin
        dbText "Name" ="Kodels Canyon Route (2) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x8dcf9513d7aad741b8d2b0b838b56105
        End
    End
    Begin
        dbText "Name" ="Kolob Scenic Drive (36) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x09b83730318dc34e9e6fe477426a185e
        End
    End
    Begin
        dbText "Name" ="La Verkin Creek Trail (4) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf2455ab8ab6b3a4aa767ef971b89640e
        End
    End
    Begin
        dbText "Name" ="Liberty Cap Trail (14) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x248f8be481699e4d9108050b5702ec4c
        End
    End
    Begin
        dbText "Name" ="Millet Canyon (3) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4225dbb284c0be46b3eef31521afb809
        End
    End
    Begin
        dbText "Name" ="Monument Canyon (13) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x5b4f43043bff6e46afdb81b941187aec
        End
    End
    Begin
        dbText "Name" ="Witkers Service Road (1) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd296405ba8b56443a816e11d475b71e0
        End
    End
    Begin
        dbText "Name" ="Yampa River, Left 1 (40) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4cc09c29be702b4daecbdd68157b34ec
        End
    End
    Begin
        dbText "Name" ="Yampa River, Right 2 (6) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7366b5bff7ef954f86f2aed06831b6da
        End
    End
    Begin
        dbText "Name" ="Rim Rock Drive (24) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x321bda462d20e645a5a571714c8b8652
        End
    End
    Begin
        dbText "Name" ="Sand Bench Trail (14) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x230d66d2028f4040a7ee68006aef76c8
        End
    End
    Begin
        dbText "Name" ="South Rim Road (16) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x641e554b0b770640b0bb527497bd5307
        End
    End
    Begin
        dbText "Name" ="Upper Kolob Terrace Road (38) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2099323badf83a4d84b0e502c1130947
        End
    End
    Begin
        dbText "Name" ="Wildcat Canyon Trail (13) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x0c46f79e8eaba74bb732102962c43f9f
        End
    End
    Begin
        dbText "Name" ="Retired Service Road (1) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x8318b31829d16b43b3cb6e29382c596c
        End
    End
    Begin
        dbText "Name" ="Facility Road (1) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa2cbf8fccd09ec4a9fab87713990ce44
        End
    End
    Begin
        dbText "Name" ="Facility Road (2) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x36a0d4b6e5cb9346b1efd0798a6d653a
        End
    End
    Begin
        dbText "Name" ="Fenceline Route (28) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4686b2709b30bc4098c144a8ad8dd7f1
        End
    End
    Begin
        dbText "Name" ="Gunnison - Cooper Neversink, NS (7) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x8babbb1e7868f3409965e43a8edf2429
        End
    End
    Begin
        dbText "Name" ="Blue Mesa Reservoir 2 (1) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2635ca1bb0c40944934b344b2e588faa
        End
    End
    Begin
        dbText "Name" ="Blue Mesa Reservoir 5 (1) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xeb82d0c076733c4798fb7bd165103db9
        End
    End
    Begin
        dbText "Name" ="Cable Mountain Trail (6) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xafe358cd8560514ab0208c43844ec01d
        End
    End
    Begin
        dbText "Name" ="Cathedral Valley Road (26) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x22f1608b1e34b64c9ca889fd85dac99e
        End
    End
    Begin
        dbText "Name" ="Chinle Trail (11) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x09a305b2e0f8c64d90ddceb8d0960daa
        End
    End
    Begin
        dbText "Name" ="Chinle Trail (19) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6c92726d03ce8c409c11b3728abc2408
        End
    End
    Begin
        dbText "Name" ="Closed road (7) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x56b77fa724f4964ea0bfcce8c37722b7
        End
    End
    Begin
        dbText "Name" ="Columbus Canyon (3) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x33abf0cd9867f2418d05cbda73b70560
        End
    End
    Begin
        dbText "Name" ="Main Park Road (15) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7b7b0b979b170e4eab1b2988b2a35c5b
        End
    End
    Begin
        dbText "Name" ="Middle Emerald Pools Trail (1) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xef12401e0ee1984c843b79b615e2d4a7
        End
    End
    Begin
        dbText "Name" ="North Vista Trail (11) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x5c0d966eae2e284cad94a0df3d32c965
        End
    End
    Begin
        dbText "Name" ="Oak Creek and Administrative (5) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x0a6a43210e565641940c43436c48d11d
        End
    End
    Begin
        dbText "Name" ="Oak Creek and Administrative (8) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x62d69fc4d576a04380238516f426e681
        End
    End
    Begin
        dbText "Name" ="Pacific Track Grade (2) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb82c1a3c7926bb43af812e0c00c2ef8d
        End
    End
    Begin
        dbText "Name" ="Pleasant Creek (20) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x56aca3265d861740b326434866e252a5
        End
    End
    Begin
        dbText "Name" ="Red Rock Canyon (10) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x54a17ef6e2784741ba07269a11372d8b
        End
        dbInteger "ColumnWidth" ="4065"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Gunnison River 2 (4) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9076e621e0cee74f9be5b1717e6699eb
        End
    End
    Begin
        dbText "Name" ="Gunnison River 6 left (1) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x836d731e8a56f547bddcca626112ce32
        End
    End
    Begin
        dbText "Name" ="Gunnison River 6 right (3) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x8567dbd4caf44947938017a65c3bb700
        End
    End
    Begin
        dbText "Name" ="Hop Valley Trail (21) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x78c519c0531c8544985a7e1f944859a9
        End
    End
    Begin
        dbText "Name" ="Hydro 11 (2) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xbb094d1e9ee2c64eb58b105d4cc528b5
        End
    End
    Begin
        dbText "Name" ="Hydro 85 (1) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x842723f01c5d6d4f886d54a9bf017179
        End
    End
    Begin
        dbText "Name" ="Hydro 95 (5) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x91281462df852c418fa58eca00b6e7f7
        End
    End
    Begin
        dbText "Name" ="Last Cut Drainage (4) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x68fa456473c4dc4fa69a2e065fc55868
        End
    End
    Begin
        dbText "Name" ="Monument Canyon Trail (15) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xdaeeb997c834d3478a2ff25967b69cbd
        End
    End
    Begin
        dbText "Name" ="No Thoroughfare Canyon (13) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xdee20be1bc2bb94d9cddc19b43bf4b4c
        End
    End
    Begin
        dbText "Name" ="Warner Point Nature Trail (3) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe6b57b32347b774781467d75514e4a4f
        End
    End
    Begin
        dbText "Name" ="Zion Scenic Drive (42) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9703bae4df7df34f86697cb3bb511843
        End
    End
    Begin
        dbText "Name" ="Riverside Walk (3) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4237685d9341fe4891842a08428e7581
        End
    End
    Begin
        dbText "Name" ="Ruby Point Road (5) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf6f5d1a4741a094697f0245a3198fc0b
        End
    End
    Begin
        dbText "Name" ="Smallpox Creek (4) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc4da051bcb42fb4283ccb7f906ae11f5
        End
    End
    Begin
        dbText "Name" ="Smith Mesa (6) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9150fa6168b105468d370e920a411c8a
        End
    End
    Begin
        dbText "Name" ="Ute Canyon (14) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2caaaa4c55bce944b8b89acf199f9a90
        End
    End
    Begin
        dbText "Name" ="Ute Canyon (16) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9f8d394a20bfd0499cab9f4b83fb79cc
        End
    End
    Begin
        dbText "Name" ="West Rim Trail (32) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xef1fb25a9e861549ac2f349a299aa22c
        End
    End
    Begin
        dbText "Name" ="County Road (4) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x0b574b89c71b814089c025263348e955
        End
    End
    Begin
        dbText "Name" ="Cub Creek extra area (6) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x366fdb21d9cd4f4eb977182e3cd0ac59
        End
    End
    Begin
        dbText "Name" ="Deertrap Mountain Trail (13) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb38c952ece2ed644bd2869163ef0ce90
        End
    End
    Begin
        dbText "Name" ="Drainage 1 (4) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x31d428bfb42ab148a5d226e0f2cdad60
        End
    End
    Begin
        dbText "Name" ="East Tour Road (13) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd651a44f5845644784e2841c374f0c8e
        End
    End
    Begin
        dbText "Name" ="Green River, Left 1 (2) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x52a32bdeb872784ab453da17280ead80
        End
    End
    Begin
        dbText "Name" ="Green River, Left 1 (33) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x499258cafd300148b8a0159644810905
        End
    End
    Begin
        dbText "Name" ="Green River, Right 2 (28) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x1961342abde7584996df59a7f3f43a98
        End
    End
    Begin
        dbText "Name" ="Green River, Right 4 (35) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x092352026797f945afbf970d5e4cc348
        End
    End
    Begin
        dbText "Name" ="Grizzly Gulch (15) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x98e3d291362a6f48961a3660ff3add0b
        End
    End
    Begin
        dbText "Name" ="Grotto Trail (2) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2907d40a75e5f1459bdf085ebc2a1e21
        End
    End
    Begin
        dbText "Name" ="Gunnison - Cooper Neversink, NN (5) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf43bcc4e3e3f554a96062e278699f5af
        End
    End
    Begin
        dbText "Name" ="sc.Species"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x06a19686c037e84092f5a4a38dc4fa67
        End
    End
    Begin
        dbText "Name" ="sc.Master_Common_Name"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xcc41d4a13be69c4487c84388612fc26e
        End
    End
    Begin
        dbText "Name" ="Blue Mesa Reservoir 1 (121) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe0f92353a409f44c87b40d19f635563e
        End
    End
    Begin
        dbText "Name" ="Cooper Ranch River Trail (1) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9695bb76e333d9468a74f9e454ce44b2
        End
    End
    Begin
        dbText "Name" ="Corral Service Road (1) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x19fc69d2460b07419f6420fa419b3895
        End
    End
    Begin
        dbText "Name" ="County Road (2) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2e70612702d5fd44858d465a274fd2f6
        End
    End
    Begin
        dbText "Name" ="County Road (3) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xade76ef88b352e43af890450b033741e
        End
    End
    Begin
        dbText "Name" ="East Glade Park Road (16) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x53c339ce54bf904da27f2b15eeb25a9a
        End
    End
    Begin
        dbText "Name" ="East Glade Park Road (18) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xbf03f8646b918347bf7e4dae1468c379
        End
    End
    Begin
        dbText "Name" ="East Red Hill drainage (1) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf0ac3f90ef68be4f9cebe615816c59ea
        End
    End
    Begin
        dbText "Name" ="Main Park Road (45) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2b45e28f50adc54ca45dbe712ee469bb
        End
    End
    Begin
        dbText "Name" ="Middle Taylor Creek Trail (4) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4ce6171ac9fec54cb747dfc45fc021be
        End
    End
    Begin
        dbText "Name" ="No Thoroughfare Trail (8) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc38578f97270a942ba69ee9887e53f01
        End
    End
    Begin
        dbText "Name" ="North Dam Fork Chicken Creek (4) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x612e4124ed76b5428e4b19e68befa61d
        End
    End
    Begin
        dbText "Name" ="North Rim Road (15) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x289281317b40dc4cb11295184f8c49b7
        End
    End
    Begin
        dbText "Name" ="Northgate Peaks Trail (4) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x17e54930b36f654cb84e8aa7a3be08fb
        End
    End
    Begin
        dbText "Name" ="Pleasant Creek (19) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6dd3fe2c4ac0b3488c7ac67d23776462
        End
    End
    Begin
        dbText "Name" ="Gunnison River 1 (3) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd011c261bf0ce648b50e56bff49d71c2
        End
    End
    Begin
        dbText "Name" ="Gunnison River 5 (6) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x45c178db6e83774baca8e3deb7c5bd03
        End
    End
    Begin
        dbText "Name" ="Hydro 11 (3) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa5bb8bb507a9224e949375c3e10b1855
        End
    End
    Begin
        dbText "Name" ="Hydro 86 (2) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xdba5945cf3d4a743880cfd787357f4e5
        End
    End
    Begin
        dbText "Name" ="Kayenta Trail (2) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6557772fda1bda43b2e358f599a462a3
        End
    End
    Begin
        dbText "Name" ="Kodels Canyon (1) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe152f931e2650f4c929065f88047d1dd
        End
    End
    Begin
        dbText "Name" ="Kodels Canyon (2) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x619ccdeff8233343a212e3129ff0ef5e
        End
    End
    Begin
        dbText "Name" ="Kodels Canyon Route (1) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7176c2434dea5a4f8231d8fbf52e8697
        End
    End
    Begin
        dbText "Name" ="Kolob Scenic Drive (34) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xec26fa78c3ec814cae5ea0e641efcf7f
        End
    End
    Begin
        dbText "Name" ="La Verkin Creek Trail (27) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x757251ccc4bb7546a8636574416ea106
        End
    End
    Begin
        dbText "Name" ="Last Cut Drainage (3) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x81274a8152f0d642973aff8a364cf8f3
        End
    End
    Begin
        dbText "Name" ="Monument Canyon (19) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x5412aaea0cbbbc48911196a91063ce86
        End
    End
    Begin
        dbText "Name" ="Museum Trail (1) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x11abcd627b2e2b4c93934e147cc6d157
        End
    End
    Begin
        dbText "Name" ="Neversink Trail (2) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb3ab3503c37f2b49a5712255d3e68036
        End
    End
    Begin
        dbText "Name" ="No Thoroughfare Canyon (18) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd301a158548dc549bcc015a97848cde9
        End
    End
    Begin
        dbText "Name" ="Watchman Housing and PTI (2) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x69e95e6894c2054f861b4740b9116d1a
        End
    End
    Begin
        dbText "Name" ="Wrangler Trail (3) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa5783bf9f242724f82f5fa620e6842dc
        End
    End
    Begin
        dbText "Name" ="Wrangler Trail (4) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xbbd994c89a2d174bb3e35fa566289228
        End
    End
    Begin
        dbText "Name" ="Yampa River, Left 1 (4) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x71d1b06104408b49979f41a111696388
        End
    End
    Begin
        dbText "Name" ="Yampa River, Right 2 (40) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd7e42ebc7a8a4646b3f966496f5985c2
        End
    End
    Begin
        dbText "Name" ="Zion Scenic Drive (32) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x36631dd6c7e79b4b8a4191affcf7d91d
        End
    End
    Begin
        dbText "Name" ="Rim Rock Drive, E Glade Park Rd to Ute Canyon Vi (18) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd668d6099a1a214ca15d4cda3516ee34
        End
    End
    Begin
        dbText "Name" ="Scenic Drive (38) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x85578752f600914eb5fd773a71094cea
        End
    End
    Begin
        dbText "Name" ="Scenic Drive (41) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf4e9064fa6523641ae71078f3c83450e
        End
    End
    Begin
        dbText "Name" ="Service Road 1 (9) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x28d77bce9ae16e439c081fe0555062e8
        End
    End
    Begin
        dbText "Name" ="Service Road 2 (3) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x997a105658ce0c45a816b7b77b4b8c6a
        End
    End
    Begin
        dbText "Name" ="Slide Draw Route (1) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xbdec7837c7edcc47adb41e627383dac7
        End
    End
    Begin
        dbText "Name" ="South Millet Canyon (2) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x06f984482ccd9f4e9ec01f92ead7a65e
        End
    End
    Begin
        dbText "Name" ="State Route 9 (39) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3707a8dc900ffe48a30dc1c6708eb1cf
        End
    End
    Begin
        dbText "Name" ="State Route 9 (44) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x06bf45a6252c1345b5a4a8a6da4696a8
        End
    End
    Begin
        dbText "Name" ="Stave Spring (2) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4d4b1fd0ec4a9d4aae3062ed8fc61f31
        End
    End
    Begin
        dbText "Name" ="Timber Creek Overlook Trail (2) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x881ed290d3ada74f97c47712591fb64a
        End
    End
    Begin
        dbText "Name" ="Tomichi Route (1) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2cf84a1f86348e4fbe50b206ec3997aa
        End
    End
    Begin
        dbText "Name" ="Upper Emerald Pools Trail (1) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3e9aa6234fe5f24c8387417a216e337b
        End
    End
    Begin
        dbText "Name" ="Wedding Canyon (6) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x0bbeae63995c1c499f156484c7b3d035
        End
    End
    Begin
        dbText "Name" ="Weeping Rock (1) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x764327c1f173d64b9ea8f8a9c69b22c8
        End
    End
    Begin
        dbText "Name" ="West Fork Chicken Creek (2) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x5884b93fb7613b4d9d8d80a48670aee1
        End
    End
    Begin
        dbText "Name" ="West Rim Trail (16) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7805cc25e207514aa5c2be117f514cad
        End
    End
    Begin
        dbText "Name" ="Wildcat Canyon Trail (19) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9608e712770e0f48b691168f6f8f53c4
        End
    End
    Begin
        dbText "Name" ="Rim Rock Drive (23) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x59c139576492be4fb991cb0f97654dfb
        End
    End
    Begin
        dbText "Name" ="Deadhorse Trail (9) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x54dafa5240c30a4ca1bf5f0cd2dbff62
        End
    End
    Begin
        dbText "Name" ="Dillon Pinnacles (6) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf6d7adf0d2945c4ab4a6d97b585e0555
        End
    End
    Begin
        dbText "Name" ="Drainage 12 (5) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc17129b9dee62f47836d2f968b5c0cdf
        End
    End
    Begin
        dbText "Name" ="Drainage 15 (4) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x5518083ba326154a958f7188f8cad9eb
        End
    End
    Begin
        dbText "Name" ="Drainage 9 (4) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2811a80878005f46afbaac245f845356
        End
    End
    Begin
        dbText "Name" ="East Tour Road (10) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7bea071eb6c84d44964264b572f292f7
        End
    End
    Begin
        dbText "Name" ="Fossil Butte east drainage (3) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x325b3ac75ea98045bdfe88a05ba1d8a6
        End
    End
    Begin
        dbText "Name" ="Gold Star Canyon (4) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb462e6c22e384e499ff013fcfce2a847
        End
    End
    Begin
        dbText "Name" ="sc.Unit_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2423c39c58663d41b4e037ceb49c79b3
        End
    End
    Begin
        dbText "Name" ="sc.Visit_Year"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9b2e458b58a58349bf909a80c0e39cfa
        End
    End
    Begin
        dbText "Name" ="sc.IsDead"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x0d676ca31185d847b0eff7522842420d
        End
    End
    Begin
        dbText "Name" ="4WD Road 1 (3) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb36647f00b041f448d1422cc4391ab21
        End
        dbInteger "ColumnWidth" ="2835"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Big Fill Section County Road (10) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd82b3aa2a0a5f94fa14c0e728c627e73
        End
    End
    Begin
        dbText "Name" ="Big Fill Section County Road (5) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x0a22a498f736634c8ae06d59497d560f
        End
    End
    Begin
        dbText "Name" ="Big Fill Section County Road (6) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4ee4e99ab4ce9244a31c94d12371d8e8
        End
    End
    Begin
        dbText "Name" ="Black Canyon Road (11) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb05aa4503365a94d8af52b287381e2a8
        End
        dbInteger "ColumnWidth" ="2595"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Blue Mesa Reservoir 1 (142) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x8ae9063037d3ae44ba94047ec8035f6d
        End
    End
    Begin
        dbText "Name" ="Blue Mesa Reservoir 3 (2) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x09b9d0bbfec6e741a7265800d7065565
        End
    End
    Begin
        dbText "Name" ="Blue Mesa Reservoir 7 (3) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe62f5e6c74144b458758b3c5632ece7d
        End
    End
    Begin
        dbText "Name" ="Cathedral Valley Road (13) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x8792125f471d214092229cf7d0691008
        End
    End
    Begin
        dbText "Name" ="Chicken Creek (11) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x93c70157880da340aca21546a87ac963
        End
    End
    Begin
        dbText "Name" ="Chicken Creek (14) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x068cf41670a5e448b1e509d3f5527c6a
        End
    End
    Begin
        dbText "Name" ="Chicken Creek (7) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7ef796b67e383841a7495a818429a307
        End
    End
    Begin
        dbText "Name" ="Chicken Creek Nature Trail (4) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9a00458847e4194997fad2de547ccd25
        End
    End
    Begin
        dbText "Name" ="Connector Trail (13) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2ce6007ef2d42344981c3d7bb9b32564
        End
    End
    Begin
        dbText "Name" ="Cooper Ranch Trail (1) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x0ba4291003ee25438a2a20005f0bb5ff
        End
    End
    Begin
        dbText "Name" ="East Portal Rd to Campground (1) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x253031ac08507a4f8e5b97b8fe097ee2
        End
    End
    Begin
        dbText "Name" ="East Portal Road (42) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x24a3c8e04f05a647b19c5e075e28d41f
        End
    End
    Begin
        dbText "Name" ="Lower Emerald Pools Trail (2) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb7abf607389d5448bd38f852fc5a79ea
        End
    End
    Begin
        dbText "Name" ="Main Park Road (30) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x46a3f0c872da8b49bfc0e828fbce670d
        End
    End
    Begin
        dbText "Name" ="Middle Taylor Creek Trail (6) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xecb186f4652a49409a2412a8b013fba3
        End
    End
    Begin
        dbText "Name" ="Parus Trail (5) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7b0620ae87670948a47acef3a3bb7749
        End
    End
    Begin
        dbText "Name" ="Pine Creek Trail (3) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe6e815be40b00f46a4f464c5d4be06b2
        End
    End
    Begin
        dbText "Name" ="Gunnison - Cooper Neversink, SS (6) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4d6f508d8436154796a3668302220ee5
        End
    End
    Begin
        dbText "Name" ="Gunnison River 5 (2) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x662991c5ae3baa4eb0aee68593d3041e
        End
    End
    Begin
        dbText "Name" ="Gunnison River 6 right (2) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xfde7ea6daa472c4c9c1facfd5ce8ab27
        End
    End
    Begin
        dbText "Name" ="Highway 24 (100) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4e537f4f316e1d4d9bc892efb57ce59d
        End
    End
    Begin
        dbText "Name" ="Highway 24 (102) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe11b31a7b3c5974e82a95b1808796395
        End
    End
    Begin
        dbText "Name" ="Highway 24 (11) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf92681e8a4cdb14d89b2c4d5d1047b79
        End
    End
    Begin
        dbText "Name" ="Highway 24 (39) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf758f400e712fa48bd54591cf308206c
        End
    End
    Begin
        dbText "Name" ="Historic Quarry Trail (7) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4b97de1f14defb4a931a4911a3824dbc
        End
    End
    Begin
        dbText "Name" ="Hydro 03 (1) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc6bfeb9b69f2564b9ecaf3b1d935fdec
        End
    End
    Begin
        dbText "Name" ="Hydro 82 (1) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd62e25ea33184c479ca38190780b709b
        End
    End
    Begin
        dbText "Name" ="Hydro 88 (1) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf82d7bb6f472e14fb2f0e6e4efc3b1c6
        End
    End
    Begin
        dbText "Name" ="Limekin Gulch (1) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x041c1dd006a54c45b89d366686c58c03
        End
    End
    Begin
        dbText "Name" ="Lizard Canyon (2) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xbe2d67426c98764389f771a276e330b7
        End
    End
    Begin
        dbText "Name" ="Lizard Canyon (3) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4226e911374bb544a25d9afa5f72bedc
        End
    End
    Begin
        dbText "Name" ="Monument Canyon (18) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc99370d69da8ad44bf88381cfc7a3983
        End
    End
    Begin
        dbText "Name" ="Moose Bones Canyon (3) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x292fd7320046f74c8891a08105be503a
        End
    End
    Begin
        dbText "Name" ="No Thoroughfare Canyon (19) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc77803594685904b9ec48e4973b5ded0
        End
    End
    Begin
        dbText "Name" ="Rim Rock Drive, E Glade Park Rd to Ute Canyon Vi (17) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x46db72a67d05e1478e3298ba01c553b4
        End
    End
    Begin
        dbText "Name" ="Sand Bench Trail (15) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb8527a130d84384e948952f47ffaf838
        End
    End
    Begin
        dbText "Name" ="South Rim Road (55) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd586d9849919c247844d71c73753345b
        End
    End
    Begin
        dbText "Name" ="Upper Kolob Terrace Road (26) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xab11a78b213d55439d769f1896383e5a
        End
    End
    Begin
        dbText "Name" ="Watchman Trail (5) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x39da0481c4e699448bf465f610066da4
        End
    End
    Begin
        dbText "Name" ="Watchman Trail (7) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x21c81d9eeaef0f4b8e6bfb98b8ed54c7
        End
    End
    Begin
        dbText "Name" ="Water Tank Road (1) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xca798630f541744ba22181c938a32821
        End
    End
    Begin
        dbText "Name" ="West Rim Trail (15) TCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x74a09af27ca32f4283e5b372997e9aba
        End
    End
    Begin
        dbText "Name" ="Upper Kolob Terrace Road (38) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Warner Point Nature Trail (3) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="West Fork Chicken Creek (2) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Green River, Right 2 (28) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="No Thoroughfare Trail (8) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Oak Creek (15) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Red Rock Canyon (10) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Residence Service Road (1) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Hydro 03 (1) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Gunnison River 1 (3) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Gunnison River 2 (4) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Gunnison River 3 (2) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Gunnison River 5 (2) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Gunnison River 5 (6) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Gunnison River 6 left (1) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Highway 24 (39) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Hop Valley Trail (21) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Hop Valley Trail (4) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Lizard Canyon (3) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="La Verkin Creek Trail (27) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="South Rim Road (16) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="South Rim Road (55) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="State Route 9 (39) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tomichi Route (1) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="East Portal Road (42) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="East Red Hill drainage (1) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Facility Road (2) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Fruita Canyon (5) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Green River, Left 1 (3) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Monument Canyon (19) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Monument Canyon Trail (15) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Neversink Trail (2) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="No Thoroughfare Canyon (13) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Riverside Walk (3) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Hydro 86 (2) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Middle Taylor Creek Trail (6) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Monument Canyon (13) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Monument Canyon (18) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Witkers Service Road (1) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Wrangler Trail (4) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Yampa River, Left 1 (40) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Archeology Trail (1) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Big Fill Section County Road (10) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Watchman Trail (5) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Grizzly Gulch (15) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Blue Mesa Reservoir 1 (142) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Pacific Track Grade (2) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Pine Creek Trail (3) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Pleasant Creek (19) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Pleasant Creek (20) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Poulsons Road (4) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Blue Mesa Reservoir 7 (3) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Eagle Nest Point Road (2) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Service Road 1 (9) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Ute Canyon (16) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Weeping Rock (1) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Green River, Right 2 (6) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="No Thoroughfare Canyon (19) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Hydro 82 (3) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Oak Creek and Administrative (8) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Retired Service Road (1) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Gunnison River 6 right (2) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Highway 24 (11) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Blue Mesa Reservoir 2 (1) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Kodels Canyon (2) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Kodels Canyon Route (2) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Timber Creek Overlook Trail (2) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="County Road (4) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cundick Ridge Road (4) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deadhorse Trail (9) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deertrap Mountain Trail (13) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Chicken Creek (11) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Chinle Trail (19) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cooper Ranch Trail (1) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Corral Service Road (1) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Drainage 12 (5) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="East Mesa Trail (14) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="East Portal Rd to Campground (1) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Moose Bones Canyon (3) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Museum Trail (1) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Millet Canyon (3) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Wrangler Trail (3) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Yampa River, Right 2 (40) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="No Thoroughfare Trail (9) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Hydro 11 (3) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Parus Trail (5) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cable Mountain Trail (6) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Chicken Creek Nature Trail (4) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="East Glade Park Road (16) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Lizard Canyon (2) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Lower Kolob Terrace Road (26) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Smith Mesa (6) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Hydro 95 (5) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="La Verkin Creek Trail (4) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Upper Kolob Terrace Road (26) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="East Rim Trail (34) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="East Tour Road (10) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="East Tour Road (13) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Facility Road (1) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Fruita Canyon (4) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Green River, Left 1 (33) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="North Rim Road (15) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="North Vista Trail (11) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Northgate Peaks Trail (3) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Oak Creek and Administrative (5) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Hydro 1 (1) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Kodels Canyon (1) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Kolob Scenic Drive (34) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Stave Spring (2) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Corral Service Road (2) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="County Road (3) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Chinle Trail (11) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Scenic Drive (41) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Hydro 88 (1) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Middle Taylor Creek Trail (4) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="West Rim Trail (16) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="West Rim Trail (32) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Wildcat Canyon Trail (13) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Ute Canyon (14) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Watchman Housing and PTI (2) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Water Tank Road (1) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Wedding Canyon (6) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="West Rim Trail (15) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Green River, Right 4 (35) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Gunnison - Cooper Neversink, NN (5) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="North Dam Fork Chicken Creek (4) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Black Canyon Road (11) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Blue Mesa Reservoir 1 (121) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Red Canyon (8) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rim Rock Drive (23) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rim Rock Drive (24) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rim Rock Drive, E Glade Park Rd to Ute Canyon Vi (17) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Gunnison River 6 right (3) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Highway 24 (102) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Blue Mesa Reservoir 3 (2) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Last Cut Drainage (4) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Liberty Cap Trail (14) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Limekin Gulch (1) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Slide Draw Route (1) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Smallpox Creek (4) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Last Cut Drainage (3) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Green River, Left 3 (57) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="No Thoroughfare Canyon (18) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Hydro 85 (1) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Millet Canyon (2) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Yampa River, Right 2 (6) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="4WD Road 1 (3) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Gunnison - Cooper Neversink, NS (7) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Northgate Peaks Trail (4) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Historic Quarry Trail (7) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Blue Mesa Reservoir 1 (240) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cathedral Valley Road (13) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Chicken Creek (7) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Columbus Canyon (3) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Columbus Canyon (4) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Connector Trail (13) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cooper Ranch River Trail (1) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Drainage 9 (4) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="East Glade Park Road (17) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Ruby Point Road (5) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sand Bench Trail (14) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sand Bench Trail (15) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Scenic Drive (38) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Wildcat Canyon Trail (19) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Yampa River, Left 1 (4) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Watchman Trail (7) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Green River, Right 4 (36) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Grotto Trail (2) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="No Thoroughfare Trail (2) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Hydro 82 (1) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.[Alive?]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Big Fill Section County Road (6) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Gunnison - Cooper Neversink, SN (7) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Poulsons Road (5) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Highway 24 (100) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Blue Mesa Reservoir 5 (1) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Lower Emerald Pools Trail (2) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rim Rock Drive, E Glade Park Rd to Ute Canyon Vi (18) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="South Millet Canyon (2) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Kayenta Trail (2) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Kolob Scenic Drive (36) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="State Route 9 (44) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Upper Emerald Pools Trail (1) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="County Road (2) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Dillon Pinnacles (6) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Fenceline Route (28) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Fenceline Route (30) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Fossil Butte east drainage (3) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Gold Star Canyon (3) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Gold Star Canyon (4) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Green River, Left 1 (2) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Drainage 1 (4) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Big Fill Section County Road (5) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Zion Scenic Drive (32) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Zion Scenic Drive (42) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Hydro 11 (2) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Gunnison - Cooper Neversink, SS (6) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Kodels Canyon Route (1) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cub Creek extra area (6) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cathedral Valley Road (26) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Chicken Creek (14) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Chicken Creek Nature Trail (5) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Closed road (7) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Drainage 15 (4) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="East Glade Park Road (18) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Main Park Road (15) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Main Park Road (23) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Main Park Road (30) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Main Park Road (45) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Service Road 2 (3) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Hydro 95 (4) SE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Middle Emerald Pools Trail (1) SE"
        dbLong "AggregateType" ="-1"
    End
End
