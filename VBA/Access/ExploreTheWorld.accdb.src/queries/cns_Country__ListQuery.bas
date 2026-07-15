Operation =1
Option =0
Begin InputTables
    Name ="cns_Country"
End
Begin OutputColumns
    Expression ="cns_Country.GUID"
    Expression ="cns_Country.Country"
    Expression ="cns_Country.Iso2"
End
Begin OrderBy
    Expression ="cns_Country.Country"
    Flag =0
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
Begin
    Begin
        dbText "Name" ="cns_Country.Country"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="cns_Country.GUID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="cns_Country.Iso2"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =1296
    Bottom =1092
    Left =-1
    Top =-1
    Right =1280
    Bottom =847
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =131
        Top =119
        Right =275
        Bottom =263
        Top =0
        Name ="cns_Country"
        Name =""
    End
End
