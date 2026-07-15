Operation =1
Option =0
Where ="(((MsOfficeEvents.Word)=True))"
Begin InputTables
    Name ="MsOfficeEvents"
End
Begin OutputColumns
    Expression ="MsOfficeEvents.Log"
    Expression ="MsOfficeEvents.Category"
    Expression ="MsOfficeEvents.Name"
End
Begin OrderBy
    Expression ="MsOfficeEvents.Category"
    Flag =0
    Expression ="MsOfficeEvents.Name"
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
dbBoolean "NoFormat" ="0"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="MsOfficeEvents.Category"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1770"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="MsOfficeEvents.Name"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="4200"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="MsOfficeEvents.Log"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="675"
        dbBoolean "ColumnHidden" ="0"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =1833
    Bottom =1238
    Left =-1
    Top =-1
    Right =1089
    Bottom =993
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =78
        Top =99
        Right =281
        Bottom =350
        Top =0
        Name ="MsOfficeEvents"
        Name =""
    End
End
