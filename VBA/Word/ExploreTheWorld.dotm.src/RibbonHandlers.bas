Attribute VB_Name = "RibbonHandlers"
Option Explicit

Public Sub SaveAsJson_OnClick(control As IRibbonControl)
    Call MSO_MsWord.WriteActiveDocument
End Sub
