Attribute VB_Name = "Samples"
Option Compare Database
Option Explicit
 
Private Sub HowtoUse()

    Dim rpt As Report
    Dim rname As String
    
    Dim vDevMode As New CPrtDevMode
    Dim vDevName As New CPrtDevNames
    Dim vMip As New CPrtMip
    
    rname = "R_Table3"
    
    DoCmd.OpenReport rname, acViewDesign

    Set rpt = Reports(rname)
    
    Call vDevName.LoadData(rpt.PrtDevNames)
    Call vMip.LoadData(rpt.PrtMip)
    Call vDevMode.LoadData(rpt.PrtDevMode)
    
    vDevMode.PaperWidth_mm = 2970
    vDevMode.PaperLength_mm = 2100
    vMip.DataOnly = 0
    vMip.BottomMargin_mm = 15
    vDevName.IsDefault = 0
    rpt.PrtDevMode = vDevMode.ToString()
    rpt.PrtMip = vMip.ToString()
    rpt.PrtDevNames = vDevName.ToString()
    
    DoCmd.Close acReport, rname, acSaveYes
    
End Sub
