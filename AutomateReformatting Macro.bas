Attribute VB_Name = "AutomateReformatting"
Sub CopyingandDeleting()
Attribute CopyingandDeleting.VB_ProcData.VB_Invoke_Func = " \n14"

Application.DisplayAlerts = False

If Application.Sheets.Count > 1 Then
  
    Sheets("Copy").Delete

End If

If Application.Sheets.Count > 1 Then
    
    Sheets("HealthScore").Delete
    Sheets("TotalCholesterol").Delete
    Sheets("HDL").Delete
    Sheets("LDL").Delete
    Sheets("TRIG").Delete
    Sheets("SBP").Delete
    Sheets("DBP").Delete
    Sheets("BP").Delete
    Sheets("Glucose").Delete
    Sheets("A1C").Delete
    Sheets("BMI").Delete
    Sheets("Waist").Delete
    Sheets("WHR").Delete
    Sheets("Nicotine").Delete
    Sheets("GGT").Delete
    Sheets("PSA").Delete
    Sheets("Age").Delete
    Sheets("Averages").Delete
    Sheets("WaistbyGender").Delete
    Sheets("MetabolicSyndromes").Delete

End If

Application.DisplayAlerts = True

    Cells.Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Select
    ActiveSheet.Name = "Copy"
    ActiveSheet.Paste

Call AutomateDeletion

End Sub

Sub AutomateDeletion()

    Sheets("Copy").Range("V:V,AA:AA,AC:AC,AN:AN").Select
    Sheets("Copy").Range("AN1").Activate
    Sheets("Copy").Range("V:V,AA:AA,AC:AC,AN:AN,BE:BE,BG:BG,BQ:BQ,BR:BR,BT:BT").Select
    Sheets("Copy").Range("BT1").Activate
    Sheets("Copy").Range("V:V,AA:AA,AC:AC,AN:AN,BE:BE,BG:BG,BQ:BQ,BR:BR,BT:CD,CF:CF").Select
    Sheets("Copy").Range("CF1").Activate
    Sheets("Copy").Range("V:V,AA:AA,AC:AC,AN:AN,BE:BE,BG:BG,BQ:BQ,BR:BR,BT:CD,CF:EC,EG:EG").Select
    Sheets("Copy").Range("EG1").Activate
    Sheets("Copy").Range("V:V,AA:AA,AC:AC,AN:AN,BE:BE,BG:BG,BQ:BQ,BR:BR,BT:CD,CF:EC,EG:EL,EN:ET").Select
    Sheets("Copy").Range("EN1").Activate
    Selection.Delete Shift:=xlToLeft
    
Call AutomateCreation
    
End Sub

Sub AutomateCreation()

Dim NCells As Integer

'Insert columns for conversions

    Sheets("Copy").Columns("AF:AF").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Sheets("Copy").Columns("AH:AH").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Sheets("Copy").Columns("AJ:AJ").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Sheets("Copy").Columns("AN:AN").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Sheets("Copy").Columns("AR:AR").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Sheets("Copy").Columns("BH:BH").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Sheets("Copy").Columns("BL:BL").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Sheets("Copy").Columns("BP:BP").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Sheets("Copy").Columns("BS:BS").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
'Label new columns
    
    Sheets("Copy").Range("AF1") = "GLU Conv."
    Sheets("Copy").Range("AH1") = "HDL Conv."
    Sheets("Copy").Range("AJ1") = "Height Conv."
    Sheets("Copy").Range("AN1") = "Hip Conv."
    Sheets("Copy").Range("AR1") = "LDL Conv."
    Sheets("Copy").Range("BH1") = "TC Conv."
    Sheets("Copy").Range("BL1") = "TRIG Conv."
    Sheets("Copy").Range("BP1") = "Waist Conv."
    Sheets("Copy").Range("BS1") = "Weight Conv."
    Sheets("Copy").Range("BZ1") = "WHR"
    Sheets("Copy").Range("CA1") = "NCells"
    Sheets("Copy").Range("CB1") = "HealthScore Cat"
    Sheets("Copy").Range("CC1") = "TC Cat"
    Sheets("Copy").Range("CD1") = "HDL Cat"
    Sheets("Copy").Range("CE1") = "LDL Cat"
    Sheets("Copy").Range("CF1") = "Tri Cat"
    Sheets("Copy").Range("CG1") = "SBP Cat"
    Sheets("Copy").Range("CH1") = "DBP Cat"
    Sheets("Copy").Range("CI1") = "BP Cat"
    Sheets("Copy").Range("CJ1") = "GLU Cat"
    Sheets("Copy").Range("CK1") = "A1C Cat"
    Sheets("Copy").Range("CL1") = "BMI Cat"
    Sheets("Copy").Range("CM1") = "Waist Cat"
    Sheets("Copy").Range("CN1") = "WHR Cat"
    Sheets("Copy").Range("CO1") = "NIC Cat"
    Sheets("Copy").Range("CP1") = "GGT Cat"
    Sheets("Copy").Range("CQ1") = "PSA Cat"
    Sheets("Copy").Range("DA1") = "HS"
    Sheets("Copy").Range("DB1") = "TC"
    Sheets("Copy").Range("DC1") = "HDL"
    Sheets("Copy").Range("DD1") = "LDL"
    Sheets("Copy").Range("DE1") = "TRIG"
    Sheets("Copy").Range("DF1") = "SBP"
    Sheets("Copy").Range("DG1") = "DBP"
    Sheets("Copy").Range("DH1") = "BP"
    Sheets("Copy").Range("DI1") = "Glucose"
    Sheets("Copy").Range("DJ1") = "A1C"
    Sheets("Copy").Range("DK1") = "BMI"
    Sheets("Copy").Range("DL1") = "Waist"
    Sheets("Copy").Range("DM1") = "WHR"
    Sheets("Copy").Range("DN1") = "GGT"
    Sheets("Copy").Range("DO1") = "PSA"
    
'Input data into new columns

    Sheets("Copy").Range("AF2") = "=IF(AE2="""","""",AE2*18)"
    Sheets("Copy").Range("AH2") = "=IF(AG2="""","""",AG2*38.67)"
    Sheets("Copy").Range("AJ2") = "=IF(AI2="""","""",AI2/2.54)"
    Sheets("Copy").Range("AN2") = "=IF(AM2="""","""",AM2/2.54)"
    Sheets("Copy").Range("AR2") = "=IF(AQ2="""","""",AQ2*38.67)"
    Sheets("Copy").Range("BH2") = "=IF(BG2="""","""",BG2*38.67)"
    Sheets("Copy").Range("BL2") = "=IF(BK2="""","""",BK2*88.57)"
    Sheets("Copy").Range("BP2") = "=IF(BO2="""","""",BO2/2.54)"
    Sheets("Copy").Range("BS2") = "=IF(BR2="""","""",BR2*2.2)"
    Sheets("Copy").Range("BZ2") = "=IF(OR(BP2="""",AN2=""""),"""",BP2/AN2)"
    Sheets("Copy").Range("CA2") = "=COUNTA(A:A)"
    Sheets("Copy").Range("CB2") = "=IFS(BU2="""","""",BU2>=85,""Ideal"",BU2>=70,""Low"",BU2>=60,""Moderate"",BU2>=50,""High"",BU2>=0,""Very High"")" 'HS
    Sheets("Copy").Range("CC2") = "=IFS(BH2="""","""",BH2<200.5,""Low"",BH2<239.5,""Moderate"",BH2>=239.5,""High"")" ' TC
    Sheets("Copy").Range("CD2") = "=IFS(AH2="""","""",AH2>=49.5,""Low"",AH2>=39.5,""Moderate"",AH2<39.5,""High"")" ' HDL
    Sheets("Copy").Range("CE2") = "=IFS(AR2="""","""",AR2<129.5,""Low"",AR2<159.5,""Moderate"",AR2>=159.5,""High"")" ' LDL
    Sheets("Copy").Range("CF2") = "=IFS(BL2="""","""",BL2<149.5,""Low"",BL2<199.5,""Moderate"",BL2>=199.5,""High"")" ' TRIG
    Sheets("Copy").Range("CG2") = "=IFS(BF2="""","""",BF2<=121,""Low"",BF2<141,""Moderate"",BF2>=141,""High"")" ' SBP
    Sheets("Copy").Range("CH2") = "=IFS(Z2="""","""",Z2<=81,""Low"",Z2<91,""Moderate"",Z2>=91,""High"")" ' DBP
    Sheets("Copy").Range("CI2") = "" ' BP
    Sheets("Copy").Range("CJ2") = "=IFS(AF2="""","""",AF2<100.5,""Low"",AF2<125.5,""Moderate"",AF2>=125.5,""High"")" ' Glucose
    Sheets("Copy").Range("CK2") = "=IFS(M2="""","""",M2<5.7,""Low"",M2<6.5,""Moderate"",M2>=6.5,""High"")" ' A1C
    Sheets("Copy").Range("CL2") = "=IFS(BT2="""","""",BT2<30,""Low"",BT2<40,""Moderate"",BT2>=40,""High"")" 'BMI
    Sheets("Copy").Range("CM2") = "=IFS(BO2="""","""",AND(G2=""Male"",BP2<40),""Low"",AND(G2=""Male"",BP2<45),""Moderate"",AND(G2=""Male"",BP2>=45),""High"",BP2<35,""Low"",BP2<40,""Moderate"",BP2>=40,""High"")" ' Waist
    Sheets("Copy").Range("CN2") = "=IFS(BZ2="""","""",BZ2<=.95,""Low"",BZ2<1,""Moderate"",BZ2>=1,""High"")" 'WHR
    Sheets("Copy").Range("CO2") = "=IFS(AND(BY2="""", BV2=""""),"""",BY2=""POS"",""POS"",BY2=""NEG"",""NEG"",AND(BY2="""",BV2=""POS""),""POS"",AND(BY2="""",BV2=""NEG""),""NEG"",BY2=""SNS"","""",BY2=""QNS"","""",BY2=""TNP"","""",BY2=""NSA"","""")" 'NIC
    Sheets("Copy").Range("CP2") = "=IFS(AC2="""","""",AC2<66,""Low"",AC2>=66,""High"")" ' GGT
    Sheets("Copy").Range("CQ2") = "=IFS(BX2="""","""",BX2<=2.5,""Low"",BX2<6.5,""Moderate"",BX2>=6.5,""High"")" ' PSA
    Sheets("Copy").Range("DA2") = "=IF(BU2="""","""",BU2)"
    Sheets("Copy").Range("DB2") = "=IF(BH2="""","""",BH2)"
    Sheets("Copy").Range("DC2") = "=IF(AH2="""","""",AH2)"
    Sheets("Copy").Range("DD2") = "=IF(AR2="""","""",AR2)"
    Sheets("Copy").Range("DE2") = "=IF(BL2="""","""",BL2)"
    Sheets("Copy").Range("DF2") = "=IF(BF2="""","""",BF2)"
    Sheets("Copy").Range("DG2") = "=IF(Z2="""","""",Z2)"
    Sheets("Copy").Range("DH2") = ""
    Sheets("Copy").Range("DI2") = "=IF(AF2="""","""",AF2)"
    Sheets("Copy").Range("DJ2") = "=IF(M2="""","""",M2)"
    Sheets("Copy").Range("DK2") = "=IF(BT2="""","""",BT2)"
    Sheets("Copy").Range("DL2") = "=IF(BP2="""","""",BP2)"
    Sheets("Copy").Range("DM2") = "=IF(BZ2="""","""",BZ2)"
    Sheets("Copy").Range("DN2") = "=IF(AC2="""","""",AC2)"
    Sheets("Copy").Range("DO2") = "=IF(BX2="""","""",BX2)"
    
    
NCells = Worksheets("Copy").Range("CA2").Value

'Autofill columns

Sheets("Copy").Range("AF2").AutoFill Destination:=Range("'Copy'!AF2:'Copy'!AF" & NCells)
Sheets("Copy").Range("AH2").AutoFill Destination:=Range("'Copy'!AH2:'Copy'!AH" & NCells)
Sheets("Copy").Range("AJ2").AutoFill Destination:=Range("'Copy'!AJ2:'Copy'!AJ" & NCells)
Sheets("Copy").Range("AN2").AutoFill Destination:=Range("'Copy'!AN2:'Copy'!AN" & NCells)
Sheets("Copy").Range("AR2").AutoFill Destination:=Range("'Copy'!AR2:'Copy'!AR" & NCells)
Sheets("Copy").Range("BH2").AutoFill Destination:=Range("'Copy'!BH2:'Copy'!BH" & NCells)
Sheets("Copy").Range("BL2").AutoFill Destination:=Range("'Copy'!BL2:'Copy'!BL" & NCells)
Sheets("Copy").Range("BP2").AutoFill Destination:=Range("'Copy'!BP2:'Copy'!BP" & NCells)
Sheets("Copy").Range("BS2").AutoFill Destination:=Range("'Copy'!BS2:'Copy'!BS" & NCells)
Sheets("Copy").Range("BZ2").AutoFill Destination:=Range("'Copy'!BZ2:'Copy'!BZ" & NCells)
Sheets("Copy").Range("CB2:CQ2").AutoFill Destination:=Range("'Copy'!CB2:'Copy'!CQ" & NCells)
Sheets("Copy").Range("DA2:DO2").AutoFill Destination:=Range("'Copy'!DA2:'Copy'!DO" & NCells)




Call InformationCalculation

End Sub

Sub InformationCalculation()

Dim NPersons As Integer, NYears As Integer, NCells As Integer

NCells = Worksheets("Copy").Range("CA2").Value
Sheets("Copy").Range("CR1") = "NPersons"
Sheets("Copy").Range("CR2") = "=CA2-1"
NPersons = Worksheets("Copy").Range("CR2").Value
Sheets("Copy").Range("CS1") = "Year"
Sheets("Copy").Range("CS2") = "=YEAR(B2)"
Sheets("Copy").Range("CS2").AutoFill Destination:=Range("'Copy'!CS2:'Copy'!CS" & NCells)
Sheets("Copy").Range("CT1") = "NYears"
NYears = Worksheets("Copy").Range("CR2").Value
Sheets("Copy").Range("CT2") = "=SUM(IF(FREQUENCY(INDIRECT(""CS2:CS""&CA2),INDIRECT(""CS2:CS""&CA2))>0,1))"
Sheets("Copy").Range("CU1") = "Years"
Sheets("Copy").Range("CV1") = "MinYear"
Sheets("Copy").Range("CW1") = "MaxYear"

Dim MaxYear As Integer, MinYear As Integer, YearPaste As Integer
MaxYear = Application.WorksheetFunction.Max(Worksheets("Copy").Range("CS:CS"))
MinYear = Application.WorksheetFunction.Min(Worksheets("Copy").Range("CS:CS"))
YearPaste = 2
Sheets("Copy").Range("CV2") = MinYear
Sheets("Copy").Range("CW2") = MaxYear

Do Until MinYear > MaxYear

Sheets("Copy").Range("CU" & YearPaste) = MinYear

MinYear = MinYear + 1
YearPaste = YearPaste + 1

Loop

    ActiveWindow.ScrollColumn = 1
    
Call ChangebwYears

End Sub

Sub ChangebwYears()

Dim MaxYear As Integer, MinYear As Integer, FinalYearStr As String, PrevYearStr As String, IDStr As String, xrowa As Integer, xrowb As Integer, NCells As Integer, ycolumna As Integer, ycolumnb As Integer
Dim Match As Boolean, xrowc As Integer, FinalYearVal As Integer, PrevYearVal As Integer

Sheets("Copy").Range("DQ1") = "HS Change"
Sheets("Copy").Range("DR1") = "TC Change"
Sheets("Copy").Range("DS1") = "HDL Change"
Sheets("Copy").Range("DT1") = "LDL Change"
Sheets("Copy").Range("DU1") = "TRIG Change"
Sheets("Copy").Range("DV1") = "SBP Change"
Sheets("Copy").Range("DW1") = "DBP Change"
Sheets("Copy").Range("DX1") = "GLU Change"
Sheets("Copy").Range("DY1") = "A1C Change"
Sheets("Copy").Range("DZ1") = "BMI Change"
Sheets("Copy").Range("EA1") = "Waist Change"
Sheets("Copy").Range("EB1") = "WHR Change"
Sheets("Copy").Range("EC1") = "NIC Change"
Sheets("Copy").Range("ED1") = "GGT Change"
Sheets("Copy").Range("EE1") = "PSA Change"

Sheets("Copy").Range("EG1") = "HS Names"
Sheets("Copy").Range("EH1") = "TC Names"
Sheets("Copy").Range("EI1") = "HDL Names"
Sheets("Copy").Range("EJ1") = "LDL Names"
Sheets("Copy").Range("EK1") = "TRIG Names"
Sheets("Copy").Range("EL1") = "SBP Names"
Sheets("Copy").Range("EM1") = "DBP Names"
Sheets("Copy").Range("EN1") = "GLU Names"
Sheets("Copy").Range("EO1") = "A1C Names"
Sheets("Copy").Range("EP1") = "BMI Names"
Sheets("Copy").Range("EQ1") = "Waist Names"
Sheets("Copy").Range("ER1") = "WHR Names"
Sheets("Copy").Range("ES1") = "NIC Names"
Sheets("Copy").Range("ET1") = "GGT Names"
Sheets("Copy").Range("EU1") = "PSA Names"

MinYear = Sheets("Copy").Range("CV2").Value
MaxYear = Sheets("Copy").Range("CW2").Value
NCells = Sheets("Copy").Range("CR2").Value + 1
xrowa = 2
xrowb = 2
xrowc = 2
ycolumna = 80
ycolumnb = 80
Match = False

Do While ycolumna <= 95

    If ycolumna = 87 Then
        ycolumna = 88
    End If
    
    If ycolumna >= 88 Then
        ycolumnb = ycolumna - 1
    End If

    Do While xrowa <= NCells
    
        If Sheets("Copy").Range("CS" & xrowa).Value = MaxYear Then
        
            FinalYearStr = Sheets("Copy").Cells(xrowa, ycolumna)
            IDStr = Sheets("Copy").Cells(xrowa, 1)
            
            Do While xrowb <= NCells And Match = False
            
                If Sheets("Copy").Cells(xrowb, 1) = IDStr And Sheets("Copy").Cells(xrowb, 97).Value = MaxYear - 1 Then
                
                    PrevYearStr = Sheets("Copy").Cells(xrowb, ycolumna)
                    Match = True
                    
                    If FinalYearStr = "Ideal" Then
                        FinalYearVal = 5
                    End If
                    If FinalYearStr = "Low" Then
                        FinalYearVal = 4
                    End If
                    If FinalYearStr = "Moderate" Then
                        FinalYearVal = 3
                    End If
                    If FinalYearStr = "High" Then
                        FinalYearVal = 2
                    End If
                    If FinalYearStr = "Very High" Then
                        FinalYearVal = 1
                    End If
                    If FinalYearStr = "NEG" Then
                        FinalYearVal = 6
                    End If
                    If FinalYearStr = "POS" Then
                        FinalYearVal = 7
                    End If
                    If FinalYearStr = "" Or FinalYearStr = "QNS" Or FinalYearStr = "SNS" Then
                        FinalYearVal = 0
                    End If
                    If PrevYearStr = "Ideal" Then
                        PrevYearVal = 5
                    End If
                    If PrevYearStr = "Low" Then
                        PrevYearVal = 4
                    End If
                    If PrevYearStr = "Moderate" Then
                        PrevYearVal = 3
                    End If
                    If PrevYearStr = "High" Then
                        PrevYearVal = 2
                    End If
                    If PrevYearStr = "Very High" Then
                        PrevYearVal = 1
                    End If
                    If PrevYearStr = "NEG" Then
                        PrevYearVal = 6
                    End If
                    If PrevYearStr = "POS" Then
                        PrevYearVal = 7
                    End If
                    If PrevYearStr = "" Or PrevYearStr = "QNS" Or PrevYearStr = "SNS" Then
                        PrevYearVal = 0
                    End If
                    
                    If PrevYearVal = 0 Or FinalYearVal = 0 Then
                        GoTo NoDataForChange
                    End If
                    
'                    If PrevYearVal > FinalYearVal Then
'                        Sheets("Copy").Cells(xrowc, ycolumnb + 41) = "Negative Change"
'                    End If
'                    If PrevYearVal < FinalYearVal Then
'                        Sheets("Copy").Cells(xrowc, ycolumnb + 41) = "Positive Change"
'                    End If
'                    If PrevYearVal = FinalYearVal Then
'                        Sheets("Copy").Cells(xrowc, ycolumnb + 41) = "No Change"
'                    End If

                    If PrevYearVal = 7 Then
                        
                        If FinalYearVal = 7 Then
                            Sheets("Copy").Cells(xrowc, ycolumnb + 41) = "POS-POS"
                        End If
                        
                        If FinalYearVal = 6 Then
                            Sheets("Copy").Cells(xrowc, ycolumnb + 41) = "POS-NEG"
                        End If
                        
                    End If
                    
                    If PrevYearVal = 6 Then
                    
                        If FinalYearVal = 7 Then
                            Sheets("Copy").Cells(xrowc, ycolumnb + 41) = "NEG-POS"
                        End If
                        
                        If FinalYearVal = 6 Then
                            Sheets("Copy").Cells(xrowc, ycolumnb + 41) = "NEG-NEG"
                        End If
                    
                    End If

                    If PrevYearVal = 5 Then
                    
                        If FinalYearVal = 5 Then
                            Sheets("Copy").Cells(xrowc, ycolumnb + 41) = "Ideal-Ideal"
                        End If
                        
                        If FinalYearVal = 4 Then
                            Sheets("Copy").Cells(xrowc, ycolumnb + 41) = "Ideal-Low"
                        End If
                        
                        If FinalYearVal = 3 Then
                            Sheets("Copy").Cells(xrowc, ycolumnb + 41) = "Ideal-Moderate"
                        End If
                        
                        If FinalYearVal = 2 Then
                            Sheets("Copy").Cells(xrowc, ycolumnb + 41) = "Ideal-High"
                        End If
                        
                        If FinalYearVal = 1 Then
                            Sheets("Copy").Cells(xrowc, ycolumnb + 41) = "Ideal-Very High"
                        End If
                    
                    End If
                    
                    If PrevYearVal = 4 Then
                    
                        If FinalYearVal = 5 Then
                            Sheets("Copy").Cells(xrowc, ycolumnb + 41) = "Low-Ideal"
                        End If
                        
                        If FinalYearVal = 4 Then
                            Sheets("Copy").Cells(xrowc, ycolumnb + 41) = "Low-Low"
                        End If
                        
                        If FinalYearVal = 3 Then
                            Sheets("Copy").Cells(xrowc, ycolumnb + 41) = "Low-Moderate"
                        End If
                        
                        If FinalYearVal = 2 Then
                            Sheets("Copy").Cells(xrowc, ycolumnb + 41) = "Low-High"
                        End If
                        
                        If FinalYearVal = 1 Then
                            Sheets("Copy").Cells(xrowc, ycolumnb + 41) = "Low-Very High"
                        End If
                    
                    End If
                    
                    If PrevYearVal = 3 Then
                    
                        If FinalYearVal = 5 Then
                            Sheets("Copy").Cells(xrowc, ycolumnb + 41) = "Moderate-Ideal"
                        End If
                        
                        If FinalYearVal = 4 Then
                            Sheets("Copy").Cells(xrowc, ycolumnb + 41) = "Moderate-Low"
                        End If
                        
                        If FinalYearVal = 3 Then
                            Sheets("Copy").Cells(xrowc, ycolumnb + 41) = "Moderate-Moderate"
                        End If
                        
                        If FinalYearVal = 2 Then
                            Sheets("Copy").Cells(xrowc, ycolumnb + 41) = "Moderate-High"
                        End If
                        
                        If FinalYearVal = 1 Then
                            Sheets("Copy").Cells(xrowc, ycolumnb + 41) = "Moderate-Very High"
                        End If
                    
                    End If
                    
                    If PrevYearVal = 2 Then
                    
                        If FinalYearVal = 5 Then
                            Sheets("Copy").Cells(xrowc, ycolumnb + 41) = "High-Ideal"
                        End If
                        
                        If FinalYearVal = 4 Then
                            Sheets("Copy").Cells(xrowc, ycolumnb + 41) = "High-Low"
                        End If
                        
                        If FinalYearVal = 3 Then
                            Sheets("Copy").Cells(xrowc, ycolumnb + 41) = "High-Moderate"
                        End If
                        
                        If FinalYearVal = 2 Then
                            Sheets("Copy").Cells(xrowc, ycolumnb + 41) = "High-High"
                        End If
                        
                        If FinalYearVal = 1 Then
                            Sheets("Copy").Cells(xrowc, ycolumnb + 41) = "High-Very High"
                        End If
                    
                    End If
                    
                    If PrevYearVal = 1 Then
                    
                        If FinalYearVal = 5 Then
                            Sheets("Copy").Cells(xrowc, ycolumnb + 41) = "Very High-Ideal"
                        End If
                        
                        If FinalYearVal = 4 Then
                            Sheets("Copy").Cells(xrowc, ycolumnb + 41) = "Very High-Low"
                        End If
                        
                        If FinalYearVal = 3 Then
                            Sheets("Copy").Cells(xrowc, ycolumnb + 41) = "Very High-Moderate"
                        End If
                        
                        If FinalYearVal = 2 Then
                            Sheets("Copy").Cells(xrowc, ycolumnb + 41) = "Very High-High"
                        End If
                        
                        If FinalYearVal = 1 Then
                            Sheets("Copy").Cells(xrowc, ycolumnb + 41) = "Very High-Very High"
                        End If
                    
                    End If
                    
                    Sheets("Copy").Cells(xrowc, ycolumnb + 57) = Application.WorksheetFunction.Concat(Sheets("Copy").Cells(xrowa, 6), " ", Sheets("Copy").Cells(xrowa, 9))
                    
                    xrowc = xrowc + 1
                    
NoDataForChange:
                    
                End If
                
                xrowb = xrowb + 1
            
            Loop
            
        End If
        
        xrowa = xrowa + 1
        xrowb = 2
        Match = False
    
    Loop
    
    xrowa = 2
    xrowc = 2
    ycolumna = ycolumna + 1
    ycolumnb = ycolumna

Loop

Call AddSheets

End Sub

Sub AddSheets()

With Sheets
    .Add().Name = "HealthScore"
    .Add().Name = "TotalCholesterol"
    .Add().Name = "HDL"
    .Add().Name = "LDL"
    .Add().Name = "TRIG"
    .Add().Name = "SBP"
    .Add().Name = "DBP"
    .Add().Name = "BP"
    .Add().Name = "Glucose"
    .Add().Name = "A1C"
    .Add().Name = "BMI"
    .Add().Name = "Waist"
    .Add().Name = "WaistbyGender"
    .Add().Name = "WHR"
    .Add().Name = "Nicotine"
    .Add().Name = "GGT"
    .Add().Name = "PSA"
    .Add().Name = "Age"
    .Add().Name = "Averages"
    .Add().Name = "MetabolicSyndromes"
End With

Dim shArray() As Variant, yrArray() As Variant, k As Integer, j As Integer, i As Double, xrow As Integer, ycolumn As Integer, NYears As Integer

NYears = Worksheets("Copy").Range("CT2").Value

xrow = 2
i = 0

ReDim yrArray(0)

Do Until Sheets("Copy").Cells(xrow, "CU").Value = ""
    yrArray(i) = Sheets("Copy").Cells(xrow, "CU").Value

i = i + 1
ReDim Preserve yrArray(i)
xrow = xrow + 1

Loop

shArray = Array("HealthScore", "TotalCholesterol", "HDL", "LDL", "TRIG", "SBP", "DBP", "BP", "Glucose", "A1C", "BMI", "Waist", "WHR", "Nicotine", "GGT", "PSA")


i = 0
j = 0
k = 0
ycolumn = 1

For i = LBound(yrArray, 1) To UBound(yrArray, 1)

    For j = LBound(shArray, 1) To UBound(shArray, 1)
        Sheets(shArray(j)).Cells(1, ycolumn) = yrArray(i)
        
    Next j

ycolumn = ycolumn + 1

Next i

i = 0
j = 0
ycolumn = NYears + 2

For i = LBound(yrArray, 1) To UBound(yrArray, 1)

    For j = LBound(shArray, 1) To UBound(shArray, 1)
        Sheets(shArray(j)).Cells(1, ycolumn) = yrArray(i)
        
    Next j

ycolumn = ycolumn + 1

Next i


'''Populate sheets w/ data

Dim CurrentYear As Integer, MaxYear As Integer, NCells As Integer
Dim CatArray() As Variant, CopyColumn As Integer, PasteColumn As Integer
Dim DataArray() As Variant, CopyColumnB As Integer, PasteColumnB As Integer
CurrentYear = Sheets("Copy").Range("CV2").Value
MaxYear = Sheets("Copy").Range("CW2").Value
CopyColumn = 80
PasteColumn = 1
CopyColumnB = 105
PasteColumnB = 2 + NYears
NCells = Worksheets("Copy").Range("CA2").Value

xrow = 2
i = 0
j = 0
k = 0

For j = LBound(shArray, 1) To UBound(shArray, 1)

    Do While CurrentYear <= MaxYear

        ReDim CatArray(0)
    
    'Fill Category Array
    
        Do While xrow <= NCells
                 
            If Sheets("Copy").Cells(xrow, "CS").Value = CurrentYear Then
                CatArray(i) = Sheets("Copy").Cells(xrow, CopyColumn).Value
                i = i + 1
                ReDim Preserve CatArray(i)
            End If
    
        xrow = xrow + 1
        
        Loop
    
        xrow = 2
        k = 0
        
        ReDim DataArray(0)
        
    'Fill Data Array
        
        If shArray(j) = "Nicotine" Then
        GoTo NicotineData:
        End If
        
        Do While xrow <= NCells
                 
            If Sheets("Copy").Cells(xrow, "CS").Value = CurrentYear Then
                DataArray(k) = Sheets("Copy").Cells(xrow, CopyColumnB).Value
                k = k + 1
                ReDim Preserve DataArray(k)
            End If
            
        xrow = xrow + 1
        
        Loop
        
NicotineData:
        
        xrow = 2
        i = 0
        
    'Paste Category Array
    
        For i = LBound(CatArray, 1) To UBound(CatArray, 1)
            Sheets(shArray(j)).Cells(xrow, PasteColumn) = CatArray(i)
            xrow = xrow + 1
        Next i
        
        xrow = 2
        k = 0
        
    'Paste Data Array
        
        If shArray(j) = "Nicotine" Then
        GoTo NicotineDataB:
        End If
        
        For k = LBound(DataArray, 1) To UBound(DataArray, 1)
            Sheets(shArray(j)).Cells(xrow, PasteColumnB) = DataArray(k)
            xrow = xrow + 1
        Next k
        
NicotineDataB:
                
        xrow = 2
        i = 0
        k = 0
    
        CurrentYear = CurrentYear + 1
        
        PasteColumn = PasteColumn + 1
        
        PasteColumnB = PasteColumnB + 1

    Loop
    
    CurrentYear = Sheets("Copy").Range("CV2").Value
    xrow = 2
    i = 0
    CopyColumn = CopyColumn + 1
    PasteColumn = 1
    
    If shArray(j) = "Nicotine" Then
    GoTo NicotineDataC:
    End If
    
    CopyColumnB = CopyColumnB + 1
    
NicotineDataC:
    
    PasteColumnB = NYears + 2

Next j

Call ChartFormatting

End Sub

Sub ChartFormatting()

Dim NYears As Integer, i As Integer, MeasureYear As Integer, MaxYear As Integer, CellShift As Integer, NCells As Integer, AverageColumn As Integer
Dim Low As Single, Moderate As Single, High As Single, POS As Single, NEG As Single, Ideal As Single, VeryHigh As Single, Average As Single
Dim shArray() As Variant

shArray = Array("HealthScore", "TotalCholesterol", "HDL", "LDL", "TRIG", "SBP", "DBP", "BP", "Glucose", "A1C", "BMI", "Waist", "WHR", "Nicotine", "GGT", "PSA")

i = 0
NYears = Worksheets("Copy").Range("CT2").Value
CellShift = 1
NCells = Worksheets("Copy").Range("CA2").Value
MaxYear = Sheets("Copy").Range("CW2").Value

For i = LBound(shArray, 1) To UBound(shArray, 1)
    Worksheets(shArray(i)).Activate
    Sheets(shArray(i)).Range(Cells(1, 1), Cells(1, NYears)).Select
    Selection.Copy
    Sheets(shArray(i)).Range(Cells(1, 2 * NYears + 4), Cells(1, 3 * NYears + 4)).PasteSpecial
    Sheets(shArray(i)).Cells(2, 2 * NYears + 3) = "Low"
    Sheets(shArray(i)).Cells(3, 2 * NYears + 3) = "Moderate"
    Sheets(shArray(i)).Cells(4, 2 * NYears + 3) = "High"
    Sheets(shArray(i)).Cells(5, 2 * NYears + 3) = "Average"
    
    MeasureYear = Sheets("Copy").Range("CV2").Value
    CellShift = 1
    AverageColumn = NYears + 2
    
    Do While MeasureYear <= MaxYear
    
        If Application.WorksheetFunction.CountA(Sheets(shArray(i)).Range(Cells(2, CellShift), Cells(NCells, CellShift))) = 0 Then
        GoTo NoData:
        End If
        
'        Low = Application.WorksheetFunction.CountIf(Sheets(shArray(i)).Range(Cells(2, CellShift), Cells(NCells, CellShift)), "Low") / Application.WorksheetFunction.CountA(Sheets(shArray(i)).Range(Cells(2, CellShift), Cells(NCells, CellShift)))
'        Moderate = Application.WorksheetFunction.CountIf(Sheets(shArray(i)).Range(Cells(2, CellShift), Cells(NCells, CellShift)), "Moderate") / Application.WorksheetFunction.CountA(Sheets(shArray(i)).Range(Cells(2, CellShift), Cells(NCells, CellShift)))
'        High = Application.WorksheetFunction.CountIf(Sheets(shArray(i)).Range(Cells(2, CellShift), Cells(NCells, CellShift)), "High") / Application.WorksheetFunction.CountA(Sheets(shArray(i)).Range(Cells(2, CellShift), Cells(NCells, CellShift)))
        
        Low = Application.WorksheetFunction.CountIf(Sheets(shArray(i)).Range(Cells(2, CellShift), Cells(NCells, CellShift)), "Low")
        Moderate = Application.WorksheetFunction.CountIf(Sheets(shArray(i)).Range(Cells(2, CellShift), Cells(NCells, CellShift)), "Moderate")
        High = Application.WorksheetFunction.CountIf(Sheets(shArray(i)).Range(Cells(2, CellShift), Cells(NCells, CellShift)), "High")
        
        If shArray(i) = "Nicotine" Then
        GoTo NicotineDataAA:
        End If
        
        If Application.WorksheetFunction.CountA(Sheets(shArray(i)).Range(Cells(2, AverageColumn), Cells(NCells, AverageColumn))) = 0 Then
        GoTo NoData:
        End If
        
        Average = Application.WorksheetFunction.Average(Sheets(shArray(i)).Range(Cells(2, AverageColumn), Cells(NCells, AverageColumn)))
        
NicotineDataAA:
    
        Sheets(shArray(i)).Cells(2, 2 * NYears + 3 + CellShift) = Low
        Sheets(shArray(i)).Cells(3, 2 * NYears + 3 + CellShift) = Moderate
        Sheets(shArray(i)).Cells(4, 2 * NYears + 3 + CellShift) = High
        Sheets(shArray(i)).Range(Cells(2, 2 * NYears + 3 + CellShift), Cells(4, 2 * NYears + 3 + CellShift)).Select
        
        If shArray(i) = "Nicotine" Then
        GoTo NoData:
        End If
        
        Sheets(shArray(i)).Cells(5, 2 * NYears + 3 + CellShift) = Average
        
NoData:
        
        MeasureYear = MeasureYear + 1
        CellShift = CellShift + 1
        AverageColumn = AverageColumn + 1
    
    Loop
    
Next i


Worksheets("Nicotine").Activate
Sheets("Nicotine").Cells(2, 2 * NYears + 3) = "POS"
Sheets("Nicotine").Cells(3, 2 * NYears + 3) = "NEG"
Sheets("Nicotine").Range(Cells(4, 2 * NYears + 3), Cells(5, 3 * NYears + 3)).Select
Selection.ClearContents
Sheets("Nicotine").Range(Cells(1, NYears + 2), Cells(1, 2 * NYears + 2)).Select
Selection.ClearContents

Worksheets("GGT").Activate
Sheets("GGT").Cells(2, 2 * NYears + 3) = "Low"
Sheets("GGT").Cells(3, 2 * NYears + 3) = "High"
Sheets("GGT").Cells(4, 2 * NYears + 3) = "Average"
Sheets("GGT").Range(Cells(5, 2 * NYears + 3), Cells(5, 3 * NYears + 3)).Select
Selection.ClearContents

Worksheets("HealthScore").Activate
Sheets("HealthScore").Cells(2, 2 * NYears + 3) = "Ideal"
Sheets("HealthScore").Cells(3, 2 * NYears + 3) = "Low"
Sheets("HealthScore").Cells(4, 2 * NYears + 3) = "Moderate"
Sheets("HealthScore").Cells(5, 2 * NYears + 3) = "High"
Sheets("HealthScore").Cells(6, 2 * NYears + 3) = "Very High"
Sheets("HealthScore").Cells(7, 2 * NYears + 3) = "Average"

Worksheets("BP").Activate
Sheets("BP").Cells(5, 2 * NYears + 3) = ""


MeasureYear = Sheets("Copy").Range("CV2").Value
CellShift = 1
    

Worksheets("Nicotine").Activate

Do While MeasureYear <= MaxYear

    If Application.WorksheetFunction.CountA(Sheets("Nicotine").Range(Cells(2, CellShift), Cells(NCells, CellShift))) = 0 Then
    GoTo NoData2:
    End If
    

    POS = Application.WorksheetFunction.CountIf(Sheets("Nicotine").Range(Cells(2, CellShift), Cells(NCells, CellShift)), "POS")
    NEG = Application.WorksheetFunction.CountIf(Sheets("Nicotine").Range(Cells(2, CellShift), Cells(NCells, CellShift)), "NEG")
    Sheets("Nicotine").Cells(2, 2 * NYears + 3 + CellShift) = POS
    Sheets("Nicotine").Cells(3, 2 * NYears + 3 + CellShift) = NEG
    Sheets("Nicotine").Range(Cells(2, 2 * NYears + 3 + CellShift), Cells(3, 2 * NYears + 3 + CellShift)).Select
    
NoData2:
    
    MeasureYear = MeasureYear + 1
    CellShift = CellShift + 1
    
Loop

MeasureYear = Sheets("Copy").Range("CV2").Value
CellShift = 1
AverageColumn = NYears + 2

Worksheets("HealthScore").Activate

Do While MeasureYear <= MaxYear

    If Application.WorksheetFunction.CountA(Sheets("HealthScore").Range(Cells(2, CellShift), Cells(NCells, CellShift))) = 0 Then
    GoTo NoData3:
    End If
    
    
    Ideal = Application.WorksheetFunction.CountIf(Sheets("HealthScore").Range(Cells(2, CellShift), Cells(NCells, CellShift)), "Ideal")
    Low = Application.WorksheetFunction.CountIf(Sheets("HealthScore").Range(Cells(2, CellShift), Cells(NCells, CellShift)), "Low")
    Moderate = Application.WorksheetFunction.CountIf(Sheets("HealthScore").Range(Cells(2, CellShift), Cells(NCells, CellShift)), "Moderate")
    High = Application.WorksheetFunction.CountIf(Sheets("HealthScore").Range(Cells(2, CellShift), Cells(NCells, CellShift)), "High")
    VeryHigh = Application.WorksheetFunction.CountIf(Sheets("HealthScore").Range(Cells(2, CellShift), Cells(NCells, CellShift)), "Very High")
    Average = Application.WorksheetFunction.Average(Sheets("HealthScore").Range(Cells(2, AverageColumn), Cells(NCells, AverageColumn)))
    
    Sheets("HealthScore").Cells(2, 2 * NYears + 3 + CellShift) = Ideal
    Sheets("HealthScore").Cells(3, 2 * NYears + 3 + CellShift) = Low
    Sheets("HealthScore").Cells(4, 2 * NYears + 3 + CellShift) = Moderate
    Sheets("HealthScore").Cells(5, 2 * NYears + 3 + CellShift) = High
    Sheets("HealthScore").Cells(6, 2 * NYears + 3 + CellShift) = VeryHigh
    Sheets("HealthScore").Cells(7, 2 * NYears + 3 + CellShift) = Average
    Sheets("HealthScore").Range(Cells(2, 2 * NYears + 3 + CellShift), Cells(6, 2 * NYears + 3 + CellShift)).Select
    
NoData3:
    
    MeasureYear = MeasureYear + 1
    CellShift = CellShift + 1
    AverageColumn = AverageColumn + 1
    
Loop

MeasureYear = Sheets("Copy").Range("CV2").Value
CellShift = 1
AverageColumn = NYears + 2

Worksheets("GGT").Activate

Do While MeasureYear <= MaxYear

    If Application.WorksheetFunction.CountA(Sheets("GGT").Range(Cells(2, CellShift), Cells(NCells, CellShift))) = 0 Then
    GoTo NoData4:
    End If

    Low = Application.WorksheetFunction.CountIf(Sheets("GGT").Range(Cells(2, CellShift), Cells(NCells, CellShift)), "Low")
    High = Application.WorksheetFunction.CountIf(Sheets("GGT").Range(Cells(2, CellShift), Cells(NCells, CellShift)), "High")
    Average = Application.WorksheetFunction.Average(Sheets("GGT").Range(Cells(2, AverageColumn), Cells(NCells, AverageColumn)))
    
    Sheets("GGT").Cells(2, 2 * NYears + 3 + CellShift) = Low
    Sheets("GGT").Cells(3, 2 * NYears + 3 + CellShift) = High
    Sheets("GGT").Cells(4, 2 * NYears + 3 + CellShift) = Average
    Sheets("GGT").Range(Cells(2, 2 * NYears + 3 + CellShift), Cells(3, 2 * NYears + 3 + CellShift)).Select
    Sheets("GGT").Range(Cells(4, 2 * NYears + 3 + CellShift), Cells(4, 2 * NYears + 3 + CellShift)).Select
    Selection.NumberFormat = "General"

NoData4:

    MeasureYear = MeasureYear + 1
    CellShift = CellShift + 1
    AverageColumn = AverageColumn + 1

Loop

MeasureYear = Sheets("Copy").Range("CV2").Value
CellShift = 1
AverageColumn = NYears + 2

Worksheets("BP").Activate

Do While MeasureYear <= MaxYear

    Sheets("BP").Cells(2, NYears * 2 + 3 + CellShift) = Application.WorksheetFunction.Sum(Sheets("SBP").Cells(2, NYears * 2 + 3 + CellShift), Sheets("DBP").Cells(2, NYears * 2 + 3 + CellShift)) / 2
    Sheets("BP").Cells(3, NYears * 2 + 3 + CellShift) = Application.WorksheetFunction.Sum(Sheets("SBP").Cells(3, NYears * 2 + 3 + CellShift), Sheets("DBP").Cells(3, NYears * 2 + 3 + CellShift)) / 2
    Sheets("BP").Cells(4, NYears * 2 + 3 + CellShift) = Application.WorksheetFunction.Sum(Sheets("SBP").Cells(4, NYears * 2 + 3 + CellShift), Sheets("DBP").Cells(4, NYears * 2 + 3 + CellShift)) / 2
    Sheets("BP").Range(Cells(2, 2 * NYears + 3 + CellShift), Cells(4, 2 * NYears + 3 + CellShift)).Select
    MeasureYear = MeasureYear + 1
    CellShift = CellShift + 1
    AverageColumn = AverageColumn + 1

Loop

Call DataChanges

End Sub

Sub DataChanges()

Dim shArray() As Variant
Dim NYears As Integer, i As Integer, ycolumn As Integer, NCells As Integer

shArray = Array("HealthScore", "TotalCholesterol", "HDL", "LDL", "TRIG", "SBP", "DBP", "Glucose", "A1C", "BMI", "Waist", "WHR", "Nicotine", "GGT", "PSA")

NCells = Worksheets("Copy").Range("CA2").Value

i = 0
NYears = Worksheets("Copy").Range("CT2").Value
ycolumn = 121

Sheets("Copy").Activate

For i = LBound(shArray, 1) To UBound(shArray, 1)

    If shArray(i) = "HealthScore" Then
    
        Sheets(shArray(i)).Cells(2, NYears * 3 + 5) = "Remained Ideal or Low"
        Sheets(shArray(i)).Cells(3, NYears * 3 + 5) = "Remained High or Very High"
        Sheets(shArray(i)).Cells(4, NYears * 3 + 5) = "Remained Moderate"
        Sheets(shArray(i)).Cells(5, NYears * 3 + 5) = "Ideal or Low to Moderate"
        Sheets(shArray(i)).Cells(6, NYears * 3 + 5) = "Ideal or Low to High or Very High"
        Sheets(shArray(i)).Cells(7, NYears * 3 + 5) = "High or Very High to Low or Ideal"
        Sheets(shArray(i)).Cells(8, NYears * 3 + 5) = "High or Very High to Moderate"
        Sheets(shArray(i)).Cells(9, NYears * 3 + 5) = "Moderate to Ideal or Low"
        Sheets(shArray(i)).Cells(10, NYears * 3 + 5) = "Moderate to High or Very High"
        
        Sheets(shArray(i)).Cells(2, NYears * 3 + 6) = Application.WorksheetFunction.CountIf(Sheets("Copy").Range(Cells(1, ycolumn), Cells(NCells, ycolumn)), "Low-Low") + Application.WorksheetFunction.CountIf(Sheets("Copy").Range(Cells(1, ycolumn), Cells(NCells, ycolumn)), "Low-Ideal") + Application.WorksheetFunction.CountIf(Sheets("Copy").Range(Cells(1, ycolumn), Cells(NCells, ycolumn)), "Ideal-Low") + Application.WorksheetFunction.CountIf(Sheets("Copy").Range(Cells(1, ycolumn), Cells(NCells, ycolumn)), "Ideal-Ideal")
        Sheets(shArray(i)).Cells(3, NYears * 3 + 6) = Application.WorksheetFunction.CountIf(Sheets("Copy").Range(Cells(1, ycolumn), Cells(NCells, ycolumn)), "High-High") + Application.WorksheetFunction.CountIf(Sheets("Copy").Range(Cells(1, ycolumn), Cells(NCells, ycolumn)), "High-Very High") + Application.WorksheetFunction.CountIf(Sheets("Copy").Range(Cells(1, ycolumn), Cells(NCells, ycolumn)), "Very High-High") + Application.WorksheetFunction.CountIf(Sheets("Copy").Range(Cells(1, ycolumn), Cells(NCells, ycolumn)), "Very High-Very High")
        Sheets(shArray(i)).Cells(4, NYears * 3 + 6) = Application.WorksheetFunction.CountIf(Sheets("Copy").Range(Cells(1, ycolumn), Cells(NCells, ycolumn)), "Moderate-Moderate")
        Sheets(shArray(i)).Cells(5, NYears * 3 + 6) = Application.WorksheetFunction.CountIf(Sheets("Copy").Range(Cells(1, ycolumn), Cells(NCells, ycolumn)), "Low-Moderate") + Application.WorksheetFunction.CountIf(Sheets("Copy").Range(Cells(1, ycolumn), Cells(NCells, ycolumn)), "Ideal-Moderate")
        Sheets(shArray(i)).Cells(6, NYears * 3 + 6) = Application.WorksheetFunction.CountIf(Sheets("Copy").Range(Cells(1, ycolumn), Cells(NCells, ycolumn)), "Ideal-High") + Application.WorksheetFunction.CountIf(Sheets("Copy").Range(Cells(1, ycolumn), Cells(NCells, ycolumn)), "Ideal-Very High") + Application.WorksheetFunction.CountIf(Sheets("Copy").Range(Cells(1, ycolumn), Cells(NCells, ycolumn)), "Low-High") + Application.WorksheetFunction.CountIf(Sheets("Copy").Range(Cells(1, ycolumn), Cells(NCells, ycolumn)), "Low-Very High")
        Sheets(shArray(i)).Cells(7, NYears * 3 + 6) = Application.WorksheetFunction.CountIf(Sheets("Copy").Range(Cells(1, ycolumn), Cells(NCells, ycolumn)), "Very High-Ideal") + Application.WorksheetFunction.CountIf(Sheets("Copy").Range(Cells(1, ycolumn), Cells(NCells, ycolumn)), "Very High-Low") + Application.WorksheetFunction.CountIf(Sheets("Copy").Range(Cells(1, ycolumn), Cells(NCells, ycolumn)), "High-Ideal") + Application.WorksheetFunction.CountIf(Sheets("Copy").Range(Cells(1, ycolumn), Cells(NCells, ycolumn)), "High-Low")
        Sheets(shArray(i)).Cells(8, NYears * 3 + 6) = Application.WorksheetFunction.CountIf(Sheets("Copy").Range(Cells(1, ycolumn), Cells(NCells, ycolumn)), "Very High-Moderate") + Application.WorksheetFunction.CountIf(Sheets("Copy").Range(Cells(1, ycolumn), Cells(NCells, ycolumn)), "High-Moderate")
        Sheets(shArray(i)).Cells(9, NYears * 3 + 6) = Application.WorksheetFunction.CountIf(Sheets("Copy").Range(Cells(1, ycolumn), Cells(NCells, ycolumn)), "Moderate-Ideal") + Application.WorksheetFunction.CountIf(Sheets("Copy").Range(Cells(1, ycolumn), Cells(NCells, ycolumn)), "Moderate-Low")
        Sheets(shArray(i)).Cells(10, NYears * 3 + 6) = Application.WorksheetFunction.CountIf(Sheets("Copy").Range(Cells(1, ycolumn), Cells(NCells, ycolumn)), "Moderate-High") + Application.WorksheetFunction.CountIf(Sheets("Copy").Range(Cells(1, ycolumn), Cells(NCells, ycolumn)), "Moderate-Very High")
    
        GoTo HSorGGTorNIC
    
    End If
    
    If shArray(i) = "GGT" Then
    
        Sheets(shArray(i)).Cells(2, NYears * 3 + 5) = "Remained Low"
        Sheets(shArray(i)).Cells(3, NYears * 3 + 5) = "Remained High"
        Sheets(shArray(i)).Cells(4, NYears * 3 + 5) = "Low to High"
        Sheets(shArray(i)).Cells(5, NYears * 3 + 5) = "High to Low"
        
        Sheets(shArray(i)).Cells(2, NYears * 3 + 6) = Application.WorksheetFunction.CountIf(Sheets("Copy").Range(Cells(1, ycolumn), Cells(NCells, ycolumn)), "Low-Low")
        Sheets(shArray(i)).Cells(3, NYears * 3 + 6) = Application.WorksheetFunction.CountIf(Sheets("Copy").Range(Cells(1, ycolumn), Cells(NCells, ycolumn)), "High-High")
        Sheets(shArray(i)).Cells(4, NYears * 3 + 6) = Application.WorksheetFunction.CountIf(Sheets("Copy").Range(Cells(1, ycolumn), Cells(NCells, ycolumn)), "Low-High")
        Sheets(shArray(i)).Cells(5, NYears * 3 + 6) = Application.WorksheetFunction.CountIf(Sheets("Copy").Range(Cells(1, ycolumn), Cells(NCells, ycolumn)), "High-Low")
        
        GoTo HSorGGTorNIC
    
    End If
    
    If shArray(i) = "Nicotine" Then
    
        Sheets(shArray(i)).Cells(2, NYears * 3 + 5) = "Remained NEG"
        Sheets(shArray(i)).Cells(3, NYears * 3 + 5) = "Remained POS"
        Sheets(shArray(i)).Cells(4, NYears * 3 + 5) = "NEG to POS"
        Sheets(shArray(i)).Cells(5, NYears * 3 + 5) = "POS to NEG"
        
        Sheets(shArray(i)).Cells(2, NYears * 3 + 6) = Application.WorksheetFunction.CountIf(Sheets("Copy").Range(Cells(1, ycolumn), Cells(NCells, ycolumn)), "NEG-NEG")
        Sheets(shArray(i)).Cells(3, NYears * 3 + 6) = Application.WorksheetFunction.CountIf(Sheets("Copy").Range(Cells(1, ycolumn), Cells(NCells, ycolumn)), "POS-POS")
        Sheets(shArray(i)).Cells(4, NYears * 3 + 6) = Application.WorksheetFunction.CountIf(Sheets("Copy").Range(Cells(1, ycolumn), Cells(NCells, ycolumn)), "NEG-POS")
        Sheets(shArray(i)).Cells(5, NYears * 3 + 6) = Application.WorksheetFunction.CountIf(Sheets("Copy").Range(Cells(1, ycolumn), Cells(NCells, ycolumn)), "POS-NEG")
        
        GoTo HSorGGTorNIC
    
    End If

    Sheets(shArray(i)).Cells(2, NYears * 3 + 5) = "Remained Low"
    Sheets(shArray(i)).Cells(3, NYears * 3 + 5) = "Low to Moderate"
    Sheets(shArray(i)).Cells(4, NYears * 3 + 5) = "Low to High"
    Sheets(shArray(i)).Cells(5, NYears * 3 + 5) = "Remained Moderate"
    Sheets(shArray(i)).Cells(6, NYears * 3 + 5) = "Moderate to Low"
    Sheets(shArray(i)).Cells(7, NYears * 3 + 5) = "Moderate to High"
    Sheets(shArray(i)).Cells(8, NYears * 3 + 5) = "Remained High"
    Sheets(shArray(i)).Cells(9, NYears * 3 + 5) = "High to Low"
    Sheets(shArray(i)).Cells(10, NYears * 3 + 5) = "High to Moderate"
    
    Sheets(shArray(i)).Cells(2, NYears * 3 + 6) = Application.WorksheetFunction.CountIf(Sheets("Copy").Range(Cells(1, ycolumn), Cells(NCells, ycolumn)), "Low-Low")
    Sheets(shArray(i)).Cells(3, NYears * 3 + 6) = Application.WorksheetFunction.CountIf(Sheets("Copy").Range(Cells(1, ycolumn), Cells(NCells, ycolumn)), "Low-Moderate")
    Sheets(shArray(i)).Cells(4, NYears * 3 + 6) = Application.WorksheetFunction.CountIf(Sheets("Copy").Range(Cells(1, ycolumn), Cells(NCells, ycolumn)), "Low-High")
    Sheets(shArray(i)).Cells(5, NYears * 3 + 6) = Application.WorksheetFunction.CountIf(Sheets("Copy").Range(Cells(1, ycolumn), Cells(NCells, ycolumn)), "Moderate-Moderate")
    Sheets(shArray(i)).Cells(6, NYears * 3 + 6) = Application.WorksheetFunction.CountIf(Sheets("Copy").Range(Cells(1, ycolumn), Cells(NCells, ycolumn)), "Moderate-Low")
    Sheets(shArray(i)).Cells(7, NYears * 3 + 6) = Application.WorksheetFunction.CountIf(Sheets("Copy").Range(Cells(1, ycolumn), Cells(NCells, ycolumn)), "Moderate-High")
    Sheets(shArray(i)).Cells(8, NYears * 3 + 6) = Application.WorksheetFunction.CountIf(Sheets("Copy").Range(Cells(1, ycolumn), Cells(NCells, ycolumn)), "High-High")
    Sheets(shArray(i)).Cells(9, NYears * 3 + 6) = Application.WorksheetFunction.CountIf(Sheets("Copy").Range(Cells(1, ycolumn), Cells(NCells, ycolumn)), "High-Low")
    Sheets(shArray(i)).Cells(10, NYears * 3 + 6) = Application.WorksheetFunction.CountIf(Sheets("Copy").Range(Cells(1, ycolumn), Cells(NCells, ycolumn)), "High-Moderate")
     
HSorGGTorNIC:

    ycolumn = ycolumn + 1
    
Next i

Call BarChartCreation

End Sub

Sub BarChartCreation()

Dim shArray() As Variant
Dim cht As ChartObject, CurrentSheetName As String, ChartCount As Integer, CurrentSheet As Worksheet
Dim NYears As Integer

shArray = Array("TotalCholesterol", "HDL", "LDL", "TRIG", "SBP", "DBP", "BP", "Glucose", "A1C", "BMI", "Waist", "WHR", "PSA")

i = 0
NYears = Worksheets("Copy").Range("CT2").Value

For i = LBound(shArray, 1) To UBound(shArray, 1)

    Worksheets(shArray(i)).Activate
    
    Set CurrentSheet = Worksheets(shArray(i))

    CurrentSheetName = shArray(i)
    
    ChartCount = ActiveSheet.ChartObjects.Count

    If ChartCount > 0 Then
        ActiveSheet.ChartObjects.Delete
    End If
    
    Set co = Sheets(shArray(i)).ChartObjects.Add(1, 1, 1, 1)
    ActiveSheet.ChartObjects(1).Activate
    co.Chart.SetSourceData Source:=CurrentSheet.Range(CurrentSheet.Cells(1, NYears * 2 + 3), CurrentSheet.Cells(4, NYears * 3 + 3))
    co.Chart.ChartType = xlColumnClustered
    co.Chart.HasTitle = True
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = CurrentSheetName
    If shArray(i) = "TotalCholesterol" Then
        ActiveChart.ChartTitle.Text = "Total Cholesterol"
    End If
    Set cht = co.Chart.Parent
    ActiveChart.ApplyDataLabels
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.ShowValue = True
    With cht
    .Left = CurrentSheet.Cells(11, NYears * 2 + 3).Left
    .Top = CurrentSheet.Cells(11, NYears * 2 + 3).Top
    .Height = CurrentSheet.Range(Cells(11, NYears * 2 + 3), Cells(27, NYears * 4 + 3)).Height
    .Width = CurrentSheet.Range(Cells(11, NYears * 2 + 3), Cells(27, NYears * 4 + 3)).Width
    End With

Next i

'HealthScore, Nicotine, and GGT charts

Worksheets("Nicotine").Activate
    
Set CurrentSheet = Worksheets("Nicotine")

CurrentSheetName = "Nicotine"

ChartCount = ActiveSheet.ChartObjects.Count

If ChartCount > 0 Then
    ActiveSheet.ChartObjects.Delete
End If

Set co = Sheets("Nicotine").ChartObjects.Add(1, 1, 1, 1)
ActiveSheet.ChartObjects(1).Activate
co.Chart.SetSourceData Source:=CurrentSheet.Range(CurrentSheet.Cells(1, NYears * 2 + 3), CurrentSheet.Cells(3, NYears * 3 + 3))
co.Chart.ChartType = xlColumnClustered
co.Chart.HasTitle = True
ActiveChart.ChartTitle.Select
ActiveChart.ChartTitle.Text = CurrentSheetName
Set cht = co.Chart.Parent
ActiveChart.ApplyDataLabels
ActiveChart.FullSeriesCollection(1).DataLabels.Select
Selection.ShowValue = True
With cht
.Left = CurrentSheet.Cells(11, NYears * 2 + 3).Left
.Top = CurrentSheet.Cells(11, NYears * 2 + 3).Top
.Height = CurrentSheet.Range(Cells(11, NYears * 2 + 3), Cells(27, NYears * 4 + 3)).Height
.Width = CurrentSheet.Range(Cells(11, NYears * 2 + 3), Cells(27, NYears * 4 + 3)).Width
End With

Worksheets("GGT").Activate
    
Set CurrentSheet = Worksheets("GGT")

CurrentSheetName = "GGT"

ChartCount = ActiveSheet.ChartObjects.Count

If ChartCount > 0 Then
    ActiveSheet.ChartObjects.Delete
End If

Set co = Sheets("GGT").ChartObjects.Add(1, 1, 1, 1)
ActiveSheet.ChartObjects(1).Activate
co.Chart.SetSourceData Source:=CurrentSheet.Range(CurrentSheet.Cells(1, NYears * 2 + 3), CurrentSheet.Cells(3, NYears * 3 + 3))
co.Chart.ChartType = xlColumnClustered
co.Chart.HasTitle = True
ActiveChart.ChartTitle.Select
ActiveChart.ChartTitle.Text = CurrentSheetName
Set cht = co.Chart.Parent
ActiveChart.ApplyDataLabels
ActiveChart.FullSeriesCollection(1).DataLabels.Select
Selection.ShowValue = True
With cht
.Left = CurrentSheet.Cells(11, NYears * 2 + 3).Left
.Top = CurrentSheet.Cells(11, NYears * 2 + 3).Top
.Height = CurrentSheet.Range(Cells(11, NYears * 2 + 3), Cells(27, NYears * 4 + 3)).Height
.Width = CurrentSheet.Range(Cells(11, NYears * 2 + 3), Cells(27, NYears * 4 + 3)).Width
End With

Worksheets("HealthScore").Activate
    
Set CurrentSheet = Worksheets("HealthScore")

CurrentSheetName = "HealthScore"

ChartCount = ActiveSheet.ChartObjects.Count

If ChartCount > 0 Then
    ActiveSheet.ChartObjects.Delete
End If

Set co = Sheets("HealthScore").ChartObjects.Add(1, 1, 1, 1)
ActiveSheet.ChartObjects(1).Activate
co.Chart.SetSourceData Source:=CurrentSheet.Range(CurrentSheet.Cells(1, NYears * 2 + 3), CurrentSheet.Cells(6, NYears * 3 + 3))
co.Chart.ChartType = xlColumnClustered
co.Chart.HasTitle = True
ActiveChart.ChartTitle.Select
ActiveChart.ChartTitle.Text = "Health Score"
Set cht = co.Chart.Parent
ActiveChart.ApplyDataLabels
ActiveChart.FullSeriesCollection(1).DataLabels.Select
Selection.ShowValue = True
With cht
.Left = CurrentSheet.Cells(11, NYears * 2 + 3).Left
.Top = CurrentSheet.Cells(11, NYears * 2 + 3).Top
.Height = CurrentSheet.Range(Cells(11, NYears * 2 + 3), Cells(29, NYears * 4 + 3)).Height
.Width = CurrentSheet.Range(Cells(11, NYears * 2 + 3), Cells(29, NYears * 4 + 3)).Width
End With

Call PieChartCreation

End Sub

Sub PieChartCreation()

Dim shArray() As Variant
Dim cht As ChartObject, CurrentSheetName As String, CurrentYear As String, ChartCount As Integer, CurrentSheet As Worksheet, riskrange As Range, datarange As Range, multirange As Range
Dim NYears As Integer

shArray = Array("TotalCholesterol", "HDL", "LDL", "TRIG", "SBP", "DBP", "BP", "Glucose", "A1C", "BMI", "Waist", "WHR", "PSA")

i = 0
NYears = Worksheets("Copy").Range("CT2").Value

For i = LBound(shArray, 1) To UBound(shArray, 1)

    Worksheets(shArray(i)).Activate
    
    Set CurrentSheet = Worksheets(shArray(i))

    CurrentSheetName = shArray(i)
    
    CurrentYear = CurrentSheet.Cells(1, NYears * 3 + 3).Value
    
    ChartCount = ActiveSheet.ChartObjects.Count
    
    Set riskrange = Range(CurrentSheet.Cells(1, NYears * 2 + 3), CurrentSheet.Cells(4, NYears * 2 + 3))
    Set datarange = Range(CurrentSheet.Cells(1, NYears * 3 + 3), CurrentSheet.Cells(4, NYears * 3 + 3))
    Set multirange = Union(riskrange, datarange)
    
    Set co = Sheets(shArray(i)).ChartObjects.Add(1, 1, 1, 1)
    ActiveSheet.ChartObjects(2).Activate
    co.Chart.SetSourceData Source:=multirange
    co.Chart.ChartType = xlPie
    co.Chart.HasTitle = True
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = CurrentYear & " " & CurrentSheetName & " Risk Level Breakdown"
    If shArray(i) = "TotalCholesterol" Then
        ActiveChart.ChartTitle.Text = CurrentYear & " Total Cholesterol Risk Level Breakdown"
    End If
    ActiveChart.SetElement (msoElementDataLabelBestFit)
    Set cht = co.Chart.Parent
    ActiveChart.ApplyDataLabels
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.ShowValue = True
    Selection.ShowPercentage = True
    With cht
    .Left = CurrentSheet.Cells(11, NYears * 2 + 11).Left
    .Top = CurrentSheet.Cells(11, NYears * 2 + 11).Top
    .Height = CurrentSheet.Range(Cells(11, NYears * 2 + 11), Cells(27, NYears * 4 + 11)).Height
    .Width = CurrentSheet.Range(Cells(11, NYears * 2 + 11), Cells(27, NYears * 4 + 11)).Width
    End With

Next i

'HealthScore, GGT, and Nicotine piecharts

Worksheets("HealthScore").Activate
    
    Set CurrentSheet = Worksheets("HealthScore")
    
    CurrentYear = CurrentSheet.Cells(1, NYears * 3 + 3).Value
    
    ChartCount = ActiveSheet.ChartObjects.Count
    
    Set riskrange = Range(CurrentSheet.Cells(1, NYears * 2 + 3), CurrentSheet.Cells(6, NYears * 2 + 3))
    Set datarange = Range(CurrentSheet.Cells(1, NYears * 3 + 3), CurrentSheet.Cells(6, NYears * 3 + 3))
    Set multirange = Union(riskrange, datarange)
    
    Set co = Sheets("HealthScore").ChartObjects.Add(1, 1, 1, 1)
    ActiveSheet.ChartObjects(2).Activate
    co.Chart.SetSourceData Source:=multirange
    co.Chart.ChartType = xlPie
    co.Chart.HasTitle = True
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = CurrentYear & " Health Score Risk Level Breakdown"
    ActiveChart.SetElement (msoElementDataLabelBestFit)
    Set cht = co.Chart.Parent
    ActiveChart.ApplyDataLabels
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.ShowValue = True
    Selection.ShowPercentage = True
    With cht
    .Left = CurrentSheet.Cells(11, NYears * 2 + 11).Left
    .Top = CurrentSheet.Cells(11, NYears * 2 + 11).Top
    .Height = CurrentSheet.Range(Cells(11, NYears * 2 + 11), Cells(27, NYears * 4 + 11)).Height
    .Width = CurrentSheet.Range(Cells(11, NYears * 2 + 11), Cells(27, NYears * 4 + 11)).Width
    End With
   
Worksheets("GGT").Activate
    
    Set CurrentSheet = Worksheets("GGT")
    
    CurrentYear = CurrentSheet.Cells(1, NYears * 3 + 3).Value
    
    ChartCount = ActiveSheet.ChartObjects.Count
    
    Set riskrange = Range(CurrentSheet.Cells(1, NYears * 2 + 3), CurrentSheet.Cells(3, NYears * 2 + 3))
    Set datarange = Range(CurrentSheet.Cells(1, NYears * 3 + 3), CurrentSheet.Cells(3, NYears * 3 + 3))
    Set multirange = Union(riskrange, datarange)
    
    Set co = Sheets("GGT").ChartObjects.Add(1, 1, 1, 1)
    ActiveSheet.ChartObjects(2).Activate
    co.Chart.SetSourceData Source:=multirange
    co.Chart.ChartType = xlPie
    co.Chart.HasTitle = True
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = CurrentYear & " GGT Risk Level Breakdown"
    ActiveChart.SetElement (msoElementDataLabelBestFit)
    Set cht = co.Chart.Parent
    ActiveChart.ApplyDataLabels
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.ShowValue = True
    Selection.ShowPercentage = True
    With cht
    .Left = CurrentSheet.Cells(11, NYears * 2 + 11).Left
    .Top = CurrentSheet.Cells(11, NYears * 2 + 11).Top
    .Height = CurrentSheet.Range(Cells(11, NYears * 2 + 11), Cells(27, NYears * 4 + 11)).Height
    .Width = CurrentSheet.Range(Cells(11, NYears * 2 + 11), Cells(27, NYears * 4 + 11)).Width
    End With
    
Worksheets("Nicotine").Activate
    
    Set CurrentSheet = Worksheets("Nicotine")
    
    CurrentYear = CurrentSheet.Cells(1, NYears * 3 + 3).Value
    
    ChartCount = ActiveSheet.ChartObjects.Count
    
    Set riskrange = Range(CurrentSheet.Cells(1, NYears * 2 + 3), CurrentSheet.Cells(3, NYears * 2 + 3))
    Set datarange = Range(CurrentSheet.Cells(1, NYears * 3 + 3), CurrentSheet.Cells(3, NYears * 3 + 3))
    Set multirange = Union(riskrange, datarange)
    
    Set co = Sheets("Nicotine").ChartObjects.Add(1, 1, 1, 1)
    ActiveSheet.ChartObjects(2).Activate
    co.Chart.SetSourceData Source:=multirange
    co.Chart.ChartType = xlPie
    co.Chart.HasTitle = True
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = CurrentYear & " Nicotine Risk Level Breakdown"
    ActiveChart.SetElement (msoElementDataLabelBestFit)
    Set cht = co.Chart.Parent
    ActiveChart.ApplyDataLabels
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.ShowValue = True
    Selection.ShowPercentage = True
    With cht
    .Left = CurrentSheet.Cells(11, NYears * 2 + 11).Left
    .Top = CurrentSheet.Cells(11, NYears * 2 + 11).Top
    .Height = CurrentSheet.Range(Cells(11, NYears * 2 + 11), Cells(27, NYears * 4 + 11)).Height
    .Width = CurrentSheet.Range(Cells(11, NYears * 2 + 11), Cells(27, NYears * 4 + 11)).Width
    End With

Call ChangeNames

End Sub

Sub ChangeNames()
Dim shArray() As Variant, i As Integer, xrowcopy As Integer, xrowP As Integer, xrowN As Integer, xrowNC As Integer, NCells As Integer, ycolumn As Integer

shArray = Array("HealthScore", "TotalCholesterol", "HDL", "LDL", "TRIG", "SBP", "DBP", "Glucose", "A1C", "BMI", "Waist", "WHR", "Nicotine", "GGT", "PSA")

i = 0

xrowcopy = 2
xrowP = 2
xrowN = 2
xrowNC = 2
ycolumn = 121
Sheets("Copy").Activate

For i = LBound(shArray, 1) To UBound(shArray, 1)

    NCells = Application.WorksheetFunction.CountA(Sheets("Copy").Columns(ycolumn))
    
    Sheets(shArray(i)).Cells(1, 21) = "Positive Change"
    Sheets(shArray(i)).Cells(1, 22) = "Negative Change"
    Sheets(shArray(i)).Cells(1, 23) = "No Change"
    
    Do While xrowcopy <= NCells
    
        If Sheets("Copy").Cells(xrowcopy, ycolumn) = "Positive Change" Then
            Sheets(shArray(i)).Cells(xrowP, 21) = Sheets("Copy").Cells(xrowcopy, ycolumn + 16)
            xrowP = xrowP + 1
        End If
        If Sheets("Copy").Cells(xrowcopy, ycolumn) = "Negative Change" Then
            Sheets(shArray(i)).Cells(xrowN, 22) = Sheets("Copy").Cells(xrowcopy, ycolumn + 16)
            xrowN = xrowN + 1
        End If
        If Sheets("Copy").Cells(xrowcopy, ycolumn) = "No Change" Then
            Sheets(shArray(i)).Cells(xrowNC, 23) = Sheets("Copy").Cells(xrowcopy, ycolumn + 16)
            xrowNC = xrowNC + 1
        End If
        
        xrowcopy = xrowcopy + 1
        
    Loop
    
    xrowcopy = 2
    xrowP = 2
    xrowN = 2
    xrowNC = 2
    
    ycolumn = ycolumn + 1

Next i

Call Averages

End Sub

Sub Averages()

Dim shArray() As Variant, i As Integer, xrow As Integer
Dim NYears As Integer

shArray = Array("HealthScore", "TotalCholesterol", "HDL", "LDL", "TRIG", "SBP", "DBP", "Glucose", "A1C", "BMI", "Waist", "WHR", "Nicotine", "GGT", "PSA")

i = 0
xrow = 2
NYears = Worksheets("Copy").Range("CT2").Value

Worksheets(shArray(i)).Activate
Range(Cells(1, 1), Cells(1, NYears)).Select
Selection.Copy
Worksheets("Averages").Activate
Worksheets("Averages").Range(Cells(1, 2), Cells(1, NYears + 1)).PasteSpecial

For i = LBound(shArray, 1) To UBound(shArray, 1)

    If shArray(i) = "HealthScore" Then
        Worksheets(shArray(i)).Activate
        Range(Cells(7, NYears * 2 + 3), Cells(7, NYears * 3 + 3)).Select
        Selection.Copy
        GoTo SpecialAverages:
    End If
    
    If shArray(i) = "GGT" Then
        Worksheets(shArray(i)).Activate
        Range(Cells(4, NYears * 2 + 3), Cells(4, NYears * 3 + 3)).Select
        Selection.Copy
        GoTo SpecialAverages:
    End If
    
    If shArray(i) = "Nicotine" Then
        Worksheets(shArray(i)).Activate
        Range(Cells(2, NYears * 2 + 3), Cells(2, NYears * 3 + 3)).Select
        Selection.Copy
        GoTo SpecialAverages:
    End If
    
    Worksheets(shArray(i)).Activate
    
    Range(Cells(5, NYears * 2 + 3), Cells(5, NYears * 3 + 3)).Select
    Selection.Copy
    
SpecialAverages:

    Worksheets("Averages").Activate
    Worksheets("Averages").Range(Cells(xrow, 1), Cells(xrow, NYears)).PasteSpecial
    xrow = xrow + 1

Next i

i = 0
xrow = 2
Worksheets("Averages").Activate

For i = LBound(shArray, 1) To UBound(shArray, 1)

    Worksheets("Averages").Cells(xrow, 1) = shArray(i)
    xrow = xrow + 1

Next i

Call Age

End Sub

Sub Age()

Dim xrow As Integer, NCells As Integer

Sheets("Copy").Activate

NCells = Range("CA2")
xrow = 2
Range("EF1") = "Age"

Do While xrow <= NCells

Cells(xrow, 136) = "=ROUNDDOWN(YEARFRAC(E" & xrow & ", B" & xrow & ", 1), 0)"

xrow = xrow + 1

Loop

'''Prep data for age sheet
Dim yrArray() As Variant, k As Integer, j As Integer, i As Double, ycolumn As Integer, NYears As Integer

NYears = Worksheets("Copy").Range("CT2").Value

xrow = 2
i = 0

ReDim yrArray(0)

Do Until Sheets("Copy").Cells(xrow, "CU").Value = ""

    yrArray(i) = Sheets("Copy").Cells(xrow, "CU").Value
    i = i + 1
    ReDim Preserve yrArray(i)
    xrow = xrow + 1

Loop

i = 0
j = 0
k = 0
ycolumn = 1

For i = LBound(yrArray, 1) To UBound(yrArray, 1)
    
    Sheets("Age").Cells(1, ycolumn) = yrArray(i)
    ycolumn = ycolumn + 1

Next i

i = 0
j = 0
ycolumn = 1


'''Populate sheet w/ data

Dim CurrentYear As Integer, MaxYear As Integer
Dim DataArray() As Variant, CopyColumnB As Integer, PasteColumnB As Integer
CurrentYear = Sheets("Copy").Range("CV2").Value
MaxYear = Sheets("Copy").Range("CW2").Value
CopyColumnB = 136
PasteColumnB = 1

xrow = 2
i = 0
j = 0
k = 0

Do While CurrentYear <= MaxYear
    
    ReDim DataArray(0)
    
'Fill Data Array
    
    Do While xrow <= NCells
             
        If Sheets("Copy").Cells(xrow, "CS").Value = CurrentYear Then
            DataArray(k) = Sheets("Copy").Cells(xrow, CopyColumnB).Value
            k = k + 1
            ReDim Preserve DataArray(k)
        End If
        
        xrow = xrow + 1
    
    Loop
    
    xrow = 2
    i = 0
    
'Paste Data Array

    Sheets("Age").Activate
    
    For k = LBound(DataArray, 1) To UBound(DataArray, 1)
        Sheets("Age").Cells(xrow, PasteColumnB) = DataArray(k)
        xrow = xrow + 1
    Next k
            
    xrow = 2
    i = 0
    k = 0

    CurrentYear = CurrentYear + 1
    
    PasteColumnB = PasteColumnB + 1

Loop

'Display averages

CurrentYear = Sheets("Copy").Range("CV2").Value
ycolumn = 0

Do While CurrentYear <= MaxYear

    Sheets("Age").Cells(1, NYears + 2 + ycolumn) = CurrentYear & " Average"
    Sheets("Age").Cells(2, NYears + 2 + ycolumn) = Application.WorksheetFunction.Average(Sheets("Age").Range(Cells(2, ycolumn + 1), Cells(NCells, ycolumn + 1)))
    ycolumn = ycolumn + 1
    CurrentYear = CurrentYear + 1
    
Loop

Sheets("Age").Range(Cells(2, NYears + 2), Cells(2, NYears + 2 + NYears)).Copy
Sheets("Averages").Activate
Sheets("Averages").Cells(17, 1) = "Age"
Sheets("Averages").Cells(17, 2).PasteSpecial

Call WaistGenders

End Sub

Sub WaistGenders()

Dim xrow As Integer, NCells As Integer

Sheets("Copy").Activate

NCells = Range("CA2")
xrow = 2

'''Prep data for waist sheet
Dim yrArray() As Variant, k As Integer, j As Integer, i As Double, ycolumn As Integer, NYears As Integer

NYears = Worksheets("Copy").Range("CT2").Value

xrow = 2
i = 0

ReDim yrArray(0)

Do Until Sheets("Copy").Cells(xrow, "CU").Value = ""

    yrArray(i) = Sheets("Copy").Cells(xrow, "CU").Value
    i = i + 1
    ReDim Preserve yrArray(i)
    xrow = xrow + 1

Loop

i = 0
j = 0
k = 0
ycolumn = 1

For i = LBound(yrArray, 1) To UBound(yrArray, 1)
    
    Sheets("WaistbyGender").Cells(1, ycolumn) = yrArray(i)
    ycolumn = ycolumn + 1

Next i

For i = LBound(yrArray, 1) To UBound(yrArray, 1)
    
    Sheets("WaistbyGender").Cells(1, ycolumn) = yrArray(i)
    ycolumn = ycolumn + 1

Next i

i = 0
j = 0
ycolumn = 1

'''Populate sheet w/ data

Dim CurrentYear As Integer, MaxYear As Integer
Dim DataArray() As Variant, CopyColumnB As Integer, PasteColumnB As Integer, LLoop As Integer
CurrentYear = Sheets("Copy").Range("CV2").Value
MaxYear = Sheets("Copy").Range("CW2").Value
CopyColumnB = 7
PasteColumnB = 1

xrow = 2
i = 0
j = 0
k = 0
LLoop = 0

Do While LLoop <= 1

    Do While CurrentYear <= MaxYear
        
        ReDim DataArray(0)
        
    'Fill Data Array
        
        Do While xrow <= NCells
                 
            If Sheets("Copy").Cells(xrow, "CS").Value = CurrentYear Then
                DataArray(k) = Sheets("Copy").Cells(xrow, CopyColumnB).Value
                k = k + 1
                ReDim Preserve DataArray(k)
            End If
            
            xrow = xrow + 1
        
        Loop
        
        xrow = 2
        i = 0
        
    'Paste Data Array
    
        Sheets("WaistbyGender").Activate
        
        For k = LBound(DataArray, 1) To UBound(DataArray, 1)
            Sheets("WaistbyGender").Cells(xrow, PasteColumnB) = DataArray(k)
            xrow = xrow + 1
        Next k
                
        xrow = 2
        i = 0
        k = 0
    
        CurrentYear = CurrentYear + 1
        
        PasteColumnB = PasteColumnB + 1
    
    Loop

    LLoop = LLoop + 1
    
    CurrentYear = Sheets("Copy").Range("CV2").Value
    CopyColumnB = 68
    PasteColumnB = PasteColumnB + 1
    
    xrow = 2
    i = 0
    j = 0
    k = 0
    
Loop

Sheets("WaistbyGender").Cells(1, NYears * 2 + 3) = MaxYear & " Males"
Sheets("WaistbyGender").Cells(1, NYears * 2 + 4) = MaxYear & " Females"

xrow = 2

Do While xrow <= NCells

    If Sheets("WaistbyGender").Cells(xrow, NYears) = "Male" Then
        Sheets("WaistbyGender").Cells(xrow, NYears * 2 + 3) = Sheets("WaistbyGender").Cells(xrow, NYears * 2 + 1)
    Else
        Sheets("WaistbyGender").Cells(xrow, NYears * 2 + 4) = Sheets("WaistbyGender").Cells(xrow, NYears * 2 + 1)
    End If
    
    xrow = xrow + 1
    
Loop

Sheets("WaistbyGender").Cells(1, NYears * 2 + 6) = MaxYear & " Male Average"
Sheets("WaistbyGender").Cells(1, NYears * 2 + 7) = MaxYear & " Female Average"
Sheets("WaistbyGender").Cells(2, NYears * 2 + 6) = Application.WorksheetFunction.Average(Sheets("WaistbyGender").Range(Cells(2, NYears * 2 + 3), Cells(NCells, NYears * 2 + 3)))
Sheets("WaistbyGender").Cells(2, NYears * 2 + 7) = Application.WorksheetFunction.Average(Sheets("WaistbyGender").Range(Cells(2, NYears * 2 + 3), Cells(NCells, NYears * 2 + 4)))
Sheets("WaistbyGender").Range(Cells(2, NYears * 2 + 6), Cells(2, NYears * 2 + 7)).Select
Selection.NumberFormat = "0.00"

'Make Graph

Dim cht As ChartObject, CurrentSheetName As String, CurrentSheet As Worksheet, MaxYrStr As String
    
Set CurrentSheet = Worksheets("WaistbyGender")

MaxYrStr = Sheets("WaistbyGender").Cells(1, NYears)

CurrentSheetName = MaxYrStr & " Average Waist by Gender"

Set co = Sheets("WaistbyGender").ChartObjects.Add(1, 1, 1, 1)
ActiveSheet.ChartObjects(1).Activate
co.Chart.SetSourceData Source:=CurrentSheet.Range(CurrentSheet.Cells(1, NYears * 2 + 6), CurrentSheet.Cells(2, NYears * 2 + 7))
co.Chart.ChartType = xlColumnClustered
co.Chart.HasTitle = True
ActiveChart.ChartTitle.Select
ActiveChart.ChartTitle.Text = CurrentSheetName
ActiveChart.Axes(xlValue).MinimumScale = 20
ActiveChart.Axes(xlValue).MaximumScale = 50
Set cht = co.Chart.Parent
ActiveChart.ApplyDataLabels
ActiveChart.FullSeriesCollection(1).DataLabels.Select
Selection.ShowValue = True
With cht
.Left = CurrentSheet.Cells(7, NYears * 3 + 7).Left
.Top = CurrentSheet.Cells(7, NYears * 3 + 7).Top
.Height = CurrentSheet.Range(Cells(7, NYears * 3 + 7), Cells(23, NYears * 5 + 7)).Height
.Width = CurrentSheet.Range(Cells(7, NYears * 3 + 7), Cells(23, NYears * 5 + 7)).Width
End With
ActiveChart.Legend.Select
Selection.Delete
ActiveChart.FullSeriesCollection(1).Points(2).Select
With Selection.Format.Fill
    .Visible = msoTrue
    .ForeColor.ObjectThemeColor = msoThemeColorAccent2
    .ForeColor.TintAndShade = 0
    .ForeColor.Brightness = 0
    .Transparency = 0
    .Solid
End With

Call MetabolicSyndromes

End Sub

Sub MetabolicSyndromes()

Sheets("Copy").Activate

Dim NCells As Integer, xrow As Integer

NCells = Range("CA2")
xrow = 2

Cells(1, 153) = "# of Metabolic Risks"

'Count Risks

Do While xrow <= NCells

    If Cells(xrow, 116) = "" Or Cells(xrow, 107) = "" Or Cells(xrow, 109) = "" Or Cells(xrow, 110) = "" Or Cells(xrow, 111) = "" Or Cells(xrow, 113) = "" Then
        GoTo IncompleteData
    End If
    
    Cells(xrow, 153) = 0
    
    If Cells(xrow, 7) = "Male" Then
        
        'Waist Risk
        If Cells(xrow, 116) > 40 Then
            Cells(xrow, 153) = Cells(xrow, 153) + 1
        End If
        
        'HDL Risk
        If Cells(xrow, 107) < 40 Then
            Cells(xrow, 153) = Cells(xrow, 153) + 1
        End If
        
    End If
    
    If Cells(xrow, 7) = "Female" Then
    
        'Waist Risk
        If Cells(xrow, 116) > 35 Then
            Cells(xrow, 153) = Cells(xrow, 153) + 1
        End If
        
        'HDL Risk
        If Cells(xrow, 107) < 50 Then
            Cells(xrow, 153) = Cells(xrow, 153) + 1
        End If
    
    End If
    
    'Trig Risk
    If Cells(xrow, 109) >= 150 Then
        Cells(xrow, 153) = Cells(xrow, 153) + 1
    End If
    
    'BP Risk
    If Cells(xrow, 110) / Cells(xrow, 111) >= 130 / 85 Then
        Cells(xrow, 153) = Cells(xrow, 153) + 1
    End If
    
    'Gluc Risk
    If Cells(xrow, 113) >= 100 Then
        Cells(xrow, 153) = Cells(xrow, 153) + 1
    End If

IncompleteData:
    
    xrow = xrow + 1

Loop

'Create MetabolicSyndromes Sheet

Dim yrArray() As Variant, k As Integer, j As Integer, i As Double, ycolumn As Integer, NYears As Integer

NYears = Worksheets("Copy").Range("CT2").Value

xrow = 2
i = 0

ReDim yrArray(0)

Do Until Sheets("Copy").Cells(xrow, "CU").Value = ""

    yrArray(i) = Sheets("Copy").Cells(xrow, "CU").Value
    i = i + 1
    ReDim Preserve yrArray(i)
    xrow = xrow + 1

Loop

i = 0
j = 0
k = 0
ycolumn = 1

For i = LBound(yrArray, 1) To UBound(yrArray, 1)
    
    Sheets("MetabolicSyndromes").Cells(1, ycolumn) = yrArray(i)
    ycolumn = ycolumn + 1

Next i

i = 0
j = 0
ycolumn = 1


'''Populate sheet w/ data

Dim CurrentYear As Integer, MaxYear As Integer
Dim DataArray() As Variant, CopyColumnB As Integer, PasteColumnB As Integer
CurrentYear = Sheets("Copy").Range("CV2").Value
MaxYear = Sheets("Copy").Range("CW2").Value
CopyColumnB = 153
PasteColumnB = 1

xrow = 2
i = 0
j = 0
k = 0

Do While CurrentYear <= MaxYear
    
    ReDim DataArray(0)
    
'Fill Data Array
    
    Do While xrow <= NCells
             
        If Sheets("Copy").Cells(xrow, "CS").Value = CurrentYear Then
            DataArray(k) = Sheets("Copy").Cells(xrow, CopyColumnB).Value
            k = k + 1
            ReDim Preserve DataArray(k)
        End If
        
        xrow = xrow + 1
    
    Loop
    
    xrow = 2
    i = 0
    
'Paste Data Array

    Sheets("MetabolicSyndromes").Activate
    
    For k = LBound(DataArray, 1) To UBound(DataArray, 1)
        Sheets("MetabolicSyndromes").Cells(xrow, PasteColumnB) = DataArray(k)
        xrow = xrow + 1
    Next k
            
    xrow = 2
    i = 0
    k = 0

    CurrentYear = CurrentYear + 1
    
    PasteColumnB = PasteColumnB + 1

Loop

'Display averages

CurrentYear = Sheets("Copy").Range("CV2").Value
ycolumn = 0

Do While CurrentYear <= MaxYear

    Sheets("MetabolicSyndromes").Cells(1, NYears + 2 + ycolumn) = CurrentYear & " Average"
    Sheets("MetabolicSyndromes").Cells(2, NYears + 2 + ycolumn) = Application.WorksheetFunction.Average(Sheets("MetabolicSyndromes").Range(Cells(2, ycolumn + 1), Cells(NCells, ycolumn + 1)))
    ycolumn = ycolumn + 1
    CurrentYear = CurrentYear + 1
    
Loop

'Count Number of members with each number of syndromes

CurrentYear = Sheets("Copy").Range("CV2").Value
ycolumn = 0

Sheets("MetabolicSyndromes").Cells(1, NYears * 2 + 3 + ycolumn) = "# of Metabolic Syndromes"
Sheets("MetabolicSyndromes").Cells(2, NYears * 2 + 3 + ycolumn) = 0
Sheets("MetabolicSyndromes").Cells(3, NYears * 2 + 3 + ycolumn) = 1
Sheets("MetabolicSyndromes").Cells(4, NYears * 2 + 3 + ycolumn) = 2
Sheets("MetabolicSyndromes").Cells(5, NYears * 2 + 3 + ycolumn) = 3
Sheets("MetabolicSyndromes").Cells(6, NYears * 2 + 3 + ycolumn) = 4
Sheets("MetabolicSyndromes").Cells(7, NYears * 2 + 3 + ycolumn) = 5
Sheets("MetabolicSyndromes").Cells(8, NYears * 2 + 3 + ycolumn) = "# w/ complete Data"

Do While CurrentYear <= MaxYear
    
    Sheets("MetabolicSyndromes").Cells(1, NYears * 2 + 4 + ycolumn) = CurrentYear & " Counts"
    Sheets("MetabolicSyndromes").Cells(2, NYears * 2 + 4 + ycolumn) = Application.WorksheetFunction.CountIf(Sheets("MetabolicSyndromes").Range(Cells(2, ycolumn + 1), Cells(NCells, ycolumn + 1)), 0)
    Sheets("MetabolicSyndromes").Cells(3, NYears * 2 + 4 + ycolumn) = Application.WorksheetFunction.CountIf(Sheets("MetabolicSyndromes").Range(Cells(2, ycolumn + 1), Cells(NCells, ycolumn + 1)), 1)
    Sheets("MetabolicSyndromes").Cells(4, NYears * 2 + 4 + ycolumn) = Application.WorksheetFunction.CountIf(Sheets("MetabolicSyndromes").Range(Cells(2, ycolumn + 1), Cells(NCells, ycolumn + 1)), 2)
    Sheets("MetabolicSyndromes").Cells(5, NYears * 2 + 4 + ycolumn) = Application.WorksheetFunction.CountIf(Sheets("MetabolicSyndromes").Range(Cells(2, ycolumn + 1), Cells(NCells, ycolumn + 1)), 3)
    Sheets("MetabolicSyndromes").Cells(6, NYears * 2 + 4 + ycolumn) = Application.WorksheetFunction.CountIf(Sheets("MetabolicSyndromes").Range(Cells(2, ycolumn + 1), Cells(NCells, ycolumn + 1)), 4)
    Sheets("MetabolicSyndromes").Cells(7, NYears * 2 + 4 + ycolumn) = Application.WorksheetFunction.CountIf(Sheets("MetabolicSyndromes").Range(Cells(2, ycolumn + 1), Cells(NCells, ycolumn + 1)), 5)
    Sheets("MetabolicSyndromes").Cells(8, NYears * 2 + 4 + ycolumn) = "=COUNTA(C[-7])"
    ycolumn = ycolumn + 1
    CurrentYear = CurrentYear + 1
    
Loop

CurrentYear = Sheets("Copy").Range("CV2").Value


Sheets("MetabolicSyndromes").Range(Cells(2, NYears + 2), Cells(2, NYears + 2 + NYears)).Copy
Sheets("Averages").Activate
Sheets("Averages").Cells(18, 1) = "# of Metabolic Syndromes"
Sheets("Averages").Cells(18, 2).PasteSpecial

MsgBox ("Complete!")

End Sub

