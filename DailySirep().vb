Sub Macro1()
'
' Macro1 Macro
'

'
    Dim tp As Range, hpgr As Range, aglo As Range, myMultiAreaRange As Range
    Set tp = Range("E5:E372,F5:F372,AD5:AD372,AM5:AM372,J5:J372,S5:S372,M5:M372,I5:I372").Select
    Set hpgr = Range("EY5:EY372,EZ5:EZ372,FK5:FK372,EJ5:EJ372,EP5:EP372,ES5:ES372,EV5:EV372,EG5:EG372").Select
    Set aglo = Range("EY5:EY372,DB5:DB372,DE5:DE372,DH5:DH372,DK5:DK372,DN5:DN372,DX5:DX372,CY5:CY372").Select
    Set myMultiAreaRange = Union(tp, hpgr, aglo)
    myMultiAreaRange.Select

    Selection.Copy
    Sheets.Add After:=ActiveSheet
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Rows("2:2").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Selection.Delete Shift:=xlUp
    Columns("A:A").Select
    Selection.NumberFormat = "m/d/yyyy"
    Columns("B:B").EntireColumn.AutoFit
    Columns("C:C").EntireColumn.AutoFit
    Columns("B:B").Select
    Selection.NumberFormat = "0.00"
    Columns("C:C").Select
    Selection.NumberFormat = "0.00"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "TP-MINERALTRITURADO"
    Range("B2").Select
End Sub




Sub Macro4()
'
' Macro4 Macro
'

'
    'Nueva hoja
    Sheets.Add After:=ActiveSheet
    'Fecha
    Range("A1") = "Fecha"
    Range("A2") = "1/1/2022"
    Range("A2").Select
    Selection.AutoFill Destination:=Range("A2:A366"), Type:=xlFillDefault
    '*********************Tritu Primaria*********************
    'Mineral triturado
    Range("B1") = "TP-MineralTriturado"
    Range("B2").Formula = "=+VLookup(A2, Procesos_Real!E8:DP372, 2, False)"
    Range("B2").Select
    Selection.AutoFill Destination:=Range("B2:B366"), Type:=xlFillDefault
    'Ley Au
    Range("C1") = "TP-LeyAu"
    Range("C2").Formula = "=+VLookup(A2, Procesos_Real!E8:DP372, 26, False)"
    Range("C2").Select
    Selection.AutoFill Destination:=Range("C2:C366"), Type:=xlFillDefault
    'Au Triturado
    Range("D1") = "TP-AuTriturado"
    Range("D2") = 0
    Range("D2").Select
    Selection.AutoFill Destination:=Range("D2:D366"), Type:=xlFillDefault
    'Productividad
    Range("E1") = "TP-Productividad"
    Range("E2").Formula = "=+VLookup(A2, Procesos_Real!E8:DP372,6,False)"
    Range("E2").Select
    Selection.AutoFill Destination:=Range("E2:E366"), Type:=xlFillDefault
    'Disponibilidad
    Range("F1") = "TP-Disponibilidad"
    Range("F2").Formula = "=+VLookup(A2, Procesos_Real!E8:DP372,15,False)"
    Range("F2").Select
    Selection.AutoFill Destination:=Range("F2:F366"), Type:=xlFillDefault
    'Utilizacion
    Range("G1") = "TP-Utilizacion"
    Range("G2").Formula = "=+VLookup(A2, Procesos_Real!E8:DP372,9,False)"
    Range("G2").Select
    Selection.AutoFill Destination:=Range("G2:G366"), Type:=xlFillDefault
    'Horas operativas
    Range("H1") = "TP-HorasOperativas"
    Range("H2").Formula = "=+VLookup(A2, Procesos_Real!E8:DP372,5,False)"
    Range("H2").Select
    Selection.AutoFill Destination:=Range("H2:H366"), Type:=xlFillDefault

    '************************* HPRG **********************************
    'Mineral Tritu
    Range("I1") = "HPGR-MineralTriturado"
    Range("I2").Formula = "=+VLookup(A2, Procesos_Real!E8:LV372,151,False)"
    Range("I2").Select
    Selection.AutoFill Destination:=Range("I2:I366"), Type:=xlFillDefault
    'Leu Au
    Range("J1") = "HPGR-LeyAu"
    Range("J2").Formula = "=+VLookup(A2, Procesos_Real!E8:LV372,152,False)"
    Range("J2").Select
    Selection.AutoFill Destination:=Range("J2:J366"), Type:=xlFillDefault
    'Au Triturado
    Range("K1") = "HPGR-AuTriturado"
    Range("K2") = 0
    Range("K2").Select
    Selection.AutoFill Destination:=Range("K2:K366"), Type:=xlFillDefault
    'Productividad
    Range("L1") = "HPGR-Productividad"
    Range("L2").Formula = "=+VLookup(A2, Procesos_Real!E8:LV372,136,False)"
    Range("L2").Select
    Selection.AutoFill Destination:=Range("L2:L366"), Type:=xlFillDefault
    'Disponibilidad
    Range("M1") = "HPGR-Disponibilidad"
    Range("M2").Formula = "=+VLookup(A2, Procesos_Real!E8:LV372,142,False)"
    Range("M2").Select
    Selection.AutoFill Destination:=Range("M2:M366"), Type:=xlFillDefault
    'Utilizacion
    Range("N1") = "HPGR-Utilizacion"
    Range("N2").Formula = "=+VLookup(A2, Procesos_Real!E8:LV372,145,False)"
    Range("N2").Select
    Selection.AutoFill Destination:=Range("N2:N366"), Type:=xlFillDefault
    'P80
    Range("O1") = "HPGR-P80"
    Range("O2").Formula = "=+VLookup(A2, Procesos_Real!E8:LV372,148,False)"
    Range("O2").Select
    Selection.AutoFill Destination:=Range("O2:O366"), Type:=xlFillDefault
    'HorasOperativas
    Range("P1") = "HPGR-HorasOperativas"
    Range("P2").Formula = "=+VLookup(A2, Procesos_Real!E8:LV372,133,False)"
    Range("P2").Select
    Selection.AutoFill Destination:=Range("P2:P366"), Type:=xlFillDefault

    '************************* Aglomeracion **********************************
    'Mineral Tritu
    Range("Q1") = "AGLOM-MineralAglomerado"
    Range("Q2").Formula = "=+VLookup(A2, Procesos_Real!E8:LV372,151,False)"
    Range("Q2").Select
    Selection.AutoFill Destination:=Range("Q2:Q366"), Type:=xlFillDefault
    'Productividad
    Range("R1") = "AGLOM-Productividad"
    Range("R2").Formula = "=+VLookup(A2, Procesos_Real!E8:DP372,102,False)"
    Range("R2").Select
    Selection.AutoFill Destination:=Range("R2:R366"), Type:=xlFillDefault
    'Utilizacion
    Range("S1") = "AGLOM-Utilizacion"
    Range("S2").Formula = "=+VLookup(A2, Procesos_Real!E8:DP372,105,False)"
    Range("S2").Select
    Selection.AutoFill Destination:=Range("S2:S366"), Type:=xlFillDefault
    'Disponibilidad
    Range("T1") = "AGLOM-Disponibilidad"
    Range("T2").Formula = "=+VLookup(A2, Procesos_Real!E8:DP372,108,False)"
    Range("T2").Select
    Selection.AutoFill Destination:=Range("T2:T366"), Type:=xlFillDefault
    'AdicionCemento
    Range("U1") = "AGLOM-AdicionCemento"
    Range("U2").Formula = "=+VLookup(A2, Procesos_Real!E8:DP372,111,False)"
    Range("U2").Select
    Selection.AutoFill Destination:=Range("U2:U366"), Type:=xlFillDefault
    'AdicionCN
    Range("V1") = "AGLOM-AdicionCN"
    Range("V2").Formula = "=+VLookup(A2, Procesos_Real!E8:DP372,114,False)"
    Range("V2").Select
    Selection.AutoFill Destination:=Range("V2:V366"), Type:=xlFillDefault
    'Humedad
    Range("W1") = "AGLOM-Humedad"
    Range("W2").Formula = "=+VLookup(A2, Procesos_Real!E8:LV372,124,False)"
    Range("W2").Select
    Selection.AutoFill Destination:=Range("W2:W366"), Type:=xlFillDefault
    'HoraOperativas
    Range("X1") = "AGLOM-HorasOperativas"
    Range("X2").Formula = "=+VLookup(A2, Procesos_Real!E8:DP372,99,False)"
    Range("X2").Select
    Selection.AutoFill Destination:=Range("X2:X366"), Type:=xlFillDefault
    'Cemento
    Range("Y1") = "AGLOM-Cemento"
    Range("Y2").Formula = "=(Q2*U2)/1000"
    Range("Y2").Select
    Selection.AutoFill Destination:=Range("Y2:Y366"), Type:=xlFillDefault

    '************************* Stacker **********************************
    'MineralStacker
    Range("Z1") = "STACKER-MineralStacker"
    Range("Z2").Formula = "=+VLookup(A2, Procesos_Real!E8:LV372,152,False)"
    Range("Z2").Select
    Selection.AutoFill Destination:=Range("Z2:Z366"), Type:=xlFillDefault
    'LeyAu
    Range("AA1") = "STACKER-LeyAu"
    Range("AA2").Formula = "=+VLookup(A2, Procesos_Real!E8:LV372,151,False)"
    Range("AA2").Select
    Selection.AutoFill Destination:=Range("AA2:AA366"), Type:=xlFillDefault
    'AuApiladoStacker
    Range("AB1") = "STACKER-AuApiladoStacker"
    Range("AB2") = 0
    Range("AB2").Select
    Selection.AutoFill Destination:=Range("AB2:AB366"), Type:=xlFillDefault
    'Recuperacion
    Range("AC1") = "STACKER-Recuperacion"
    Range("AC2").Formula = "=+VLookup(A2, Procesos_Real!E8:LV372,154,False)*100"
    Range("AC2").Select
    Selection.AutoFill Destination:=Range("AC2:AC366"), Type:=xlFillDefault
    'AuExtraibleApilado
    Range("AD1") = "STACKER-AuExtraibleApilado"
    Range("AD2") = 0
    Range("AD2").Select
    Selection.AutoFill Destination:=Range("AD2:AD366"), Type:=xlFillDefault
    'Productividad
    Range("AE1") = "STACKER-Productividad"
    Range("AE2").Formula = "=+VLookup(A2, Procesos_Real!E8:LV372,158,False)"
    Range("AE2").Select
    Selection.AutoFill Destination:=Range("AE2:AE366"), Type:=xlFillDefault
    'Disponibilidad
    Range("AF1") = "STACKER-Disponibilidad"
    Range("AF2").Formula = "=+VLookup(A2, Procesos_Real!E8:LV372,159,False)"
    Range("AF2").Select
    Selection.AutoFill Destination:=Range("AF2:AF366"), Type:=xlFillDefault
    'Utilizacion
    Range("AG1") = "STACKER-Utilizacion"
    Range("AG2").Formula = "=+VLookup(A2, Procesos_Real!E8:LV372,160,False)"
    Range("AG2").Select
    Selection.AutoFill Destination:=Range("AG2:AG366"), Type:=xlFillDefault
    'HorasOperativas
    Range("AH1") = "STACKER-HorasOperativas"
    Range("AH2").Formula = "=+VLookup(A2, Procesos_Real!E8:LV372,156,False)"
    Range("AH2").Select
    Selection.AutoFill Destination:=Range("AH2:AH366"), Type:=xlFillDefault

    '************************* Apilamiento **********************************
    'MineralApilado
    Range("AI1") = "AP-MineralApilado"
    Range("AI2").Formula = "=+VLookup(A2, Procesos_Real!E8:LV372,196,False)"
    Range("AI2").Select
    Selection.AutoFill Destination:=Range("AI2:AI366"), Type:=xlFillDefault
    'LeyAu
    Range("AJ1") = "AP-LeyAu"
    Range("AJ2").Formula = "=+VLookup(A2, Procesos_Real!E8:LV372,197,False)"
    Range("AJ2").Select
    Selection.AutoFill Destination:=Range("AJ2:AJ366"), Type:=xlFillDefault
    'TotalAuApilado
    Range("AK1") = "AP-TotalAuApilado"
    Range("AK2").Formula = "=+VLookup(A2, Procesos_Real!E8:LV372,198,False)"
    Range("AK2").Select
    Selection.AutoFill Destination:=Range("AK2:AK366"), Type:=xlFillDefault
    'Recuperacion
    Range("AL1") = "AP-Recuperacion"
    Range("AL2").Formula = "=IFERROR(+VLookup(A2, Procesos_Real!E8:LV372,199,False)*100,0)"
    Range("AL2").Select
    Selection.AutoFill Destination:=Range("AL2:AL366"), Type:=xlFillDefault
    'TotalAuExApilado
    Range("AM1") = "AP-TotalAuExApilado"
    Range("AM2").Formula = "=+VLookup(A2, Procesos_Real!E8:LV372,200,False)*100"
    Range("AM2").Select
    Selection.AutoFill Destination:=Range("AM2:AM366"), Type:=xlFillDefault

    '************************* Lixiviacion **********************************
    'SolucionBarren
    Range("AN1") = "LIXI-SolucionBarren"
    Range("AN2").Formula = "=+VLookup(A2, Procesos_Real!E8:LV372,302,False)"
    Range("AN2").Select
    Selection.AutoFill Destination:=Range("AN2:AN366"), Type:=xlFillDefault
    'CNSolucionBarren
    Range("AO1") = "LIXI-CNSolucionBarren"
    Range("AO2").Formula = "=+VLookup(A2, Procesos_Real!E8:LV372,303,False)"
    Range("AO2").Select
    Selection.AutoFill Destination:=Range("AO2:AO366"), Type:=xlFillDefault
    'SolucionILS
    Range("AP1") = "LIXI-SolucionILS"
    Range("AP2").Formula = "=+VLookup(A2, Procesos_Real!E8:LV372,304,False)"
    Range("AP2").Select
    Selection.AutoFill Destination:=Range("AP2:AP366"), Type:=xlFillDefault
    'CNSolucionILS
    Range("AQ1") = "LIXI-CNSolucionILS"
    Range("AQ2").Formula = "=+VLookup(A2, Procesos_Real!E8:LV372,305,False)"
    Range("AQ2").Select
    Selection.AutoFill Destination:=Range("AQ2:AQ366"), Type:=xlFillDefault
    'SolucionPLS
    Range("AR1") = "LIXI-SolucionPLS"
    Range("AR2").Formula = "=+VLookup(A2, Procesos_Real!E8:LV372,306,False)"
    Range("AR2").Select
    Selection.AutoFill Destination:=Range("AR2:AR366"), Type:=xlFillDefault
    'LeyAuSolucionPLS
    Range("AS1") = "LIXI-LeyAuSolucionPLS"
    Range("AS2").Formula = "=+VLookup(A2, Procesos_Real!E8:LV372,307,False)"
    Range("AS2").Select
    Selection.AutoFill Destination:=Range("AS2:AS366"), Type:=xlFillDefault
    'AuLixiviado
    Range("AT1") = "LIXI-AuLixiviado"
    Range("AT2").Formula = "=+VLookup(A2, Procesos_Real!E8:OM372,397,False)"
    Range("AT2").Select
    Selection.AutoFill Destination:=Range("AT2:AT366"), Type:=xlFillDefault
    'CNSolucionPLS
    Range("AU1") = "LIXI-CNSolucionPLS"
    Range("AU2").Formula = "=+VLookup(A2, Procesos_Real!E8:LV372,309,False)"
    Range("AU2").Select
    Selection.AutoFill Destination:=Range("AU2:AU366"), Type:=xlFillDefault
    'pHSolucionPLS
    Range("AV1") = "LIXI-pHSolucionPLS"
    Range("AV2").Formula = "=+VLookup(A2, Procesos_Real!E8:LV372,310,False)"
    Range("AV2").Select
    Selection.AutoFill Destination:=Range("AV2:AV366"), Type:=xlFillDefault

    '************************* SART **********************************
    'PLSaSART
    Range("AW1") = "SART-PLSaSART"
    Range("AW2").Formula = "=+VLookup(A2, Procesos_Real!E8:DP372,82,False)"
    Range("AW2").Select
    Selection.AutoFill Destination:=Range("AW2:AW366"), Type:=xlFillDefault
    'LeyCuAlimentada
    Range("AX1") = "SART-LeyCuAlimentada"
    Range("AX2").Formula = "=+VLookup(A2, Procesos_Real!E8:DP372,87,False)"
    Range("AX2").Select
    Selection.AutoFill Destination:=Range("AX2:AX366"), Type:=xlFillDefault
    'LeyCuSalida
    Range("AY1") = "SART-LeyCuSalida"
    Range("AY2").Formula = "=+VLookup(A2, Procesos_Real!E8:DP372,90,False)"
    Range("AY2").Select
    Selection.AutoFill Destination:=Range("AY2:AY366"), Type:=xlFillDefault
    'LeyAuAlimentada
    Range("AZ1") = "SART-LeyAuAlimentada"
    Range("AZ2").Formula = "=+VLookup(A2, Procesos_Real!E8:MB372,331,False)"
    Range("AZ2").Select
    Selection.AutoFill Destination:=Range("AZ2:AZ366"), Type:=xlFillDefault
    'LeyAuSalida
    Range("BA1") = "SART-LeyAuSalida"
    Range("BA2").Formula = "=+VLookup(A2, Procesos_Real!E8:MB372,334,False)"
    Range("BA2").Select
    Selection.AutoFill Destination:=Range("BA2:BA366"), Type:=xlFillDefault
    'Eficiencia
    Range("BB1") = "SART-Eficiencia"
    Range("BB2").Formula = "=+VLookup(A2, Procesos_Real!E8:DP372,93,False)*100"
    Range("BB2").Select
    Selection.AutoFill Destination:=Range("BB2:BB366"), Type:=xlFillDefault

    '************************* ADR **********************************
    'PLSaCarbones
    Range("BC1") = "ADR-PLSaCarbones"
    Range("BC2").Formula = "=+VLookup(A2, Procesos_Real!E8:DP372,62,False)"
    Range("BC2").Select
    Selection.AutoFill Destination:=Range("BC2:BC366"), Type:=xlFillDefault
    'LeyAuPLS+ILS
    Range("BD1") = "ADR-LeyAuPLS+ILS"
    Range("BD2").Formula = "=+VLookup(A2, Procesos_Real!E8:DP372,67,False)"
    Range("BD2").Select
    Selection.AutoFill Destination:=Range("BD2:BD366"), Type:=xlFillDefault
    'LeyAuBLS
    Range("BE1") = "ADR-LeyAuBLS"
    Range("BE2").Formula = "=+VLookup(A2, Procesos_Real!E8:DP372,70,False)"
    Range("BE2").Select
    Selection.AutoFill Destination:=Range("BE2:BE366"), Type:=xlFillDefault
    'Eficiencia
    Range("BF1") = "ADR-Eficiencia"
    Range("BF2").Formula = "=+VLookup(A2, Procesos_Real!E8:DP372,73,False)"
    Range("BF2").Select
    Selection.AutoFill Destination:=Range("BF2:BF366"), Type:=xlFillDefault
    'AuAdsorbido
    Range("BG1") = "ADR-AuAdsorbido"
    Range("BG2").Formula = "=+VLookup(A2, Procesos_Real!E8:DP372,76,False)"
    Range("BG2").Select
    Selection.AutoFill Destination:=Range("BG2:BG366"), Type:=xlFillDefault
    'AuDesorbido
    Range("BH1") = "ADR-AuDesorbido"
    Range("BH2").Formula = "=+VLookup(A2, Procesos_Real!E8:DV372,120,False)"
    Range("BH2").Select
    Selection.AutoFill Destination:=Range("BH2:BH366"), Type:=xlFillDefault
    'AuDore
    Range("BI1") = "ADR-AuDore"
    Range("BI2").Formula = "=+VLookup(A2, Procesos_Real!E8:DP372,79,False)"
    Range("BI2").Select
    Selection.AutoFill Destination:=Range("BI2:BI366"), Type:=xlFillDefault
End Sub