Sub Macro1()
'
' Macro1 Macro
' 宏由 Administrator 录制，时间: 2019/04/19
'

'
    Range("B1").Select
    Selection.Formula = "=LEFT(A1,9)"
    Range("B1").Select
    Selection.AutoFill Destination:=Range("B1:B15062"), Type:=xlFillDefault
    Range("B1:B15062").Select
    Range("C1").Select
    Selection.Formula = "=MID(A1,12,47)"
    Range("C1").Select
    Selection.AutoFill Destination:=Range("C1:C15062"), Type:=xlFillDefault
    Range("C1:C15062").Select
End Sub
Sub Macro2()
'
' Macro2 Macro
' 宏由 Administrator 录制，时间: 2019/04/19
'

'
    Application.CutCopyMode = False
    Selection.Formula = "=PHONETIC(D1:D1000)"
    Range("E2").Select
End Sub
Sub Macro3()
'
' Macro3 Macro
' 宏由 Administrator 录制，时间: 2019/04/19
'

'
    Range("A2").Select
    Selection.Formula = "1"
    Selection.AutoFill Destination:=Range("A2:A326"), Type:=xlFillDefault
    
    Range("B2").Select
    Selection.Formula = "head"
    Range("C2").Select
    Selection.Formula = "X/HEX"
    Range("D2").Select
    Selection.Formula = "Y/HEX"
    Range("E2").Select
    Selection.Formula = "Z/HEX"
    Range("F2").Select
    Selection.Formula = "X/°"
    Range("G2").Select
    Selection.Formula = "Y/°"
    Range("H2").Select
    Selection.Formula = "Z/°"
    
    Range("B3").Select
    Selection.Formula = "=MID(A$1,1+98*(A3-2),6)"
    Selection.AutoFill Destination:=Range("B3:B230"), Type:=xlFillDefault
                        
                        
    Range("C3").Select
    Selection.Formula = "=MID(A$1,7+98*(A3-2),8)"
    Selection.AutoFill Destination:=Range("C3:C230"), Type:=xlFillDefault
    
    Range("D3").Select
    Selection.Formula = "=MID(A$1,15+98*(A3-2),8)"
    Selection.AutoFill Destination:=Range("D3:D230"), Type:=xlFillDefault

    Range("E3").Select
    Selection.Formula = "=MID(A$1,23+98*(A3-2),8)"
    Selection.AutoFill Destination:=Range("E3:E230"), Type:=xlFillDefault
                    
                    
    Range("F3").Select
    Selection.Formula = "=IF(HEX2DEC(C3)>2147483647,-(HEX2DEC(C3)/2),HEX2DEC(C3))/2147483647*400"
    Selection.AutoFill Destination:=Range("F3:F230"), Type:=xlFillDefault
    
    Range("G3").Select
    Selection.Formula = "=IF(HEX2DEC(D3)>2147483647,-(HEX2DEC(D3)/2),HEX2DEC(D3))/2147483647*400"
    Selection.AutoFill Destination:=Range("G3:G230"), Type:=xlFillDefault

    Range("H3").Select
    Selection.Formula = "=IF(HEX2DEC(E3)>2147483647,-(HEX2DEC(E3)/2),HEX2DEC(E3))/2147483647*400"
    Selection.AutoFill Destination:=Range("H3:H230"), Type:=xlFillDefault
    
    
End Sub
