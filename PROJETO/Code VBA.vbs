'codigo vba dentro da planilha

Sub teste()
Dim ws As Worksheet
Dim linha_soma As String
Dim nome As String
Set ws = ThisWorkbook.Worksheets("dados")

linha_soma = ws.Range("d5").End(xlDown).Row

ws.Range("A5:j5").Copy

' Cola apenas a formatação no intervalo B1:B10
ws.Range("a5:j" & linha_soma + 1).PasteSpecial Paste:=xlPasteFormats

' Limpa a área de transferência
Application.CutCopyMode = False

Call formulas(ws, linha_soma)

nome = ws.Range("j3").Value


 With ws.PageSetup
        .LeftMargin = Application.InchesToPoints(0.5) ' Margem esquerda
        .RightMargin = Application.InchesToPoints(0.5) ' Margem direita
        .TopMargin = Application.InchesToPoints(0.5) ' Margem superior
        .BottomMargin = Application.InchesToPoints(0.5) ' Margem inferior
        .FitToPagesWide = 1 ' Ajustar para caber em uma página de largura
        .FitToPagesTall = False ' Não ajustar a altura
    End With


    ws.Columns("A").Hidden = True
    ws.Columns("F").Hidden = True
    ws.Columns("H").Hidden = True

    
    ' Selecionar o intervalo e imprimir
    ws.Range("A1:J" & linha_soma + 1).Select
    Selection.PrintOut Copies:=1, Collate:=True, PrintToFile:=True, PrToFileName:="C:\AUTOMACAO\RELATORIOS PARCELAS\RELATORIOS\" & nome & ".pdf"
    
    
    ws.Columns("A").Hidden = False
    ws.Columns("F").Hidden = False
    ws.Columns("H").Hidden = False



''antigo
'Range("A1:j" & linha_soma + 1).Select
'Selection.PrintOut Copies:=1, Collate:=True, PrintToFile:=True, PrToFileName:="C:\AUTOMACAO\RELATORIOS PARCELAS\RELATORIOS\" & nome & ".pdf"

End Sub

Sub formulas(ws As Worksheet, linha As String)

''formula dias corrido
'ws.Range("a5:a" & linha).Formula = "=IF(IF(RC[2]<=R2C13,0,RC[2]-R2C13)<0,0,IF(RC[2]<=R2C13,0,RC[2]-R2C13))"
Call dias_corridos(ws, linha)

'formula juros
ws.Range("h5:h" & linha).Formula = "=if(RC[-3]=""AlteracaoDeVencimento"",0,if(RC[1]=0,""0"",RC[-4]-RC[1]))"

'formula saldo remanescente
ws.Range("i5:i" & linha).Formula = "=if(RC[-4]=""AlteracaoDeVencimento"",0,RC[-5]/(1+R2C14/100)^(RC[-8]/30)-RC[-2])"

'formula liquidado ou em aberto
ws.Range("j5:j" & linha).Formula = "=IF(RC[-1]<>0, ""Em aberto"", ""Liquidado"")"

'ws.Range("j1").Value = Format(Now, "dd/mm/yyyy")
'ws.Range("j2").Value = ws.Range("j1").Value + 7

'SOMAS DO DOCUMENTO
ws.Range("d" & linha + 1).Formula = "=SUM(R5C4:R" & linha & "C4)"
ws.Range("h" & linha + 1).Formula = "=SUM(R5C8:R" & linha & "C8)"
ws.Range("i" & linha + 1).Formula = "=SUM(R5C9:R" & linha & "C9)"


ws.Range("a" & linha + 1 & ":j" & linha + 1).Interior.Color = RGB(212, 212, 212)

End Sub

Sub dias_corridos(ws As Worksheet, linha As String)
Dim dt As Date
Dim dt1 As Date
Dim dias As Integer

dt = Format(ws.Range("m2").Value, "dd/MM/yyyy")

    For i = 5 To linha
    
        dt1 = Format(ws.Range("c" & i).Value, "dd/MM/yyyy")
        
        dias = dt1 - dt
        
        If dias > 0 Then
        
            ws.Range("a" & i).Value = dias
        
        Else
        
            ws.Range("a" & i).Value = 0
        
        End If
    
    Next i

End Sub



