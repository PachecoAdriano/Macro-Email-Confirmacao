Attribute VB_Name = "Módulo3"
Sub confirm_Aplicacao()
    Dim MyOlapp     As Object, MeuItem As Object
    Dim Cliente     As String
    Dim EmailCopia  As String
    Dim Email       As String
    Dim Linha       As Integer
    Dim PauseTime   As Integer
    Dim Start       As Single
    Dim TipoOp      As String
    Dim TipoFundo   As String
    Dim Valor       As String
    Dim Corpo       As String
    
    
         Linha = Sheets("Confirmação Aplicação").Cells(Sheets("Confirmação Aplicação").Rows.Count, 1).End(xlUp).Row
         
         Set MyOlapp = CreateObject("Outlook.Application")
         PauseTime = Range("M7")
         
        
        Do While Linha >= 4
             EmailCopia = Range("J" & Linha)
             Email = Range("H" & Linha)
             Cliente = Range("G" & Linha)
             TipoOp = Range("K" & Linha)
             TipoFundo = Range("F" & Linha)
             Valor = Range("D" & Linha).Value
         
             Set MeuItem = MyOlapp.CreateItem(olMailItem)
             With MeuItem
             
                 .Display
                 .to = Email
                 .CC = EmailCopia & ";" & "investimento@fiduc.com.br"
                 .Subject = "CONFIRMAÇÃO ORDEM DE APLICAÇÃO FIDUC"
                 Corpo = "<font size=3 color=1F497D face=calibri>Olá, <br >" & Cliente
                 Corpo = Corpo & "<br>"
                 Corpo = Corpo & "<br><font size=3 color=1F497D face=calibri>Ordem recebida e executada conforme abaixo:"
                 Corpo = Corpo & "<br><font size=3 color=1F497D face=calibri>Aplicação no Superfundo FIDUC " & TipoFundo
                 Corpo = Corpo & "<br><font size=3 color=1F497D face=calibri>Valor: " & FormatCurrency(Replace(Valor, ".", ","))
                 Corpo = Corpo & "<br><font size=3 color=1F497D face=calibri>Liquidação: " & TipoOp
                 Corpo = Corpo & "<br>"
                 Corpo = Corpo & "<br>Atenciosamente"
                 .HTMLBody = Corpo & .HTMLBody
                 .Send
                 
             End With
             Start = Timer    ' Set start time.
             Do While Timer < Start + PauseTime
                 DoEvents    ' Yield to other processes.
             Loop
             
             
             Linha = Linha - 1
             
         Loop
         
         MsgBox "Troxa!"


End Sub



