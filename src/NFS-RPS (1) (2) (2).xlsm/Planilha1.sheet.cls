
Sub CriarArquivoTextoCabecalho()
    
    On Error GoTo TratarErro
    
    Dim lsCaminho   As String
    Dim llArquivo   As Long
    Dim linha  As Integer
    Dim ultimaLinha As Integer
    Dim COL As String
    
    Sheets("CABECALHO").Select
    
    
    'Path where the file will be saved
    lsCaminho = "C:\Users\pribeiro\OneDrive\Documentos\olá.txt"
  
    'Identifies whether the file already exists
    
        llArquivo = FreeFile
        Open lsCaminho For Output As #llArquivo
        
        ultimaLinha = CABECALHO.Cells(CABECALHO.Rows.Count, 2).End(xlUp).Row
        
        COL = 1
        For i = 3 To ultimaLinha
                
                For J = 1 To 7
                
                Print #llArquivo, FormataDadosCabecalho(CABECALHO.Cells(i, J))
                
                Next J
                
        
        Next i
  
        
        Close #llArquivo
   
  
'Error treatment
Sair:
    Close #llArquivo
    Exit Sub
TratarErro:
    MsgBox "There was an error writing the file!"
    GoTo Sair
    Resume
End Sub


Public Function FormataDadosCabecalho(ByVal Celula As Range) As String

    Dim tamanho As Long
    Dim tipo As String
    Dim ultimaLinha As Long
    Dim linha As Long
    Dim valor As String
    Dim LtValor As String
    Dim ColLayout As String
    
    
    
    
    ultimaLinha = Layout.Cells(Layout.Rows.Count, 1).End(xlUp).Row
    
    
    For linha = 3 To ultimaLinha
    
    
    tamanho = Layout.Cells(linha, 3).Value
    tipo = Layout.Cells(linha, 2).Value
    ColLayout = Layout.Cells(linha, 4).Value
    
    If tipo <> "Fixo" Then
    
        LtValor = CABECALHO.Range(Layout.Range(linha, 6).Value & Celula.Row).Value
                
                
        
    Else
    
        LtValor = ""
        
    End If
    
        Select Case tipo
        
            Case "FIXO"
            
                FormataDadosCabecalho = FormataDadosCabecalho & ColLayout
            
            Case "NUM"
            
                FormataDadosCabecalho = FormataDadosCabecalho & Left(Replace(Replace(Replace(LtValor.Value, "R$", ""), " ", ""), ",", ""))
                FormataDadosCabecalho = FormataDadosCabecalho & Format(FormataDadosCabecalho, String(tamanho, "0"))
            
            Case "TEXT"
            
                FormataDadosCabecalho = FormataDadosCabecalho & Left(LtValor & String(tamanho, " "), tamanho)
        
        End Select
    
    
    Next linha




End Function
Sub formatadados()
    Dim valor As String
    Dim valorarq As String
    Dim linhaLay As Integer
    Dim ultimaLinhaLay As Long
    Dim colLay As Integer
    Dim colCab As Integer
    Dim linha As Integer
    Dim tipo As String
    Dim tamanho As Long
    Dim lsArquivo As String
    Dim arquivoTexto As Object
    Dim ultimaLinhaDt As String
    Dim ultimaColDet As String
    Dim linhapreenchimentodet As String
    Dim Campo As String
    
    
    On Error GoTo trataerro
    
    
    
    Sheets("Detalhe").Cells(2, 4) = InputBox(vbCrLf & _
    "Insira o numero da sua última nota gerada", "Numeração da nota") + 1
    
    NumeroCorreto = MsgBox("Certifique-se de que o número '" & Sheets("Detalhe").Cells(2, 4).Value - 1 & "' confere." & vbCrLf & _
    "O envio incorreto compromete a emissão de notas em lote no site da Prefeitura." & vbCrLf & "Deseja mesmo continuar?", vbYesNo + vbQuestion, "Importante!")
    
    
    If NumeroCorreto = vbYes Then
    
    
    
    lsCaminho = InputBox("Insira o caminho da pasta onde deseja salvar", "Salvar") & "\RPS" & Sheets("Detalhe").Cells(2, 5).Value & ".txt"
    
    
    
    Set arquivoTexto = CreateObject("Scripting.FileSystemObject").CreateTextFile(lsCaminho, True)
    
    
    ultimaLinhaLay = Sheets("Layout").Cells(Rows.Count, 1).End(xlUp).Row
    
    
    For linha = 2 To ultimaLinhaLay
    
        tamanho = Sheets("Layout").Cells(linha, 4).Value
        
        tipo = Sheets("Layout").Cells(linha, 3).Value
        
       
            
        valor = Sheets("Layout").Cells(linha, 2)
        
        
        
        
        
        Campo = Sheets("Layout").Cells(linha, 1)
        
        
        Select Case tipo
        
        
            Case "FIXO"
                valorarq = valor
            
            Case "NUM"
                valorarq = Format(valor, String(tamanho, "0"))
                
            
            Case "TEXT"
                valorarq = valor & String(tamanho - Len(valor), " ")
                
            Case "DATA"
                valorarq = valor
                
            Case "LISTA"
                valorarq = Left(valor, 1)
        
        End Select
        
          
    
    
    arquivoTexto.Write valorarq
        
        
    valor = ""
    
    Next linha
    
    arquivoTexto.writeline ""
    
    ultimaLinhaDet = Sheets("Detalhe").Cells(Rows.Count, 8).End(xlUp).Row
    ultimaColDet = 52
    
    
    For i = 2 To ultimaLinhaDet + 1
    
        For J = 2 To ultimaColDet + 1
            
            Sheets("Layout").Cells(J, 7) = Sheets("Detalhe").Cells(i, J - 1)
            valor = ""
            
            tipo = Sheets("Layout").Cells(J, 8)
            
            If tipo = "NUMVALOR" Then
        
                valor = Format(Sheets("Layout").Cells(J, 7).Value, "Currency")
            
            
            Else
            
            
            valor = Sheets("Layout").Cells(J, 7).Value
            
            End If
            
            tamanho = Sheets("Layout").Cells(J, 9).Value
            
            
            Select Case tipo
            
                Case "FIXO"
                    valorarq = valor
                Case "NUM"
                    valorarq = Format(valor, String(tamanho, "0"))
                    
                Case "NUMVALOR"
                valorarq = Replace(Replace(Replace(valor, "R$", ""), " ", ""), ",", "")
                valorarq = Right(Format(valorarq, String(tamanho, "0")), tamanho)
                    
                Case "TEXT"
                'consertar aqui >>>>>>>>>>>
                    valorarq = valor & String(tamanho - Len(valor), " ")
                    
                Case "TEXTDESC"
                'consertar aqui >>>>>>>>>>>
                    tamanho = Len(valor)
                    valorarq = valor & String(tamanho - Len(valor), " ")
                    
                Case "LISTA"
                    valorarq = Left(valor, 1)
                
                Case "LISTA2"
                    valorarq = Left(valor, 2)
            
            End Select
            
        arquivoTexto.Write valorarq
            
        
        Next J
        
        If i < ultimaLinhaDet Then
        
            arquivoTexto.writeline ""
            
        End If
        
    Next i
    
    
    arquivoTexto.writeline ""
     
    For i = 2 To 7
    
        tipo = Sheets("Layout").Cells(i, 16).Value
        
        tamanho = Sheets("Layout").Cells(i, 17).Value
        
        Campo = Sheets("Layout").Cells(i, 14).Value
        
        valor = ""
        
        If tipo = "NUMVALOR" Then
        
            valor = Format(Sheets("Layout").Cells(i, 15).Value, "Currency")
            
            
        Else
        
        
    
            valor = Sheets("Layout").Cells(i, 15).Value
        
    
        End If
        
        Select Case tipo
            
                Case "FIXO"
                    valorarq = valor
                Case "NUM"
                    valorarq = Format(valor, String(tamanho, "0"))
                    
                Case "NUMVALOR"
                valorarq = Replace(Replace(Replace(valor, "R$", ""), " ", ""), ",", "")
                valorarq = Right(Format(valorarq, String(tamanho, "0")), tamanho)
                    
                Case "TEXT"
                'consertar aqui >>>>>>>>>>>
                    valorarq = valor & String(tamanho - Len(valor), " ")
                    
                Case "TEXTDESC"
                'consertar aqui >>>>>>>>>>>
                    tamanho = Len(valor)
                    valorarq = valor & String(tamanho - Len(valor), " ")
                    
                Case "LISTA"
                    valorarq = Left(valor, 1)
                
                Case "LISTA2"
                    valorarq = Left(valor, 2)
            
            End Select
            
        arquivoTexto.Write valorarq
        
        
        
        
    
    Next i
    
     
     
     
     
    MsgBox "Valor Total dos Serviços: R$ " & Sheets("Layout").Cells(4, 15).Value, vbOKOnly
    
    MsgBox "O arquivo foi salvo em: " & lsCaminho
    
    Exit Sub
    
    Else
    
trataerro:
    MsgBox "Operação Cancelada", vbOKOnly
    Sheets("Detalhe").Cells(2, 4) = ""
    End If

End Sub