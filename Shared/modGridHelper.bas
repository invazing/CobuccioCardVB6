Attribute VB_Name = "modGridHelper"
Option Explicit

Public Sub ConfigurarGridTransacoes(ByRef grid As MSFlexGrid)

        '<EhHeader>
        On Error GoTo ConfigurarGridTransacoes_Err

        '</EhHeader>

100     With grid

102         .Rows = 2
104         .Cols = 6
106         .FixedRows = 1
108         .FixedCols = 0
110         .Rows = 1

112         .FocusRect = flexFocusNone
114         .SelectionMode = flexSelectionByRow
116         .AllowUserResizing = flexResizeColumns

118         .Row = 0
120         .Col = 0: .Text = "ID"
122         .Col = 1: .Text = "Número Cartão"
124         .Col = 2: .Text = "Valor"
126         .Col = 3: .Text = "Data"
128         .Col = 4: .Text = "Descrição"
130         .Col = 5: .Text = "Status"

132         .ColWidth(0) = 800
134         .ColWidth(1) = 1800
136         .ColWidth(2) = 1200
138         .ColWidth(3) = 1400
140         .ColWidth(4) = 3000
142         .ColWidth(5) = 1200

144         .Row = 0

            Dim i As Integer

146         For i = 0 To .Cols - 1
148             .Col = i
150             .CellAlignment = flexAlignCenterCenter
152             .CellFontBold = True
154             .CellBackColor = &H8000000F
156         Next i

158         .GridLines = flexGridFlat
160         .GridColor = &HC0C0C0
        End With

        '<EhFooter>
        Exit Sub

ConfigurarGridTransacoes_Err:
        RegistrarErro "modGridHelper", Err.Description
        MsgBox Err.Description & vbCrLf & "in CobuccioCard.modGridHelper.ConfigurarGridTransacoes " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Public Sub AjustarLarguraAutomatica(ByRef grid As MSFlexGrid)

        '<EhHeader>
        On Error GoTo AjustarLarguraAutomatica_Err

        '</EhHeader>

        On Error Resume Next

        Dim larguraDisponivel As Long

        Dim larguraColuna     As Long

        Dim i                 As Integer

        Dim margem            As Long

        Dim totalCols         As Integer
    
100     If grid Is Nothing Then Exit Sub
102     If grid.Cols = 0 Then Exit Sub

104     margem = 300
106     totalCols = grid.Cols

108     With grid
        
110         larguraDisponivel = .Width - margem
       
112         larguraColuna = larguraDisponivel / totalCols
      
114         For i = 0 To totalCols - 1
116             .ColWidth(i) = larguraColuna
118         Next i
        
        End With

        '<EhFooter>
        Exit Sub

AjustarLarguraAutomatica_Err:
        RegistrarErro "modGridHelper", Err.Description
        MsgBox Err.Description & vbCrLf & "in CobuccioCard.modGridHelper.AjustarLarguraAutomatica " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Public Sub AdicionarLinhaGrid(ByRef grid As MSFlexGrid, _
                              ByVal idTransacao As Long, _
                              ByVal numeroCartao As String, _
                              ByVal valor As Currency, _
                              ByVal dataTransacao As Date, _
                              ByVal descricao As String, _
                              ByVal status As String)

        '<EhHeader>
        On Error GoTo AdicionarLinhaGrid_Err

        '</EhHeader>

        On Error Resume Next

        Dim linha As Long
    
100     With grid
102         .Rows = .Rows + 1
104         linha = .Rows - 1

106         .TextMatrix(linha, 0) = idTransacao
108         .TextMatrix(linha, 1) = numeroCartao
110         .TextMatrix(linha, 2) = FormatCurrency(valor, 2)
112         .TextMatrix(linha, 3) = Format(dataTransacao, "dd/mm/yyyy HH:nn")
114         .TextMatrix(linha, 4) = descricao
116         .TextMatrix(linha, 5) = status

118         AplicarCorPorStatus grid, linha, status

120         .Row = linha
122         .Col = 0: .CellAlignment = flexAlignCenterCenter
124         .Col = 1: .CellAlignment = flexAlignLeftCenter
126         .Col = 2: .CellAlignment = flexAlignRightCenter
128         .Col = 3: .CellAlignment = flexAlignCenterCenter
130         .Col = 4: .CellAlignment = flexAlignLeftCenter
132         .Col = 5: .CellAlignment = flexAlignCenterCenter
        End With

        '<EhFooter>
        Exit Sub

AdicionarLinhaGrid_Err:
        RegistrarErro "modGridHelper", Err.Description
        MsgBox Err.Description & vbCrLf & "in CobuccioCard.modGridHelper.AdicionarLinhaGrid " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Private Sub AplicarCorPorStatus(ByRef grid As MSFlexGrid, _
                                ByVal linha As Long, _
                                ByVal status As String)

        '<EhHeader>
        On Error GoTo AplicarCorPorStatus_Err

        '</EhHeader>

        Dim corFundo As Long

        Dim corTexto As Long

        Dim i        As Integer

100     Select Case UCase(Trim(status))

            Case "APROVADA"
102             corFundo = RGB(144, 238, 144)
104             corTexto = RGB(0, 100, 0)

106         Case "PENDENTE"
108             corFundo = RGB(255, 255, 153)
110             corTexto = RGB(153, 102, 0)

112         Case "CANCELADA"
114             corFundo = RGB(255, 182, 193)
116             corTexto = RGB(139, 0, 0)

118         Case Else
120             corFundo = vbWhite
122             corTexto = vbBlack
        End Select
    
124     With grid
126         .Row = linha

128         For i = 0 To .Cols - 1
130             .Col = i
132             .CellBackColor = corFundo
134             .CellForeColor = corTexto
136         Next i

        End With

        '<EhFooter>
        Exit Sub

AplicarCorPorStatus_Err:
        RegistrarErro "modGridHelper", Err.Description
        MsgBox Err.Description & vbCrLf & "in CobuccioCard.modGridHelper.AplicarCorPorStatus " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Public Sub LimparGrid(ByRef grid As MSFlexGrid)

        '<EhHeader>
        On Error GoTo LimparGrid_Err

        '</EhHeader>

100     With grid
102         .Rows = 2
104         .FixedRows = 1
106         .Rows = 1
        End With

        '<EhFooter>
        Exit Sub

LimparGrid_Err:
        RegistrarErro "modGridHelper", Err.Description
        MsgBox Err.Description & vbCrLf & "in CobuccioCard.modGridHelper.LimparGrid " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Public Function ObterIdSelecionado(ByRef grid As MSFlexGrid) As Long

        '<EhHeader>
        On Error GoTo ObterIdSelecionado_Err

        '</EhHeader>

100     ObterIdSelecionado = 0
    
102     With grid

104         If .Row > 0 And .Rows > 1 Then
106             ObterIdSelecionado = CLng(.TextMatrix(.Row, 0))
            End If

        End With

        '<EhFooter>
        Exit Function

ObterIdSelecionado_Err:
        RegistrarErro "modGridHelper", Err.Description
        MsgBox Err.Description & vbCrLf & "in CobuccioCard.modGridHelper.ObterIdSelecionado " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Function

Public Sub PreencherGridComRecordset(ByRef grid As MSFlexGrid, _
                                     ByRef rs As ADODB.Recordset)

        '<EhHeader>
        On Error GoTo PreencherGridComRecordset_Err

        '</EhHeader>

        Dim i            As Long, linha As Long

        Dim temCabecalho As Boolean

        Dim colStatus    As Integer

        Dim colCategoria As Integer

        Dim temStatus    As Boolean

        Dim temCategoria As Boolean

100     If rs Is Nothing Then Exit Sub
102     If rs.EOF Then
104         MsgBox "Nenhum dado encontrado.", vbInformation

            Exit Sub

        End If

106     With grid
108         temCabecalho = (.Rows > 0 And .FixedRows = 1)
110         colStatus = -1
112         colCategoria = -1

114         For i = 0 To rs.Fields.Count - 1

116             Select Case LCase(rs.Fields(i).Name)

                    Case "status", "status_transacao"
118                     colStatus = i
120                     temStatus = True

122                 Case "categoria"
124                     colCategoria = i
126                     temCategoria = True
                End Select

128         Next i

130         Do Until rs.EOF
132             .Rows = .Rows + 1
134             linha = .Rows - 1

136             For i = 0 To rs.Fields.Count - 1

138                 If i < .Cols Then
140                     .TextMatrix(linha, i) = Nz(rs.Fields(i).Value, "")
                    End If

142             Next i

144             If temStatus Then
146                 AplicarCorPorStatus grid, linha, rs.Fields(colStatus).Value
                End If

148             rs.MoveNext
            Loop

150         AplicarCoresAlternadas grid
        End With

        '<EhFooter>
        Exit Sub

PreencherGridComRecordset_Err:
        RegistrarErro "modGridHelper", Err.Description
        MsgBox Err.Description & vbCrLf & "in CobuccioCard.modGridHelper.PreencherGridComRecordset " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Public Sub AplicarCoresAlternadas(ByRef grid As MSFlexGrid)

        '<EhHeader>
        On Error GoTo AplicarCoresAlternadas_Err

        '</EhHeader>

        On Error Resume Next

        Dim i   As Long

        Dim cor As Long

        Dim j   As Integer
    
100     With grid

102         For i = 1 To .Rows - 1

104             If i Mod 2 = 0 Then
106                 cor = RGB(245, 245, 245)
                Else
108                 cor = vbWhite
                End If
            
110             .Row = i
112             .Col = 0

114             If .CellBackColor = vbWhite Or .CellBackColor = RGB(245, 245, 245) Then

116                 For j = 0 To .Cols - 1
118                     .Col = j
120                     .CellBackColor = cor
122                 Next j

                End If

124         Next i

        End With

        '<EhFooter>
        Exit Sub

AplicarCoresAlternadas_Err:
        RegistrarErro "modGridHelper", Err.Description
        MsgBox Err.Description & vbCrLf & "in CobuccioCard.modGridHelper.AplicarCoresAlternadas " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Public Sub ConfigurarGridResumoFinanceiro(ByRef grid As MSFlexGrid)

        '<EhHeader>
        On Error GoTo ConfigurarGridResumoFinanceiro_Err

        '</EhHeader>

100     With grid
102         .Clear
104         .Rows = 1
106         .Cols = 7
108         .FixedRows = 1
110         .FixedCols = 0

112         .TextMatrix(0, 0) = "ID"
114         .TextMatrix(0, 1) = "Número Cartão"
116         .TextMatrix(0, 2) = "Descrição"
118         .TextMatrix(0, 3) = "Valor"
120         .TextMatrix(0, 4) = "Data"
122         .TextMatrix(0, 5) = "Status"
124         .TextMatrix(0, 6) = "Categoria"

126         .ColWidth(0) = 800
128         .ColWidth(1) = 1600
130         .ColWidth(2) = 2600
132         .ColWidth(3) = 1200
134         .ColWidth(4) = 1400
136         .ColWidth(5) = 1100
138         .ColWidth(6) = 1600

            Dim i As Integer

140         For i = 0 To .Cols - 1
142             .Row = 0
144             .Col = i
146             .CellAlignment = flexAlignCenterCenter
148             .CellFontBold = True
150             .CellBackColor = &H8000000F
152         Next i

154         .FocusRect = flexFocusNone
156         .SelectionMode = flexSelectionByRow
158         .AllowUserResizing = flexResizeColumns
160         .GridLines = flexGridFlat
162         .GridColor = &HC0C0C0
        End With

        '<EhFooter>
        Exit Sub

ConfigurarGridResumoFinanceiro_Err:
        RegistrarErro "modGridHelper", Err.Description
        MsgBox Err.Description & vbCrLf & "in CobuccioCard.modGridHelper.ConfigurarGridResumoFinanceiro " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Public Sub ConfigurarGridAnaliseFinanceira(ByRef grid As MSFlexGrid)

        '<EhHeader>
        On Error GoTo ConfigurarGridAnaliseFinanceira_Err

        '</EhHeader>

        On Error Resume Next

100     With grid
102         .Clear
104         .Rows = 1
106         .Cols = 6
108         .FixedRows = 1
110         .FixedCols = 0

112         .TextMatrix(0, 0) = "Status Transação"
114         .TextMatrix(0, 1) = "Categoria"
116         .TextMatrix(0, 2) = "Quantidade"
118         .TextMatrix(0, 3) = "Valor Total"
120         .TextMatrix(0, 4) = "Data Inicial"
122         .TextMatrix(0, 5) = "Data Final"

124         .ColWidth(0) = 1600
126         .ColWidth(1) = 1800
128         .ColWidth(2) = 1200
130         .ColWidth(3) = 1400
132         .ColWidth(4) = 1400
134         .ColWidth(5) = 1400

            Dim i As Integer

136         For i = 0 To .Cols - 1
138             .Row = 0
140             .Col = i
142             .CellAlignment = flexAlignCenterCenter
144             .CellFontBold = True
146             .CellBackColor = &H8000000F
148         Next i

150         .FocusRect = flexFocusNone
152         .SelectionMode = flexSelectionByRow
154         .AllowUserResizing = flexResizeColumns
156         .GridLines = flexGridFlat
158         .GridColor = &HC0C0C0
        End With

        '<EhFooter>
        Exit Sub

ConfigurarGridAnaliseFinanceira_Err:
        RegistrarErro "modGridHelper", Err.Description
        MsgBox Err.Description & vbCrLf & "in CobuccioCard.modGridHelper.ConfigurarGridAnaliseFinanceira " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Private Function Nz(ByVal valor As Variant, ByVal ValorPadrao As Variant) As Variant

        '<EhHeader>
        On Error GoTo Nz_Err

        '</EhHeader>

100     If IsNull(valor) Or IsEmpty(valor) Then
102         Nz = ValorPadrao
        Else
104         Nz = valor
        End If

        '<EhFooter>
        Exit Function

Nz_Err:
        RegistrarErro "modGridHelper", Err.Description
        MsgBox Err.Description & vbCrLf & "in CobuccioCard.modGridHelper.Nz " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Function

