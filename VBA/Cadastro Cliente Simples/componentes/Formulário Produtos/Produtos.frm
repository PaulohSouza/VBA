VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Produtos 
   Caption         =   "Cadastro de produtos"
   ClientHeight    =   7620
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10200
   OleObjectBlob   =   "Produtos.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Produtos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cancelar_Click()
On Error Resume Next
Range("a1").Select
 
Columns("a:a").Select
Selection.Find(What:=codigo.Value, After:=ActiveCell, LookIn:= _
        xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:= _
        xlNext, MatchCase:=False, SearchFormat:=False).Activate
        
    ActiveCell.Offset(0, 1).Activate
 
Selection.EntireRow.Delete
cancelar.Locked = True
salvar.Locked = True

 codigo.Text = ""

End Sub


Private Sub editar_Click()
'Habilitar Controles
Resposta = MsgBox("Deseja Editar produto Selecionado?", 4 + vbQuestion, "Editar Registro")
    Select Case Resposta
           Case vbYes

           coddun.Locked = False
           codean.Locked = False
           embalagem.Locked = False
           quant.Locked = False
           custo.Locked = False
           obs.Locked = False
           

salvar.Locked = False
excluir.Locked = True
ActiveCell.Offset(0, -8).Activate

Case vbNo
    End Select

End Sub

Private Sub excluir_Click()

Resposta = MsgBox("Deseja Excluir produto Selecionado?", 4 + 16, "Excluir Registro")
    Select Case Resposta
           Case vbYes
ActiveCell.Offset(0, 1).Activate
      ActiveCell = "Inativo"
      
MsgBox "produto Excluído com Sucesso?", 0 + vbInformation, "Registro Excluído"

            Sheets("Temp_Produtos").Activate
    Range("A1").Select
    Selection.CurrentRegion.Select
    Selection.EntireRow.Delete

Sheets("Produtos").Select
Range("A1").Select
ActiveSheet.ShowAllData

' Filtrar Status
    ActiveSheet.ListObjects("Tabela2").Range.AutoFilter Field:=11, Criteria1:="Ativo"
'Selecionar Região Atual
    Range("a1").Select
    Selection.CurrentRegion.Select
'Copiar
    Selection.Copy
'Colar em Plan Temporária
    Sheets("Temp_Produtos").Select
    Range("a1").Select
    Selection.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
        , SkipBlanks:=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
 'Deletar Colunas
   
      
    
    Columns("D:E").Select
    Selection.Delete Shift:=xlToLeft
    Columns("H:I").Select
    Selection.Delete Shift:=xlToLeft
    Range("A2").Select


            Case vbNo
            

          
    End Select

End Sub


Private Sub limpar_Click()
' Limpar caixas de texto
            
         codigo.Text = ""
         codpro.Text = ""
         descricao.Text = ""
         coddun.Text = ""
         codean.Text = ""
         fornecedor.Text = ""
         embalagem.Text = ""
         quant.Text = ""
         custo.Text = ""
         obs.Text = ""
         
End Sub

Private Sub ListBox1_Click()
On Error Resume Next
Sheets("Produtos").Select
'Limpar os Filtros, para que não haja problemas ao lançar novos registros
 Range("A1").Select
 ActiveSheet.ShowAllData
 
 
ActiveSheet.ListObjects("Tabela2").Range.AutoFilter Field:=1, Criteria1:="=" & ListBox1.Value
   
Columns("a:a").Select
Selection.Find(What:=ListBox1.Value, After:=ActiveCell, LookIn:= _
        xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:= _
        xlNext, MatchCase:=False, SearchFormat:=False).Activate
            
'Carregar Registros
    
    codigo.Text = ActiveCell
    ActiveCell.Offset(0, 1).Activate
    codpro.Text = ActiveCell
    ActiveCell.Offset(0, 1).Activate
    descricao.Text = ActiveCell
    ActiveCell.Offset(0, 1).Activate
    coddun.Text = ActiveCell
    ActiveCell.Offset(0, 1).Activate
    codean.Text = ActiveCell
    ActiveCell.Offset(0, 1).Activate
    fornecedor.Text = ActiveCell
    ActiveCell.Offset(0, 1).Activate
    embalagem.Text = ActiveCell
    ActiveCell.Offset(0, 1).Activate
    quant.Text = ActiveCell
    ActiveCell.Offset(0, 1).Activate
    custo.Text = ActiveCell
    ActiveCell.Offset(0, 1).Activate
    obs.Text = ActiveCell
    
    
editar.Locked = False
excluir.Locked = False


End Sub

Private Sub novo_Click()
On Error Resume Next
Resposta = MsgBox("Deseja Incluir um novo produto?", 4 + vbQuestion, "Incluir Registro")

    Select Case Resposta
    
           Case vbYes
         
           ' Desbloqueio das caixas de texto
           
           coddun.Locked = False
           codean.Locked = False
           embalagem.Locked = False
           quant.Locked = False
           custo.Locked = False
           obs.Locked = False
           
            ' Desbloqueio dos botões
           
           salvar.Locked = False
           cancelar.Locked = False
           
            ' Limpar caixas de texto
            
         codigo.Text = ""
         codpro.Text = ""
         descricao.Text = ""
         coddun.Text = ""
         codean.Text = ""
         fornecedor.Text = ""
         embalagem.Text = ""
         quat.Text = ""
         custo.Text = ""
         obs.Text = ""
                           
   'Ativa Plan Produtos
   Sheets("Produtos").Activate
   Range("a1").Select
   ActiveSheet.ShowAllData
   'Determina a proxima linha vazia
   ProximaLinha = Application.WorksheetFunction. _
       CountA(Range("A:A")) + 1
   Cells(ProximaLinha, 1) = ProximaLinha - 1

         codigo.Text = Cells(ProximaLinha, 1)
         
            Case vbNo
           
                  
           End Select
           
           End Sub


Private Sub pesquisar_Click()
'On Error Resume Next
Sheets("Produtos").Activate
Range("a1").Select

'Limpar Filtros
 ActiveSheet.ShowAllData
 
' Filtrar Codigo do produto
If codpro.Text <> "" Then
    ActiveSheet.ListObjects("Tabela2").Range.AutoFilter Field:=2, Criteria1:="=" & codpro.Text
Else
    ActiveSheet.ListObjects("Tabela2").Range.AutoFilter Field:=2
End If
'--------------------------------------------------------------------------------------------------
'Filtrar Descrição
If descricao.Text <> "" Then
    ActiveSheet.ListObjects("Tabela2").Range.AutoFilter Field:=3, Criteria1:="=" & descricao.Text
Else
    ActiveSheet.ListObjects("Tabela2").Range.AutoFilter Field:=3
End If
'--------------------------------------------------------------------------------------------------
'Filtrar Fornecedor
If fornecedor.Text <> "" Then
    ActiveSheet.ListObjects("Tabela2").Range.AutoFilter Field:=6, Criteria1:="=" & fornecedor.Text
Else
    ActiveSheet.ListObjects("Tabela2").Range.AutoFilter Field:=6
End If
'-------------------------------------------------------------------------------------------------


Sheets("Temp_Produtos").Activate
    Range("A1").Select
    Selection.CurrentRegion.Select
    Selection.EntireRow.Delete

Sheets("Produtos").Select
Range("A1").Select

' Filtrar Status
    ActiveSheet.ListObjects("Tabela2").Range.AutoFilter Field:=11, Criteria1:="Ativo"
'Selecionar Região Atual
    Range("a1").Select
    Selection.CurrentRegion.Select
'Copiar
    Selection.Copy
'Colar em Plan Temporária
    Sheets("Temp_Produtos").Select
    Range("a1").Select
    Selection.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
        , SkipBlanks:=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
 'Deletar Colunas
   
    
    Columns("D:E").Select
    Selection.Delete Shift:=xlToLeft
    Columns("H:I").Select
    Selection.Delete Shift:=xlToLeft
    Range("A2").Select
          


End Sub



Private Sub salvar_Click()
On Error Resume Next
Sheets("Produtos").Activate
 Range("A1").Select
 ActiveSheet.ShowAllData
 
  If codpro = "" Then
        MsgBox "Insira o código do produto", vbCritical, "Atenção,"
        codpro.SetFocus
        Exit Sub
        End If
        
        If descricao = "" Then
        MsgBox "Insira a descrição do produto", vbCritical, "Atenção,"
        descricao.SetFocus
        Exit Sub
        End If

If fornecedor = "" Then
        MsgBox "Insira o fornecedor", vbCritical, "Atenção,"
        fornecedor.SetFocus
        Exit Sub
        End If

If embalagem = "" Then
        MsgBox "Insira o tipo de embalagem", vbCritical, "Atenção,"
        embalagem.SetFocus
        Exit Sub
        End If

If quant = "" Then
        MsgBox "Insira a quantidade do produto", vbCritical, "Atenção,"
        quant.SetFocus
        Exit Sub
        End If


  
Columns("a:a").Select
Selection.Find(What:=codigo.Value, After:=ActiveCell, LookIn:= _
        xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:= _
        xlNext, MatchCase:=False, SearchFormat:=False).Activate
        
        ActiveCell.Offset(0, 1).Activate
        ActiveCell = codpro.Text
        ActiveCell.Offset(0, 1).Activate
        ActiveCell = descricao.Text
        ActiveCell.Offset(0, 1).Activate
        ActiveCell = coddun.Text
        ActiveCell.Offset(0, 1).Activate
        ActiveCell = codean.Text
        ActiveCell.Offset(0, 1).Activate
        ActiveCell = fornecedor.Text
        ActiveCell.Offset(0, 1).Activate
        ActiveCell = embalagem.Text
        ActiveCell.Offset(0, 1).Activate
        ActiveCell = quant.Text
        ActiveCell.Offset(0, 1).Activate
        ActiveCell = custo.Text
        ActiveCell.Offset(0, 1).Activate
        ActiveCell = obs.Text
        ActiveCell.Offset(0, 1).Activate
        ActiveCell = "Ativo"
        
        salvar.Locked = True

MsgBox "Registro Salvo com Sucesso!", 0 + vbInformation, "Registro Salvo"

Sheets("Temp_Produtos").Activate
    Range("A1").Select
    Selection.CurrentRegion.Select
    Selection.EntireRow.Delete

Sheets("Produtos").Select
Range("A1").Select
ActiveSheet.ShowAllData

' Filtrar Status
    ActiveSheet.ListObjects("Tabela2").Range.AutoFilter Field:=11, Criteria1:="Ativo"
'Selecionar Região Atual
    Range("a1").Select
    Selection.CurrentRegion.Select
'Copiar
    Selection.Copy
'Colar em Plan Temporária
    Sheets("Temp_Produtos").Select
    Range("a1").Select
    Selection.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
        , SkipBlanks:=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
 'Deletar Colunas
  
    Columns("D:E").Select
    Selection.Delete Shift:=xlToLeft
    Columns("H:I").Select
    Selection.Delete Shift:=xlToLeft
    Range("A2").Select
    
    ' Limpar caixas de texto
            
         codigo.Text = ""
         codpro.Text = ""
         descricao.Text = ""
         coddun.Text = ""
         codean.Text = ""
         fornecedor.Text = ""
         embalagem.Text = ""
         quant.Text = ""
         custo.Text = ""
         obs.Text = ""
          
End Sub


Private Sub UserForm_Initialize()
On Error Resume Next

With embalagem
    .AddItem "CX"
    .AddItem "UND"
End With

Sheets("Temp_Produtos").Activate
    Range("A1").Select
    Selection.CurrentRegion.Select
    Selection.EntireRow.Delete

Sheets("Produtos").Select
Range("A1").Select
ActiveSheet.ShowAllData

' Filtrar Status
    ActiveSheet.ListObjects("Tabela2").Range.AutoFilter Field:=11, Criteria1:="Ativo"
'Selecionar Região Atual
    Range("a1").Select
    Selection.CurrentRegion.Select
'Copiar
    Selection.Copy
'Colar em Plan Temporária
    Sheets("Temp_Produtos").Select
    Range("a1").Select
    Selection.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
        , SkipBlanks:=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
 'Deletar Colunas
   
    Columns("D:E").Select
    Selection.Delete Shift:=xlToLeft
    Columns("H:I").Select
    Selection.Delete Shift:=xlToLeft
    Range("A2").Select

End Sub


