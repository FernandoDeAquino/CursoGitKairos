VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClMembros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim Sql As String
Dim mMatricula As Long
Public Property Let Matricula(kVal As Long)
   mMatricula = kVal
End Property
Public Property Get Matricula() As Long
   Matricula = mMatricula
End Property
Public Function ProxMatricula(BdNomeCompleto As String) As Double
    Dim Rs  As New ADODB.Recordset
    Dim Bd  As New ADODB.Connection
'Abre uma conexão
    Bd.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & BdNomeCompleto
    Sql = "SELECT Max(membros.ID_Matr) AS ID_Matr FROM membros;"
    Rs.Open Sql, Bd
    If Not Rs.EOF Then
    
    End If
    ProxMatricula = IIf(Not Rs.EOF, (Rs.Fields(0).Value + 1), 1)
End Function
Public Function PesquisaDataAniversario(BdNomeCompleto As String, MesDeAniversario As String) As Recordset
' Recordset somente-leitura
    Dim Rs  As New ADODB.Recordset
    Dim Bd  As New ADODB.Connection
'   On Error GoTo Erro

    'Abre uma conexão
    Bd.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & BdNomeCompleto
     MesDeAniversario = Format(MesDeAniversario, "00")
     Sql = "SELECT C_Nome, Day([DT_Nascimento]) AS Nascimento From membros "
   '  Sql = Sql & "Where  and Month [DT_Nascimento] = " & MesDeAniversario & " ORDER BY Day([DT_Nascimento]);"
     Sql = Sql & "Where C_Situacao<>'Excluido' AND(((Month([DT_Nascimento])) = " & MesDeAniversario & "))ORDER BY Day([DT_Nascimento]),C_Nome;"
     'Sql = "SELECT membros.C_Nome, Format([DT_Nascimento],'dd/mm/yyyy') AS DT_Nascimento FROM membros WHERE (((Format([DT_Nascimento],'mm'))='" & MesDeAniversario & "')) ORDER BY Format([DT_Nascimento],'dd/mm/yyyy') DESC;"

    
    Rs.Open Sql, Bd
    Set PesquisaDataAniversario = Rs
   
 
    Set Rs = Nothing
    Set Bd = Nothing
    
 '   Exit Function
'Erro:
'    Call ErrorMessage(Err)

End Function
Public Function CriaCarteira(BdNomeCompleto As String, Optional Funcao As String) As Recordset
    Dim Rs  As New ADODB.Recordset
    Dim Bd  As New ADODB.Connection
    
    If Funcao <> "" Then
        Sql = "SELECT * From membros WHERE C_FuncaoEcle='" & Funcao & "' and C_Situacao='ATIVO'"
    End If
    
    
    
    
    Bd.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & BdNomeCompleto
    
    Sql = Sql & " ORDER BY C_nome"
        Rs.Open Sql, Bd
    Set CriaCarteira = Rs
   
 
    Set Rs = Nothing
    Set Bd = Nothing


End Function

Public Function pesquisa(BdNomeCompleto As String, Optional Matricula As Long, Optional Nome As String) As Recordset
' Recordset somente-leitura
    Dim Rs  As New ADODB.Recordset
    Dim Bd  As New ADODB.Connection
'   On Error GoTo Erro

    'Abre uma conexão
    Bd.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & BdNomeCompleto
    If Matricula = 0 And Nome = "" Then
        Sql = "SELECT * From Membros "
    ElseIf Matricula <> 0 Then
        Sql = "SELECT * From membros WHERE Id_Matr =" & Matricula
    Else
        Sql = "SELECT * From membros WHERE C_nome Like '" & Nome & "%'"
    End If
    
    Sql = Sql & " ORDER BY C_nome"
        Rs.Open Sql, Bd
    Set pesquisa = Rs
   
 
    Set Rs = Nothing
    Set Bd = Nothing
    
 '   Exit Function
'Erro:
'    Call ErrorMessage(Err)

End Function

Public Sub Status(BdNomeCompleto As String, Matricula As Long, Status As Boolean)

Dim Bd  As New ADODB.Connection
On Error GoTo Erro
 '   Sql = "DELETE * FROM Aluno WHERE ID_Mat=" & Matricula
Sql = "Update membros Set F_Desativado =" & Status & " WHERE ID_Matr=" & Matricula
    Bd.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & BdNomeCompleto
    Bd.Execute Sql

 
 Set Bd = Nothing
 
   Exit Sub
Erro:
    Call ErrorMessage(Err)

End Sub


Public Sub Excluir(BdNomeCompleto As String, Matricula As Long)

Dim Bd  As New ADODB.Connection
On Error GoTo Erro
 '   Sql = "DELETE * FROM Aluno WHERE ID_Mat=" & Matricula
Sql = "Update membros Set F_Desativado = TRUE WHERE ID_Matr=" & Matricula
    Bd.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & BdNomeCompleto
    Bd.Execute Sql

 
 Set Bd = Nothing
 
   Exit Sub
Erro:
    Call ErrorMessage(Err)

End Sub

Public Function Localizar(BdNomeCompleto As String, Matricula As Long) As Recordset
' Recordset somente-leitura
    Dim Rs  As New ADODB.Recordset
    Dim Bd  As New ADODB.Connection
   On Error GoTo Erro

    'Abre uma conexão
    Bd.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & BdNomeCompleto
    
    Sql = "SELECT * From membros WHERE Id_Matr =" & Matricula
    
    'Abre o Recordse somente-leitura

    Rs.Open Sql, Bd

    Set Localizar = Rs
    'Imprime os valores dos campos para o primeiro registro

    'For Each fld In rs.Fields

     '   Debug.Print fld.Value & ";";

    'Next

    'Fecha o Recordset
 
    Set Rs = Nothing
    Set Bd = Nothing
    
    Exit Function
Erro:
    Call ErrorMessage(Err)

End Function
Public Sub Grava(BdNomeCompleto As String, Matricula As Long, Nome As String, _
                 Endereco As String, Complemento As String, Numero As String, _
                 Bairro As String, Municipio As String, Estado As String, _
                 CEP As String, Celular As String, Telefone As String, _
                 DTNascimento As String, DTBatismo As String, Cpf As String, _
                 Indentidade As String, OrgaoEmissor As String, DTEmisao As String, _
                 M_Obs As String, DTConversao As String, Profissao As String, Situacao As String, Funcao_Eclesiastica As String)
                 


Dim Bd  As New ADODB.Connection
On Error GoTo Erro
Dim Nascimento As Date
Nascimento = IIf(DTNascimento = "", Empty, DTNascimento)

 If Localizar(BdNomeCompleto, Matricula).EOF Then
    ' Insere Novo Usuario
    Sql = "INSERT INTO membros (ID_Matr, C_nome, C_End, C_Comp, "
    Sql = Sql & "C_Numero, C_Bairro, C_Municipio, C_Estado, C_CEP, "
    Sql = Sql & "C_Celular, C_Tel, DT_Nascimento, DT_Batismo, C_CPF, "
    Sql = Sql & "C_Int, C_Orgao, DT_Emissao,DT_Conversao, M_OBS,C_Profissao,C_FuncaoEcle,C_Situacao)"
    
    
    Sql = Sql & " VALUES (" & Matricula & ",'" & Nome & "','" & Endereco & "','" & Complemento
    Sql = Sql & "','" & Numero & "','" & Bairro & "','" & Municipio & "','" & Estado
    Sql = Sql & "','" & CEP & "','" & Celular & "','" & Telefone & "','" & Nascimento
    Sql = Sql & "','" & DTBatismo & "','" & Cpf & "','" & Indentidade
    Sql = Sql & "','" & OrgaoEmissor & "','" & DTEmisao & "','" & DTConversao & "','" & M_Obs & "','" & Profissao & "','" & Funcao_Eclesiastica & "','" & Situacao & "')"
    
     
    
 Else
    ' Altera Cooperado
    
    Sql = "Update membros Set C_Nome = '" & Nome & "',C_End = '" & Endereco & "',"
    Sql = Sql & " C_Bairro= '" & Bairro & "',C_Municipio='" & Municipio
    Sql = Sql & "', C_Estado = '" & Estado & "',C_CEP = '" & CEP & "',C_Tel ='" & Telefone & "'"
    Sql = Sql & ",C_Celular='" & Celular & "',DT_Nascimento = '" & Nascimento & "',C_Cpf = '" & Cpf
    Sql = Sql & "',C_Int='" & Indentidade & "',C_Orgao ='" & OrgaoEmissor
    Sql = Sql & "',M_OBS='" & M_Obs & "', C_comp ='" & Complemento & "', C_numero= '" & Numero
    Sql = Sql & "',dt_Batismo='" & DTBatismo & "', dt_emissao='" & DTEmisao & "'"
    Sql = Sql & ", DT_conversao='" & DTConversao & "', C_Profissao='" & Profissao & "'"
    Sql = Sql & ", C_FuncaoEcle='" & Funcao_Eclesiastica & "', C_Situacao='" & Situacao & "'"
    Sql = Sql & " Where  ID_Matr=" & Matricula

 End If
 'Abre uma conexão
 Bd.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & BdNomeCompleto
    Bd.Execute Sql
    
 Set Bd = Nothing
 
   Exit Sub
Erro:
    Call ErrorMessage(Err)
End Sub
Private Sub ErrorMessage(mError As Variant)
'Objetivo: Rotina de tratamento e exibição de msgs de erro
   
   Dim Msgstr As String
   Msgstr = ""
   
   Select Case mError
     Case 3008
       Msgstr = "Você tentou abrir uma tabela que já foi aberta para uso exclusivo por outro usuário"
     Case 3009 Or 3211
       Msgstr = "Você tentou bloquear uma tabela que está sendo usada por outro usuário"
     Case 3046 Or 3158
       Msgstr = "Você tentou gravar um registro bloqueado por outro usuário"
     Case 3056
       Msgstr = "Restaure o seu último backup pois não foi possível reparar o Banco de Dados"
     Case 3187
       Msgstr = "Você tentou ler um registro bloqueado por outro usuário"
     Case 3188
       Msgstr = "Você tentou atualizar um registro que outro programa no seu equipamento já bloqueou"
     Case 3189
       Msgstr = "Você tentou obter acesso a uma tabela sobre a qual outro usuário impôs um bloqueio exclusivo"
     Case 3218
       Msgstr = "Você tentou atualizar um registro bloqueado"
     Case 3260
       Msgstr = "Você tentou gravar um registro bloqueado por outro usuário"
     Case 3049
    '   Call RepairDb(dBName)
       Resume
     Case 3420
       'tentou executar ação sobre algum objeto
       'que não estava aberto.
     Case Else
       Msgstr = "Erro número: " & mError & Chr(10) & mError.Description
   End Select
   
   If Msgstr <> "" Then Msgstr = MsgBox(Msgstr, vbExclamation, "Problemas com o Banco de Dados")
   
End Sub

Public Function MoveAnterior(BdNomeCompleto As String, pChave As String) As Recordset
   Dim Rs  As New ADODB.Recordset
    Dim Bd  As New ADODB.Connection
    On Error GoTo Erro
    Bd.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & BdNomeCompleto
    Rs.Open "membros", Bd, adOpenDynamic
    Rs.Find "ID_Matr=" & pChave
   With Rs
     If Not .EOF Then
       .MovePrevious
       If .BOF Then
       .MoveFirst
       End If
     End If
     Set MoveAnterior = Rs
   End With
  Set Bd = Nothing
    Set Rs = Nothing
   Exit Function
Erro:
   Call ErrorMessage(Err)

End Function
Public Function MoveSeguinte(BdNomeCompleto As String, pChave As String) As Recordset
    Dim Rs  As New ADODB.Recordset
    Dim Bd  As New ADODB.Connection
    On Error GoTo Erro
    Bd.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & BdNomeCompleto
    Rs.Open "membros", Bd, adOpenDynamic
    Rs.Find "ID_Matr=" & pChave
   With Rs
     If Not .EOF Then
       .MoveNext
       If .EOF Then
         .MoveLast
       End If
       Set MoveSeguinte = Rs
     End If
   End With
   Set Bd = Nothing
    Set Rs = Nothing
   Exit Function
Erro:
   Call ErrorMessage(Err)
End Function
Public Function Primeiro(BdNomeCompleto As String) As Recordset
    Dim Rs  As New ADODB.Recordset
    Dim Bd  As New ADODB.Connection
    On Error GoTo Erro
    Bd.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & BdNomeCompleto
    Rs.Open "membros", Bd
   With Rs
     If Not .BOF Then
        .MoveFirst
        If Not .BOF Then
          Set Primeiro = Rs
        End If
     End If
   End With
    Set Bd = Nothing
    Set Rs = Nothing
    Exit Function
Erro:
    Call ErrorMessage(Err)
End Function
Public Function Ultimo(BdNomeCompleto As String) As Recordset
    Dim Rs  As New ADODB.Recordset
    Dim Bd  As New ADODB.Connection
    On Error GoTo Erro
    Bd.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & BdNomeCompleto
    Rs.Open "membros", Bd, adOpenDynamic
   With Rs
     If Not .EOF Then
     
       .MoveLast
       If Not .EOF Then
          Set Ultimo = Rs
       End If
     End If
   End With
   Set Bd = Nothing
    Set Rs = Nothing
   Exit Function
Erro:
   Call ErrorMessage(Err)
End Function
