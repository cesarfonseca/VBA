Attribute VB_Name = "Módulo1"
Option Explicit
Public Cn As ADODB.Connection
Public Rs As ADODB.Recordset
Public Inicio As String
Public Final As String
Public Banco As String
Public Server As String
Private Function dCr(T As Variant) As Variant ' Mesclado T & S
Dim i As Long, k As Byte
Const Ass = "##CrptbyCF1##"
On Error GoTo Impar
For i = 1 To 2 * Len(Ass) Step 2
dCr = dCr & Chr(Asc(Mid(T, i, 1)) Xor Asc(Mid(T, 1 + i, 1)))
Next i
If dCr <> Ass Then GoTo Impar
dCr = ""
For i = 1 + 2 * Len(Ass) To Len(T) Step 2
dCr = dCr & Chr(Asc(Mid(T, i, 1)) Xor Asc(Mid(T, 1 + i, 1)))
Next i
Exit Function
Impar:
dCr = ""
T = Ass & T
Randomize
For i = 1 To Len(T)
Repete:
k = Int(64 * Rnd()) + 192
If Asc(Mid(T, i, 1)) = k Then GoTo Repete
dCr = dCr & Chr(Asc(Mid(T, i, 1)) Xor k) & Chr(k)
Next i
End Function
Public Function AbreConexao() 'Identifica o servidor e o banco e abre a conexão
Origem = Environ$("Logonserver")
Select Case Origem
Case "\\JURERE"
Banco = "Focus": Server = "192.168.83.14"
Case Else
Banco = "Focus": Server = "Anatelbdro01"
End Select
Set Cn = New ADODB.Connection
Set Rs = New ADODB.Recordset
Cn.ConnectionTimeout = 15   'Espera 15 segundos pela conexão.
Cn.CommandTimeout = 300     'Espera no máximo 5 minutos pela consulta
Cn.Open "Provider=SQLOLEDB.1;Password=s*a*r*u*l*o*c*a*l;User ID=userSaruLocal;Initial Catalog= " & Banco & " ;Data Source=" & Server
Rs.CursorType = adOpenKeyset
End Function
Public Function FechaConexao()
On Error Resume Next
Rs.Close
Set Rs = Nothing
Cn.Close
Set Cn = Nothing
End Function
