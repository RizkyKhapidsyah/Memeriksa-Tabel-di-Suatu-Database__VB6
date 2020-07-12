VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Memeriksa Tabel di Suatu Database"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   5175
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   1680
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function TableExists(DatabaseName$, _
TableName$) As Boolean
'DataBaseName adalah nama file database yang akan 'diperiksa apakah TableName ada di dalamnya
Dim oDB As Database, td As TableDef
On Error GoTo ErrorHandler
'Sesuaikan dengan cara membuka file database yang 'dipassword jika file database tersebut dipassword...
Set oDB = Workspaces(0).OpenDatabase(DatabaseName)
On Error Resume Next
  Set td = oDB.TableDefs(TableName)
  TableExists = Err.Number = 0
  oDB.Close
  Exit Function
ErrorHandler:
  Select Case Err.Number
         Case 3024
              MsgBox "Database tidak ada!", _
                     vbCritical, "Database Error"
              End
         Case Else
              MsgBox Err.Number & " - " & _
                     Err.Description
  End Select
  Exit Function
End Function

Private Sub Command1_Click()
'Ganti "Akademik.mdb" di bawah dengan nama database 'Anda dengan catatan masih terdapat dalam direktori 'yang sama dengan program ini berada.
DatabaseName$ = App.Path & "\Akademik.mdb"
'Ganti "Mahasiswa" dengan nama tabel yang ingin Anda 'periksa.
TableName$ = "Mahasiswa"
 Call TableExists(DatabaseName$, TableName$)
 If TableExists(DatabaseName$, TableName$) = True Then
    MsgBox "Tabel " & TableName$ & " ada!", _
            vbInformation, "Tabel Ada"
 Else
    MsgBox "Tabel " & TableName$ & " tidak ada!", _
           vbCritical, "Tidak Ada"
 End If
End Sub


