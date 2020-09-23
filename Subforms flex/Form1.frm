VERSION 5.00
Object = "{5EE7D624-A906-4612-98E5-6552BD25E927}#2.0#0"; "SITSSubForm2.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   ScaleHeight     =   6060
   ScaleWidth      =   7350
   StartUpPosition =   3  'Windows Default
   Begin SITSSubForms.SITSSubForm SITSSubForm1 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   6165
      BackColorBkg    =   12632256
      Caption         =   "    | Name                   |^Sex             | Phone"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Cols            =   4
      Format          =   $"Form1.frx":0000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------- Read before used -----------
'  1/befor you use it, declared 1 array dynamic variable as follows,
' Dim c() As SITSSubForms.coltype
'add information to this array 1 by 1
' like the following code
'This code is not for sale
'For any comment please send mail to me roathvb@yahoo.com

Private Sub Form_Load()
Dim cnn As New ADODB.Connection
Dim rst As New ADODB.Recordset
Dim c() As SITSSubForms.coltype
    'Set SITSSubForm1.Container = Picture1
    SITSSubForm1.Cols = 4
    cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data source=c:\db1.mdb"
    'number of column in your subform
    rst.ActiveConnection = cnn


    'SITSSubForm1.PutText "fdsfs" 'use to put text into cell
    'MsgBox SITSSubForm1.GetText(1, 1) 'get text from a specified cell

    'put index of listbox you want to assign and recordset you want to retrieved from
    rst.Open "Select Name from student"
    SITSSubForm1.AssignCombo 0, rst
'SITSSubForm1
         
End Sub





Private Sub SITSSubForm1_BeforeDelete(cancel As Integer, ByVal row As Integer)
    'cancel = 1
    If MsgBox("Are you sure? you want to delete this record?", vbYesNo) = vbNo Then
        cancel = 1
    End If
End Sub



Private Sub SITSSubForm1_ButtonClick(ByVal col As Long, ByVal left As Long, ByVal top As Long)
  If col = 3 Then
    MonthView1.left = left
    MonthView1.top = top + SITSSubForm1.top
    MonthView1.Visible = True
  End If
End Sub

Private Sub SITSSubForm1_RecordChanged(oldrec() As String, newrec() As String, ByVal row As Integer)
    MsgBox oldrec(1)
End Sub

Private Sub SITSSubForm1_RowColChange()
'    MonthView1.Visible = False
End Sub
