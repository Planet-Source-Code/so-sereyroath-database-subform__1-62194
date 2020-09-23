VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl SITSSubForm 
   BackStyle       =   0  'Transparent
   ClientHeight    =   3060
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6075
   PropertyPages   =   "UserControl1.ctx":0000
   ScaleHeight     =   3060
   ScaleWidth      =   6075
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   1005
      Index           =   0
      Left            =   1440
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Left            =   2520
      Picture         =   "UserControl1.ctx":002D
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Left            =   4080
      TabIndex        =   3
      Top             =   2280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3600
      TabIndex        =   1
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   5318
      _Version        =   393216
      BackColor       =   16777215
      BackColorFixed  =   12632256
      BackColorBkg    =   12632256
      ScrollTrack     =   -1  'True
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   0
      Left            =   1320
      TabIndex        =   2
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Menu Mnudel 
      Caption         =   "Delete"
      Visible         =   0   'False
      Begin VB.Menu Mnuremove 
         Caption         =   "Remove"
      End
   End
End
Attribute VB_Name = "SITSSubForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'// Write By So Sereyroath
'// For Using In Smart IT Solution
'// Any using without authorize will be illegal
'// ©Copyright by So Sereyroath
'// Any Comment sent mail to: roathvb@yahoo.com
'// 09/07/2005
'Private WithEvents lst1 As ListBox
Public Enum Sorts
    SitsNoSorted = 0
    SitsSortGenAsc = 1
    SitsSortGenDesc = 2
    SitsSortNumAsc = 3
    SitsSortNumDesc = 4
    SitsSortStrNoCaseAsc = 5
    SitsSortStrNoCaseDesc = 6
    SitsSortStringAscending = 7
    SitsSortStringDescending = 8
End Enum
Dim pformat As String
Dim sorta As Sorts
Public Enum celltype
    sitstext = 0
    Sitscombo = 1
    SitsCheck = 2
    Sitsdate = 3
End Enum
'Public WithEvents Lst As ListBox
'Dim list1() As ListBox
Dim oldvalue() As String
Dim newvalue() As String
Dim p As Integer
Dim rowchange As Boolean
Dim oldrow As Integer, oldval As String, q As Boolean
Public Event RecordChanged(oldrec() As String, newrec() As String, ByVal row As Integer)
Attribute RecordChanged.VB_Description = "Process when user change value row or col"
Public Event BeforeDelete(cancel As Integer, ByVal row As Integer)
Attribute BeforeDelete.VB_Description = "This event will process after user pressed delete key and before the item will remove"
Public Event AfterDelete()
Public Event RowColChange()
Attribute RowColChange.VB_Description = "Process when user change row or col"
Public Event ButtonClick(ByVal col As Long, ByVal left As Long, ByVal top As Long)
Attribute ButtonClick.VB_Description = "Execute when user click on button"
Dim VRows As Double, VCols As Double
'MSFlexGrid1.MergeCells


Public Type coltype
    col As Long
    combo As Byte
    duplicate As Boolean
End Type

Public Type coltypev
    col As Long
    combo As Byte
    comboindex As Integer
    duplicate As Boolean
End Type

Public Enum SfMerge
    SfMergeNever = 0
    SfMergeFree = 1
    SfMergeRestrictRows = 2
    SfMergeRestrictColumns = 3
    SfMergeRestrictBoth = 4
End Enum

Public Enum SfAlignment
    SfAlignLeftTop = 0
    SfAlignLeftCenter = 1
    SfAlignLeftBottom = 2
    SfAlignCenterTop = 3
    SfAlignCenterCenter = 4
    SfAlignCenterBottom = 5
    SfAlignRightTop = 6
    SfAlignRightCenter = 7
    SfAlignRightBottom = 8
    SfAlignGeneral = 9
End Enum
Dim VcAlignment As SfAlignment
Dim Vcoltype() As coltypev
'Property Variables:


Private Sub Check1_Click()
    MSFlexGrid1.text = IIf(Check1.Value = 1, Chr(254), Chr(110))
End Sub

'Private Sub Combo1_Click(Index As Integer)
'    With MSFlexGrid1
'        If .row = .Rows - 1 Then
'            .text = Combo1(Index).text
'            .Rows = .Rows + 1
'        End If
'    End With
'
'End Sub

Private Sub Command1_Click()
    For i = List1.LBound To List1.UBound
        List1(i).Visible = False
    Next
With MSFlexGrid1
    If Vcoltype(.col - 1).combo = Sitscombo Then
        List1(Vcoltype(MSFlexGrid1.col - 1).comboindex).Visible = True
    

            List1(Vcoltype(.col - 1).comboindex).Move .CellLeft + .left, .CellTop + .top + .CellHeight, .CellWidth
            List1(Vcoltype(.col - 1).comboindex).ZOrder 0
            MSFlexGrid1.Height = UserControl.Height - List1(Vcoltype(.col - 1).comboindex).Height
            RaiseEvent ButtonClick(.col, .left + .CellLeft, .top + .CellTop + .CellHeight)
 '       End With
    Else
            RaiseEvent ButtonClick(.col, .left + .CellLeft, .top + .CellTop + .CellHeight)
  
    End If
End With
End Sub




Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        With MSFlexGrid1
            If .row < .Rows - 1 Then
                'RaiseEvent BeforeDelete(cancel, .row)
                Mnuremove_Click
            End If
        End With
    End If
End Sub

Private Sub list1_click(Index As Integer)
Dim i As Integer, jj As Integer
With MSFlexGrid1
 
  If Vcoltype(.col - 1).duplicate = True Then
    For i = 1 To .Rows - 1
        If .row <> i And LCase(List1(Index).text) = LCase(.TextMatrix(i, 1)) Then
                MsgBox "This item was already exits, Please try another 1"
                Exit Sub
        End If
     
    Next
  End If
 ' Static oldrow As Double
  If oldrow = 0 Then oldrow = 1
  If MSFlexGrid1.row <> MSFlexGrid1.Rows - 1 Then
    oldval = MSFlexGrid1.text
    'MSFlexGrid1.text = Text1.text
    .text = List1(Index).text
    If .text <> oldval Then rowchange = True
    If p = 0 And rowchange = True Then
       p = 1
      If MSFlexGrid1.row <> oldrow And oldrow <> 0 Then MSFlexGrid1_RowColChange
    End If
  Else
    
  End If
  '
  .text = List1(Index).text
  List1(Index).Visible = False
  If .row = .Rows - 1 Then .Rows = .Rows + 1
  oldrow = .row
End With
End Sub



Private Sub lst_Click()
        MsgBox "Hello"
End Sub

Private Sub Mnuremove_Click()
Dim cancel As Integer
    With MSFlexGrid1
        If .Rows > 2 Then
           RaiseEvent BeforeDelete(cancel, .row)
           If cancel = 1 Then Exit Sub
           .RemoveItem .row
           RaiseEvent AfterDelete
        End If
    End With
End Sub

Private Sub MSFlexGrid1_Click()

  If Vcoltype(MSFlexGrid1.col - 1).combo = SitsCheck Then
        MSFlexGrid1.text = IIf(Asc(IIf(MSFlexGrid1.text = "", 0, MSFlexGrid1.text)) <> "254", Chr(254), Chr(111))
  End If
    'loadcontrol
    MSFlexGrid1_RowColChange
    
End Sub

'Private Sub MSFlexGrid1_GotFocus()
    'MSFlexGrid1_RowColChange
'End Sub



Private Sub MSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton And MSFlexGrid1.row > 0 Then
        PopupMenu Mnudel
    End If
End Sub

Private Sub MSFlexGrid1_RowColChange()
Static oldcol As Long
Dim s As Boolean
'Static r As Boolean

    With MSFlexGrid1
    If oldrow <> .row Then
      On Error GoTo ss
        If UBound(oldvalue) > 0 And UBound(newvalue) > 0 Then
            For i = 1 To .Cols - 1
                'newvalue(i) = .TextMatrix(oldrow, i)
                'If oldvalue(i) <> newvalue(i) And oldvalue(i) <> "" Then s = True: Exit For
                If oldvalue(i) <> .TextMatrix(oldrow, i) And oldvalue(i) <> "" Then s = True: Exit For
            Next
            If s = True Then RaiseEvent RecordChanged(oldvalue, newvalue, oldrow)
            s = False
        End If
ss:
        ReDim oldvalue(.Cols - 1)
        For i = 1 To .Cols - 1
            oldvalue(i) = .TextMatrix(.row, i)
        Next
        oldrow = .row
        oldcol = .col
        
    Else
        If oldcol <> .col Then
            ReDim newvalue(.Cols - 1)
            For i = 1 To .Cols - 1
                newvalue(i) = .TextMatrix(.row, i)
            'If oldvalue(i) <> newvalue(i) Then s = True
            Next
        End If
    End If
          
         If Vcoltype(.col - 1).combo = Sitscombo Then
            Command1.Move .CellLeft + .left + (.CellWidth - 255), .CellTop + .top, 255, .CellHeight + Screen.TwipsPerPixelY
            Command1.Visible = True
            Command1.SetFocus
            Command1.ZOrder 0
            For i = List1.LBound To List1.UBound
                 List1(i).Visible = False
            Next
            Text1.Visible = False
            Check1.Visible = False
         ElseIf Vcoltype(.col - 1).combo = SitsCheck Then
            MSFlexGrid1.CellAlignment = 4
            MSFlexGrid1.CellFontName = "Wingdings"
            'Check1.Visible = True
            'Check1.Move .CellLeft + .Left + .CellWidth / 2, .CellTop + .Top
            'Check1.SetFocus
            'Check1.Value = Abs(1 - Val(.text))
            For i = List1.LBound To List1.UBound
                  List1(i).Visible = False
            Next
            Text1.Visible = False
            Command1.Visible = False
         ElseIf Vcoltype(.col - 1).combo = sitstext Then
            Text1.Move .CellLeft + .left, .CellTop + .top, .CellWidth, .CellHeight
            Text1.Visible = True
            Text1.SetFocus
            Text1.text = .text
             For i = List1.LBound To List1.UBound
                List1(i).Visible = False
            Next
            Check1.Visible = False
            Command1.Visible = False
         ElseIf Vcoltype(.col - 1).combo = Sitsdate Then
            Command1.Move .CellLeft + .left + (.CellWidth - 255), .CellTop + .top, 255, .CellHeight + Screen.TwipsPerPixelY
            Command1.Visible = True
            Command1.SetFocus
            Command1.ZOrder 0
            For i = List1.LBound To List1.UBound
                 List1(i).Visible = False
            Next
            Text1.Visible = False
            Check1.Visible = False
         End If
    End With
    RaiseEvent RowColChange
End Sub

Private Sub MSFlexGrid1_Scroll()
    For i = LBound(Vcoltype) To UBound(Vcoltype)
        If Vcoltype(i).combo = 1 Then List1(Vcoltype(i).comboindex).Visible = False
    Next
    Text1.Visible = False
    Check1.Visible = False
    Command1.Visible = False
End Sub

Private Sub Text1_change()
 If Text1.text <> "" Then
    With MSFlexGrid1
    If .row = .Rows - 1 Then .Rows = .Rows + 1
    .text = Text1.text
    End With
 End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
     rowchange = True
End Sub

Public Property Get Rows() As Long
Attribute Rows.VB_Description = "Determines the total number of columns or rows in a subform"
    Rows = MSFlexGrid1.Rows
End Property

Public Property Let Rows(ByVal vNewValue As Long)
    If vNewValue > CLng(70000) Then
        MsgBox "Value can not longer than 70000", vbCritical
        Exit Property
    End If
        MSFlexGrid1.Rows = vNewValue
        PropertyChanged "Rows"
End Property

Public Property Get Cols() As Double
Attribute Cols.VB_Description = "Determines the total number of columns in a subform"
    Cols = MSFlexGrid1.Cols
End Property

Public Property Let Cols(ByVal vNewValue As Double)
    If vNewValue > 100000 Then
        MsgBox "Value can not longer than 100000", vbCritical
        Exit Property
    End If
        MSFlexGrid1.Cols = vNewValue
        PropertyChanged "Cols"
End Property



Public Function GetText(Optional ByVal row As Long = -1, Optional ByVal col As Long = -1) As String
Attribute GetText.VB_Description = "Read text from a specified cell"
On Error Resume Next
    With MSFlexGrid1
        If row = -1 Then row = .row
        If col = -1 Then col = .col
        GetText = .TextMatrix(row, col)
        If Err.Number > 0 Then GetText = ""
    End With
End Function

Public Sub PutText(ByVal text As String, Optional ByVal row As Long = -1, Optional ByVal col As Long = -1, Optional ByVal addnew As Boolean = True)
Attribute PutText.VB_Description = "Write text to a specified cell"
On Error Resume Next
    With MSFlexGrid1
        If row = -1 Then row = .row
        If col = -1 Then col = .col
        .TextMatrix(row, col) = text
         If addnew = True Then .Rows = .Rows + 1
    '    If Err.Number > 0 Then GetText = ""
    End With
End Sub

Public Property Get CellAlignment() As SfAlignment
Attribute CellAlignment.VB_Description = "Specified the alignment for cell"
Attribute CellAlignment.VB_MemberFlags = "400"
     CellAlignment = MSFlexGrid1.CellAlignment
    
End Property

Public Property Let CellAlignment(ByVal vNewValue As SfAlignment)
    MSFlexGrid1.CellAlignment = vNewValue
    PropertyChanged "CellAlignment"
End Property

Public Property Let Cellbackcolor(ByVal vNewValue As Variant)
Attribute Cellbackcolor.VB_Description = "Specified the cellbackcolor"
Attribute Cellbackcolor.VB_MemberFlags = "400"
    MSFlexGrid1.Cellbackcolor = vNewValue
End Property

Public Property Let CellFontName(ByVal vNewValue As String)
Attribute CellFontName.VB_Description = "Change font for current cell"
Attribute CellFontName.VB_MemberFlags = "400"
    MSFlexGrid1.CellFontName = vNewValue
    PropertyChanged "CellFontName"
End Property

Public Property Let CellFontSize(ByVal vNewValue As Byte)
Attribute CellFontSize.VB_Description = "Change font size for cell"
Attribute CellFontSize.VB_MemberFlags = "400"
    MSFlexGrid1.CellFontSize = vNewValue
End Property

Public Property Get CellLeft() As Long
Attribute CellLeft.VB_Description = "Specified cell left"
Attribute CellLeft.VB_MemberFlags = "400"
    CellLeft = MSFlexGrid1.CellLeft
    PropertyChanged "CellLeft"
End Property

Public Property Get CellTop() As Long
Attribute CellTop.VB_Description = "Specified cell top position"
Attribute CellTop.VB_MemberFlags = "400"
    CellTop = MSFlexGrid1.CellTop
End Property

Public Property Get CellWidth() As Long
Attribute CellWidth.VB_Description = "specified the cell width"
Attribute CellWidth.VB_MemberFlags = "400"
    CellWidth = MSFlexGrid1.CellWidth
    
End Property

Public Property Get CellHeight() As Long
Attribute CellHeight.VB_Description = "Specified the cell height"
Attribute CellHeight.VB_MemberFlags = "400"
    CellHeight = MSFlexGrid1.CellHeight
    
End Property




Public Property Get BackColorBkg() As OLE_COLOR
Attribute BackColorBkg.VB_Description = "Returns/sets the background color of various elements of the subform"
Attribute BackColorBkg.VB_ProcData.VB_Invoke_Property = "StandardColor;Appearance"
    BackColorBkg = MSFlexGrid1.BackColorBkg
End Property

Public Property Let BackColorBkg(ByVal vNewValue As OLE_COLOR)
    MSFlexGrid1.BackColorBkg = vNewValue
    PropertyChanged "BackColorBkg"
End Property


Public Property Get BackColorSel() As OLE_COLOR
Attribute BackColorSel.VB_Description = "Returns/set the back  color of fixed column of various element of the subform"
Attribute BackColorSel.VB_ProcData.VB_Invoke_Property = "StandardColor;Appearance"
Attribute BackColorSel.VB_UserMemId = 0
    BackColorSel = MSFlexGrid1.BackColorSel
End Property

Public Property Let BackColorSel(ByVal vNewValue As OLE_COLOR)
    MSFlexGrid1.BackColorSel = vNewValue
    PropertyChanged "BackColorSel"
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=MSFlexGrid1,MSFlexGrid1,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returens/sets the background color of various elements of the Subform"
Attribute BackColor.VB_ProcData.VB_Invoke_Property = "StandardColor;Appearance"
Attribute BackColor.VB_UserMemId = -501
    BackColor = MSFlexGrid1.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    MSFlexGrid1.BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property



'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  

    MSFlexGrid1.BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
    MSFlexGrid1.BackColorSel = PropBag.ReadProperty("BackColorSel", &H8000000D)
    MSFlexGrid1.BackColorBkg = PropBag.ReadProperty("BackColorBkg", &H808080)
    MSFlexGrid1.BackColorFixed = PropBag.ReadProperty("BackColorFixed", &HC0C0C0)
    MSFlexGrid1.FormatString = PropBag.ReadProperty("Caption", "")
    MSFlexGrid1.MergeCells = PropBag.ReadProperty("MergeCell", SfMergeNever)
    'Set MSFlexGrid1.Font = PropBag.ReadProperty("BackColor", Ambient.Font)
    Set MSFlexGrid1.Font = PropBag.ReadProperty("Font", Ambient.Font)
    MSFlexGrid1.MergeCells = PropBag.ReadProperty("MergeCell", SfMergeNever)
    MSFlexGrid1.Rows = PropBag.ReadProperty("Rows", 2)
    MSFlexGrid1.Cols = PropBag.ReadProperty("Cols", 2)
    MSFlexGrid1.Sort = PropBag.ReadProperty("Sorted", 0)
    pformat = PropBag.ReadProperty("Format", "")
    If Ambient.UserMode Then
        If pformat <> "" Then
            Format = pformat
        End If
        
    End If
End Sub

Private Sub UserControl_Resize()
    MSFlexGrid1.Width = UserControl.Width
    MSFlexGrid1.Height = UserControl.Height
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", MSFlexGrid1.BackColor, &HFFFFFF)
    Call PropBag.WriteProperty("BackColorSel", MSFlexGrid1.BackColorSel, &H8000000D)
    Call PropBag.WriteProperty("BackColorBkg", MSFlexGrid1.BackColorBkg, &H808080)
    Call PropBag.WriteProperty("BackColorFixed", MSFlexGrid1.BackColorFixed, &HC0C0C0)
    Call PropBag.WriteProperty("Caption", MSFlexGrid1.FormatString, "")
    Call PropBag.WriteProperty("MergeCell", MSFlexGrid1.MergeCells, SfMergeNever)
    'Call PropBag.WriteProperty("BackColor", MSFlexGrid1.Font, Ambient.Font)
    Call PropBag.WriteProperty("Font", MSFlexGrid1.Font, Ambient.Font)
    Call PropBag.WriteProperty("MergeCell", MSFlexGrid1.MergeCells, SfMergeNever)
    Call PropBag.WriteProperty("Rows", MSFlexGrid1.Rows, 2)
    Call PropBag.WriteProperty("Cols", MSFlexGrid1.Cols, 2)
    Call PropBag.WriteProperty("Format", pformat, "")
    'Call PropBag.WriteProperty("Sorted", MSFlexGrid1.Sort, 0)
End Sub

Public Property Get BackColorFixed() As OLE_COLOR
Attribute BackColorFixed.VB_Description = "Returns/set the back  color of fixed column of various element of the subform"
Attribute BackColorFixed.VB_ProcData.VB_Invoke_Property = "StandardColor;Appearance"
Attribute BackColorFixed.VB_UserMemId = -520
    BackColorFixed = MSFlexGrid1.BackColorFixed
End Property

Public Property Let BackColorFixed(ByVal New_BackColor As OLE_COLOR)
    MSFlexGrid1.BackColorFixed = New_BackColor
    PropertyChanged "BackColorFixed"
End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Specified the column name for subform"
Attribute Caption.VB_ProcData.VB_Invoke_Property = "StandardDataFormat;Text"
Attribute Caption.VB_UserMemId = -518
Attribute Caption.VB_MemberFlags = "200"
    Caption = MSFlexGrid1.FormatString
End Property

Public Property Let Caption(ByVal New_BackColor As String)
    MSFlexGrid1.FormatString = New_BackColor
    PropertyChanged "Caption"
End Property

Public Property Get MergeCell() As SfMerge
Attribute MergeCell.VB_Description = "Merge cell "
Attribute MergeCell.VB_MemberFlags = "400"
    MergeCell = MSFlexGrid1.MergeCells
End Property

Public Property Let MergeCell(ByVal New_BackColor As SfMerge)
    MSFlexGrid1.MergeCells = New_BackColor
    PropertyChanged "MergeCell"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=6,0,0,0
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = MSFlexGrid1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set MSFlexGrid1.Font = New_Font
    Set Text1.Font = New_Font
    Text1.Font.Size = IIf(New_Font.Size >= 14, New_Font.Size - 2, New_Font.Size)
    PropertyChanged "Font"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    Set MSFlexGrid1.Font = Ambient.Font
End Sub

Private Property Let Celltypes(vNewValue() As coltype)
Dim mv As MonthView, lst As ListBox
    ReDim Vcoltype(UBound(vNewValue))
    For i = LBound(vNewValue) To UBound(vNewValue)
         Vcoltype(i).col = vNewValue(i).col
         Vcoltype(i).combo = vNewValue(i).combo
         Vcoltype(i).duplicate = vNewValue(i).duplicate
         If Vcoltype(i).combo = 1 Then
            'Dim cmb As ListBox
            On Error Resume Next
            's = List1.UBound + 1
            'If Err.Number > 0 Then s = 0
            'MSFlexGrid1.ScrollBars = flexScrollBarHorizontal
            'ReDim Preserve List1(s)
            'loadcontrol
            Vcoltype(i).comboindex = List1.UBound
            Set lst = List1(List1.UBound + 1)
            Load lst
            lst.Visible = False
         End If
    Next
End Property


Public Sub AssignCombo(ByVal col As Long, rst As Variant)
  If Vcoltype(col).combo = 1 Then
    List1(Vcoltype(col).comboindex).clear
    Do While Not rst.EOF
        List1(Vcoltype(col).comboindex).AddItem rst(0)
        rst.movenext
    Loop
    'List1(Vcoltype(col).comboindex).AddItem " "
  End If
End Sub

'
'Sub loadcontrol()
'      Dim cform As Form
''Dim li As ListBox
'   If Ambient.UserMode Then
'    'Set mParent = UserControl.Parent
'    Set cform = UserControl.Parent
'    Set Lists = cform.Controls.Add("Vb.Listbox", "List1")
'
'    Set list1(list1.Ubound) = Lists
'        list1(list1.Ubound).Appearance = 0
'    'Set Lists = Nothing
'   End If
'End Sub



Public Property Get Sorted() As Sorts
Attribute Sorted.VB_Description = "Sort data in subform"
Attribute Sorted.VB_MemberFlags = "400"
     Sorted = sorta
End Property

Public Property Let Sorted(ByVal vNewValue As Sorts)
    MSFlexGrid1.Sort = vNewValue
    PropertyChanged "Sorted"
End Property

Public Sub RemoveItem(ByVal row As Long)
Attribute RemoveItem.VB_Description = "Delete item from a specified row"
    If row < MSFlexGrid1.Rows Then
        MSFlexGrid1.RemoveItem row
    End If
End Sub

Public Property Get Format() As String
Attribute Format.VB_Description = "Returns/set the subform cell format"
Attribute Format.VB_MemberFlags = "40"
    Format = pformat
End Property

Public Property Let Format(ByVal formats As String)
  Dim s() As String, c() As coltype, t() As String
  On Error GoTo x
    pformat = formats
    
    s = Split(formats, Chr(1))
    
    ReDim c(UBound(s))
    For i = 0 To UBound(s)
        t = Split(s(i), Chr(2))
        c(i).col = t(0)
        c(i).combo = t(1)
        c(i).duplicate = t(2)
    Next
    
    Celltypes = c
    PropertyChanged "Format"
    Exit Property
x:
    PropertyChanged "Format"
    MsgBox "Invalid format"
End Property



Public Sub About()
Attribute About.VB_Description = "Design By So Sereyroath"
Attribute About.VB_UserMemId = -552
    MsgBox "Design By So Sereyroath, Call to 012908812 now for any information."
End Sub
