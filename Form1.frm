VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MSFlexGrid Move Column Example"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   5670
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid fg 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   4895
      _Version        =   393216
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
   End
   Begin VB.Label lblfgwidth 
      Caption         =   "Label3"
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Flex Grid Width"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblMousexy 
      Caption         =   "Label2"
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Mouse (x,y)"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim colDown As Long ' column when mouse went down
Dim colUp As Long   ' column when mouse came up
Dim ColLeftRight As Integer

Private Const IDC_ARROW = 32512&  ' normal (arrow) cursor
Private Const IDC_CROSS = 32516&  ' cross cursor

Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long

Private Sub cmdReset_Click()
 Init
End Sub

Private Sub fg_Click()
'
' Select all cells in a row
'
  With fg
    If .Row > 0 Then
      .Col = 0
      .ColSel = .Cols - 1
    End If
  End With
Exit Sub

End Sub

Private Sub fg_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'
' Track column where mouse went down for column moves
'
  colDown = FindColumn(x, y)
Exit Sub

End Sub

Private Sub fg_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'
' Set cross cursor when moving columns
'
  If colDown <> -1 Then
    SetCursor LoadCursor(0, IDC_CROSS)
  End If
  Me.lblMousexy = "(" & x & "," & y & ")"
  If x + 346 >= fg.Width Then
   If fg.LeftCol <> fg.Cols Then
    fg.LeftCol = fg.LeftCol + 1
    ColLeftRight = 0
   End If
  End If
  If x <= 345 Then
   If fg.LeftCol <> fg.Cols Then
    If fg.LeftCol <> 0 Then
     fg.LeftCol = fg.LeftCol - 1
    End If
    ColLeftRight = 1
   End If
  End If

Exit Sub

End Sub

Private Sub fg_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'
' If column move, exchange the columns
'
  Dim t As Variant
  
  colUp = FindColumn(x, y)
  If colDown <> -1 And colUp <> -1 Then
    If colDown <> colUp Then
      ' reverse columns
      With fg
        For y = 0 To .Rows - 1
          t = .TextMatrix(y, colDown)
          .TextMatrix(y, colDown) = .TextMatrix(y, colUp)
          .TextMatrix(y, colUp) = t
        Next y
      End With
    End If
  End If
  colDown = -1
  SetCursor LoadCursor(0, IDC_ARROW)
Exit Sub

End Sub

Private Function FindColumn(x As Single, y As Single)
'
' If x,y identifies a header column, return the column else -1
'
  Dim i, j As Long
  Dim cellX As Long
  
  
  With fg
    If .TopRow = 1 And y >= 0 And y < .RowHeight(0) Then 'in the header
      For i = .LeftCol To .Cols - 1
       If x >= cellX And x < cellX + .ColWidth(i) Then
         Select Case ColLeftRight
          Case 0
           FindColumn = i
           Exit Function
          Case 1
           FindColumn = i
           Exit Function
          End Select
        End If
        cellX = cellX + .ColWidth(i)
      Next i
    End If
  End With
  FindColumn = -1
  
Exit Function

End Function
Private Sub Form_Load()
 Init
End Sub
Private Sub Init()

  Dim y As Long
  Dim x As Long
  
  colDown = -1  ' init to 'none'
  fg.Clear
  '
  ' set some random data
  '
  With fg
    .Cols = 4
    .Rows = 51
    .TextMatrix(0, 0) = "One"
    .TextMatrix(0, 1) = "Two"
    .TextMatrix(0, 2) = "Three"
    .TextMatrix(0, 3) = "Four"
    For y = 1 To 50
      For x = 0 To 3
        .TextMatrix(y, x) = Rnd(1000)
      Next x
    Next y
  End With
  Me.lblfgwidth = fg.Width
 
  
Exit Sub

End Sub
