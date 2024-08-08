VERSION 5.00
Begin VB.Form CalculatorGUI 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calculator"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   3030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton ButtonDelete 
      Appearance      =   0  'Flat
      Caption         =   "C"
      Height          =   615
      Left            =   840
      TabIndex        =   16
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton ButtonResult 
      Appearance      =   0  'Flat
      Caption         =   "="
      Height          =   615
      Left            =   1560
      TabIndex        =   15
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton ButtonDivide 
      Caption         =   "/"
      Height          =   615
      Left            =   2280
      TabIndex        =   14
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton ButtonMultiply 
      Caption         =   "*"
      Height          =   615
      Left            =   2280
      TabIndex        =   13
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton ButtonSubtract 
      Caption         =   "-"
      Height          =   615
      Left            =   2280
      TabIndex        =   12
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton ButtonSum 
      Caption         =   "+"
      Height          =   615
      Left            =   2280
      TabIndex        =   11
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton Button0 
      Caption         =   "0"
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton Button9 
      Caption         =   "9"
      Height          =   615
      Left            =   1560
      TabIndex        =   9
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton Button8 
      Caption         =   "8"
      Height          =   615
      Left            =   840
      TabIndex        =   8
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton Button7 
      Caption         =   "7"
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton Button6 
      Caption         =   "6"
      Height          =   615
      Left            =   1560
      TabIndex        =   6
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton Button5 
      Caption         =   "5"
      Height          =   615
      Left            =   840
      TabIndex        =   5
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton Button4 
      Appearance      =   0  'Flat
      Caption         =   "4"
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton Button3 
      Appearance      =   0  'Flat
      Caption         =   "3"
      Height          =   615
      Left            =   1560
      TabIndex        =   3
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton Button2 
      Appearance      =   0  'Flat
      Caption         =   "2"
      Height          =   615
      Left            =   840
      TabIndex        =   2
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton Button1 
      Appearance      =   0  'Flat
      Caption         =   "1"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox TextResult 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "CalculatorGUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Button0_Click()
    TextResult.Text = TextResult.Text + "0"
End Sub

Private Sub Button1_Click()
    TextResult.Text = TextResult.Text + "1"
End Sub

Private Sub Button2_Click()
    TextResult.Text = TextResult.Text + "2"
End Sub

Private Sub Button3_Click()
    TextResult.Text = TextResult.Text + "3"
End Sub

Private Sub Button4_Click()
    TextResult.Text = TextResult.Text + "4"
End Sub

Private Sub Button5_Click()
    TextResult.Text = TextResult.Text + "5"
End Sub

Private Sub Button6_Click()
    TextResult.Text = TextResult.Text + "6"
End Sub

Private Sub Button7_Click()
    TextResult.Text = TextResult.Text + "7"
End Sub

Private Sub Button8_Click()
    TextResult.Text = TextResult.Text + "8"
End Sub

Private Sub Button9_Click()
    TextResult.Text = TextResult.Text + "9"
End Sub

Private Sub ButtonDelete_Click()
    If TextResult.Text <> "" Then
        TextResult.Text = Left$(TextResult.Text, Len(TextResult.Text) - 1)
    End If
End Sub

Private Sub ButtonDivide_Click()
    If TextResult.Text <> "" Then
        TextResult.Text = EnsureOperand(TextResult.Text) + "/"
    End If
End Sub

Private Sub ButtonMultiply_Click()
    If TextResult.Text <> "" Then
        TextResult.Text = EnsureOperand(TextResult.Text) + "*"
    End If
End Sub

Private Sub ButtonResult_Click()
    Dim Expression() As String
    Dim Value As String
    
    Value = TextResult.Text
    
    On Error Resume Next
    If InStr(Value, "+") > 0 Then
        Expression = Split(Value, "+")
        TextResult.Text = Sum(CInt(Expression(0)), CInt(Expression(1)))
    ElseIf InStr(Value, "-") > 0 Then
        Expression = Split(Value, "-")
        TextResult.Text = Subtract(CInt(Expression(0)), CInt(Expression(1)))
    ElseIf InStr(Value, "*") > 0 Then
        Expression = Split(Value, "*")
        TextResult.Text = Multiply(CInt(Expression(0)), CInt(Expression(1)))
    ElseIf InStr(Value, "/") > 0 Then
        Expression = Split(Value, "/")
        TextResult.Text = Divide(CInt(Expression(0)), CInt(Expression(1)))
    End If
End Sub

Private Sub ButtonSubtract_Click()
    If TextResult.Text <> "" Then
        TextResult.Text = EnsureOperand(TextResult.Text) + "-"
    End If
End Sub

Private Sub ButtonSum_Click()
    If TextResult.Text <> "" Then
        TextResult.Text = EnsureOperand(TextResult.Text) + "+"
    End If
End Sub
