VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Rebecca's Grocery List"
   ClientHeight    =   13485
   ClientLeft      =   4140
   ClientTop       =   1530
   ClientWidth     =   11730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   899
   ScaleMode       =   0  'User
   ScaleWidth      =   782
   Begin VB.ListBox lstRight 
      Height          =   2400
      Index           =   14
      ItemData        =   "Grocery.frx":0000
      Left            =   8880
      List            =   "Grocery.frx":0002
      TabIndex        =   46
      Top             =   10200
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ListBox lstRight 
      Height          =   2400
      Index           =   10
      ItemData        =   "Grocery.frx":0004
      Left            =   9720
      List            =   "Grocery.frx":0006
      TabIndex        =   45
      Top             =   10200
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtListBox 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   12135
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   44
      Top             =   960
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.ListBox lstRight 
      Height          =   3375
      Index           =   19
      ItemData        =   "Grocery.frx":0008
      Left            =   9840
      List            =   "Grocery.frx":000A
      TabIndex        =   34
      Top             =   9720
      Width           =   1575
   End
   Begin VB.ListBox lstRight 
      Height          =   3375
      Index           =   18
      ItemData        =   "Grocery.frx":000C
      Left            =   8160
      List            =   "Grocery.frx":000E
      TabIndex        =   33
      Top             =   9720
      Width           =   1575
   End
   Begin VB.ListBox lstRight 
      Height          =   3375
      Index           =   17
      ItemData        =   "Grocery.frx":0010
      Left            =   6480
      List            =   "Grocery.frx":0012
      TabIndex        =   32
      Top             =   9720
      Width           =   1575
   End
   Begin VB.ListBox lstRight 
      Height          =   3375
      Index           =   16
      ItemData        =   "Grocery.frx":0014
      Left            =   4800
      List            =   "Grocery.frx":0016
      TabIndex        =   31
      Top             =   9720
      Width           =   1575
   End
   Begin VB.ListBox lstRight 
      Height          =   3375
      Index           =   15
      ItemData        =   "Grocery.frx":0018
      Left            =   3120
      List            =   "Grocery.frx":001A
      TabIndex        =   30
      Top             =   9720
      Width           =   1575
   End
   Begin VB.ListBox lstRight 
      Height          =   2400
      Index           =   13
      ItemData        =   "Grocery.frx":001C
      Left            =   9840
      List            =   "Grocery.frx":001E
      TabIndex        =   29
      Top             =   6720
      Width           =   1575
   End
   Begin VB.ListBox lstRight 
      Height          =   2400
      Index           =   12
      ItemData        =   "Grocery.frx":0020
      Left            =   8160
      List            =   "Grocery.frx":0022
      TabIndex        =   28
      Top             =   6720
      Width           =   1575
   End
   Begin VB.ListBox lstRight 
      Height          =   2400
      Index           =   11
      ItemData        =   "Grocery.frx":0024
      Left            =   6480
      List            =   "Grocery.frx":0026
      TabIndex        =   26
      Top             =   6720
      Width           =   1575
   End
   Begin VB.ListBox lstRight 
      Height          =   2400
      Index           =   9
      ItemData        =   "Grocery.frx":0028
      Left            =   4800
      List            =   "Grocery.frx":002A
      TabIndex        =   15
      Top             =   6720
      Width           =   1575
   End
   Begin VB.ListBox lstRight 
      Height          =   2400
      Index           =   8
      ItemData        =   "Grocery.frx":002C
      Left            =   3120
      List            =   "Grocery.frx":002E
      TabIndex        =   14
      Top             =   6720
      Width           =   1575
   End
   Begin VB.ListBox lstRight 
      Height          =   2400
      Index           =   7
      ItemData        =   "Grocery.frx":0030
      Left            =   8160
      List            =   "Grocery.frx":0032
      TabIndex        =   13
      Top             =   3840
      Width           =   1575
   End
   Begin VB.ListBox lstRight 
      Height          =   5325
      Index           =   6
      ItemData        =   "Grocery.frx":0034
      Left            =   9840
      List            =   "Grocery.frx":0036
      TabIndex        =   12
      Top             =   960
      Width           =   1575
   End
   Begin VB.ListBox lstRight 
      Height          =   2400
      Index           =   5
      ItemData        =   "Grocery.frx":0038
      Left            =   4800
      List            =   "Grocery.frx":003A
      TabIndex        =   11
      Top             =   3840
      Width           =   1575
   End
   Begin VB.ListBox lstRight 
      Height          =   2400
      Index           =   4
      ItemData        =   "Grocery.frx":003C
      Left            =   3120
      List            =   "Grocery.frx":003E
      TabIndex        =   10
      Top             =   3840
      Width           =   1575
   End
   Begin VB.ListBox lstRight 
      Height          =   5325
      Index           =   3
      ItemData        =   "Grocery.frx":0040
      Left            =   6480
      List            =   "Grocery.frx":0042
      TabIndex        =   9
      Top             =   960
      Width           =   1575
   End
   Begin VB.ListBox lstRight 
      Height          =   2400
      Index           =   2
      ItemData        =   "Grocery.frx":0044
      Left            =   4800
      List            =   "Grocery.frx":0046
      TabIndex        =   8
      Top             =   960
      Width           =   1575
   End
   Begin VB.ListBox lstRight 
      Height          =   2400
      Index           =   1
      ItemData        =   "Grocery.frx":0048
      Left            =   3120
      List            =   "Grocery.frx":004A
      TabIndex        =   7
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton CmdPrintList 
      Caption         =   "Print List"
      Height          =   495
      Left            =   600
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdBackAll 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   5
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   4
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton cmdMoveALL 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   1680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton CmdMove 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   1080
      Width           =   495
   End
   Begin VB.ListBox lstRight 
      Height          =   2400
      Index           =   0
      ItemData        =   "Grocery.frx":004C
      Left            =   8160
      List            =   "Grocery.frx":004E
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
   Begin VB.ListBox lstLeft 
      Height          =   12150
      ItemData        =   "Grocery.frx":0050
      Left            =   120
      List            =   "Grocery.frx":0052
      TabIndex        =   0
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label lblOwner 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   240
      TabIndex        =   43
      Top             =   120
      Width           =   105
   End
   Begin VB.Label lblDate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   10800
      TabIndex        =   42
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "19"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   19
      Left            =   10440
      TabIndex        =   41
      Top             =   9360
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   18
      Left            =   8880
      TabIndex        =   40
      Top             =   9360
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "17"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   17
      Left            =   7200
      TabIndex        =   39
      Top             =   9360
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "16"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   16
      Left            =   5400
      TabIndex        =   38
      Top             =   9360
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "15"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   15
      Left            =   3720
      TabIndex        =   37
      Top             =   9360
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "13"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   13
      Left            =   10440
      TabIndex        =   36
      Top             =   6360
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   12
      Left            =   8760
      TabIndex        =   35
      Top             =   6360
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   11
      Left            =   7080
      TabIndex        =   27
      Top             =   6360
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   9
      Left            =   5400
      TabIndex        =   25
      Top             =   6360
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   8
      Left            =   3915
      TabIndex        =   24
      Top             =   6360
      Width           =   150
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   7
      Left            =   8880
      TabIndex        =   23
      Top             =   3480
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   6
      Left            =   10635
      TabIndex        =   22
      Top             =   600
      Width           =   150
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   5
      Left            =   5475
      TabIndex        =   21
      Top             =   3480
      Width           =   150
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   4
      Left            =   3720
      TabIndex        =   20
      Top             =   3480
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   3
      Left            =   7185
      TabIndex        =   19
      Top             =   600
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   2
      Left            =   5505
      TabIndex        =   18
      Top             =   600
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   1
      Left            =   3825
      TabIndex        =   17
      Top             =   600
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   0
      Left            =   8880
      TabIndex        =   16
      Top             =   600
      Width           =   300
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBack_Click()
    For x = 0 To 19
        If lstRight(x).ListIndex >= 0 Then
            If x < 10 Then
                lstLeft.AddItem "0" & x & "," & lstRight(x).Text
            Else
                lstLeft.AddItem x & "," & lstRight(x).Text
            End If
            lstRight(x).RemoveItem lstRight(x).ListIndex
        End If
    Next x
End Sub

Private Sub cmdBackAll_Click()
    For x = 0 To 19
        Do While lstRight(x).ListCount ' Other
            If x < 10 Then
                lstLeft.AddItem "0" & x & "," & lstRight(x).List(0)
            Else
                lstLeft.AddItem x & "," & lstRight(x).List(0)
            End If
            lstRight(x).RemoveItem 0
        Loop
    Next x
End Sub

Private Sub CmdMove_click()
    If lstLeft.ListIndex >= 0 Then
        If Left$(lstLeft.Text, 2) = "01" Then ' Dairy
            x = Len(lstLeft.Text)
            lstRight(1).AddItem Right$(lstLeft.Text, x - 3)
        ElseIf Left$(lstLeft.Text, 2) = "02" Then ' Bakery
            x = Len(lstLeft.Text)
            lstRight(2).AddItem Right$(lstLeft.Text, x - 3)
        ElseIf Left$(lstLeft.Text, 2) = "03" Then ' Produce
            x = Len(lstLeft.Text)
            lstRight(3).AddItem Right$(lstLeft.Text, x - 3)
        ElseIf Left$(lstLeft.Text, 2) = "04" Then ' Canned Goods
            x = Len(lstLeft.Text)
            lstRight(4).AddItem Right$(lstLeft.Text, x - 3)
        ElseIf Left$(lstLeft.Text, 2) = "05" Then ' Cereal
            x = Len(lstLeft.Text)
            lstRight(5).AddItem Right$(lstLeft.Text, x - 3)
        ElseIf Left$(lstLeft.Text, 2) = "06" Then ' Baking
            x = Len(lstLeft.Text)
            lstRight(6).AddItem Right$(lstLeft.Text, x - 3)
        ElseIf Left$(lstLeft.Text, 2) = "07" Then ' Condiments
            x = Len(lstLeft.Text)
            lstRight(7).AddItem Right$(lstLeft.Text, x - 3)
        ElseIf Left$(lstLeft.Text, 2) = "08" Then ' Beverages
            x = Len(lstLeft.Text)
            lstRight(8).AddItem Right$(lstLeft.Text, x - 3)
        ElseIf Left$(lstLeft.Text, 2) = "09" Then ' Frozen
            x = Len(lstLeft.Text)
            lstRight(9).AddItem Right$(lstLeft.Text, x - 3)
        ElseIf Left$(lstLeft.Text, 2) = "10" Then ' Non-foods
            x = Len(lstLeft.Text)
            lstRight(10).AddItem Right$(lstLeft.Text, x - 3)
        ElseIf Left$(lstLeft.Text, 2) = "11" Then '
            x = Len(lstLeft.Text)
            lstRight(11).AddItem Right$(lstLeft.Text, x - 3)
        ElseIf Left$(lstLeft.Text, 2) = "12" Then ' Meat
            x = Len(lstLeft.Text)
            lstRight(12).AddItem Right$(lstLeft.Text, x - 3)
        ElseIf Left$(lstLeft.Text, 2) = "13" Then ' Snacks
            x = Len(lstLeft.Text)
            lstRight(13).AddItem Right$(lstLeft.Text, x - 3)
        ElseIf Left$(lstLeft.Text, 2) = "14" Then '
            x = Len(lstLeft.Text)
            lstRight(14).AddItem Right$(lstLeft.Text, x - 3)
        ElseIf Left$(lstLeft.Text, 2) = "15" Then ' Office
            x = Len(lstLeft.Text)
            lstRight(15).AddItem Right$(lstLeft.Text, x - 3)
        ElseIf Left$(lstLeft.Text, 2) = "16" Then ' Cleaning
            x = Len(lstLeft.Text)
            lstRight(16).AddItem Right$(lstLeft.Text, x - 3)
        ElseIf Left$(lstLeft.Text, 2) = "17" Then ' Paper Products
            x = Len(lstLeft.Text)
            lstRight(17).AddItem Right$(lstLeft.Text, x - 3)
        ElseIf Left$(lstLeft.Text, 2) = "18" Then ' Health/Beauty
            x = Len(lstLeft.Text)
            lstRight(18).AddItem Right$(lstLeft.Text, x - 3)
        ElseIf Left$(lstLeft.Text, 2) = "19" Then ' personal
            x = Len(lstLeft.Text)
            lstRight(19).AddItem Right$(lstLeft.Text, x - 3)
        Else ' everrything else in other list
            x = Len(lstLeft.Text)
            lstRight(0).AddItem Right$(lstLeft.Text, x - 3)
        End If
        lstLeft.RemoveItem lstLeft.ListIndex
    End If
End Sub

Private Sub cmdMoveALL_Click()
    'If lstLeft.ListIndex >= 0 Then
        'Do While lstLeft.ListCount
            'Text = lstLeft.Text
            'If Left$(Text, 2) = "00" Then
            'If Left$(lstLeft.Text, 2) = "01" Then
            '    x = Len(lstLeft.Text)
            '    MsgBox "help"
            '    lstRight(3).AddItem Right$(lstLeft.Text, x - 3)
            'End If
            'If Left$(lstLeft.Text, 2) = "01" Then ' Condiments
            '    x = Len(lstLeft.Text)
            '    lstRight(1).AddItem Right$(lstLeft.Text, x - 3)
            'End If
            'For y = 0 To 19
            '    If Left$(Text, 2) = "01" Then
         '           lstRight(0).AddItem lstLeft.List
          '          lstLeft.RemoveItem 0
            '    End If
            'Next y
        'Loop
    'End If
End Sub

Private Sub CmdPrintList_Click()
    lstLeft.Visible = False ' don't print this list box
    CmdMove.Visible = False
    cmdMoveALL.Visible = False
    cmdBack.Visible = False
    cmdBackAll.Visible = False
    CmdPrintList.Visible = False
    lblOwner.Visible = True
    txtListBox.Visible = True
    txtListBox.Text = "_______________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________"
    Form1.PrintForm ' print the form as it is
    txtListBox.Visible = False
    lstLeft.Visible = True  ' make it visible
    CmdMove.Visible = True
    cmdMoveALL.Visible = True
    cmdBack.Visible = True
    cmdBackAll.Visible = True
    CmdPrintList.Visible = True
    'On Error GoTo error_handler
    'Open "C:\test.txt" For Output As #2
    '   Do While lstLeft.ListCount ' Index
    '        Print #2, lstLeft.List(0)
    '        lstLeft.RemoveItem (0)
    '    Loop
    '   Close #2
    'Exit Sub
'error_handler:
    'Close
    'MsgBox "Unable to write to file"
End Sub

Private Sub Form_Load()
    Dim item As String
    txtListBox.Visible = False
    Label1(0).Caption = "Other"
    Label1(1).Caption = "Dairy"
    Label1(2).Caption = "Bakery"
    Label1(3).Caption = "Produce"
    Label1(4).Caption = "Canned Goods"
    Label1(5).Caption = "Cereal"
    Label1(6).Caption = "Baking"
    Label1(7).Caption = "Condiments"
    Label1(8).Caption = "Beverage"
    Label1(9).Caption = "Frozen"
    'Label1(10).Caption = "Available"
    Label1(11).Caption = "Pasta/Rice"
    Label1(12).Caption = "Meat"
    Label1(13).Caption = "Snacks"
    'Label1(14).Caption = "Available"
    Label1(15).Caption = "Office"
    Label1(16).Caption = "Laundry/Clean"
    Label1(17).Caption = "Paper Products"
    Label1(18).Caption = "Health Beauty"
    Label1(19).Caption = "Non-Foods"
    lblOwner.Visible = False
    lblOwner.Caption = "Rebecca's Shopping List"
    lblDate.Caption = FormatDateTime(Now, vbLongDate) 'format(now, "dddd mmm dd, yyyy")
    On Error GoTo error_handler
    Open "c:\grocery1.txt" For Input As #1
    Do Until EOF(1)
        Line Input #1, item
        'x = Len(item) ' delete numbers if needed
        'lstLeft.AddItem Right$(item, x - 3)
        lstLeft.AddItem item
    Loop
    Close #1
    Exit Sub
error_handler:
    MsgBox "Unable to load Data File"
End Sub

Private Sub lstLeft_DblClick()
    CmdMove.Value = True
End Sub
