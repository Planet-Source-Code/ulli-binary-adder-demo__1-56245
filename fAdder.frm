VERSION 5.00
Begin VB.Form fAdder 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Five Bit Binary Adder Demo"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13845
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   13845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.VScrollBar scrB 
      Height          =   285
      Left            =   8475
      Max             =   -16
      Min             =   15
      TabIndex        =   38
      Top             =   165
      Value           =   -16
      Width           =   240
   End
   Begin VB.VScrollBar scrA 
      Height          =   285
      Left            =   5865
      Max             =   -16
      Min             =   15
      TabIndex        =   37
      Top             =   165
      Value           =   -16
      Width           =   240
   End
   Begin VB.CommandButton btExit 
      Caption         =   "Exit"
      Height          =   360
      Left            =   12540
      TabIndex        =   36
      Top             =   5775
      Width           =   1140
   End
   Begin VB.PictureBox picAdder 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'Kein
      Height          =   4590
      Index           =   4
      Left            =   330
      ScaleHeight     =   4590
      ScaleWidth      =   2580
      TabIndex        =   28
      Top             =   855
      Width           =   2580
   End
   Begin VB.CheckBox ckA 
      BackColor       =   &H00E0E0E0&
      Height          =   225
      Index           =   4
      Left            =   1710
      Style           =   1  'Grafisch
      TabIndex        =   1
      Top             =   630
      Width           =   195
   End
   Begin VB.CheckBox ckB 
      BackColor       =   &H00E0E0E0&
      Height          =   225
      Index           =   4
      Left            =   2055
      Style           =   1  'Grafisch
      TabIndex        =   6
      Top             =   630
      Width           =   195
   End
   Begin VB.CheckBox ckB 
      BackColor       =   &H00E0E0E0&
      Height          =   225
      Index           =   3
      Left            =   4635
      Style           =   1  'Grafisch
      TabIndex        =   7
      Top             =   630
      Width           =   195
   End
   Begin VB.CheckBox ckA 
      BackColor       =   &H00E0E0E0&
      Height          =   225
      Index           =   3
      Left            =   4290
      Style           =   1  'Grafisch
      TabIndex        =   2
      Top             =   630
      Width           =   195
   End
   Begin VB.PictureBox picAdder 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'Kein
      Height          =   4590
      Index           =   3
      Left            =   2910
      ScaleHeight     =   4590
      ScaleWidth      =   2580
      TabIndex        =   25
      Top             =   855
      Width           =   2580
      Begin VB.Label lbCout 
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fest Einfach
         Height          =   225
         Index           =   3
         Left            =   30
         TabIndex        =   27
         Top             =   3810
         Width           =   240
      End
   End
   Begin VB.PictureBox picAdder 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'Kein
      Height          =   4590
      Index           =   1
      Left            =   8070
      ScaleHeight     =   4590
      ScaleWidth      =   2580
      TabIndex        =   13
      Top             =   855
      Width           =   2580
      Begin VB.Label lbCout 
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fest Einfach
         Height          =   225
         Index           =   1
         Left            =   15
         TabIndex        =   20
         Top             =   3780
         Width           =   240
      End
   End
   Begin VB.PictureBox picAdder 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'Kein
      Height          =   4590
      Index           =   2
      Left            =   5490
      ScaleHeight     =   4590
      ScaleWidth      =   2580
      TabIndex        =   14
      Top             =   855
      Width           =   2580
      Begin VB.Label lbCout 
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fest Einfach
         Height          =   225
         Index           =   2
         Left            =   30
         TabIndex        =   24
         Top             =   3795
         Width           =   240
      End
   End
   Begin VB.CheckBox ckA 
      BackColor       =   &H00E0E0E0&
      Height          =   225
      Index           =   2
      Left            =   6870
      Style           =   1  'Grafisch
      TabIndex        =   3
      Top             =   630
      Width           =   195
   End
   Begin VB.CheckBox ckB 
      BackColor       =   &H00E0E0E0&
      Height          =   225
      Index           =   2
      Left            =   7215
      Style           =   1  'Grafisch
      TabIndex        =   8
      Top             =   630
      Width           =   195
   End
   Begin VB.CheckBox ckA 
      BackColor       =   &H00E0E0E0&
      Height          =   225
      Index           =   1
      Left            =   9435
      Style           =   1  'Grafisch
      TabIndex        =   4
      Top             =   630
      Width           =   195
   End
   Begin VB.CheckBox ckB 
      BackColor       =   &H00E0E0E0&
      Height          =   225
      Index           =   1
      Left            =   9795
      Style           =   1  'Grafisch
      TabIndex        =   9
      Top             =   630
      Width           =   195
   End
   Begin VB.CheckBox ckA 
      BackColor       =   &H00E0E0E0&
      Height          =   225
      Index           =   0
      Left            =   12015
      Style           =   1  'Grafisch
      TabIndex        =   5
      Top             =   615
      Width           =   195
   End
   Begin VB.CheckBox ckB 
      BackColor       =   &H00E0E0E0&
      Height          =   225
      Index           =   0
      Left            =   12390
      Style           =   1  'Grafisch
      TabIndex        =   10
      Top             =   615
      Width           =   195
   End
   Begin VB.CheckBox ckC 
      BackColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   13230
      Style           =   1  'Grafisch
      TabIndex        =   12
      Top             =   3855
      Width           =   210
   End
   Begin VB.CheckBox ckM 
      BackColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   13230
      Style           =   1  'Grafisch
      TabIndex        =   11
      Top             =   1365
      Width           =   210
   End
   Begin VB.PictureBox picAdder 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   4590
      Index           =   0
      Left            =   10650
      Picture         =   "fAdder.frx":0000
      ScaleHeight     =   4590
      ScaleWidth      =   2580
      TabIndex        =   0
      Top             =   855
      Width           =   2580
      Begin VB.Label lb 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Mode"
         Height          =   195
         Index           =   5
         Left            =   2130
         TabIndex        =   39
         Top             =   615
         Width           =   405
      End
      Begin VB.Label lbCout 
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fest Einfach
         Height          =   225
         Index           =   0
         Left            =   15
         TabIndex        =   19
         Top             =   3795
         Width           =   240
      End
   End
   Begin VB.Label lbMode 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "M = 0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   13230
      TabIndex        =   42
      Top             =   2565
      Width           =   555
   End
   Begin VB.Label lbOP 
      Alignment       =   1  'Rechts
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   13575
      TabIndex        =   41
      Top             =   60
      Width           =   105
   End
   Begin VB.Label lb 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Incect Carry"
      Height          =   405
      Index           =   6
      Left            =   13260
      TabIndex        =   40
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label lb 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Not B"
      Height          =   195
      Index           =   4
      Left            =   13245
      TabIndex        =   35
      Top             =   1155
      Width           =   405
   End
   Begin VB.Label lb 
      Alignment       =   2  'Zentriert
      Appearance      =   0  '2D
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Legend:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   90
      TabIndex        =   34
      Top             =   330
      Width           =   600
   End
   Begin VB.Label lb 
      Alignment       =   2  'Zentriert
      Appearance      =   0  '2D
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "XOr"
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   2
      Left            =   765
      TabIndex        =   33
      Top             =   555
      Width           =   540
   End
   Begin VB.Label lb 
      Alignment       =   2  'Zentriert
      Appearance      =   0  '2D
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Or"
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   1
      Left            =   765
      TabIndex        =   32
      Top             =   345
      Width           =   540
   End
   Begin VB.Label lb 
      Alignment       =   2  'Zentriert
      Appearance      =   0  '2D
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "And"
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   0
      Left            =   765
      TabIndex        =   31
      Top             =   135
      Width           =   540
   End
   Begin VB.Label lbCout 
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fest Einfach
      Height          =   225
      Index           =   4
      Left            =   105
      TabIndex        =   30
      Top             =   3855
      Width           =   240
   End
   Begin VB.Label lbS 
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fest Einfach
      Height          =   225
      Index           =   4
      Left            =   2205
      TabIndex        =   29
      Top             =   5445
      Width           =   240
   End
   Begin VB.Label lbS 
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fest Einfach
      Height          =   225
      Index           =   3
      Left            =   4785
      TabIndex        =   26
      Top             =   5445
      Width           =   240
   End
   Begin VB.Label lbVals 
      Alignment       =   2  'Zentriert
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fest Einfach
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   6750
      TabIndex        =   23
      Top             =   5820
      Width           =   165
   End
   Begin VB.Label lbValB 
      Alignment       =   1  'Rechts
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fest Einfach
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   8310
      TabIndex        =   22
      Top             =   180
      Width           =   135
   End
   Begin VB.Label lbValA 
      Alignment       =   1  'Rechts
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fest Einfach
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   5700
      TabIndex        =   21
      Top             =   180
      Width           =   135
   End
   Begin VB.Label lbO 
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fest Einfach
      Height          =   225
      Left            =   1410
      TabIndex        =   18
      Top             =   5445
      Width           =   240
   End
   Begin VB.Label lbS 
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fest Einfach
      Height          =   225
      Index           =   2
      Left            =   7365
      TabIndex        =   17
      Top             =   5445
      Width           =   240
   End
   Begin VB.Label lbS 
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fest Einfach
      Height          =   225
      Index           =   1
      Left            =   9945
      TabIndex        =   16
      Top             =   5445
      Width           =   240
   End
   Begin VB.Label lbS 
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fest Einfach
      Height          =   225
      Index           =   0
      Left            =   12525
      TabIndex        =   15
      Top             =   5460
      Width           =   240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H008080FF&
      X1              =   13320
      X2              =   13320
      Y1              =   390
      Y2              =   3960
   End
End
Attribute VB_Name = "fAdder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub InitCommonControls Lib "comctl32" ()

Private Const NumAdders As Long = 5
Private Const vbGray    As Long = &HE0E0E0
Private AdderDisabled   As Boolean

Private Sub ADC(Which As String)

  'analog/digital converter
  'converts analog scrollbar values to digital adder input

  Dim i     As Long

    AdderDisabled = True 'disables adder during conversion
    For i = 0 To NumAdders - 1
        Select Case Which
          Case "A"
            ckA(i) = IIf(scrA And 2 ^ i, vbChecked, vbUnchecked)
          Case "B"
            ckB(i) = IIf(scrB And 2 ^ i, vbChecked, vbUnchecked)
        End Select
    Next i
    AdderDisabled = False

End Sub

Private Function Add(ByVal A As Boolean, ByVal B As Boolean, ByVal CarryIn As Boolean, _
                     ByVal Mode As Boolean, _
                     ByRef CarryOut As Boolean, ByRef OverflowOut As Boolean) As Boolean

  'adds two bits A and B (B is mode dependent) and carry-in
  'returns sum, carry-out and overflow

    Add = (A Xor (B Xor Mode)) Xor CarryIn
    CarryOut = (A And (B Xor Mode)) Or (CarryIn And (A Xor (B Xor Mode)))
    OverflowOut = CarryIn Xor CarryOut

End Function

Private Sub AddAll()

  'gets input; adds all bits with ripple carry, displays result

  Dim A         As Boolean
  Dim B         As Boolean
  Dim CarryIn   As Boolean
  Dim Mode      As Boolean
  Dim R         As Boolean
  Dim CarryOut  As Boolean
  Dim Overflow  As Boolean
  Dim i         As Long
  Dim DecA      As Long
  Dim DecB      As Long
  Dim DecR      As Long

    If Not AdderDisabled Then
        On Error Resume Next
            picAdder(0).SetFocus 'hide focus rect
        On Error GoTo 0

        Mode = (ckM = vbChecked) 'get mode
        CarryIn = (ckC = vbChecked) 'get injected carry

        Select Case True
          Case Not Mode And CarryIn
            lbMode = "M = 1"
            lbOP = " A + B + 1 "
          Case Mode And Not CarryIn
            lbMode = "M = 2"
            lbOP = " A + NOT(B) "
          Case Mode And CarryIn
            lbMode = "M = 3"
            lbOP = " A - B "
          Case Else
            lbOP = " A + B "
            lbMode = "M = 0"
        End Select

        For i = 0 To NumAdders - 1

            A = (ckA(i) = vbChecked) 'get operand A bit

            B = (ckB(i) = vbChecked) 'get operand B bit

            R = Add(A, B, CarryIn, Mode, CarryOut, Overflow) 'add

            lbS(i).BackColor = IIf(R, vbRed, vbBlack) ' display result bit
            lbCout(i).BackColor = IIf(CarryOut, vbRed, vbBlack) 'display carry fwd bit

            CarryIn = CarryOut 'ripple carry

            DecA = DecA - A * 2 ^ i 'add up decimal
            DecB = DecB - B * 2 ^ i
            DecR = DecR - R * 2 ^ i

        Next i

        lbO.BackColor = IIf(Overflow, vbRed, vbBlack) 'display final overflow

        'convert unsigned decimal values to signed decimal values
        If DecA >= 2 ^ i / 2 Then
            DecA = DecA - 2 ^ i
        End If
        scrA = DecA
        If DecB >= 2 ^ i / 2 Then
            DecB = DecB - 2 ^ i
        End If
        scrB = DecB
        If DecR >= 2 ^ i / 2 Then
            DecR = DecR - 2 ^ i
        End If

        'display decimal Decues
        lbValA = " A = " & DecA & " "
        lbValB = " B = " & DecB & " "
        lbVals = " R = " & DecR & IIf(Overflow, " (Overflow) ", " ")
    End If

End Sub

Private Sub btExit_Click()

    Unload Me

End Sub

Private Sub ckA_Click(Index As Integer)

    ckA(Index).BackColor = IIf(ckA(Index) = vbChecked, vbRed, vbGray)
    AddAll

End Sub

Private Sub ckB_Click(Index As Integer)

    ckB(Index).BackColor = IIf(ckB(Index) = vbChecked, vbRed, vbGray)
    AddAll

End Sub

Private Sub ckC_Click()

    ckC.BackColor = IIf(ckC = vbChecked, vbRed, vbGray)
    AddAll

End Sub

Private Sub ckM_Click()

    ckM.BackColor = IIf(ckM = vbChecked, vbRed, vbGray)
    AddAll

End Sub

Private Sub Form_Initialize()

    InitCommonControls

End Sub

Private Sub Form_Load()

    picAdder(1).Picture = picAdder(0).Picture 'propagate picture
    picAdder(2).Picture = picAdder(0).Picture
    picAdder(3).Picture = picAdder(0).Picture
    picAdder(4).Picture = picAdder(0).Picture
    AddAll

End Sub

Private Sub scrA_Change()

    ADC "A"
    AddAll

End Sub

Private Sub scrB_Change()

    ADC "B"
    AddAll

End Sub

':) Ulli's VB Code Formatter V2.17.4 (2004-Sep-18 12:32)  Decl: 7  Code: 181  Total: 188 Lines
':) CommentOnly: 7 (3,7%)  Commented: 13 (6,9%)  Empty: 56 (29,8%)  Max Logic Depth: 3
