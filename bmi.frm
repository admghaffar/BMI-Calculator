VERSION 5.00
Begin VB.Form bmi 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Bmi Adam"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdexit 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8520
      TabIndex        =   9
      Top             =   7200
      Width           =   2055
   End
   Begin VB.CommandButton cmdreset 
      Caption         =   "RESET"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6360
      TabIndex        =   8
      Top             =   7200
      Width           =   2055
   End
   Begin VB.CommandButton cmdcalculate 
      BackColor       =   &H00FFFF00&
      Caption         =   "CALCULATE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1080
      TabIndex        =   7
      Top             =   7320
      Width           =   2055
   End
   Begin VB.TextBox txttotal 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   855
      Left            =   5040
      TabIndex        =   6
      Top             =   5640
      Width           =   2295
   End
   Begin VB.TextBox txtheight 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   5040
      TabIndex        =   4
      Top             =   4080
      Width           =   2295
   End
   Begin VB.TextBox txtmass 
      Alignment       =   2  'Center
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   855
      Left            =   5040
      TabIndex        =   2
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Label lbltotal 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "TOTAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   840
      TabIndex        =   5
      Top             =   5640
      Width           =   2535
   End
   Begin VB.Label lblheight 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "HEIGHT (METER)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   840
      TabIndex        =   3
      Top             =   3960
      Width           =   2535
   End
   Begin VB.Label lblmass 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "MASS (KILOGRAM)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   840
      TabIndex        =   1
      Top             =   2400
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "BODY MASS INDEX (BMI) CALCULATION "
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3120
      TabIndex        =   0
      Top             =   360
      Width           =   5055
   End
End
Attribute VB_Name = "bmi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcalculate_Click()
Dim Mass As Double
Dim Height As Double


Mass = txtmass.Text
Height = txtheight.Text

txttotal.Text = Mass / (Height * Height)

End Sub

Private Sub cmdexit_Click()
End
End Sub

Private Sub cmdreset_Click()
txtmass.Text = " "
txtheight.Text = " "
txttotal.Text = " "



End Sub
