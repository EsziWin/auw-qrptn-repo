VERSION 5.00
Begin VB.Form Sztorno 
   BackColor       =   &H0091E9FB&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Sztornó beállítása és nyugtázása"
   ClientHeight    =   2748
   ClientLeft      =   36
   ClientTop       =   300
   ClientWidth     =   4560
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2748
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H0091E9FB&
      Caption         =   "A sztornózás oka"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1092
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   1932
      Begin VB.OptionButton Option2 
         BackColor       =   &H0091E9FB&
         Caption         =   "Hibás teljesítés"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   7.8
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1692
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H0091E9FB&
         Caption         =   "Téves adatbevitel"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   7.8
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Value           =   -1  'True
         Width           =   1692
      End
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   288
      Left            =   2160
      TabIndex        =   9
      Top             =   120
      Width           =   1092
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   288
      Left            =   2160
      TabIndex        =   6
      Top             =   480
      Width           =   1092
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mégsem"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2160
      Width           =   2292
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sztornózás végrehajtása"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1560
      Width           =   2292
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   2160
      MaxLength       =   15
      TabIndex        =   3
      Top             =   1200
      Width           =   2052
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   2160
      MaxLength       =   6
      TabIndex        =   1
      Top             =   840
      Width           =   1092
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Mai dátum: "
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   8
      Top             =   120
      Width           =   1812
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Eredeti könyvelés kelte: "
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   240
      TabIndex        =   7
      Top             =   480
      Width           =   1932
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Sztornó számla száma: "
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   2052
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Sztornó kelte: "
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   840
      TabIndex        =   0
      Top             =   840
      Width           =   1332
   End
End
Attribute VB_Name = "Sztorno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  sztornovegrehajt = 1
  Call langclos
  Sztorno.Hide
End Sub

Private Sub Command2_Click()
  sztornovegrehajt = 0
  Call langclos
  Sztorno.Hide
End Sub

Private Sub Form_Load()
  Call langinit("sztorno", 2)
  Call szkriptel("sztorno")
End Sub

Private Sub Option1_Click()
  If Option1.Value = True Then Text1.Text = Text3.Text
End Sub

Private Sub Option2_Click()
  If Option2.Value = True Then Text1.Text = maidatum$
End Sub

