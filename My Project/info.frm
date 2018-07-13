VERSION 5.00
Begin VB.Form info 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Информация"
   ClientHeight    =   5115
   ClientLeft      =   2370
   ClientTop       =   2625
   ClientWidth     =   10215
   Icon            =   "info.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   10215
   Begin VB.Frame frminfo 
      Caption         =   """Графический редактор"" ver. 3.0"
      BeginProperty Font 
         Name            =   "Arial Unicode MS"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   5115
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10215
      Begin VB.OptionButton udalitinfo 
         Height          =   735
         Left            =   5520
         Picture         =   "info.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Для корзины"
         Top             =   1260
         Width           =   975
      End
      Begin VB.OptionButton sohraninfo 
         Height          =   735
         Left            =   4440
         Picture         =   "info.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Для сохранения"
         Top             =   1260
         Width           =   975
      End
      Begin VB.OptionButton okruginfo 
         Height          =   735
         Left            =   3360
         Picture         =   "info.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Для окружности"
         Top             =   1260
         Width           =   975
      End
      Begin VB.OptionButton pryminfo 
         Height          =   735
         Left            =   2280
         Picture         =   "info.frx":1590
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Для прямоугольника"
         Top             =   1260
         Width           =   975
      End
      Begin VB.OptionButton sterkainfo 
         Height          =   735
         Left            =   1200
         Picture         =   "info.frx":1E5A
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Для стирательной резинки"
         Top             =   1260
         Width           =   975
      End
      Begin VB.OptionButton risinfo 
         Height          =   735
         Left            =   120
         Picture         =   "info.frx":229C
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Для кисти"
         Top             =   1260
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.Label instruckciya 
         Caption         =   "Инструкция:"
         BeginProperty Font 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   6375
      End
      Begin VB.Label infoaboutinst 
         BeginProperty Font 
            Name            =   "Arial Unicode MS"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   120
         TabIndex        =   6
         Top             =   2280
         Width           =   6375
      End
      Begin VB.Label razrab 
         Caption         =   "Разработчик: Пономаренко Игнатий."
         BeginProperty Font 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   6375
      End
      Begin VB.Image myfoto 
         Appearance      =   0  'Flat
         Height          =   4455
         Left            =   6720
         Picture         =   "info.frx":26DE
         Stretch         =   -1  'True
         ToolTipText     =   "Это я"
         Top             =   480
         Width           =   3255
      End
   End
End
Attribute VB_Name = "info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
info.Top = (Screen.Height - info.Height) / 2 'ставим форму посередине экрана
info.Left = (Screen.Width - info.Width) / 2
infoaboutinst.Caption = main.risinform
End Sub

Private Sub Form_Unload(Cancel As Integer)
main.Enabled = True 'при закрытии формы "info" остальные формы становятся доступными
inst.Enabled = True
cvet.Enabled = True
End Sub

Private Sub okruginfo_Click()
infoaboutinst.Caption = main.okruginform
End Sub

Private Sub pryminfo_Click()
infoaboutinst.Caption = main.pryminform
End Sub

Private Sub risinfo_Click()
infoaboutinst.Caption = main.risinform
End Sub

Private Sub sohraninfo_Click()
infoaboutinst.Caption = main.sohraninform
End Sub

Private Sub sterkainfo_Click()
infoaboutinst.Caption = main.sterkainform
End Sub

Private Sub udalitinfo_Click()
infoaboutinst.Caption = main.udalitinform
End Sub
