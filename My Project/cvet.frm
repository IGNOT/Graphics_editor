VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form cvet 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Палитра"
   ClientHeight    =   720
   ClientLeft      =   1320
   ClientTop       =   1350
   ClientWidth     =   10425
   ControlBox      =   0   'False
   Icon            =   "cvet.frx":0000
   LinkTopic       =   "cvet"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   720
   ScaleWidth      =   10425
   Begin VB.Frame frmzakr 
      Caption         =   "Закрасить всё"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   6960
      TabIndex        =   7
      Top             =   0
      Width           =   3480
      Begin VB.OptionButton cvetfona 
         Height          =   375
         Index           =   5
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Свой цвет"
         Top             =   240
         Width           =   375
      End
      Begin MSComDlg.CommonDialog CommonDialog2 
         Left            =   2880
         Top             =   600
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Color           =   -2147483633
      End
      Begin VB.CommandButton svoycvetfona 
         Height          =   375
         Left            =   3000
         Picture         =   "cvet.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Выбрать свой цвет"
         Top             =   240
         Width           =   375
      End
      Begin VB.OptionButton cvetfona 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Value           =   -1  'True
         Width           =   375
      End
      Begin VB.OptionButton cvetfona 
         BackColor       =   &H000000FF&
         Height          =   375
         Index           =   1
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   375
      End
      Begin VB.OptionButton cvetfona 
         BackColor       =   &H0000FFFF&
         Height          =   375
         Index           =   2
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   375
      End
      Begin VB.OptionButton cvetfona 
         BackColor       =   &H0000FF00&
         Height          =   375
         Index           =   3
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   375
      End
      Begin VB.OptionButton cvetfona 
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   4
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame frmfonfiguri 
      Caption         =   "Фон фигуры"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   3480
      TabIndex        =   6
      Top             =   0
      Width           =   3480
      Begin MSComDlg.CommonDialog CommonDialog3 
         Left            =   2880
         Top             =   600
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Color           =   -2147483633
      End
      Begin VB.CommandButton svoycvetfonfiguri 
         Height          =   375
         Left            =   3000
         Picture         =   "cvet.frx":0E54
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Выбрать свой цвет"
         Top             =   240
         Width           =   375
      End
      Begin VB.OptionButton cvetfonfiguri 
         Height          =   375
         Index           =   5
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Свой цвет"
         Top             =   240
         Width           =   375
      End
      Begin VB.OptionButton cvetfonfiguri 
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   4
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   240
         Width           =   375
      End
      Begin VB.OptionButton cvetfonfiguri 
         BackColor       =   &H0000FF00&
         Height          =   375
         Index           =   3
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   240
         Width           =   375
      End
      Begin VB.OptionButton cvetfonfiguri 
         BackColor       =   &H0000FFFF&
         Height          =   375
         Index           =   2
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         Width           =   375
      End
      Begin VB.OptionButton cvetfonfiguri 
         BackColor       =   &H000000FF&
         Height          =   375
         Index           =   1
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         Width           =   375
      End
      Begin VB.OptionButton cvetfonfiguri 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Value           =   -1  'True
         Width           =   375
      End
   End
   Begin VB.Frame frmkist 
      Caption         =   "Кисть"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3480
      Begin VB.CommandButton svoycvetkist 
         Height          =   375
         Left            =   3000
         Picture         =   "cvet.frx":13DE
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Выбрать свой цвет"
         Top             =   240
         Width           =   375
      End
      Begin VB.OptionButton cvetkisti 
         Height          =   375
         Index           =   5
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Свой цвет"
         Top             =   240
         Width           =   375
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   2880
         Top             =   600
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Color           =   -2147483633
      End
      Begin VB.OptionButton cvetkisti 
         BackColor       =   &H00000000&
         Height          =   375
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   375
      End
      Begin VB.OptionButton cvetkisti 
         BackColor       =   &H000000FF&
         Height          =   375
         Index           =   1
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   375
      End
      Begin VB.OptionButton cvetkisti 
         BackColor       =   &H0000FFFF&
         Height          =   375
         Index           =   2
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   375
      End
      Begin VB.OptionButton cvetkisti 
         BackColor       =   &H0000FF00&
         Height          =   375
         Index           =   3
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   375
      End
      Begin VB.OptionButton cvetkisti 
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   4
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   375
      End
   End
End
Attribute VB_Name = "cvet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cvetfona_Click(Index As Integer)
main.Pic.BackColor = cvetfona(Index).BackColor 'меняем цвет фона на цвет фона нажатой кнопки
End Sub

Private Sub cvetfonfiguri_Click(Index As Integer)
main.Pic.FillColor = cvetfonfiguri(Index).BackColor
End Sub

Private Sub cvetkisti_Click(Index As Integer)
main.cvetkistidlykisti = cvetkisti(Index).BackColor 'меняем цвет, хранящийся в переменной, на цвет фона нажатой кнопки
End Sub
Private Sub svoycvetfona_Click()
CommonDialog2.ShowColor 'аналогия с "svoycvetkist_Click"
cvetfona(5).BackColor = CommonDialog2.Color
main.Pic.BackColor = CommonDialog2.Color
cvetfona(5).Value = True
End Sub

Private Sub svoycvetfonfiguri_Click()
CommonDialog3.ShowColor 'аналогия с "svoycvetkist_Click"
cvetfonfiguri(5).BackColor = CommonDialog3.Color
main.Pic.FillColor = CommonDialog3.Color
cvetfonfiguri(5).Value = True
End Sub

Private Sub svoycvetkist_Click()
CommonDialog1.ShowColor 'вызываем цвета Windows
cvetkisti(5).BackColor = CommonDialog1.Color 'меняем цвет фона "Свой цвет" на выбранный цвет
main.cvetkistidlykisti = CommonDialog1.Color 'меняем цвет, хранящийся в переменной, на выбранный цвет
cvetkisti(5).Value = True 'делаем "Свой цвет" включённым, чтобы цвет кисти совпадал с цветом фона действительной кнопки
End Sub
