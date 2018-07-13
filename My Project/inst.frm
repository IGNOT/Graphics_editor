VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form inst 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Настройки рисования"
   ClientHeight    =   4455
   ClientLeft      =   1665
   ClientTop       =   7125
   ClientWidth     =   3105
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FontTransparent =   0   'False
   HasDC           =   0   'False
   Icon            =   "inst.frx":0000
   LinkTopic       =   "inst"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   3105
   Begin VB.Frame frmrazmerkisti 
      Caption         =   "Размер кисти"
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
      Height          =   3735
      Left            =   1680
      TabIndex        =   9
      Top             =   0
      Width           =   1425
      Begin VB.VScrollBar VSrazmerkisti 
         Height          =   2775
         LargeChange     =   50
         Left            =   405
         Max             =   1
         Min             =   100
         TabIndex        =   10
         Top             =   360
         Value           =   1
         Width           =   615
      End
      Begin VB.Label znachenierazmerkisti 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   3240
         Width           =   1155
      End
   End
   Begin VB.Frame frminst 
      Caption         =   "Инструменты"
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
      Height          =   3735
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1680
      Begin MSComDlg.CommonDialog CommonDialog0 
         Left            =   120
         Top             =   3360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton sohran 
         Height          =   735
         Left            =   105
         Picture         =   "inst.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Сохранить как..."
         Top             =   2640
         Width           =   735
      End
      Begin VB.CommandButton udalit 
         Height          =   735
         Left            =   825
         Picture         =   "inst.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Стереть всё"
         Top             =   2640
         Width           =   735
      End
      Begin VB.CheckBox checkfonfiguri 
         Caption         =   "Фон фигуры"
         BeginProperty Font 
            Name            =   "Arial Unicode MS"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   352
         TabIndex        =   6
         Top             =   2040
         Width           =   975
      End
      Begin VB.OptionButton okrug 
         Height          =   735
         Left            =   825
         Picture         =   "inst.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Нарисовать окружность"
         Top             =   1200
         Width           =   735
      End
      Begin VB.OptionButton prym 
         Height          =   735
         Left            =   105
         Picture         =   "inst.frx":1590
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Нарисовать прямоугольник"
         Top             =   1200
         Width           =   735
      End
      Begin VB.OptionButton sterka 
         Height          =   735
         Left            =   825
         Picture         =   "inst.frx":1E5A
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Стереть"
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton ris 
         Height          =   735
         Left            =   105
         Picture         =   "inst.frx":229C
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Нарисовать"
         Top             =   360
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.CommandButton information 
      Height          =   735
      Left            =   0
      Picture         =   "inst.frx":26DE
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Информация"
      Top             =   3720
      Width           =   3135
   End
End
Attribute VB_Name = "inst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub checkfonfiguri_Click()
If checkfonfiguri.Value = 1 Or checkfonfiguri.Value = 2 Then 'даём/убираем доступность "Фон фигуры"
cvet.frmfonfiguri.Enabled = True
Else
cvet.frmfonfiguri.Enabled = False
End If
End Sub

Private Sub Form_Load()
Call VSrazmerkisti_Scroll 'вызываем VSrazmerkisti_Scroll
End Sub

Private Sub information_Click()
info.Show 'при нажатии на кнопку показывается форма "info"
main.Enabled = False 'при загрузке формы "info" остальные формы становятся недоступными
inst.Enabled = False
cvet.Enabled = False
End Sub

Private Sub sohran_Click()
CommonDialog0.InitDir = App.Path + "\Мои рисунки" 'указываем путь к папке по умолчанию, где будут храниться рисунки
CommonDialog0.FileName = "Мой рисунок" 'задаём начальное имя рисунка
CommonDialog0.Filter = "Файлы данных (*.jpg)|*.jpg|Файлы данных (*.png)|*.png" 'сохранять ихображения можно будет только с помощью этих форматов
CommonDialog0.Flags = 2 'при сохранении файла с уже существующим именем, пользователю выдаётся вопрос о замене оригинального файла этим файлом
CommonDialog0.DialogTitle = "Сохранить рисунок как..." 'меняем оригинальный заголовок окна сохранения
CommonDialog0.ShowSave 'окно сохранения появляется
If CommonDialog0.Flags <> 2 Then SavePicture main.Pic.Image, CommonDialog0.FileName '???

'MsgBox "Вы не ввели название!", vbInformation + vbOKOnly, "Ошибочка сохранения!"

End Sub

Private Sub udalit_Click()
main.Pic.Cls 'стереть всё
End Sub

Private Sub VSrazmerkisti_Change()
Call VSrazmerkisti_Scroll 'вызываем VSrazmerkisti_Scroll
End Sub

Private Sub VSrazmerkisti_Scroll()
znachenierazmerkisti.Caption = VSrazmerkisti.Value 'указываем пользователю размер кисти
main.Pic.DrawWidth = VSrazmerkisti.Value 'меняем размер кисти
End Sub
