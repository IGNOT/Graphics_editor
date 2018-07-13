VERSION 5.00
Begin VB.Form main 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Графический редактор - Полотно"
   ClientHeight    =   5580
   ClientLeft      =   15000
   ClientTop       =   6030
   ClientWidth     =   10455
   ClipControls    =   0   'False
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "@Arial Unicode MS"
      Size            =   8.25
      Charset         =   204
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   HasDC           =   0   'False
   Icon            =   "main.frx":0000
   LinkTopic       =   "main"
   ScaleHeight     =   5580
   ScaleWidth      =   10455
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5295
      Left            =   0
      MouseIcon       =   "main.frx":0442
      MousePointer    =   99  'Custom
      ScaleHeight     =   5295
      ScaleWidth      =   10095
      TabIndex        =   0
      Top             =   0
      Width           =   10095
      Begin VB.Shape punktirokrug 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00000000&
         BorderStyle     =   2  'Dash
         Height          =   888
         Left            =   2640
         Shape           =   3  'Circle
         Top             =   3120
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Line lineokrug1 
         BorderColor     =   &H00000000&
         BorderStyle     =   2  'Dash
         Visible         =   0   'False
         X1              =   3600
         X2              =   5754
         Y1              =   1433
         Y2              =   1320
      End
      Begin VB.Shape punktirprym 
         BorderColor     =   &H00000000&
         BorderStyle     =   2  'Dash
         Height          =   630
         Left            =   3360
         Top             =   2040
         Visible         =   0   'False
         Width           =   2175
      End
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cvetkisti, cvetkistidlykisti, cvetkistidlyfiguri As Single 'задаём переменные для всех форм
Public risinform, sterkainform, pryminform, okruginform, sohraninform, udalitinform As String
Dim xpunktirprym, ypunktirprym As Single 'задаём переменные для этой формы

Private Sub Form_Load()
cvetkistidlykisti = cvet.cvetkisti(0).BackColor 'задаём начальный цвет кисти
Call Form_Resize 'задаём начальный размер PictureBox
inst.Show 'делаем видимыми все формы
cvet.Show
'К сожалению, размеры рамок форм не совпадают. Поэтому пришлось наобум прибавить/отнять 200, а главная форма неточно совпадает с размерами других форм. Получилось, как получилось:
main.Height = inst.Height 'задаём начальный размер формы
main.Width = cvet.Width
main.Top = (Screen.Height - main.Height - cvet.Height) / 2 - 100 'ставим формы ("main" в совокупности с "cvet" и пробелом между ними) посередине экрана
main.Left = (Screen.Width - main.Width) / 2
inst.Top = main.Top 'устанавливаем расположение формы "inst"
inst.Left = main.Left - inst.Width - 200
cvet.Top = main.Top + main.Height + 200 'устанавливаем расположение формы "cvet"
cvet.Left = main.Left
'Pic.Print main.Top 'для проверки
'Pic.Print Screen.Height - (cvet.Top + cvet.Height) 'для проверки
risinform = "- Вы можете рисовать на полотне;" + vbCrLf + "- Цвет рисования зависит от цвета, который был выбран на палитре кисти;" + vbCrLf + "- Размер кисти зависит от положения полосы прокрутки, отвечающей за размер кисти."
sterkainform = "- Вы можете стирать с полотна ненужные части Вашего рисунка;" + vbCrLf + "- Размер стирательной резинки зависит от положения полосы прокрутки, отвечающей за размер кисти."
pryminform = "- Вы можете рисовать на полотне прямоугольники;" + vbCrLf + "- Цвет границ зависит от цвета, который был выбран  на палитре кисти;" + vbCrLf + "- Наличие и цвет заливки зависят от галочки `Фон фигуры` и цвета, который был выбран на палитре фона фигуры."
okruginform = "- Вы можете рисовать на полотне окружности;" + vbCrLf + "- Цвет границ зависит от цвета, который был выбран  на палитре кисти;" + vbCrLf + "- Наличие и цвет заливки зависят от галочки `Фон фигуры` и цвета, который был выбран на палитре фона фигуры."
sohraninform = "- Вы выбираете куда сохранить (с каким разрешением и именем) Ваш рисунок."
udalitinform = "- Полотно очищается."
End Sub

Private Sub Form_Resize()
If Me.WindowState = 1 Then 'при сворачивании окна рисования сворачиваются все окна
cvet.WindowState = 1
inst.WindowState = 1
End If
If Me.WindowState = 0 Then 'при развёртывании окна рисования разворачиваются все окна
cvet.WindowState = 0
inst.WindowState = 0
End If
If Me.WindowState = 0 Or Me.WindowState = 2 Then 'изменяем размер PictureBox при изменении размера формы
Pic.Height = Me.Height - 570
Pic.Width = Me.Width - 240
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
End 'закрываем программу
End Sub

Private Sub Pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If inst.prym.Value = True Or inst.okrug.Value = True Then 'работает только для "Прямоугольник" и "Окружность"
xpunktirprym = X 'задаём значения образцовых переменных
ypunktirprym = Y
punktirprym.Top = ypunktirprym 'макушка вспомогательного прямоугольника равна Y места, где нажали
punktirprym.Left = xpunktirprym 'левый бок вспомогательного прямоугольника равна X места, где нажали
punktirprym.Height = 0 'начальные размеры вспомогательного прямоугольника равны нулю (размеры этого элемента не могут быть меньше 8, но я решил подстраховаться)
punktirprym.Width = 0
If inst.okrug.Value = True Then 'работает тольуо для "Окружность"
lineokrug1.X1 = X 'концы вспомогательной линии, указывающей радиус, теперь лежат на середине окружности
lineokrug1.Y1 = Y
lineokrug1.X2 = X
lineokrug1.Y2 = Y
lineokrug1.Visible = True 'теперь вспомогательная линия видна
punktirokrug.Width = 0 'начальные размеры вспомогательной окружности равны нулю (размеры этого элемента не могут быть меньше 8, но я решил подстраховаться)
punktirokrug.Height = 0
punktirokrug.Left = xpunktirprym - punktirokrug.Width / 2 'место нажатия - середина вспомогательной окружности
punktirokrug.Top = ypunktirprym - punktirokrug.Height / 2
punktirokrug.Visible = True 'теперь вспомогательная окружность видна
End If
If inst.prym.Value = True Then punktirprym.Visible = True 'делаем вспомогательный прямоугольник видимым (работает тольуо для "Прямоугльник")
End If

If inst.ris.Value Or inst.sterka.Value Then  'работает только для "Нарисовать" и "Стереть всё"
Pic.CurrentX = X 'зафиксировать X и Y в месте, где нажали
Pic.CurrentY = Y
If inst.ris.Value Then cvetkisti = cvetkistidlykisti 'меняем переменную для цвета кисти на цвет фона нажатой кнопки
If inst.sterka.Value Then cvetkisti = Pic.BackColor 'меняем переменную для цвета кисти на цвет фона PictureBox
If Button = 1 Or Button = 2 Then Pic.PSet (X, Y), cvetkisti 'рисуем точку в месте нажатия
End If

Pic.MouseIcon = LoadPicture(App.Path + "\img\cursor2.ico") 'изменяем курсор при нажатии
End Sub

Private Sub Pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Or Button = 2 Then 'работает только при зажатии любой из двух клавиш мыши

If inst.prym.Value = True Or inst.okrug.Value = True Then 'работает только для "Прямоугольник" и "Окружность"
 If X >= xpunktirprym And Y >= ypunktirprym Then 'длина и высота вспомогательного прямоугольника не может быть отрицательной (появится ошибка), поэтому пишем код для случая, когда длина и высота >= 0
   punktirprym.Height = Y - punktirprym.Top 'высота вспомогательного прямоугольника равна расстоянию от его макушки до места нахождения указателя
   punktirprym.Width = X - punktirprym.Left 'аналогия с высотой
 Else
  If X >= xpunktirprym Then 'разбираемся с случаем, когда пользователь всё-таки решил сделать ширину "отрицательной"
    punktirprym.Width = X - punktirprym.Left 'ширина вспомогательного прямоугольника равна расстоянию от его левого бока до места нахождения указателя
    punktirprym.Top = Y 'теперь макушка равна положению указателя
    punktirprym.Height = ypunktirprym - Y 'высота равна расстоянию от образцового Y до макушки (положения курсора)
  End If
  If Y >= ypunktirprym Then 'аналогично
    punktirprym.Height = Y - punktirprym.Top
    punktirprym.Left = X
    punktirprym.Width = xpunktirprym - X
  End If
  If X < xpunktirprym And Y < ypunktirprym Then 'если всё совсем не так, как хотелось бы
    punktirprym.Top = Y 'совмещаем предыдущие коды
    punktirprym.Left = X
    punktirprym.Height = ypunktirprym - Y
    punktirprym.Width = xpunktirprym - X
  End If
 End If
 If inst.okrug.Value = True Then 'работает только для "Окружность"
   lineokrug1.X2 = X 'теперь вспомогательная линия указывает радиус окружности
   lineokrug1.Y2 = Y
   punktirokrug.Left = xpunktirprym - Sqr(punktirprym.Width ^ 2 + punktirprym.Height ^ 2) 'отступаем от места нажатия расстояние, равное радиусу будущей окружности
   punktirokrug.Top = ypunktirprym - Sqr(punktirprym.Width ^ 2 + punktirprym.Height ^ 2)
   punktirokrug.Width = Sqr(punktirprym.Width ^ 2 + punktirprym.Height ^ 2) * 2 'теперь диаметр (ширина) вспомогательной окружности равен двум радиусам будущей окружности (возмещаем отступ (смотр. выше) и прибавляем радиус). Таким образом место нажатия продолжает оставаться серединой окружности.
   punktirokrug.Height = punktirokrug.Width
 End If
End If

If inst.ris.Value = True Or inst.sterka.Value = True Then 'работает только для "Нарисовать" и "Стереть всё"
Pic.Line -(X, Y), cvetkisti 'рисуется линия с цветом и шириной, которые мы задали
End If

End If
End Sub

Private Sub Pic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If inst.prym.Value = True Or inst.okrug.Value = True Then 'работает только для "Прямоугольник" и "Окружность"
 If inst.checkfonfiguri.Value = 1 Or inst.checkfonfiguri.Value = 2 Then 'делаем прямоугольник/окружность закрашенными/незакрашенными в зависимости от выбора пользователя
  Pic.FillStyle = 0
   Else
  Pic.FillStyle = 1
 End If
 If inst.prym.Value = True Then 'работает только для "Прямоугольник"
 punktirprym.Visible = False 'делаем вспомогательный прямоугольник невидимым в случае, если он был видимым
 Pic.Line (xpunktirprym, ypunktirprym)-(X, Y), cvetkistidlykisti, B 'рисуем прямоугольник, если выбран "Прямоугольник"
 End If
 If inst.okrug.Value = True Then 'работает только для "Окружность"
 punktirokrug.Visible = False 'делаем вспомогательную окружность невидимой в случае, если она была видимой
 lineokrug1.Visible = False 'делаем вспомогательную линию невидимой в случае, если он был видимым
 Pic.Circle (xpunktirprym, ypunktirprym), Sqr(punktirprym.Width ^ 2 + punktirprym.Height ^ 2), cvetkistidlykisti 'рисуем окружность, если выбрана "Окружность". Радиус вычисляется с помощью вспомогательного прямоугольника "punktirprym", который невидим во время рисования, чтобы не сбить с толку пользователя, и теориемы Пифагора.
 End If
End If

Pic.MouseIcon = LoadPicture(App.Path + "\img\cursor1.ico") 'изменяем курсор при отжатии
End Sub
