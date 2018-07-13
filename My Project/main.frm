VERSION 5.00
Begin VB.Form main 
   BackColor       =   &H00C0C0FF&
   Caption         =   "����������� �������� - �������"
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
Public cvetkisti, cvetkistidlykisti, cvetkistidlyfiguri As Single '����� ���������� ��� ���� ����
Public risinform, sterkainform, pryminform, okruginform, sohraninform, udalitinform As String
Dim xpunktirprym, ypunktirprym As Single '����� ���������� ��� ���� �����

Private Sub Form_Load()
cvetkistidlykisti = cvet.cvetkisti(0).BackColor '����� ��������� ���� �����
Call Form_Resize '����� ��������� ������ PictureBox
inst.Show '������ �������� ��� �����
cvet.Show
'� ���������, ������� ����� ���� �� ���������. ������� �������� ������ ���������/������ 200, � ������� ����� ������� ��������� � ��������� ������ ����. ����������, ��� ����������:
main.Height = inst.Height '����� ��������� ������ �����
main.Width = cvet.Width
main.Top = (Screen.Height - main.Height - cvet.Height) / 2 - 100 '������ ����� ("main" � ������������ � "cvet" � �������� ����� ����) ���������� ������
main.Left = (Screen.Width - main.Width) / 2
inst.Top = main.Top '������������� ������������ ����� "inst"
inst.Left = main.Left - inst.Width - 200
cvet.Top = main.Top + main.Height + 200 '������������� ������������ ����� "cvet"
cvet.Left = main.Left
'Pic.Print main.Top '��� ��������
'Pic.Print Screen.Height - (cvet.Top + cvet.Height) '��� ��������
risinform = "- �� ������ �������� �� �������;" + vbCrLf + "- ���� ��������� ������� �� �����, ������� ��� ������ �� ������� �����;" + vbCrLf + "- ������ ����� ������� �� ��������� ������ ���������, ���������� �� ������ �����."
sterkainform = "- �� ������ ������� � ������� �������� ����� ������ �������;" + vbCrLf + "- ������ ������������ ������� ������� �� ��������� ������ ���������, ���������� �� ������ �����."
pryminform = "- �� ������ �������� �� ������� ��������������;" + vbCrLf + "- ���� ������ ������� �� �����, ������� ��� ������  �� ������� �����;" + vbCrLf + "- ������� � ���� ������� ������� �� ������� `��� ������` � �����, ������� ��� ������ �� ������� ���� ������."
okruginform = "- �� ������ �������� �� ������� ����������;" + vbCrLf + "- ���� ������ ������� �� �����, ������� ��� ������  �� ������� �����;" + vbCrLf + "- ������� � ���� ������� ������� �� ������� `��� ������` � �����, ������� ��� ������ �� ������� ���� ������."
sohraninform = "- �� ��������� ���� ��������� (� ����� ����������� � ������) ��� �������."
udalitinform = "- ������� ���������."
End Sub

Private Sub Form_Resize()
If Me.WindowState = 1 Then '��� ������������ ���� ��������� ������������� ��� ����
cvet.WindowState = 1
inst.WindowState = 1
End If
If Me.WindowState = 0 Then '��� ������������ ���� ��������� ��������������� ��� ����
cvet.WindowState = 0
inst.WindowState = 0
End If
If Me.WindowState = 0 Or Me.WindowState = 2 Then '�������� ������ PictureBox ��� ��������� ������� �����
Pic.Height = Me.Height - 570
Pic.Width = Me.Width - 240
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
End '��������� ���������
End Sub

Private Sub Pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If inst.prym.Value = True Or inst.okrug.Value = True Then '�������� ������ ��� "�������������" � "����������"
xpunktirprym = X '����� �������� ���������� ����������
ypunktirprym = Y
punktirprym.Top = ypunktirprym '������� ���������������� �������������� ����� Y �����, ��� ������
punktirprym.Left = xpunktirprym '����� ��� ���������������� �������������� ����� X �����, ��� ������
punktirprym.Height = 0 '��������� ������� ���������������� �������������� ����� ���� (������� ����� �������� �� ����� ���� ������ 8, �� � ����� ���������������)
punktirprym.Width = 0
If inst.okrug.Value = True Then '�������� ������ ��� "����������"
lineokrug1.X1 = X '����� ��������������� �����, ����������� ������, ������ ����� �� �������� ����������
lineokrug1.Y1 = Y
lineokrug1.X2 = X
lineokrug1.Y2 = Y
lineokrug1.Visible = True '������ ��������������� ����� �����
punktirokrug.Width = 0 '��������� ������� ��������������� ���������� ����� ���� (������� ����� �������� �� ����� ���� ������ 8, �� � ����� ���������������)
punktirokrug.Height = 0
punktirokrug.Left = xpunktirprym - punktirokrug.Width / 2 '����� ������� - �������� ��������������� ����������
punktirokrug.Top = ypunktirprym - punktirokrug.Height / 2
punktirokrug.Visible = True '������ ��������������� ���������� �����
End If
If inst.prym.Value = True Then punktirprym.Visible = True '������ ��������������� ������������� ������� (�������� ������ ��� "������������")
End If

If inst.ris.Value Or inst.sterka.Value Then  '�������� ������ ��� "����������" � "������� ��"
Pic.CurrentX = X '������������� X � Y � �����, ��� ������
Pic.CurrentY = Y
If inst.ris.Value Then cvetkisti = cvetkistidlykisti '������ ���������� ��� ����� ����� �� ���� ���� ������� ������
If inst.sterka.Value Then cvetkisti = Pic.BackColor '������ ���������� ��� ����� ����� �� ���� ���� PictureBox
If Button = 1 Or Button = 2 Then Pic.PSet (X, Y), cvetkisti '������ ����� � ����� �������
End If

Pic.MouseIcon = LoadPicture(App.Path + "\img\cursor2.ico") '�������� ������ ��� �������
End Sub

Private Sub Pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Or Button = 2 Then '�������� ������ ��� ������� ����� �� ���� ������ ����

If inst.prym.Value = True Or inst.okrug.Value = True Then '�������� ������ ��� "�������������" � "����������"
 If X >= xpunktirprym And Y >= ypunktirprym Then '����� � ������ ���������������� �������������� �� ����� ���� ������������� (�������� ������), ������� ����� ��� ��� ������, ����� ����� � ������ >= 0
   punktirprym.Height = Y - punktirprym.Top '������ ���������������� �������������� ����� ���������� �� ��� ������� �� ����� ���������� ���������
   punktirprym.Width = X - punktirprym.Left '�������� � �������
 Else
  If X >= xpunktirprym Then '����������� � �������, ����� ������������ ��-���� ����� ������� ������ "�������������"
    punktirprym.Width = X - punktirprym.Left '������ ���������������� �������������� ����� ���������� �� ��� ������ ���� �� ����� ���������� ���������
    punktirprym.Top = Y '������ ������� ����� ��������� ���������
    punktirprym.Height = ypunktirprym - Y '������ ����� ���������� �� ����������� Y �� ������� (��������� �������)
  End If
  If Y >= ypunktirprym Then '����������
    punktirprym.Height = Y - punktirprym.Top
    punktirprym.Left = X
    punktirprym.Width = xpunktirprym - X
  End If
  If X < xpunktirprym And Y < ypunktirprym Then '���� �� ������ �� ���, ��� �������� ��
    punktirprym.Top = Y '��������� ���������� ����
    punktirprym.Left = X
    punktirprym.Height = ypunktirprym - Y
    punktirprym.Width = xpunktirprym - X
  End If
 End If
 If inst.okrug.Value = True Then '�������� ������ ��� "����������"
   lineokrug1.X2 = X '������ ��������������� ����� ��������� ������ ����������
   lineokrug1.Y2 = Y
   punktirokrug.Left = xpunktirprym - Sqr(punktirprym.Width ^ 2 + punktirprym.Height ^ 2) '��������� �� ����� ������� ����������, ������ ������� ������� ����������
   punktirokrug.Top = ypunktirprym - Sqr(punktirprym.Width ^ 2 + punktirprym.Height ^ 2)
   punktirokrug.Width = Sqr(punktirprym.Width ^ 2 + punktirprym.Height ^ 2) * 2 '������ ������� (������) ��������������� ���������� ����� ���� �������� ������� ���������� (��������� ������ (�����. ����) � ���������� ������). ����� ������� ����� ������� ���������� ���������� ��������� ����������.
   punktirokrug.Height = punktirokrug.Width
 End If
End If

If inst.ris.Value = True Or inst.sterka.Value = True Then '�������� ������ ��� "����������" � "������� ��"
Pic.Line -(X, Y), cvetkisti '�������� ����� � ������ � �������, ������� �� ������
End If

End If
End Sub

Private Sub Pic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If inst.prym.Value = True Or inst.okrug.Value = True Then '�������� ������ ��� "�������������" � "����������"
 If inst.checkfonfiguri.Value = 1 Or inst.checkfonfiguri.Value = 2 Then '������ �������������/���������� ������������/�������������� � ����������� �� ������ ������������
  Pic.FillStyle = 0
   Else
  Pic.FillStyle = 1
 End If
 If inst.prym.Value = True Then '�������� ������ ��� "�������������"
 punktirprym.Visible = False '������ ��������������� ������������� ��������� � ������, ���� �� ��� �������
 Pic.Line (xpunktirprym, ypunktirprym)-(X, Y), cvetkistidlykisti, B '������ �������������, ���� ������ "�������������"
 End If
 If inst.okrug.Value = True Then '�������� ������ ��� "����������"
 punktirokrug.Visible = False '������ ��������������� ���������� ��������� � ������, ���� ��� ���� �������
 lineokrug1.Visible = False '������ ��������������� ����� ��������� � ������, ���� �� ��� �������
 Pic.Circle (xpunktirprym, ypunktirprym), Sqr(punktirprym.Width ^ 2 + punktirprym.Height ^ 2), cvetkistidlykisti '������ ����������, ���� ������� "����������". ������ ����������� � ������� ���������������� �������������� "punktirprym", ������� ������� �� ����� ���������, ����� �� ����� � ����� ������������, � �������� ��������.
 End If
End If

Pic.MouseIcon = LoadPicture(App.Path + "\img\cursor1.ico") '�������� ������ ��� �������
End Sub
