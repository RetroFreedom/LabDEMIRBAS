VERSION 5.00
Begin VB.Form frmDEMIRBAS 
   BackColor       =   &H00C0C0C0&
   Caption         =   "LabDEMIRBAS - Kiþisel Demirbaþ takip Programý"
   ClientHeight    =   7080
   ClientLeft      =   2325
   ClientTop       =   1935
   ClientWidth     =   13890
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7080
   ScaleWidth      =   13890
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\LabDEMIRBAS\DEMIRBAS.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   510
      Left            =   6255
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Demirbas"
      Top             =   6255
      Width           =   4245
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1125
      Top             =   5085
   End
   Begin VB.CommandButton KayitEkle 
      Caption         =   "Kayýt Ekle"
      Height          =   555
      Left            =   12645
      Picture         =   "DEMIRBAS.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   585
      Width           =   1140
   End
   Begin VB.CommandButton KayitSil 
      Caption         =   "Kayit Sil"
      Height          =   555
      Left            =   12645
      Picture         =   "DEMIRBAS.frx":1708A
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1485
      Width           =   1140
   End
   Begin VB.CommandButton IptalEt 
      Caption         =   "Ýptal Et"
      Height          =   555
      Left            =   12645
      Picture         =   "DEMIRBAS.frx":2E114
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2340
      Width           =   1140
   End
   Begin VB.CommandButton Cikis 
      Caption         =   "Çýkýþ"
      Height          =   555
      Left            =   11295
      Picture         =   "DEMIRBAS.frx":4519E
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6300
      Width           =   2310
   End
   Begin VB.CommandButton First 
      Caption         =   "<< First"
      Height          =   570
      Left            =   720
      Picture         =   "DEMIRBAS.frx":5C228
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6345
      Width           =   930
   End
   Begin VB.CommandButton Previous 
      Caption         =   "<  Previous"
      Height          =   570
      Left            =   2115
      Picture         =   "DEMIRBAS.frx":732B2
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6345
      Width           =   945
   End
   Begin VB.CommandButton Next 
      Caption         =   "Next  >"
      Height          =   570
      Left            =   3150
      Picture         =   "DEMIRBAS.frx":8A33C
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6345
      Width           =   1035
   End
   Begin VB.CommandButton Last 
      Caption         =   "Last >>"
      Height          =   570
      Left            =   4590
      Picture         =   "DEMIRBAS.frx":A13C6
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6345
      Width           =   1065
   End
   Begin VB.CommandButton Find 
      Caption         =   "Kayýt Bul"
      Height          =   525
      Left            =   12645
      Picture         =   "DEMIRBAS.frx":B8450
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3195
      Width           =   1125
   End
   Begin VB.TextBox txtDemirbasYer 
      DataField       =   "DYeri"
      DataSource      =   "data1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2115
      TabIndex        =   5
      Top             =   3150
      Width           =   4095
   End
   Begin VB.TextBox txtDemirbasSeri 
      DataField       =   "DSeriNo"
      DataSource      =   "data1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2115
      TabIndex        =   4
      Top             =   2610
      Width           =   2175
   End
   Begin VB.TextBox txtDemirbasModel 
      DataField       =   "DModel"
      DataSource      =   "data1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2115
      TabIndex        =   3
      Top             =   2115
      Width           =   2175
   End
   Begin VB.TextBox txtDemirbasKyeri 
      DataField       =   "DResim"
      DataSource      =   "data1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2115
      TabIndex        =   6
      Top             =   3690
      Width           =   4095
   End
   Begin VB.TextBox txtDemirbasAdi 
      DataField       =   "DAdi"
      DataSource      =   "data1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2115
      TabIndex        =   1
      Top             =   1035
      Width           =   4095
   End
   Begin VB.TextBox txtDemirbasMarka 
      DataField       =   "DMarka"
      DataSource      =   "data1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2115
      TabIndex        =   2
      Top             =   1575
      Width           =   2175
   End
   Begin VB.TextBox txtDemirbasNo 
      DataField       =   "DNo"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2115
      TabIndex        =   0
      Top             =   540
      Width           =   2175
   End
   Begin VB.TextBox txtDemirbasAciklama 
      DataField       =   "DAçýklama"
      DataSource      =   "data1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   2115
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   4230
      Width           =   4095
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      Caption         =   "Açýklama"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   660
      TabIndex        =   24
      Top             =   4365
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      Caption         =   "Kayýt yeri"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   420
      TabIndex        =   23
      Top             =   3765
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   5535
      Left            =   6300
      Stretch         =   -1  'True
      Top             =   495
      Width           =   6240
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      Caption         =   "Model"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   420
      TabIndex        =   22
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      Caption         =   "Demirbaþ No"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   450
      TabIndex        =   21
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      Caption         =   "Demirbaþ Adý"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   420
      TabIndex        =   20
      Top             =   1170
      Width           =   1455
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      Caption         =   "Seri No"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   420
      TabIndex        =   19
      Top             =   2685
      Width           =   1455
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      Caption         =   "Marka"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   420
      TabIndex        =   18
      Top             =   1665
      Width           =   1455
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      Caption         =   "Yeri"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   420
      TabIndex        =   17
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BorderColor     =   &H00400040&
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   4455
      Left            =   180
      Shape           =   4  'Rounded Rectangle
      Top             =   405
      Width           =   1815
   End
End
Attribute VB_Name = "frmDEMIRBAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cikis_Click()
    Unload Me
End Sub

Private Sub GridRfresh_Click()
    Data1.Refresh
    MSFlexGrid1.Refresh
End Sub


Private Sub data1_Validate(Action As Integer, Save As Integer)
  Select Case Action
    Case vbDataActionMoveFirst
    Case vbDataActionMovePrevious
    Case vbDataActionMoveNext
    Case vbDataActionMoveLast
    Case vbDataActionAddNew
    Case vbDataActionUpdate
    Case vbDataActionDelete
    Case vbDataActionFind
    Case vbDataActionBookmark
    Case vbDataActionClose
      Screen.MousePointer = vbDefault
  End Select
  Screen.MousePointer = vbNormal
  Load_image
  Data1.Caption = txtDemirbasNo.Text & " : " & txtDemirbasAdi
End Sub

Private Sub Find_Click()
    On Error GoTo exitlbl
    Dim mesajstr$
    mesajstr$ = InputBox("Demirbaþ No Giriniz : ", "Demirbaþ No Arama")
    Data1.Recordset.Index = "DNo"
    Data1.Recordset.Seek "=", mesajstr$
    If Data1.Recordset.NoMatch = True Then
        Data1.Recordset.MoveFirst
        MsgBox ("ÜZGÜNÜM BÖYLE BÝR KAYIT BULUNAMADI")
    Else
        MsgBox ("KAYIT BULUNDU")
    End If
exitlbl:
    Exit Sub
End Sub




Private Sub Form_Load()
'    Set ws = DBEngine.Workspaces(0)
'    dbfile = "c:\LabDEMIRBAS\DEMIRBAS.mdb"
'    Set db = DBEngine.OpenDatabase(dbfile, False, False, ";pwd=")
'    Set rs = db.OpenRecordset("data", dbOpenTable)
    
End Sub

Private Sub IptalEt_Click()
    On Error GoTo exitlbl
    Data1.UpdateControls
    If Data1.Recordset.EditMode = dbEditAdd Then
        Data1.Recordset.CancelUpdate
    End If
exitlbl:
    Exit Sub
End Sub

Private Sub KayitEkle_Click()
    On Error GoTo exitlbl
    Data1.Recordset.AddNew
    txtDemirbasNo.SetFocus
exitlbl:
    Exit Sub
End Sub

Private Sub KayitSil_Click()
    On Error GoTo exitlbl
    Dim mesaj As String
    mesaj = "Kaydý silmek istediðinizden eminmisiniz "
    If MsgBox(mesaj, vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
        Data1.Recordset.Delete ' Kayýt sil
        Data1.Recordset.MoveNext 'Sonraki kayda git
        If Data1.Recordset.EOF Then 'eðer sonraki kayýt yoksa
            Data1.Recordset.MoveLast 'son kaydý göster
        End If
    End If
exitlbl:
    Exit Sub
End Sub

Private Sub First_Click()
    On Error GoTo exitlbl
    Data1.Recordset.MoveFirst
    Load_image
exitlbl:
    Exit Sub
End Sub

Private Sub Last_Click()
    On Error GoTo exitlbl
    Data1.Recordset.MoveLast
    Load_image
exitlbl:
    Exit Sub
End Sub

Private Sub Next_Click()
    On Error GoTo exitlbl
    If Data1.Recordset.EOF Then
        MsgBox ("Son kayýttassýnýz.")
        Data1.Recordset.MoveLast
    Else
        Data1.Recordset.MoveNext
        Load_image
    End If
exitlbl:
    Exit Sub
End Sub

Private Sub Previous_Click()
    On Error GoTo exitlbl
    If Data1.Recordset.BOF Then
        MsgBox ("ilk kayýttassýnýz.")
        Data1.Recordset.MoveFirst
    Else
        Data1.Recordset.MovePrevious
        Load_image
    End If
exitlbl:
    Exit Sub
End Sub

Sub Load_image()
    On Error GoTo exitlbl
    Image1.Picture = LoadPicture("C:\LabDEMIRBAS\FOTO\" & txtDemirbasKyeri.Text & ".JPG")
exitlbl:
    Exit Sub
End Sub

Private Sub Timer1_Timer()
    Data1.Caption = txtDemirbasNo.Text & " : " & txtDemirbasAdi
End Sub

Private Sub txtDemirbasKyeri_Change()
    On Error GoTo exitlbl
    Image1.Picture = LoadPicture("C:\labDEMIRBAS\FOTO\" + txtDemirbasKyeri.Text + ".JPG")
exitlbl:
    Exit Sub
End Sub

Private Sub txtDemirbasNo_Change()
    txtDemirbasKyeri.Text = txtDemirbasNo.Text
End Sub
