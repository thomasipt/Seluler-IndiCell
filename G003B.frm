VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form G003B 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PENARIKAN TUNAI"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8160
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   8160
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1958
      TabIndex        =   0
      Text            =   "1"
      Top             =   150
      Width           =   2685
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1958
      TabIndex        =   1
      Text            =   "2"
      Top             =   645
      Width           =   2685
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1958
      TabIndex        =   2
      Text            =   "3"
      Top             =   1065
      Width           =   5940
   End
   Begin VB.CommandButton cmdCLOSE 
      Caption         =   "KELUAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   5966
      TabIndex        =   4
      Top             =   2460
      Width           =   1890
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "SIMPAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   304
      TabIndex        =   3
      Top             =   2475
      Width           =   1890
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   900
      TabIndex        =   7
      Text            =   "4"
      Top             =   1800
      Width           =   1740
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3675
      TabIndex        =   6
      Text            =   "5"
      Top             =   1800
      Width           =   1740
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6345
      TabIndex        =   5
      Text            =   "6"
      Top             =   1800
      Width           =   1740
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   6510
      OleObjectBlob   =   "G003B.frx":0000
      Top             =   4575
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   240
      Left            =   263
      OleObjectBlob   =   "G003B.frx":0234
      TabIndex        =   8
      Top             =   210
      Width           =   1590
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   240
      Left            =   263
      OleObjectBlob   =   "G003B.frx":02AA
      TabIndex        =   9
      Top             =   705
      Width           =   1590
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   240
      Left            =   263
      OleObjectBlob   =   "G003B.frx":0316
      TabIndex        =   10
      Top             =   1125
      Width           =   1590
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   240
      Left            =   75
      OleObjectBlob   =   "G003B.frx":0388
      TabIndex        =   11
      Top             =   1800
      Width           =   960
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   240
      Left            =   2745
      OleObjectBlob   =   "G003B.frx":03F0
      TabIndex        =   12
      Top             =   1800
      Width           =   960
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
      Height          =   240
      Left            =   5520
      OleObjectBlob   =   "G003B.frx":045A
      TabIndex        =   13
      Top             =   1800
      Width           =   960
   End
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   -255
      TabIndex        =   15
      Top             =   1530
      Width           =   8610
   End
   Begin VB.PictureBox Picture1 
      Height          =   1890
      Left            =   -90
      ScaleHeight     =   1830
      ScaleWidth      =   10980
      TabIndex        =   14
      Top             =   2295
      Width           =   11040
   End
End
Attribute VB_Name = "G003B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Lokasi As String
Dim A, Isi As String

Private RDOE As rdoEnvironment
Private RDCO As rdoConnection
Private RSLNO As rdoResultset

Private RSL, RSLUser, RCari, RCari2, RCari3, RCari4, RCari5, RSave, RSave2, RSave3, RSave4, RSave5, REdit As rdoResultset
Private SQL, SQLUser, SCari, SCari2, SCari3, SCari4, SCari5, SSave, SSave2, SSave3, SSave4, SSave5, SEdit As String

Private RJual1, RJual2, RJual3, RJual4, RJual5, RJual6, RJual7, RJual8, RJual9, RJual10 As rdoResultset
Private SJual1, SJual2, SJual3, SJual4, SJual5, SJual6, SJual7, SJual8, SJual9, SJual10 As String

Private RBahan1, RBahan2, RBahan3, RBahan4, RBahan5, RBahan6, RBahan7, RBahan8, RBahan9, RBahan10 As rdoResultset
Private SBahan1, SBahan2, SBahan3, SBahan4, SBahan5, SBahan6, SBahan7, SBahan8, SBahan9, SBahan10 As String

Private RDEl As rdoResultset
Private SDel As String

Private RLR, RLR2 As rdoResultset
Private SLR, SLR2 As String

Private RJS As rdoResultset
Private SJS As String

Private SqlNo As String

Private Sub cmdCLOSE_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Then
    MsgBox "DATA TIDAK BOLEH KOSONG", vbCritical, "KONFIRMASI"
    Text1.SetFocus
Else
    SSave = "Select * From G003 where CodeSL = '101001'"
    Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdConcurRowVer)
    If RSave.RowCount <> 0 Then
        RSave.Edit
        RSave("MutasiC") = CCur(Text5) + CCur(Text2)
        RSave("Saldo") = CCur(Text6) - CCur(Text2)

        SSave2 = "Select * From G005 ORDER BY NOURUT"
        Set RSave2 = RDCO.OpenResultset(SSave2, rdOpenKeyset, rdConcurRowVer)
        RSave2.AddNew
            RSave2("CodeCab") = CodeCab
            RSave2("CodeSl") = "KAS"
            RSave2("NamaSl") = "KAS"
            RSave2("Nobukti") = Trim(Text1)
            RSave2("Keterangan") = Trim(Text3)
            RSave2("NominalD") = 0
            RSave2("NominalC") = CCur(Text2)
            RSave2("Saldo") = CCur(Text6) - CCur(Text2)
            RSave2("Tanggal") = Date
            RSave2("Jam") = Time
            RSave2("UserCode") = Operator
        RSave2.Update
        RSave2.Close
        Set RSave2 = Nothing
        
        RSave.Update
        RSave.Close
        Set RSave = Nothing
    End If
    
    Unload Me
    G003B.Show 1
    
End If
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=SELULER", rdDriverNoPrompt, False, CN)
Lokasi = App.Path
Skin1.LoadSkin Lokasi + "\" + Skin + ".skn"
Skin1.ApplySkin hWnd
ClearTextBoxes Me

SSave3 = "Select * From G003 where CODESL = '101001'"
Set RSave3 = RDCO.OpenResultset(SSave3, rdOpenDynamic, rdConcurRowVer)
If RSave3("SALDO") <= 0 Then
    MsgBox "SALDO KAS Rp 0.00", vbCritical, "KONFIRMASI"
    Text1.Enabled = False
    Text2.Enabled = False
    Text3.Enabled = False
    Picture1.ZOrder
    cmdCLOSE.ZOrder
End If
RSave3.Close
Set RSave3 = Nothing

Call KAS

End Sub

Private Sub KAS()
SCari = "Select * From G003 where CodeSL = '101001'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset)
If RCari.RowCount <> 0 Then
    Text4 = Format(RCari("MutasiD"), "##,###.00")
    Text5 = Format(RCari("MutasiC"), "##,###.00")
    Text6 = Format(RCari("Saldo"), "##,###.00")
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text1_Lostfocus()
Text1 = Format(Text1, ">")
Call CekData
End Sub

Private Sub CekData()
If Text1.Text = "" Then Exit Sub

SCari = "Select * From G005 where NOBUKTI = '" + Trim(Text1) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
    If RCari.RowCount <> 0 Then
        MsgBox " NO FAKTUR NOTA TELAH DIGUNAKAN", vbCritical, "KONFIRMASI"
        Text1 = ""
        Text1.SetFocus
    Else
        Text2.SetFocus
    Exit Sub
    End If

RCari.Close
Set RCari = Nothing
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Text2_LostFocus()
Text2 = Format(Text2, "##,###.00")
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text3_LostFocus()
    Text3 = Format(Text3, ">")
End Sub

