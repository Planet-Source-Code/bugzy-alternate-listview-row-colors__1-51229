VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Alt Listview Colors"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   4740
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4140
      Top             =   3690
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Start At Odd Row"
      Height          =   225
      Left            =   1170
      TabIndex        =   2
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Gridlines"
      Height          =   225
      Left            =   120
      TabIndex        =   1
      Top             =   3960
      Width           =   915
   End
   Begin MSComctlLib.ListView lv 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   4485
      _ExtentX        =   7911
      _ExtentY        =   6376
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Column 1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "sub"
         Text            =   "Column 2"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub AltLVBackground(lv As ListView, _
    ByVal BackColorOne As OLE_COLOR, _
    ByVal BackColorTwo As OLE_COLOR)
'---------------------------------------------------------------------------------
' Purpose   : Alternates row colors in a ListView control
' Method    : Creates a picture box and draws the desired color scheme in it, then
'             loads the drawn image as the listviews picture.
'---------------------------------------------------------------------------------
Dim lH      As Long
Dim lSM     As Byte
Dim picAlt  As PictureBox
    With lv
        If .View = lvwReport And .ListItems.Count Then
            Set picAlt = Me.Controls.Add("VB.PictureBox", "picAlt")
            lSM = .Parent.ScaleMode
            .Parent.ScaleMode = vbTwips
            .PictureAlignment = lvwTile
            lH = .ListItems(1).Height
            With picAlt
                .BackColor = BackColorOne
                .AutoRedraw = True
                .Height = lH * 2
                .BorderStyle = 0
                .Width = 10 * Screen.TwipsPerPixelX
                picAlt.Line (0, lH)-(.ScaleWidth, lH * 2), BackColorTwo, BF
                Set lv.Picture = .Image
            End With
            Set picAlt = Nothing
            Me.Controls.Remove "picAlt"
            lv.Parent.ScaleMode = lSM
        End If
    End With
End Sub

Private Sub Form_Load()
Dim x As Long

With lv.ListItems
    For x = 1 To 50
        .Add(, "|" & x, "ListItem " & x, , 1).ListSubItems.Add , "|", "ListSubItem " & x
    Next
End With

AltLVBackground lv, vbWhite, &HC0FFFF

End Sub

Private Sub Check1_Click()
    Me.lv.GridLines = Not Me.lv.GridLines
End Sub

Private Sub Check2_Click()
    AltLVBackground lv, IIf(Me.Check2.Value, &HC0FFFF, vbWhite), IIf(Me.Check2.Value, vbWhite, &HC0FFFF)
End Sub

