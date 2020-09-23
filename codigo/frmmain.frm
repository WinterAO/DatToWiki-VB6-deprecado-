VERSION 5.00
Begin VB.Form frmmain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DatToWiki"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11160
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   11160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraInformacionA 
      Caption         =   "Informacion a convertir"
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   5280
      Width           =   10935
      Begin VB.OptionButton OptDat 
         Caption         =   "Armas"
         Height          =   195
         Index           =   3
         Left            =   3240
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton OptDat 
         Caption         =   "Escudos"
         Height          =   195
         Index           =   2
         Left            =   2280
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton OptDat 
         Caption         =   "Armaduras"
         Height          =   195
         Index           =   1
         Left            =   1080
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton OptDat 
         Caption         =   "Cascos"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdConvertir 
      Caption         =   "Convertir"
      Height          =   360
      Left            =   4080
      TabIndex        =   1
      Top             =   6240
      Width           =   2655
   End
   Begin VB.TextBox txtTxtResultado 
      Appearance      =   0  'Flat
      Height          =   5055
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmmain.frx":0000
      Top             =   120
      Width           =   10935
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdConvertir_Click()
    If OptDat(0).value = True Then
        Call ConvertCascos
    ElseIf OptDat(1).value = True Then
        Call ConvertArmaduras
    ElseIf OptDat(2).value = True Then
        Call ConvertEscudos
    ElseIf OptDat(3).value = True Then
        Call ConvertArmas
    End If
End Sub

Private Sub Form_Load()

End Sub
