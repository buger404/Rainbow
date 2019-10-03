VERSION 5.00
Begin VB.Form LogWindow 
   BackColor       =   &H00212121&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LogWindow"
   ClientHeight    =   6864
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   11700
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6864
   ScaleWidth      =   11700
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.TextBox LogBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00212121&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00CEDB1A&
      Height          =   6900
      Left            =   96
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   11604
   End
End
Attribute VB_Name = "LogWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    NewLog "Server started ."
End Sub

