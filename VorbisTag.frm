VERSION 5.00
Begin VB.Form frmVorbisExample 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "OGG Vorbis Tag Reader"
   ClientHeight    =   1080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Get Tags"
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   600
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Simple OGG Vorbis tag-reader example by Intimidated of UniteTheCows.com"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   135
      TabIndex        =   3
      Top             =   615
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Enter .OGG filename here:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmVorbisExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim p As VorbisTag
Dim tmp As Integer
If LCase$(Right$(Text1.Text, 3)) <> "ogg" Then
tmp = MsgBox("File does not appear to be a Vorbis file! Sure to continue?", 36, "Not .OGG!")
If tmp = 7 Then Exit Sub
End If
p = GetTag(Text1.Text)
If p.Error = "" Then
MsgBox "Title: " + p.Title
MsgBox "Artist: " + p.Artist
MsgBox "Genre: " + p.Genre
MsgBox "TrackNo: " + Format$(p.TrackNumber)
MsgBox "Year: " + p.Date
MsgBox "Album: " + p.Album
MsgBox "EncodedUsing: " + p.EncodedUsing
Else
MsgBox "Error: " + Right$(p.Error, Len(p.Error) - 4)
End If
End Sub

