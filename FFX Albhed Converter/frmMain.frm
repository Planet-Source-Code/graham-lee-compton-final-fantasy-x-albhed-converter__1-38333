VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " :: Final Fantasy X : Albhed converter (1.30)"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4815
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   Icon            =   "frmMain.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   4815
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtAlbhed 
      Height          =   1185
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmMain.frx":08CA
      Top             =   360
      Width           =   4605
   End
   Begin VB.CheckBox chkAuto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "Auto"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2790
      TabIndex        =   11
      ToolTipText     =   "Automatically convert"
      Top             =   90
      Width           =   735
   End
   Begin VB.CommandButton cmdMin 
      Caption         =   "_"
      Height          =   465
      Left            =   3510
      TabIndex        =   10
      Top             =   3150
      Width           =   465
   End
   Begin VB.TextBox txtEnglish 
      Height          =   1185
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1890
      Width           =   4605
   End
   Begin VB.CheckBox chkUpdate2 
      Appearance      =   0  'Flat
      Caption         =   "Update"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3600
      TabIndex        =   9
      Top             =   1620
      Width           =   1005
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   465
      Left            =   2610
      TabIndex        =   8
      Top             =   3150
      Width           =   825
   End
   Begin VB.CommandButton cmdAlgEng 
      Caption         =   "Albhed -> English"
      Height          =   465
      Left            =   1350
      TabIndex        =   7
      Top             =   3150
      Width           =   1185
   End
   Begin VB.CommandButton cmdEngAlg 
      Caption         =   "English -> Albhed"
      Height          =   465
      Left            =   90
      TabIndex        =   6
      Top             =   3150
      Width           =   1185
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   465
      Left            =   4050
      TabIndex        =   5
      Top             =   3150
      Width           =   645
   End
   Begin VB.CheckBox chkUpdate 
      Appearance      =   0  'Flat
      Caption         =   "Update"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3600
      TabIndex        =   2
      Top             =   90
      Width           =   1005
   End
   Begin VB.Label Label2 
      Caption         =   "Type your English text here"
      Height          =   285
      Left            =   90
      TabIndex        =   4
      Top             =   1620
      Width           =   2715
   End
   Begin VB.Label Label1 
      Caption         =   "Type your Al bhed text here"
      Height          =   285
      Left            =   90
      TabIndex        =   1
      Top             =   90
      Width           =   2715
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Made by Graham Lee Compton
'----------------------------
' If you know any better way of doing this process
' I would love to know how.
' Sources appriciated

Private Sub chkUpdate2_Click()
 ' If Albhed update is checked then uncheck it
 If chkUpdate2.Value = 1 Then chkUpdate.Value = 0
End Sub

Private Sub chkUpdate_Click()
 ' If English update is checked then uncheck it
 If chkUpdate.Value = 1 Then chkUpdate2.Value = 0
End Sub

Private Sub cmdAbout_Click()
 ' My little about message
 Dim msg As String
 msg = "made by GraZC" & vbCrLf & vbCrLf
 msg = msg + "www.grazc.com"
 MsgBox msg
End Sub

Private Sub cmdAlgEng_Click()
 ' If theres no text then do nothing but beep
 If txtAlbhed.Text = "" Then Exit Sub: Beep
 
 ' Clears the english box
 txtEnglish.Text = ""
  
  Dim a As Long
  ' Create a loop to the lengh of the Albhed message
  For a = 1 To Len(txtAlbhed.Text)
   Dim strX As String
   Dim strOk As Boolean
   ' This reads every single character in the loop
   strX = LCase(Mid(txtAlbhed.Text, a, 1))
   strOk = False
  '           /--- Albhed Letters            to English --\
   If strX = "y" Then txtEnglish.Text = txtEnglish.Text & "a": strOk = True
   If strX = "p" Then txtEnglish.Text = txtEnglish.Text & "b": strOk = True
   If strX = "l" Then txtEnglish.Text = txtEnglish.Text & "c": strOk = True
   If strX = "t" Then txtEnglish.Text = txtEnglish.Text & "d": strOk = True
   If strX = "a" Then txtEnglish.Text = txtEnglish.Text & "e": strOk = True
   If strX = "v" Then txtEnglish.Text = txtEnglish.Text & "f": strOk = True
   If strX = "k" Then txtEnglish.Text = txtEnglish.Text & "g": strOk = True
   If strX = "r" Then txtEnglish.Text = txtEnglish.Text & "h": strOk = True
   If strX = "e" Then txtEnglish.Text = txtEnglish.Text & "i": strOk = True
   If strX = "z" Then txtEnglish.Text = txtEnglish.Text & "j": strOk = True
   If strX = "g" Then txtEnglish.Text = txtEnglish.Text & "k": strOk = True
   If strX = "m" Then txtEnglish.Text = txtEnglish.Text & "l": strOk = True
   If strX = "s" Then txtEnglish.Text = txtEnglish.Text & "m": strOk = True
   If strX = "h" Then txtEnglish.Text = txtEnglish.Text & "n": strOk = True
   If strX = "u" Then txtEnglish.Text = txtEnglish.Text & "o": strOk = True
   If strX = "b" Then txtEnglish.Text = txtEnglish.Text & "p": strOk = True
   If strX = "x" Then txtEnglish.Text = txtEnglish.Text & "q": strOk = True
   If strX = "n" Then txtEnglish.Text = txtEnglish.Text & "r": strOk = True
   If strX = "c" Then txtEnglish.Text = txtEnglish.Text & "s": strOk = True
   If strX = "d" Then txtEnglish.Text = txtEnglish.Text & "t": strOk = True
   If strX = "i" Then txtEnglish.Text = txtEnglish.Text & "u": strOk = True
   If strX = "j" Then txtEnglish.Text = txtEnglish.Text & "v": strOk = True
   If strX = "f" Then txtEnglish.Text = txtEnglish.Text & "w": strOk = True
   If strX = "q" Then txtEnglish.Text = txtEnglish.Text & "x": strOk = True
   If strX = "o" Then txtEnglish.Text = txtEnglish.Text & "y": strOk = True
   If strX = "w" Then txtEnglish.Text = txtEnglish.Text & "z": strOk = True
   ' If strOK isnt true, then its not a letter to translate
   ' So just put the character across
   If strOk = False Then txtEnglish.Text = txtEnglish.Text & strX
   ' Move onto the next character
  Next a
End Sub

Private Sub cmdEngAlg_Click()
 ' If theres no text then do nothing but beep
 If txtEnglish.Text = "" Then Exit Sub: Beep
 
 ' Read above, Its just the same process but reversed
 txtAlbhed.Text = ""
 
  Dim a As Long
  For a = 1 To Len(txtEnglish.Text)
   Dim strX As String
   Dim strOk As Boolean
   strX = LCase(Mid(txtEnglish.Text, a, 1))
   strOk = False
   If strX = "a" Then txtAlbhed.Text = txtAlbhed.Text & "y": strOk = True
   If strX = "b" Then txtAlbhed.Text = txtAlbhed.Text & "p": strOk = True
   If strX = "c" Then txtAlbhed.Text = txtAlbhed.Text & "l": strOk = True
   If strX = "d" Then txtAlbhed.Text = txtAlbhed.Text & "t": strOk = True
   If strX = "e" Then txtAlbhed.Text = txtAlbhed.Text & "a": strOk = True
   If strX = "f" Then txtAlbhed.Text = txtAlbhed.Text & "v": strOk = True
   If strX = "g" Then txtAlbhed.Text = txtAlbhed.Text & "k": strOk = True
   If strX = "h" Then txtAlbhed.Text = txtAlbhed.Text & "r": strOk = True
   If strX = "i" Then txtAlbhed.Text = txtAlbhed.Text & "e": strOk = True
   If strX = "j" Then txtAlbhed.Text = txtAlbhed.Text & "z": strOk = True
   If strX = "k" Then txtAlbhed.Text = txtAlbhed.Text & "g": strOk = True
   If strX = "l" Then txtAlbhed.Text = txtAlbhed.Text & "m": strOk = True
   If strX = "m" Then txtAlbhed.Text = txtAlbhed.Text & "s": strOk = True
   If strX = "n" Then txtAlbhed.Text = txtAlbhed.Text & "h": strOk = True
   If strX = "o" Then txtAlbhed.Text = txtAlbhed.Text & "u": strOk = True
   If strX = "p" Then txtAlbhed.Text = txtAlbhed.Text & "b": strOk = True
   If strX = "q" Then txtAlbhed.Text = txtAlbhed.Text & "x": strOk = True
   If strX = "r" Then txtAlbhed.Text = txtAlbhed.Text & "n": strOk = True
   If strX = "s" Then txtAlbhed.Text = txtAlbhed.Text & "c": strOk = True
   If strX = "t" Then txtAlbhed.Text = txtAlbhed.Text & "d": strOk = True
   If strX = "u" Then txtAlbhed.Text = txtAlbhed.Text & "i": strOk = True
   If strX = "v" Then txtAlbhed.Text = txtAlbhed.Text & "j": strOk = True
   If strX = "w" Then txtAlbhed.Text = txtAlbhed.Text & "f": strOk = True
   If strX = "x" Then txtAlbhed.Text = txtAlbhed.Text & "q": strOk = True
   If strX = "y" Then txtAlbhed.Text = txtAlbhed.Text & "o": strOk = True
   If strX = "z" Then txtAlbhed.Text = txtAlbhed.Text & "w": strOk = True
   If strOk = False Then txtAlbhed.Text = txtAlbhed.Text & strX
  Next a
End Sub

Private Sub cmdExit_Click()
 End
End Sub

Private Sub cmdMin_Click()
 ' Minimise
 Me.WindowState = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
 End
End Sub

Private Sub txtAlbhed_Change()
 ' Clicks the Albhed to English Button
 If chkUpdate.Value = 1 Then Call cmdAlgEng_Click
End Sub

Private Sub txtEnglish_Change()
 ' Clicks the English to Albhed button
 If chkUpdate2.Value = 1 Then Call cmdEngAlg_Click
End Sub

Private Sub txtAlbhed_GotFocus()
 ' If auto update is checked then tick Update (so this automatically converts it)
 If chkAuto.Value = 1 Then
  chkUpdate.Value = 1
  chkUpdate2.Value = 0
  Call cmdAlgEng_Click
 End If
End Sub

Private Sub txtEnglish_GotFocus()
 ' If auto update is checked then tick Update (so this automatically converts it)
 If chkAuto.Value = 1 Then
  chkUpdate.Value = 0
  chkUpdate2.Value = 1
  Call cmdEngAlg_Click
 End If
End Sub

