VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "System Tray Icon v2.03"
   ClientHeight    =   4425
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   6180
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   295
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   412
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Lisensing"
      Height          =   1455
      Left            =   675
      TabIndex        =   1
      Top             =   2295
      Width           =   4785
      Begin VB.TextBox txtUser 
         Height          =   330
         Left            =   1080
         TabIndex        =   3
         Top             =   495
         Width           =   3480
      End
      Begin VB.TextBox txtCode 
         Height          =   330
         Left            =   1080
         TabIndex        =   2
         Top             =   945
         Width           =   3480
      End
      Begin VB.Label lblLicence 
         AutoSize        =   -1  'True
         Caption         =   "Not Lisensed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   2205
         TabIndex        =   7
         Top             =   150
         Width           =   1380
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Control Status:"
         Height          =   195
         Left            =   1125
         TabIndex        =   6
         Top             =   180
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Code:"
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Top             =   990
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   195
         Left            =   360
         TabIndex        =   4
         Top             =   585
         Width           =   465
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Registered"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2340
      TabIndex        =   0
      Top             =   3870
      Width           =   1530
   End
   Begin VB.Label Label9 
      Caption         =   "Questions or Comments? Email to: tuniks@hotmail.com"
      Height          =   285
      Left            =   1080
      TabIndex        =   12
      Top             =   2070
      Width           =   3975
   End
   Begin VB.Label Label8 
      Caption         =   "http://www.lens.spb.ru~tunik"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   1530
      TabIndex        =   11
      Top             =   1800
      Width           =   2850
   End
   Begin VB.Label Label7 
      Caption         =   $"frmAboutPbarS.frx":0000
      Height          =   375
      Left            =   315
      TabIndex        =   10
      Top             =   1350
      Width           =   5460
   End
   Begin VB.Label Label6 
      Caption         =   $"frmAboutPbarS.frx":009D
      Height          =   600
      Left            =   225
      TabIndex        =   9
      Top             =   675
      Width           =   5640
   End
   Begin VB.Label Label5 
      Caption         =   "ProgressBarSlider Pro ActiveX © 2000 by Nik Tupkalov          This ActiveX Control was writen by Nik Tupkalov"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   585
      TabIndex        =   8
      Top             =   90
      Width           =   4920
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
'------------------------------------------------------------------------
' Êîä ïðè ðåãèñòðàöèè ïîëó÷àåòñÿ ïóò¸ì ïðåîáðàçîâàíèÿ ñèìâîëîâ
' èìåíè ïîëüçîâàòåëÿ â äåñÿòè÷íûé ýêâèâàëåíò è ïðåîáðàçîâàíèåì
' â ñèìâîëüíóþ ñòðîêó ñ îòñå÷åíèåì 12 ñèìâîëîâ ñëåâà. Ïðèìåð:
' Nik Tupkalov (781051073284......) - îñòàëüíûå íå íàäî.
'-------------------------------------------------------------------------
'
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal szData As String, ByVal cbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long

Const HKCR = &H80000000

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
    End Type

Public Function bSetRegValue(ByVal hKey As Long, ByVal lpszSubKey As String, ByVal sSetValue As String, ByVal sValue As String) As Boolean
    On Error Resume Next
    
    Dim phkResult As Long
    Dim lResult As Long
    Dim SA As SECURITY_ATTRIBUTES
    Dim lCreate As Long
    
    RegCreateKeyEx hKey, lpszSubKey, 0, "", 0, &H3F, SA, phkResult, lCreate
    lResult = RegSetValueEx(phkResult, sSetValue, 0, (1), sValue, CLng(Len(sValue) + 1))
    RegCloseKey phkResult
    bSetRegValue = (lResult = 0)
End Function

Public Function bGetRegValue(ByVal hKey As Long, ByVal sKey As String, ByVal sSubKey As String) As String
Dim lResult As Long, phkResult As Long, dWReserved As Long
Dim szBuffer As String, lBuffSize As Long, szBuffer2 As String
Dim lBuffSize2 As Long, lIndex As Long, lType As Long, sCompKey As String
    
lIndex = 0
lResult = RegOpenKeyEx(hKey, sKey, 0, 1, phkResult)

    Do While lResult = 0 'And Not (bFound)
        szBuffer = Space(255)
        lBuffSize = Len(szBuffer)
        szBuffer2 = Space(255)
        lBuffSize2 = Len(szBuffer2)
        lResult = RegEnumValue(phkResult, lIndex, szBuffer, lBuffSize, dWReserved, lType, szBuffer2, lBuffSize2)

        If (lResult = 0) Then
            sCompKey = Left(szBuffer, lBuffSize)

            If (sCompKey = sSubKey) Then
                bGetRegValue = Left(szBuffer2, lBuffSize2 - 1)
            End If
        End If
        lIndex = lIndex + 1
    Loop
    RegCloseKey phkResult
End Function

Private Sub cmdOK_Click()
If Not bGetRegValue(HKCR, "CLSID\{00000000-0000-0078-1051-073284000000}", "Licence") = Empty Then
    Unload Me
Else

Static i As Integer, strR As String

   If txtUser.Text = Empty Then Unload Me: Exit Sub
If Len(txtUser.Text) < 6 Then MsgBox "User Name min 6 - Simbols", vbCritical: txtUser.SetFocus: Exit Sub
                For i = 1 To 6
        strR = strR + Mid(Str(Asc(Mid(txtUser.Text, i, 1))), 2)
                Next i
            If txtCode.Text = Left(strR, 12) Then
                MsgBox "Registered OK"
bSetRegValue HKCR, "CLSID\{00000000-0000-0078-1051-073284000000}", "Licence", 1
                Unload Me
            Else
                MsgBox "Registered Information Not Correct", vbCritical
            txtUser.SetFocus
            End If
   End If
End Sub

Private Sub Form_Load()
Me.Caption = App.Title & " v" & App.Major & "." & App.Minor & "." & App.Revision
If Not bGetRegValue(HKCR, "CLSID\{00000000-0000-0078-1051-073284000000}", "Licence") = Empty Then
txtUser.Enabled = False
txtUser.BackColor = vbButtonFace
txtCode.Enabled = False
txtCode.BackColor = vbButtonFace
lblLicence.Caption = "LICENCED"
cmdOK.Caption = "Exit"
Else
lblLicence.Caption = "Not LICENCED"
End If
End Sub
