VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmGenerator 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ABC ReadMe Generator"
   ClientHeight    =   10320
   ClientLeft      =   4140
   ClientTop       =   465
   ClientWidth     =   6975
   Icon            =   "frmGenerator.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10320
   ScaleWidth      =   6975
   Begin VB.ListBox lstSoftwareType 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmGenerator.frx":0442
      Left            =   1680
      List            =   "frmGenerator.frx":0458
      TabIndex        =   26
      Top             =   3000
      Width           =   2055
   End
   Begin VB.TextBox txtDescription 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   5160
      Width           =   6735
   End
   Begin VB.TextBox txtReleaseDate 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   1080
      Width           =   5175
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7440
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Reset All Information"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   23
      Top             =   9840
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generate ReadMe File"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   22
      Top             =   9840
      Width           =   3015
   End
   Begin VB.TextBox txtWebsite 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   2520
      Width           =   5175
   End
   Begin VB.ListBox lstUncompatible 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      ItemData        =   "frmGenerator.frx":0491
      Left            =   2400
      List            =   "frmGenerator.frx":04BC
      TabIndex        =   6
      Top             =   3840
      Width           =   2175
   End
   Begin VB.ListBox lstCompatible 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   4680
      TabIndex        =   7
      Top             =   3840
      Width           =   2175
   End
   Begin VB.TextBox txtAdditionalFeatures 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   8520
      Width           =   6735
   End
   Begin VB.TextBox txtNewFeatures 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   6840
      Width           =   6735
   End
   Begin VB.TextBox txtEmail 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   2040
      Width           =   5175
   End
   Begin VB.TextBox txtCompanyName 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   1560
      Width           =   5175
   End
   Begin VB.TextBox txtVersion 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   600
      Width           =   5175
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Software Description :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   25
      Top             =   4800
      Width           =   2085
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Release Date :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   24
      Top             =   1080
      Width           =   1290
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Compatible"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4800
      TabIndex        =   21
      Top             =   3480
      Width           =   990
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Uncompatible"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2520
      TabIndex        =   20
      Top             =   3480
      Width           =   1230
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "New Features :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   19
      Top             =   6480
      Width           =   1380
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Additional Features :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   18
      Top             =   8160
      Width           =   1920
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Company Name :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   17
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Version :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   16
      Top             =   600
      Width           =   795
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Email :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   15
      Top             =   2040
      Width           =   600
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Software Type :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   14
      Top             =   3000
      Width           =   1500
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Website :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   13
      Top             =   2520
      Width           =   870
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Software Compability :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   12
      Top             =   3480
      Width           =   2100
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Name :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

'Declare variables
Dim Name As String
Dim Version As String
Dim CompanyName As String
Dim Email As String
Dim WebSite As String
Dim SoftwareType As String
Dim NewFeatures As String
Dim ReleaseDate As String
Dim AdditionalFeatures As String
Dim FileName As String
Dim Description As String
Dim Compatible As New Collection
Dim Num As Integer

'Check if the text boxes are empty or not
If txtName.Text = "" Then
MsgBox "Please enter the name of the software.", , "ABC ReadMe Generator"
GoTo SaveError
ElseIf txtVersion.Text = "" Then
MsgBox "Please enter the version of the software.", , "ABC ReadMe Generator"
GoTo SaveError
ElseIf txtCompanyName.Text = "" Then
MsgBox "Please the company name of the software.", , "ABC ReadMe Generator"
GoTo SaveError
ElseIf txtEmail.Text = "" Then
MsgBox "Please enter the email address of the company.", , "ABC ReadMe Generator"
GoTo SaveError
ElseIf txtWebsite.Text = "" Then
MsgBox "Please enter the website of the software.", , "ABC ReadMe Generator"
GoTo SaveError
ElseIf txtDescription.Text = "" Then
MsgBox "Please enter the description of the software.", , "ABC ReadMe Generator"
GoTo SaveError
ElseIf txtNewFeatures.Text = "" Then
MsgBox "Please enter the new features of the software.", , "ABC ReadMe Generator"
GoTo SaveError
ElseIf txtAdditionalFeatures.Text = "" Then
MsgBox "Please enter the additional features of the software.", , "ABC ReadMe Generator"
GoTo SaveError
ElseIf txtReleaseDate.Text = "" Then
MsgBox "Please enter the release date of the software.", , "ABC ReadMe Generator"
GoTo SaveError
End If

'Assign value for variables
Name = txtName.Text
Version = txtVersion.Text
CompanyName = txtCompanyName.Text
Email = txtEmail.Text
WebSite = txtWebsite.Text
Description = txtDescription.Text
Num = lstSoftwareType.ListIndex
SoftwareType = lstSoftwareType.List(Num)
NewFeatures = txtNewFeatures.Text
ReleaseDate = txtReleaseDate.Text
AdditionalFeatures = txtAdditionalFeatures.Text

'Display a input box to get the path to save the readme file.
FileName = InputBox("Please enter the path that this readme file should be saved :", "ABC ReadMe Generator")
'Write file
Open FileName & "\ReadMe.txt" For Output As #1
Print #1, ""
Print #1, "             ***************************************************"
Print #1, "                     " & Trim(Name) & " " & Version & " (" & SoftwareType & ")"
Print #1, ""
Print #1, "              Release Date : " & ReleaseDate
Print #1, "              By : " & Trim(CompanyName)
Print #1, "              Homepage : " & Trim(WebSite)
Print #1, "              Contact : " & Trim(Email)
Print #1, "             ***************************************************"
Print #1, ""
Print #1, "Thank you for using " & Trim(Name) & "!"
Print #1, "This readme file contains the lastest information about " & Trim(Name) & " " & Version
Print #1, ""
Print #1, "- Description"
Print #1, "- What's New in This Version"
Print #1, "- Additional Features"
Print #1, "- Software Compability"
Print #1, "- End User License Agreement"
Print #1, "- Bug Reports"
Print #1, "- Contact"
Print #1, ""
Print #1, ""
Print #1, "*****************************************"
Print #1, "DESCRIPTION"
Print #1, "*****************************************"
Print #1, Description
Print #1, ""
Print #1, "*****************************************"
Print #1, "What's New in This Version"
Print #1, "*****************************************"
Print #1, NewFeatures
Print #1, ""
Print #1, "*****************************************"
Print #1, "Additional Features"
Print #1, "*****************************************"
Print #1, AdditionalFeatures
Print #1, ""
Print #1, "*****************************************"
Print #1, "Software Compability"
Print #1, "*****************************************"
Print #1, "Compatible Operating System : "
For I = 0 To lstCompatible.ListCount
    Compatible.Add lstCompatible.List(I)
Next
For I = 1 To Compatible.Count - 1
Print #1, "*" & Compatible.Item(I)
Next
Print #1, ""
Print #1, "*****************************************"
Print #1, "End User License Agreement"
Print #1, "*****************************************"
Print #1, "  THIS SOFTWARE IS PROVIDED AS-IS, WITHOUT WARRANTY OF ANY KIND, " _
; "EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE " _
; "IMPLIED WARRANTIES OF MERCHANT ABILITY AND/OR FITNESS FOR " _
; "A PARTICULAR PURPOSE. THE AUTHOR SHALL NOT BE HELD LIABLE FOR " _
; "ANY DAMAGE TO YOU, YOUR COMPUTER, OR TO ANYONE OR ANYTHING ELSE, " _
; "THAT MAY RESULT FROM ITS USE, OR MISUSE."
Print #1, ""
Print #1, "All trademarks and other registered names contained in the " _
; appname & " package are the property of their " _
; "respective owners."
Print #1, ""
Print #1, "*****************************************"
Print #1, "Bug Reports"
Print #1, "*****************************************"
Print #1, "*Please report any bug by e-mail to:"
Print #1, "           " & Email
Print #1, ""
Print #1, "*****************************************"
Print #1, "Contacts"
Print #1, "*****************************************"
Print #1, "Contact : " & CompanyName
Print #1, "Email : " & Email
Print #1, "Website : " & WebSite
Close
Exit Sub
SaveError:
End Sub

Private Sub Command2_Click()
txtName.Text = ""
txtVersion = ""
txtEmail.Text = ""
txtWebsite.Text = ""
txtDescription.Text = ""
txtNewFeatures.Text = ""
txtAdditionalFeatures.Text = ""
txtReleaseDate.Text = ""
End Sub

Private Sub lstCompatible_DblClick()
On Error Resume Next
        Dim a As String
a = lstCompatible.Text
If a <> "" Then
lstUncompatible.AddItem lstCompatible.Text
For I = 0 To lstCompatible.ListCount
If lstCompatible.List(I) = a Then lstCompatible.RemoveItem (I)
Next
Else
End If
End Sub

Private Sub lstUncompatible_DblClick()
On Error Resume Next
Dim a As String
a = lstUncompatible.Text
If a <> "" Then
lstCompatible.AddItem lstUncompatible.Text
For I = 0 To lstUncompatible.ListCount
If lstUncompatible.List(I) = a Then lstUncompatible.RemoveItem (I)
Next
Else
End If
End Sub

Private Sub txtVersion_KeyPress(KeyAscii As Integer)
 Select Case KeyAscii
    Case 8, 46, 48 To 57: Exit Sub             'Allows only numbers to be typed
    Case Else: KeyAscii = 0
End Select
End Sub
