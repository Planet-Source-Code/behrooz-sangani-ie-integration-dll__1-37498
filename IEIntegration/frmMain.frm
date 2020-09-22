VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IE Integration [Imported Information From IE Page]"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8415
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   8415
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Height          =   550
      Left            =   7680
      Picture         =   "frmMain.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton cmdAbout 
      Height          =   550
      Left            =   6960
      Picture         =   "frmMain.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3000
      Width           =   615
   End
   Begin VB.TextBox txtSel 
      Height          =   1215
      Left            =   960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   1560
      Width           =   7335
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   1200
      Width           =   7095
      Begin VB.OptionButton Opt 
         Caption         =   "Image"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   7
         Top             =   0
         Width           =   1335
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Link"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.TextBox txtURL 
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   840
      Width           =   7335
   End
   Begin VB.TextBox txtPageURL 
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   7335
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000011&
      X1              =   1800
      X2              =   6840
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "IE Integration DLL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3360
      Width           =   2655
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   6840
      X2              =   1800
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Label Label4 
      Caption         =   "Selected Text:"
      Height          =   975
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Type:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "URL:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Page URL:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================================
'  Main Form
'  This form is loaded by ShowInfo Sub in the class
'=========================================================================================
'  Created By: Behrooz Sangani <bs20014@yahoo.com>
'  Published Date: 01/08/2002
'  WebSite: http://www.geocities.com/bs20014
'  Legal Copyright: Behrooz Sangani © 01/08/2002
'=========================================================================================

Private Sub cmdAbout_Click()
    MsgBox "Programmed by: Behrooz Sangani" & vbCrLf & _
        "Email: bs20014@yahoo.com" & vbCrLf & _
        "Web: http://www.geocities.com/bs20014/" & vbCrLf & vbCrLf & _
        "You can use this method of getting information to integrate your program with IE." & vbCrLf & _
        "Use for free, no obligation." & vbCrLf & "Read more information in article", vbInformation, "IE Integration Demo"
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub
