VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmAPI 
   BackColor       =   &H00C0C0C0&
   Caption         =   "API"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "FrmAPI.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Caption         =   "APIs"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   3375
      Begin VB.ListBox LstAPI 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6300
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Progress"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   3600
      TabIndex        =   4
      Top             =   0
      Width           =   3375
      Begin VB.ListBox LstPrg 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6300
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   3135
      End
   End
   Begin MSComctlLib.ProgressBar PrgBar 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   7920
      Visible         =   0   'False
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton CmdGetList 
      Caption         =   "Get APIs List && Save Them"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7200
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label lblDes 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   7320
      TabIndex        =   8
      Top             =   1560
      Width           =   45
   End
   Begin VB.Shape Shp 
      BackStyle       =   1  'Opaque
      Height          =   5535
      Left            =   7200
      Top             =   1320
      Width           =   4455
   End
   Begin VB.Label lblCounter 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   7560
      Width           =   45
   End
   Begin VB.Label lblNum 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   7200
      Width           =   45
   End
End
Attribute VB_Name = "FrmAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'                       Get API-Guide Functions
'                       Written by: Behrouz Rad
'                       Copyright: November 2004
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'Disclaimer:
'I Take No Responsibility For Use This Code.
'
'Description:
'API-Guide is a Program with a API-database which contain all the API-functions.
'API-Guide was created by the KPD-Team 1998-2004
'For more information, visit the AllAPI.net at http://www.allapi.net/
'or The AllAPI Network website at http://www.allapi.net/php/redirect/redirect.php?1
'The API-Guide download page is located at:
'http://www.allapi.net/php/redirect/redirect.php?10
'
'This Program Can Get All API Functions of API-Guide Program With All Details of
'Them and Save into an Access Database.
'i'm Sorry From KPD-Team For This Job But i Need to API-Guide's Database.
'Before Running This Program, Run API-Guide Program.
'You Can Get it From: http://www.allapi.net/
'ENJOY AND PLEASE VOTE...
'*********************************************************************

Public Function IsAppOpen(ByVal XClassName As String) As Boolean
Dim CheckX As Long
CheckX = FindWindow(XClassName, vbNullString)
'Simple >>> IsAppOpen = CheckX
If CheckX = 0 Then
   IsAppOpen = False
Else
   IsAppOpen = True
End If
End Function

Public Function GetTextContain(ByVal HandleX As Long) As String
Dim XParent, XChild As Long, XSubChild As Long, Res As Long
Dim Result As String
Dim MyText As String

Res = SendMessageLong(HandleX&, WM_GETTEXTLENGTH, 0&, 0&)
Result = String(Res + 1, " ")
Call SendMessageByString(HandleX&, WM_GETTEXT, Res + 1, Result)
MyText = Left(Result, Res)
GetTextContain = MyText
End Function

Private Sub CmdGetList_Click()
On Error GoTo ErrDef
Dim Cnn As ADODB.Connection
Dim Rst As ADODB.Recordset
Dim I As Integer
Dim XParent As Long, XChild As Long, XChild2 As Long, _
    XDet1 As Long, XDet2 As Long, TL As Long
Dim XList As Long
Dim XMinOS As Long
Dim XReturn As Long
Dim XLibrary As Long
Dim XExample As Long
Dim XParameter As Long
Dim XDeclaration As Long
Dim XDescription As Long
Dim TheText As String

CmdGetList.Enabled = False
LstPrg.Clear
LstAPI.Clear
Screen.MousePointer = vbHourglass
LstPrg.AddItem "Getting Primary Windows Handles..."

XParent = FindWindow("ThunderRT5Form", vbNullString)
XChild = FindWindowEx(XParent, 0&, "ThunderRT5PictureBox", vbNullString)
XChild = FindWindowEx(XParent, XChild, "ThunderRT5PictureBox", vbNullString)
XChild = FindWindowEx(XParent, XChild, "ThunderRT5PictureBox", vbNullString)
XDet1 = XChild
XChild = FindWindowEx(XParent, XChild, "ThunderRT5PictureBox", vbNullString)
XChild = FindWindowEx(XParent, XChild, "ThunderRT5PictureBox", vbNullString)
XDet2 = XChild
XChild = FindWindowEx(XParent, XChild, "ThunderRT5PictureBox", vbNullString)
LstPrg.AddItem "Done"
LstPrg.AddItem "Getting List Box Handle..."
XChild2 = FindWindowEx(XParent, XChild, "ThunderRT5PictureBox", vbNullString)
XChild2 = FindWindowEx(XParent, XChild2, "ThunderRT5PictureBox", vbNullString)
'Get Stored Items in List Box
XList = FindWindowEx(XChild, 0&, "ThunderRT5ListBox", vbNullString)
LstPrg.AddItem "Done"
'Get Return Values
LstPrg.AddItem "Getting Return Values Handle..."
XReturn = FindWindowEx(XDet1, ByVal 0&, "ThunderRT5TextBox", vbNullString)
LstPrg.AddItem "Done"
'Get Parameters
LstPrg.AddItem "Getting Parameters Handle..."
XParameter = FindWindowEx(XDet1, XReturn, "ThunderRT5TextBox", vbNullString)
LstPrg.AddItem "Done"
'Get Declaration
LstPrg.AddItem "Getting Declaration Handle..."
XDeclaration = FindWindowEx(XDet1, XParameter, "ThunderRT5TextBox", vbNullString)
LstPrg.AddItem "Done"
'Get Minimum Required OS (1)
LstPrg.AddItem "Getting Min Required OS Handle..."
XMinOS = FindWindowEx(XChild2, ByVal 0&, "ThunderRT5TextBox", vbNullString)
'Get API Description
XDescription = XMinOS
'Get Minimum Required OS (2)
XMinOS = FindWindowEx(XChild2, XMinOS, "ThunderRT5TextBox", vbNullString)
LstPrg.AddItem "Done"
'Get Library
LstPrg.AddItem "Getting Library Handle..."
XLibrary = FindWindowEx(XChild2, XMinOS, "ThunderRT5TextBox", vbNullString)
LstPrg.AddItem "Done"
'Get Example
LstPrg.AddItem "Getting Example Handle..."
XExample = FindWindowEx(XDet2, ByVal 0&, "ThunderRT5UserControl", vbNullString)
XExample = FindWindowEx(XExample, ByVal 0&, "RichEdit20A", vbNullString)
LstPrg.AddItem "Done"
LstPrg.AddItem "Getting Number of Items..."
'Get Number of Functions
TL = SendMessageLong(XList, LB_GETCOUNT, 0&, 0&)

LstPrg.AddItem "Done"
LstPrg.AddItem "Set Position..."
'-------------------------------------------------------
'Set Position to First Item of List Box
Call SendMessageLong(XList, LB_SETCURSEL, 1, 0&)
Call SendMessageLong(XList, WM_KEYDOWN, VK_UP, 0&) 'Require (Test This!)
Call SendMessageLong(XList, WM_KEYUP, VK_UP, 0&)   'Require (Test This!)
'-------------------------------------------------------
LstPrg.AddItem "Done"
LstPrg.AddItem "Creating Buffer..."
'Create a Buffer (MAX = 30 Byte >>> Fill With Space Character)
TheText = String(30, " ")
LstPrg.AddItem "Done"
PrgBar.Max = TL
lblNum.Caption = CStr(TL) & " Functions Found!"
LstPrg.AddItem "Creating Connection..."
DoEvents
Set Cnn = New ADODB.Connection
Set Rst = New ADODB.Recordset
Cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & ".\DBAPI.mdb;"
Rst.Open "tblAPI", Cnn, adOpenKeyset, adLockOptimistic
LstPrg.AddItem "Done"
LstPrg.AddItem "Proccessing..."
PrgBar.Visible = True
DoEvents
For I = 0 To TL - 1
   
   Call SendMessageByString(XList, LB_GETTEXT, I, TheText)
   
   TheText = Left(TheText, 30)
   LstAPI.AddItem TheText

   Rst.AddNew
   Rst!StrName = TheText
   Rst!StrLibrary = GetTextContain(XLibrary)
   Rst!StrMinOS = GetTextContain(XMinOS)
   Rst!StrDeclaration = GetTextContain(XDeclaration)
   Rst!StrParameter = GetTextContain(XParameter)
   Rst!StrReturn = GetTextContain(XReturn)
   Rst!StrBody = GetTextContain(XDescription)
   Rst!StrExample = GetTextContain(XExample)
   Rst.Update

   lblCounter.Caption = "Item " & I + 1 & " From " & Str(TL)
   PrgBar.Value = I
   'Move to the Next Item
   Call SendMessageLong(XList, WM_KEYDOWN, VK_DOWN, 0&)
   DoEvents
Next
Rst.Close
Cnn.Close
Set Rst = Nothing
Set Cnn = Nothing
LstPrg.AddItem "Done Successfuly"
Screen.MousePointer = vbDefault
CmdGetList.Enabled = True
MsgBox "Done Successfuly", vbInformation, "API"
Exit Sub
ErrDef:
    LstPrg.AddItem "Error! Operation Was Cancelled"
    Screen.MousePointer = vbDefault
    CmdGetList.Enabled = True
    If Not Rst Is Nothing Then
       If Rst.State = adStateOpen Then Rst.Close
    End If
    Set Rst = Nothing

    If Not Cnn Is Nothing Then
       If Cnn.State = adStateOpen Then Cnn.Close
    End If
    Set Cnn = Nothing

    MsgBox Err.Description, vbCritical, "API"
    End
End Sub

Private Sub Form_Initialize()
'CALL THIS BEFORE ANY CODE:

'Either call it in the Startup forms Form_Initialize
'Event or better still from Sub Main

InitCommonControlsXP
End Sub

Private Sub Form_Load()
If IsAppOpen("ThunderRT5Form") = False Then
   If MsgBox("API-Guide Program is not running." & vbCr & _
          "Please run it." & vbCr & _
          "you can get it from: http://www.allapi.net/" & vbCr & vbCr & _
          "Do You Want to go to the ALL-API?", vbExclamation + vbYesNo, "API") = vbYes Then
          ShellExecute Me.hwnd, "open", "http://www.allapi.net/", vbNullString, vbNullString, 1
   End If
   End
End If

lblDes.Caption = "Disclaimer:" & vbCr & _
"I Take No Responsibility For Use This Code." & vbCr & _
vbCr & "Description:" & vbCr & _
"API-Guide is a Program with a API-database which contain" & vbCr & _
"all the API-functions." & vbCr & _
"API-Guide was created by the KPD-Team 1998-2004" & vbCr & _
"For more information, visit the AllAPI.net at:" & vbCr & _
" http://www.allapi.net/" & vbCr & _
"or The AllAPI Network website at:" & vbCr & _
"http://www.allapi.net/php/redirect/redirect.php?1" & vbCr & _
"The API-Guide download page is located at:" & vbCr & _
"http://www.allapi.net/php/redirect/redirect.php?10" & vbCr & _
 vbCr & _
"This Program Can Get All API Functions of API-Guide" & vbCr & _
" Program With All Details of Them and Save into an Access" & vbCr & _
"Database." & vbCr & _
"i'm Sorry From KPD-Team For This Job But i Need to" & vbCr & _
"API-Guide's Database." & vbCr & _
"Before Running This Program, Run API-Guide Program." & vbCr & _
"You Can Get it From: http://www.allapi.net/" & vbCr & vbCr & _
"ENJOY AND PLEASE VOTE..."
End Sub
