VERSION 5.00
Begin VB.Form frmrgtclkdll 
   BackColor       =   &H80000004&
   Caption         =   "Right Click Dll"
   ClientHeight    =   2475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   ScaleHeight     =   2475
   ScaleWidth      =   4590
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt_menu_name 
      Height          =   285
      Left            =   2280
      TabIndex        =   5
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox txt_prog_name 
      Height          =   285
      Left            =   2280
      TabIndex        =   3
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox txt_file_ext 
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Top             =   360
      Width           =   2175
   End
   Begin VB.CommandButton cmdreg 
      Caption         =   "Press this to proceed  :"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label lblmenuname 
      Caption         =   "Enter the menu name :"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label lblprogname 
      Caption         =   "Enter the program name :"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label lblfileext 
      Caption         =   "Enter file ext :"
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frmrgtclkdll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ===========================================================================
'    DATE      NAME                      DESCRIPTION
' -----------  ------------------------  ------------------------------------
' 20-MAR-2002  Himanshu Dhami              Written by himansh_dhami@yahoo.com
' References  ::Help in API for regsitry handling is borrowed from VBAPI.com
                ' Read the readme file to get the proper use of it.
               
' ---------------------------------------------------------------------------

Public file_ext As String
'Public file_type As String
Public prog_name As String
Public menu_name As String
Public filetype As String
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const KEY_WRITE = &H20006
Private Const KEY_READ = &H20019
Private Const REG_SZ = 1
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
'************ for accessing the registry value
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal _
    hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired _
    As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
    (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
    lpType As Long, lpData As Any, lpcbData As Long) As Long
'***********************************************
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal _
    hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass _
    As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes _
    As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
    
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal _
    hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType _
    As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Sub FindFileType()
Dim hKey As Long  ' receives a handle to the newly created or opened registry key
    Dim subkey As String  ' name of the subkey to open
    Dim stringbuffer As String  ' receives data read from the registry
    Dim datatype As Long  ' receives data type of read value
    Dim slength As Long  ' receives length of returned data
    Dim retval As Long  ' return value

    ' Set the name of the new key and the default security settings
    'subkey = "Software\MyCorp\MyProgram\Config"
     'subkey = "dllfile\shell\Register\command"
     subkey = "" & file_ext & ""
    ' Create or open the registry key
    retval = RegOpenKeyEx(HKEY_CLASSES_ROOT, subkey, 0, KEY_READ, hKey)
    If retval <> 0 Then
        Debug.Print "ERROR: Unable to open registry key!"
        Exit Sub
    End If

    ' Make room in the buffer to receive the incoming data.
    stringbuffer = Space(255)
    slength = 255
    ' Read the "username" value from the registry key.
    retval = RegQueryValueEx(hKey, "", 0, datatype, ByVal stringbuffer, slength)
    ' Only attempt to display the data if it is in fact a string.
    If datatype = REG_SZ Then
        ' Remove empty space from the buffer and display the result.
        stringbuffer = Left(stringbuffer, slength - 1)
        Debug.Print "Username: "; stringbuffer
        MsgBox stringbuffer, vbOKOnly
        filetype = stringbuffer
    Else
        ' Don't bother trying to read any other data types.
        Debug.Print "Data not in string format.  Unable to interpret data."
    End If

    ' Close the registry key.
    retval = RegCloseKey(hKey)
End Sub
' *** Place the following code inside the form. ***
Private Sub RegCreator()
    Dim hKey As Long            ' receives handle to the registry key
    Dim secattr As SECURITY_ATTRIBUTES  ' security settings for the key
    Dim subkey As String        ' name of the subkey to create or open
    Dim neworused As Long       ' receives flag for if the key was created or opened
    Dim stringbuffer As String  ' the string to put into the registry
    Dim retval As Long          ' return value
    Dim unretval As Long
    Dim test As String
    
    ' Set the name of the new key and the default security settings
    subkey = "" & filetype & "\shell\" & menu_name & "\command"
    secattr.nLength = Len(secattr)
    secattr.lpSecurityDescriptor = 0
    secattr.bInheritHandle = 1
    
    ' Create (or open) the registry key.
    retval = RegCreateKeyEx(HKEY_CLASSES_ROOT, subkey, 0, "", 0, KEY_WRITE, _
        secattr, hKey, neworused)
    If retval <> 0 Then
        Debug.Print "Error opening or creating registry key -- aborting."
        Exit Sub
    End If
    ' Write the string to the registry.  Note the use of ByVal in the second-to-last
    ' parameter because we are passing a string.
    stringbuffer = prog_name & " " & """%1""" & vbNullChar    ' the terminating null is necessary"
    retval = RegSetValueEx(hKey, "", 0, REG_SZ, ByVal stringbuffer, _
             Len(stringbuffer))
    retval = RegCloseKey(hKey)
End Sub
Private Sub cmdreg_Click()
file_ext = CStr(txt_file_ext.Text)
prog_name = txt_prog_name.Text
menu_name = txt_menu_name.Text
If file_ext = "" Or prog_name = "" Or menu_name = "" Then
MsgBox "Who is going to fill all the entries", vbOKOnly
txt_file_ext.SetFocus
Else
MsgBox " Menu name is  " & menu_name & " ", vbOKOnly
FindFileType
If filetype = "" Then
MsgBox ("Sorry no file type associated with your extension"), vbOKOnly
End If
RegCreator
MsgBox ("Successfully created string in the registry"), vbOKOnly
End If
End
End Sub
