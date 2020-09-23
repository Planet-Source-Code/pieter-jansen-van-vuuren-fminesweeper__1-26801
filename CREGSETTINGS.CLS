VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRegSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' *********************************************************************
'  Copyright ©1997 Karl E. Peterson, All Rights Reserved
' *********************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code without prior written consent.
' *********************************************************************
' Changes: Pieter van Vuuren
' *********************************************************************
Option Explicit

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Any, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Const HKEY_CURRENT_USER = &H80000001
Private Const REG_OPTION_NON_VOLATILE = 0
Private Const READ_CONTROL = &H20000
Private Const SYNCHRONIZE = &H100000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Private Const ERROR_SUCCESS = 0&
Private Const REG_SZ = 1

Private m_Company As String
Private m_AppName As String

'*******************************************************************************
' Class_Initialize (SUB)
'
' PARAMETERS:
' None
'
' DESCRIPTION:
' ***Description goes here***
'*******************************************************************************
Private Sub Class_Initialize()
   m_Company = App.CompanyName
   m_AppName = App.ProductName
End Sub

'*******************************************************************************
' Company (PROPERTY LET)
'*******************************************************************************
Public Property Let Company(ByVal NewVal As String)
   If Len(NewVal) Then
      m_Company = Trim(NewVal)
   Else
      m_Company = App.CompanyName
   End If
End Property

'*******************************************************************************
' Company (PROPERTY GET)
'*******************************************************************************
Public Property Get Company() As String
   Company = m_Company
End Property

'*******************************************************************************
' AppName (PROPERTY LET)
'*******************************************************************************
Public Property Let AppName(ByVal NewVal As String)
   If Len(NewVal) Then
      m_AppName = Trim(NewVal)
   Else
      m_AppName = App.ProductName
   End If
End Property

'*******************************************************************************
' AppName (PROPERTY GET)
'*******************************************************************************
Public Property Get AppName() As String
   AppName = m_AppName
End Property

'*******************************************************************************
' GetSetting (FUNCTION)
'
' PARAMETERS:
' (In) - Section - String - Required. String expression containing the name of the section where the key setting is found.
'                           If omitted, key setting is assumed to be in default subkey.
' (In) - Key     - String - Required. String expression containing the name of the key setting to return.
' (In) - Default - String - Optional. Expression containing the value to return if no value is set in the key setting.
'                           If omitted, default is assumed to be a zero-length string ("").
'
' RETURN VALUE:
' String -
'
' DESCRIPTION:
' ***Description goes here***
'*******************************************************************************
Public Function GetSetting(ByVal Section As String, ByVal Key As String, Optional ByVal Default As String = "") As String
Dim nRet   As Long
Dim hKey   As Long
Dim nType  As Long
Dim nBytes As Long
Dim Buffer As String
   
   ' Assume failure and set return to Default
   GetSetting = Default

   ' Open key
   nRet = RegOpenKeyEx(HKEY_CURRENT_USER, SubKey(Section), 0&, KEY_ALL_ACCESS, hKey)
   If nRet = ERROR_SUCCESS Then
      ' Set appropriate value for default query
      If Key = "*" Then Key = vbNullString
      
      ' Determine how large the buffer needs to be
      nRet = RegQueryValueEx(hKey, Key, 0&, nType, ByVal Buffer, nBytes)
      If nRet = ERROR_SUCCESS Then
         ' Build buffer and get data
         If nBytes > 0 Then
            Buffer = Space(nBytes)
            nRet = RegQueryValueEx(hKey, Key, 0&, nType, ByVal Buffer, Len(Buffer))
            If nRet = ERROR_SUCCESS Then
               ' Trim NULL and return successful query!
               GetSetting = Left(Buffer, nBytes - 1)
            End If
         End If
      End If
   End If
End Function

'*******************************************************************************
' SaveSetting (FUNCTION)
'
' PARAMETERS:
' (In) - Section - String - Required. String expression containing the name of the section where the key setting is being saved.
' (In) - Key     - String - Required. String expression containing the name of the key setting being saved.
' (In) - Setting - String - Required. Expression containing the value that key is being set to.
'
' RETURN VALUE:
' Boolean - True=Success
'
' DESCRIPTION:
' ***Description goes here***
'*******************************************************************************
Public Function SaveSetting(ByVal Section As String, ByVal Key As String, ByVal Setting As String) As Boolean
Dim nRet    As Long
Dim hKey    As Long
Dim nResult As Long
   
   ' Open (or create and open) key
   nRet = RegCreateKeyEx(HKEY_CURRENT_USER, SubKey(Section), 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, ByVal 0&, hKey, nResult)
   If nRet = ERROR_SUCCESS Then
      ' Set appropriate value for default query
      If Key = "*" Then Key = vbNullString
      ' Write new value to registry
      nRet = RegSetValueEx(hKey, Key, 0&, REG_SZ, ByVal Setting, Len(Setting))
      Call RegCloseKey(hKey)
   End If
   SaveSetting = (nRet = ERROR_SUCCESS)
End Function

'*******************************************************************************
' SubKey (FUNCTION)
'
' PARAMETERS:
' (In) - Section - String -
'
' RETURN VALUE:
' String - the subkey
'
' DESCRIPTION:
' Build SubKey from known values
'*******************************************************************************
Private Function SubKey(Optional ByVal Section As String = "") As String
   SubKey = "Software\" & m_Company & "\" & m_AppName
   If Len(Section) Then
      SubKey = SubKey & "\" & Section
   End If
End Function
