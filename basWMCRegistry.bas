Attribute VB_Name = "basWMCRegistry"
Option Explicit

Enum RegistryClass
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_CURRENT_CONFIG = &H80000005
End Enum

Private Sub UnitTest()
    
   Dim sKey As String
   
   Dim arr As Variant
   Dim i As Integer
   
   Dim arr2 As Variant
   Dim j As Integer
   
   
   '--------------------'
   ' Test Reading Value '
   '--------------------'
   sKey = "HKCU\Software\Microsoft\OneDrive\Accounts\Business1\SPOResourceID"
   Debug.Print "OneDrive Personal URI: " & ReadRegistryValue(sKey)
    
   sKey = "HKCU\Software\Microsoft\OneDrive\Accounts\Business1\TeamSiteSPOResourceID"
   Debug.Print "OneDrive Sharepoint URI: " & ReadRegistryValue(sKey)
    
   Debug.Print ""
   Debug.Print "OneDrive Local Folders"
   Debug.Print "----------------------"
   
   sKey = "Software\Microsoft\OneDrive\Accounts\Business1\Tenants\"
   arr = EnumerateRegistryFolders(HKEY_CURRENT_USER, sKey)
   For i = LBound(arr) To UBound(arr)
      Debug.Print arr(i)
      arr2 = EnumerateRegistryValues(HKEY_CURRENT_USER, sKey & arr(i) & "\")
      For j = LBound(arr2) To UBound(arr2)
         Debug.Print "--- " & arr2(j)
      Next
   Next
      
End Sub

Public Function ReadRegistryValue(sRegPath) As Variant
   
   '-----------------------------------------'
   ' HKEY_CURRENT_USER   HKCU                '
   ' HKEY_LOCAL_MACHINE  HKLM                '
   ' HKEY_CLASSES_ROOT   HKCR                '
   ' HKEY_USERS          HKEY_USERS          '
   ' HKEY_CURRENT_CONFIG HKEY_CURRENT_CONFIG '
   '-----------------------------------------'
   
   Dim o As Object
   Set o = CreateObject("WScript.Shell")
   ReadRegistryValue = o.regread(sRegPath)
   
End Function

Public Function EnumerateRegistryValues(vHkey As Variant, sKey As String) As Variant
   
   Dim o As Object
   Dim arrKeys As Variant
   
   Set o = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & "." & "\root\default:StdRegProv")
   Call o.EnumValues(vHkey, sKey, arrKeys)
   EnumerateRegistryValues = arrKeys
   
End Function

Public Function EnumerateRegistryFolders(vHkey As Variant, sKey As String) As Variant
   
   Dim o As Object
   Dim arrSubKeys As Variant
   
   Set o = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & "." & "\root\default:StdRegProv")
   Call o.EnumKey(vHkey, sKey, arrSubKeys)
   EnumerateRegistryFolders = arrSubKeys
   
End Function
