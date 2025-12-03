VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "COneDrive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Initially, I wanted to write this as a collection of procedures and functions.
' However, the code became too confusing, so it seemed best to encapsulate that
' complexity into a class object.
'
' Currently, there are two external properties and one external method:
'
' * URI (property read/write) e.g. https://<your-organization>.sharepoint.com/sites/<your-microsoft-team>/Shared Documents/<your-microsoft-team-channel>/<your-folder> '
' * LocalPath (property read-only)
' * IsURI (method true/false)

Dim msURI As String
Dim msLocalPath As String

'|---------------------------------------------------------------------------------------|------------------------------|--------------|
'|                                        SITE                                           |           CHANNEL            |    FOLDER    |
'|---------------------------------------------------------------------------------------|------------------------------|--------------|
'| https://<your-organization>.sharepoint.com/sites/<your-microsoft-team>/Shared Documents/<your-microsoft-team-channel>/<your-folder> |
'|---------------------------------------------------------------------------------------|------------------------------|--------------|

Private mcolTenants As Collection
Private mcolChannels As Collection

Public Property Get URI() As String
   URI = msURI
End Property

Public Property Let URI(ByVal sNew As String)
   
   '---------------'
   ' Store new URI '
   '---------------'
   msURI = sNew
   
   '--------------------------------------------------'
   ' Check to see if we even need to convert the path '
   '--------------------------------------------------'
   
   If Not Me.IsURI Then
      msLocalPath = sNew
      Exit Property
   End If
      
   Select Case OneDriveType()
      
      Case "Consumer"
         ParseConsumerURI
         
      Case "Commercial-Personal"
         ParseCommercialPersonalURI
      
      Case "Commercial-Sharepoint"
         ParseCommercialSharepointURI
         
      'ParseURI

      'GetLocalPath
   End Select
      
End Property

Public Property Get LocalPath() As String
   LocalPath = msLocalPath
End Property

Private Sub ParseCommercialSharepointURI()
   
   Dim sSite As String
   Dim sURI As String
   Dim sChannel As String
   Dim sFolder As String
   Dim v As Variant
   
   Call GetTenants
   Call GetChannels
   
   sSite = gfHead(msURI, "Shared Documents") + "Shared Documents"
   sURI = gfTail(msURI, sSite & "/")
   sChannel = gfHead(sURI, "/")
   sFolder = gfTail(sURI, "/")
        
   For Each v In mcolChannels
      If sChannel = gfReverseHead(v, " - ") Then
         msLocalPath = v & "\" & sFolder
         Exit For
      End If
   Next
   
End Sub

Private Sub ParseCommercialPersonalURI()
   
   Dim sServiceEndPointURI As String
   Dim sFolder As String
   
   sServiceEndPointURI = Me.OneDriveURI
   sServiceEndPointURI = gfHead(sServiceEndPointURI, "_api") & "Documents"
   sFolder = gfTail(msURI, sServiceEndPointURI)
   sFolder = Replace(sFolder, "/", "\")
   
   msLocalPath = OneDriveCommercialPath & sFolder
   
End Sub

Private Sub ParseConsumerURI()
   
   Dim sCID As String
   Dim sFolder As String
   
   sCID = Me.OneDriveCID
   
   sFolder = gfTail(msURI, sCID)
   sFolder = Replace(sFolder, "/", "\")
   msLocalPath = OneDriveConsumerPath & sFolder
   
End Sub

Public Function OneDriveType() As String
   
   If InStr(1, msURI, "docs.live.net") > 1 Then
      OneDriveType = "Consumer"
   ElseIf InStr(1, msURI, "my.sharepoint.com") > 1 Then
      OneDriveType = "Commercial-Personal"
   Else
      OneDriveType = "Commercial-Sharepoint"
   End If
   
End Function

Public Function Tenants() As Collection
   Call GetTenants
   Set Tenants = mcolTenants
End Function

Public Function Channels() As Collection
   Call GetChannels
   Set Channels = mcolChannels
End Function

Public Function IsURI() As Boolean
   '-----------------------------------------------------------------------'
   ' A simple function to determine if it is necessary to convert the path '
   '-----------------------------------------------------------------------'
   If UCase(Left(msURI, 4)) = "HTTP" Then
      IsURI = True
   Else
      IsURI = False
   End If

End Function

Private Sub GetTenants()
   
   Dim sKey As String
   
   Dim arrTenants As Variant
   Dim i As Integer

   '-----------------------------------------------------------------------'
   ' Get Tenants.                                                          '
   ' OneDrive Personal in form of "OneDrive - <company/org name>"          '
   ' OneDrive Sharepoint in form of "<company/org name>"                   '
   ' Keys under OneDrive Sharepoint are local paths for each shared folder '
   '-----------------------------------------------------------------------'
   
   sKey = "Software\Microsoft\OneDrive\Accounts\Business1\Tenants\"
   arrTenants = EnumerateRegistryFolders(HKEY_CURRENT_USER, sKey)
      
   Set mcolTenants = New Collection
   For i = LBound(arrTenants) To UBound(arrTenants)
      mcolTenants.Add arrTenants(i)
   Next
   
End Sub

Private Sub GetChannels()
   
   Dim sKey As String
   Dim v As Variant
   Dim arr As Variant
   Dim i As Integer
   
   '-----------------------------------------------------------------------'
   ' "Channels" are the local paths under the Tenants registry folder.     '
   ' I call the channels because they coorespond to the channels in Teams. '
   ' I suspect that if you created your Sharepoint site in Sharepoint it   '
   ' might be called something a little bit different.                     '
   '-----------------------------------------------------------------------'
   sKey = "Software\Microsoft\OneDrive\Accounts\Business1\Tenants\"
   
   Set mcolChannels = New Collection
   
   For Each v In mcolTenants
      arr = EnumerateRegistryValues(HKEY_CURRENT_USER, sKey & v & "\")
      For i = LBound(arr) To UBound(arr)
         mcolChannels.Add arr(i)
      Next
   Next
         
End Sub

Public Function OneDriveConsumerPath() As String
   OneDriveConsumerPath = CStr(ReadRegistryValue("HKCU\Environment\OneDriveConsumer"))
End Function

Public Function OneDriveCommercialPath() As String
   OneDriveCommercialPath = CStr(ReadRegistryValue("HKCU\Environment\OneDriveCommercial"))
End Function

'********************************************'
' Not Used, but might be useful at somepoint '
'********************************************'
Public Function OneDriveCID() As String
   OneDriveCID = ReadRegistryValue("HKCU\Software\Microsoft\OneDrive\Accounts\Personal\cid")
End Function

Public Function OneDriveURI() As String
   OneDriveURI = ReadRegistryValue("HKCU\Software\Microsoft\OneDrive\Accounts\Business1\ServiceEndpointUri")
End Function

Public Function TeamsURI() As String
   TeamsURI = ReadRegistryValue("HKCU\Software\Microsoft\OneDrive\Accounts\Business1\TeamSiteSPOResourceId")
End Function

