VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NetScape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim path As String
Dim cachepath As String
Dim historyfile As String
Dim bookmarksfile As String
Dim cookiesfile As String
Dim programfilespath As String
Public Sub GetNSProfiles()
Dim AppDataFolder As String
Dim MozillaProfilesDefaultFolder As String
Dim sltFolderName
Dim reg As New cRegistry
Dim x As Integer
Dim g As Variant
On Error GoTo localError

reg.ClassKey = HKEY_CURRENT_USER
reg.SectionKey = "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
reg.ValueKey = "AppData"
AppDataFolder = reg.Value
    'AppDataFolder = GetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders+ Value`= AppData")
    MozillaProfilesDefaultFolder = AppDataFolder & "\Mozilla\Profiles"
    
    sltFolderName = Dir(MozillaProfilesDefaultFolder & "\*", vbDirectory)
            Do While sltFolderName <> ""
           Dim strList() As String
           ReDim Preserve strList(x)
                'Got Folder Name now
            If sltFolderName <> "." And sltFolderName <> ".." Then
                strList(x) = sltFolderName
                x = x + 1
            End If
            
            sltFolderName = Dir
            Loop
    For Each g In strList
        getNSProfileSubProfieSearch g, MozillaProfilesDefaultFolder
    Next
    
    
    NS6Exists = True

    Exit Sub

localError:
    NS6Exists = False

  
End Sub

Private Sub getNSProfileSubProfieSearch(sltFolderName As Variant, MozillaProfilesDefaultFolder As String)

Dim profileSubFolderName As String


On Error GoTo localError

profileSubFolderName = Dir(MozillaProfilesDefaultFolder & "\" & sltFolderName & "\*", vbDirectory)
Do While profileSubFolderName <> ""
    If Right(LCase(profileSubFolderName), 4) = ".slt" Then
        'we got the profile sub folder .slt name
        If TestPath(MozillaProfilesDefaultFolder & "\" & sltFolderName & "\" & profileSubFolderName) = True Then
            'cmbSelect.AddItem MozillaProfilesDefaultFolder & "\" & sltFolderName & "\" & profileSubFolderName
            'cmbProfile.AddItem sltFolderName
            path = MozillaProfilesDefaultFolder & "\" & sltFolderName & "\" & profileSubFolderName
        End If
                        
    End If
    profileSubFolderName = Dir
    Loop
                
    NS6Exists = True

    Exit Sub

localError:
    NS6Exists = False

End Sub
Private Function TestPath(szPath As String) As Boolean
    On Error GoTo PathErr
    Open szPath + "\Cookies.txt" For Input As #2
    Close #2
    'If path exists that means the profile exists.
    TestPath = True
    Exit Function
PathErr:
    Close #2
    If Err.Number = 76 Then
        'If Path error that means the profile is invalid
        TestPath = False
    Else
        'If any other error that means that path exists
        'That means it is a valid path
        TestPath = True
    End If
End Function
 Property Get GetCachePath()
GetCachePath = cachepath
End Property
 Property Get GetHistoryFile()
GetHistoryFile = historyfile
End Property
 Property Get GetBookmarksFile()
GetBookmarksFile = bookmarksfile
End Property
Public Function Delete_History()
On Error Resume Next
Kill historyfile
End Function
Public Function Delete_Cache()
On Error Resume Next
Kill cachepath & "\*.*"
End Function
Public Function Delete_Cookies()
On Error Resume Next
Kill cookiesfile
End Function

Private Sub Class_Initialize()
GetNSProfiles
cachepath = path & "\cache"
historyfile = path & "\history.dat"
bookmarksfile = path & "\bookmarks.html"
cookiesfile = path & "\cookies.txt"
End Sub


