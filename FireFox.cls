VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FireFox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim path As String
Dim cachepath As String
Dim historyfile As String
Dim bookmarksfile1 As String
Dim bookmarksfile2 As String
Dim cookiesfile As String

Private Function GetFireFoxPath()
Dim tmp As New NetScape
Dim ini1 As New INI
Dim drive As String
Dim reg As New cRegistry
Dim dir1 As String
drive = Left(tmp.GetCachePath, 2)
reg.ClassKey = HKEY_CURRENT_USER
reg.SectionKey = "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
reg.ValueKey = "AppData"
path = reg.Value & "\Mozilla\Firefox"
dir1 = ini1.GetFromINI(path & "\profiles.ini", "Profile0", "Path")
path = path & "\" & dir1
Set reg = Nothing
Set tmp = Nothing
Set ini1 = Nothing
End Function
Private Sub Class_Initialize()

GetFireFoxPath
cachepath = path & "\cache"
historyfile = path & "\history.dat"
bookmarksfile1 = path & "\bookmarks.html"
bookmarksfile2 = path & "\bookmarks.bak"
cookiesfile = path & "\cookies.txt"
End Sub

Public Function Delete_History()
On Error Resume Next
Kill historyfile

End Function
Public Function Delete_Cache()
On Error Resume Next
Kill cachepath & "\*.*"
End Function
Public Function Delete_Bookmarks()
On Error Resume Next
Kill bookmarksfile1
Kill bookmarksfile2
End Function
Public Function Delete_Cookies()
On Error Resume Next
Kill cookiesfile
End Function

