
'Attribute VB_Name = "GetLocalOneDrivePath"
' Cross-platform VBA Function to get the local path of OneDrive/SharePoint
' synchronized Microsoft Office files (Works on Windows and on macOS)
'
' Author: Guido Witt-Dörring
' Created: 2022/07/01
' Updated: 2023/06/29
' License: MIT
'
' ————————————————————————————————————————————————————————————————
' https://gist.github.com/guwidoe/038398b6be1b16c458365716a921814d
' https://stackoverflow.com/a/73577057/12287457
' ————————————————————————————————————————————————————————————————
'
' Copyright (c) 2023 Guido Witt-Dörring
'
' MIT License:
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to
' deal in the Software without restriction, including without limitation the
' rights to use, copy, modify, merge, publish, distribute, sublicense, and/or
' sell copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in
' all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
' FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
' IN THE SOFTWARE.
'
'———————————————————————————————————————————————————————————————————————————————
' COMMENTS REGARDING THE IMPLEMENTATION:
' 1) Background and Alternative
'    This function was intended to be written as a single procedure without any
'    dependencies, for maximum portability between projects, as it implements a
'    functionality that is very commonly needed for many VBA applications
'    working inside OneDrive/SharePoint synchronized directories. I followed
'    this paradigm because it was not clear to me how complicated this simple
'    sounding endeavour would turn out to be.
'    Unfortunately, more and more complications arose, and little by little,
'    the procedure turned incredibly complicated. I do not condone the coding
'    style applied here, and that this is not how I usually write code,
'    nevertheless, I'm not open to rewriting this code in a different style,
'    because a clean implementation of this algorithm already exists, as pointed
'    out in the following.
'
'    If you would like to understand the underlying algorithm of how the local
'    path can be found with only the Url-path as input, I recommend following
'    the much cleaner implementation by Cristian Buse:
'    https://github.com/cristianbuse/VBA-FileTools
'    We developed the algorithm together and wrote separate implementations
'    concurrently. His solution is contained inside a module-level library,
'    split into many procedures and using features like private types and API-
'    functions, that are not available when trying to create a single procedure
'    without dependencies like below. This makes his code more readable.
'
'    Both of our solutions are well tested and actively supported with bugfixes
'    and improvements, so both should be equally valid choices for use in your
'    project. The differences in performance/features are marginal and they can
'    often be used interchangeably. If you need more file-system interaction
'    functionality, use Cristians library, and if you only need GetLocalPath,
'    just copy this function to any module in your project and it will work.
'
' 2) How does this function NOT work?
'    Most other solutions for this problem circulating online (A list of most
'    can be found here: https://stackoverflow.com/a/73577057/12287457) are using
'    one of two approaches, :
'     1. they use the environment variables set by OneDrive:
'         - Environ(OneDrive)
'         - Environ(OneDriveCommercial)
'         - Environ(OneDriveConsumer)
'        and replace part of the URL with it. There are many problems with this
'        approach:
'         1. They are not being set by OneDrive on MacOS.
'         2. It is unclear exactly which part of the URL needs to be replaced.
'         3. Environment variables can be changed by the user.
'         4. Only there three exist. If more onedrive accounts are logged in,
'            they just overwrite the previous ones.
'        or,
'     2. they use the mount points OneDrive writes to the registry here:
'         - \HKEY_CURRENT_USER\Software\SyncEngines\Providers\OneDrive\
'        this also has several drawbacks:
'         1. The registry is not available on MacOS.
'         2. It's still unclear exactly what part of the URL should be replaced.
'         3. These registry keys can contain mistakes, like for example, when:
'             - Synchronizing a folder called "Personal" from someone else's
'               personal OneDrive
'             - Synchronizing a folder called "Business1" from someone else's
'               personal OneDrive and then relogging your own first Business
'               OneDrive account
'             - Relogging you personal OneDrive can change the "CID" property
'               from a folderID formatted cid (e.g. 3DEA8A9886F05935!125) to a
'               regular private cid (e.g. 3dea8a9886f05935) for synced folders
'               from other people's OneDrives
'
'    For these reasons, this solution uses a completely different approach to
'    solve this problem.
'
' 3) How does this function work?
'    This function builds the Web to Local translation dictionary by extracting
'    the mount points from the OneDrive settings files.
'    It reads files from...
'    On Windows:
'        - the "...\AppData\Local\Microsoft" directory
'    On Mac:
'        - the "~/Library/Containers/com.microsoft.OneDrive-mac/Data/" & _
'              "Library/Application Support" directory
'        - and/or the "~/Library/Application Support"
'    It reads the following files:
'      - \OneDrive\settings\Personal\ClientPolicy.ini
'      - \OneDrive\settings\Personal\????????????????.dat
'      - \OneDrive\settings\Personal\????????????????.ini
'      - \OneDrive\settings\Personal\global.ini
'      - \OneDrive\settings\Personal\GroupFolders.ini
'      - \OneDrive\settings\Business#\????????-????-????-????-????????????.dat
'      - \OneDrive\settings\Business#\????????-????-????-????-????????????.ini
'      - \OneDrive\settings\Business#\ClientPolicy*.ini
'      - \OneDrive\settings\Business#\global.ini
'      - \Office\CLP\* (just the filename)
'
'    Where:
'     - "*" ... 0 or more characters
'     - "?" ... one character [0-9, a-f]
'     - "#" ... one digit
'     - "\" ... path separator, (= "/" on MacOS)
'     - The "???..." filenames represent CIDs)
'
'    On MacOS, the \Office\CLP\* exists for each Microsoft Office application
'    separately. Depending on whether the application was already used in
'    active syncing with OneDrive it may contain different/incomplete files.
'    In the code, the path of this directory is stored inside the variable
'    "clpPath". On MacOS, the defined clpPath might not exist or not contain
'    all necessary files for some host applications, because Environ("HOME")
'    depends on the host app.
'    This is not a big problem as the function will still work, however in
'    this case, specifying a preferredMountPointOwner will do nothing.
'    To make sure this directory and the necessary files exist, a file must
'    have been actively synchronized with OneDrive by the application whose
'    "HOME" folder is returned by Environ("HOME") while being logged in
'    to that application with the account whose email is given as
'    preferredMountPointOwner, at some point in the past!
'
'    If you are usually working with Excel but are using this function in a
'    different app, you can instead use an alternative (Excels CLP folder) as
'    the clpPath as it will most likely contain all the necessary information
'    The alternative clpPath is commented out in the code, if you prefer to
'    use Excels CLP folder per default, just un-comment the respective line
'    in the code.
'———————————————————————————————————————————————————————————————————————————————

'———————————————————————————————————————————————————————————————————————————————
' COMMENTS REGARDING THE USAGE:
' This function can be used as a User Defined Function (UDF) from the worksheet.
' (More on that, see "USAGE EXAMPLES")
'
' This function offers three optional parameters to the user, however using
' these should only be necessary in extremely rare situations.
' The best rule regarding their usage: Don't use them.
'
' In the following these parameters will still be explained.
'
'1) returnAll
'   In some exceptional cases it is possible to map one OneDrive WebPath to
'   multiple different localPaths. This can happen when multiple Business
'   OneDrive accounts are logged in on one device, and multiple of these have
'   access to the same OneDrive folder and they both decide to synchronize it or
'   add it as link to their MySite library.
'   Calling the function with returnAll:=True will return all valid localPaths
'   for the given WebPath, separated by two forward slashes (//). This should be
'   used with caution, as the return value of the function alone is, should
'   multiple local paths exist for the input webPath, not a valid local path
'   anymore.
'   An example of how to obtain all of the local paths could look like this:
'   Dim localPath as String, localPaths() as String
'   localPath = GetLocalPath(webPath, True)
'   If Not localPath Like "http*" Then
'       localPaths = Split(localPath, "//")
'   End If
'
'2) preferredMountPointOwner
'   This parameter deals with the same problem as 'returnAll'
'   If the function gets called with returnAll:=False (default), and multiple
'   localPaths exist for the given WebPath, the function will just return any
'   one of them, as usually, it shouldn't make a difference, because the result
'   directories at both of these localPaths are mirrored versions of the same
'   webPath. Nevertheless, this option lets the user choose, which mountPoint
'   should be chosen if multiple localPaths are available. Each localPath is
'  'owned' by an OneDrive Account. If a WebPath is synchronized twice, this can
'   only happen by synchronizing it with two different accounts, because
'   OneDrive prevents you from synchronizing the same folder twice on a single
'   account. Therefore, each of the different localPaths for a given WebPath
'   has a unique 'owner'. preferredMountPointOwner lets the user select the
'   localPath by specifying the account the localPath should be owned by.
'   This is done by passing the Email address of the desired account as
'   preferredMountPointOwner.
'   For example, you have two different Business OneDrive accounts logged in,
'   foo.bar@business1.com and foo.bar@business2.com
'   Both synchronize the WebPath:
'   webPath = "https://business1.sharepoint.com/sites/TestLib/Documents/" & _
              "Test/Test/Test/test.xlsm"
'
'   The first one has added it as a link to his personal OneDrive, the local
'   path looks like this:
'   C:\Users\username\OneDrive - Business1\TestLinkParent\Test - TestLinkLib\...
'   ...Test\test.xlsm
'
'   The second one just synchronized it normally, the localPath looks like this:
'   C:\Users\username\Business1\TestLinkLib - Test\Test\test.xlsm
'
'   Calling GetLocalPath like this:
'   GetLocalPath(webPath,,, "foo.bar@business1.com") will return:
'   C:\Users\username\OneDrive - Business1\TestLinkParent\Test - TestLinkLib\...
'   ...Test\test.xlsm
'
'   Calling it like this:
'   GetLocalPath(webPath,,, "foo.bar@business2.com") will return:
'   C:\Users\username\Business1\TestLinkLib - Test\Test\test.xlsm
'
'   And calling it like this:
'   GetLocalPath(webPath,, True) will return:
'   C:\Users\username\OneDrive - Business1\TestLinkParent\Test - TestLinkLib\...
'   ...Test\test.xlsm//C:\Users\username\Business1\TestLinkLib - Test\Test\...
'   ...test.xlsm
'
'   Calling it normally like this:
'   GetLocalPath(webPath) will return any one of the two localPaths, so:
'   C:\Users\username\OneDrive - Business1\TestLinkParent\Test - TestLinkLib\...
'   ...Test\test.xlsm
'   OR
'   C:\Users\username\Business1\TestLinkLib - Test\Test\test.xlsm
'
'3) rebuildCache
'   The function creates a "translation" dictionary from the OneDrive settings
'   files and then uses this dictionary to "translate" WebPaths to LocalPaths.
'   This dictionary is implemented as a static variable to the function doesn't
'   have to recreate it every time it is called. It is written on the first
'   function call and reused on all the subsequent calls, making them faster.
'   If the function is called with rebuildCache:=True, this dictionary will be
'   rewritten, even if it was already initialized.
'   Note that it is not necessary to use this parameter manually, even if a new
'   MountPoint was added to the OneDrive, or a new OneDrive account was logged
'   in since the last function call because the function will automatically
'   determine if any of those cases occurred, without sacrificing performance.
'———————————————————————————————————————————————————————————————————————————————
Option Explicit

''——————————————————————————————————————————————————————————————————————————————
'' USAGE EXAMPLES:
'' Excel:
'Private Sub TestGetLocalPathExcel()
'    Debug.Print GetLocalPath(ThisWorkbook.FullName)
'    Debug.Print GetLocalPath(ThisWorkbook.path)
'End Sub
'
' Usage as User Defined Function (UDF):
' You might have to replace ; with , in the formulas depending on your settings.
' Add this formula to any cell, to get the local path of the workbook:
' =GetLocalPath(LEFT(CELL("filename";A1);FIND("[";CELL("filename";A1))-1))
'
' To get the local path including the filename (the FullName), use this formula:
' =GetLocalPath(LEFT(CELL("filename";A1);FIND("[";CELL("filename";A1))-1) &
' TEXTAFTER(TEXTBEFORE(CELL("filename";A1);"]");"["))
'
''Word:
'Private Sub TestGetLocalPathWord()
'    Debug.Print GetLocalPath(ThisDocument.FullName)
'    Debug.Print GetLocalPath(ThisDocument.path)
'End Sub
'
''PowerPoint:
'Private Sub TestGetLocalPathPowerPoint()
'    Debug.Print GetLocalPath(ActivePresentation.FullName)
'    Debug.Print GetLocalPath(ActivePresentation.path)
'End Sub
''——————————————————————————————————————————————————————————————————————————————


'This Function will convert a OneDrive/SharePoint Url path, e.g. Url containing
'https://d.docs.live.net/; .sharepoint.com/sites; my.sharepoint.com/personal/...
'to the locally synchronized path on your current pc or mac, e.g. a path like
'C:\users\username\OneDrive\ on Windows; or /Users/username/OneDrive/ on MacOS,
'if you have the remote directory locally synchronized with the OneDrive app.
'If no local path can be found, the input value will be returned unmodified.
'Author: Guido Witt-Dörring
'Source: https://gist.github.com/guwidoe/038398b6be1b16c458365716a921814d
'        https://stackoverflow.com/a/73577057/12287457
Public Function GetLocalPath(ByVal path As String, _
                    Optional ByVal returnAll As Boolean = False, _
                    Optional ByVal preferredMountPointOwner As String = "", _
                    Optional ByVal rebuildCache As Boolean = False) _
                             As String
    #If Mac Then
        Const vbErrPermissionDenied            As Long = 70
        Const vbErrInvalidFormatInResourceFile As Long = 325
        Const syncIDFileName As String = ".849C9593-D756-4E56-8D6E-42412F2A707B"
        Const isMac As Boolean = True
        Const ps As String = "/" 'Application.PathSeparator doesn't work
    #Else 'Windows               'in all host applications (e.g. Outlook), hence
        Const ps As String = "\" 'conditional compilation is preferred here.
        Const isMac As Boolean = False
    #End If
    Const methodName As String = "GetLocalPath"
    Const vbErrFileNotFound     As Long = 53
    Const vbErrOutOfMemory      As Long = 7
    Const vbErrKeyAlreadyExists As Long = 457
    Static locToWebColl As Collection, lastCacheUpdate As Date

    If Not Left(path, 8) = "https://" Then GetLocalPath = path: Exit Function

    Dim webRoot As String, locRoot As String, s As String, vItem As Variant
    Dim pmpo As String: pmpo = LCase$(preferredMountPointOwner)
    If Not locToWebColl Is Nothing And Not rebuildCache Then
        Dim resColl As Collection: Set resColl = New Collection
        'If the locToWebColl is initialized, this logic will find the local path
        For Each vItem In locToWebColl
            locRoot = vItem(0): webRoot = vItem(1)
            If InStr(1, path, webRoot, vbTextCompare) = 1 Then _
                resColl.Add key:=vItem(2), _
                   Item:=Replace(Replace(path, webRoot, locRoot, , 1), "/", ps)
        Next vItem
        If resColl.count > 0 Then
            If returnAll Then
                For Each vItem In resColl: s = s & "//" & vItem: Next vItem
                GetLocalPath = Mid$(s, 3): Exit Function
            End If
            On Error Resume Next: GetLocalPath = resColl(pmpo): On Error GoTo 0
            If GetLocalPath <> "" Then Exit Function
            GetLocalPath = resColl(1): Exit Function
        End If
        'Local path was not found with cached mountpoints
        GetLocalPath = path 'No Exit Function here! Check if cache needs rebuild
    End If

    Dim settPaths As Collection: Set settPaths = New Collection
    Dim settPath As Variant, clpPath As String
    #If Mac Then 'The settings directories can be in different locations
        Dim cloudStoragePath As String, cloudStoragePathExists As Boolean
        s = Environ("HOME")
        clpPath = s & "/Library/Application Support/Microsoft/Office/CLP/"
        s = Left$(s, InStrRev(s, "/Library/Containers/", , vbBinaryCompare))
        settPaths.Add s & _
                      "Library/Containers/com.microsoft.OneDrive-mac/Data/" & _
                      "Library/Application Support/OneDrive/settings/"
        settPaths.Add s & "Library/Application Support/OneDrive/settings/"
        cloudStoragePath = s & "Library/CloudStorage/"

        'Excels CLP folder:
        'clpPath = Left$(s, InStrRev(s, "/Library/Containers", , 0)) & _
                  "Library/Containers/com.microsoft.Excel/Data/" & _
                  "Library/Application Support/Microsoft/Office/CLP/"
    #Else 'On Windows, the settings directories are always in this location:
        settPaths.Add Environ("LOCALAPPDATA") & "\Microsoft\OneDrive\settings\"
        clpPath = Environ("LOCALAPPDATA") & "\Microsoft\Office\CLP\"
    #End If

    Dim i As Long
    #If Mac Then 'Request access to all possible directories at once
        Dim arrDirs() As Variant: ReDim arrDirs(1 To settPaths.count * 11 + 1)
        For Each settPath In settPaths
            For i = i + 1 To i + 9
                arrDirs(i) = settPath & "Business" & i Mod 11
            Next i
            arrDirs(i) = settPath: i = i + 1
            arrDirs(i) = settPath & "Personal"
        Next settPath
        arrDirs(i + 1) = cloudStoragePath
        Dim accessRequestInfoMsgShown As Boolean
        accessRequestInfoMsgShown = GetSetting("GetLocalPath", _
                        "AccessRequestInfoMsg", "Displayed", "False") = "True"
        If Not accessRequestInfoMsgShown Then MsgBox "The current " _
            & "VBA Project requires access to the OneDrive settings files to " _
            & "translate a OneDrive URL to the local path of the locally " & _
            "synchronized file/folder on your Mac. Because these files are " & _
            "located outside of Excels sandbox, file-access must be granted " _
            & "explicitly. Please approve the access requests following this " _
            & "message.", vbInformation
        If Not GrantAccessToMultipleFiles(arrDirs) Then _
            Err.Raise vbErrPermissionDenied, methodName
    #End If

    'Find all subdirectories in OneDrive settings folder:
    Dim oneDriveSettDirs As Collection: Set oneDriveSettDirs = New Collection
    For Each settPath In settPaths
        Dim dirName As String: dirName = Dir(settPath, vbDirectory)
        Do Until dirName = vbNullString
            If dirName = "Personal" Or dirName Like "Business#" Then _
                oneDriveSettDirs.Add Item:=settPath & dirName & ps
            dirName = Dir(, vbDirectory)
        Loop
    Next settPath

    If Not locToWebColl Is Nothing Or isMac Then
        Dim requiredFiles As Collection: Set requiredFiles = New Collection
        'Get collection of all required files
        Dim vDir As Variant
        For Each vDir In oneDriveSettDirs
            Dim cid As String: cid = IIf(vDir Like "*" & ps & "Personal" & ps, _
                                         "????????????*", _
                                         "????????-????-????-????-????????????")
            Dim FileName As String: FileName = Dir(vDir, vbNormal)
            Do Until FileName = vbNullString
                If FileName Like cid & ".ini" _
                Or FileName Like cid & ".dat" _
                Or FileName Like "ClientPolicy*.ini" _
                Or StrComp(FileName, "GroupFolders.ini", vbTextCompare) = 0 _
                Or StrComp(FileName, "global.ini", vbTextCompare) = 0 Then _
                    requiredFiles.Add Item:=vDir & FileName
                FileName = Dir
            Loop
        Next vDir
    End If

    'This part should ensure perfect accuracy despite the mount point cache
    'while sacrificing almost no performance at all by querying FileDateTimes.
    If Not locToWebColl Is Nothing And Not rebuildCache Then
        'Check if a settings file was modified since the last cache rebuild
        Dim vFile As Variant
        For Each vFile In requiredFiles
            If FileDateTime(vFile) > lastCacheUpdate Then _
                rebuildCache = True: Exit For 'full cache refresh is required!
        Next vFile
        If Not rebuildCache Then Exit Function
    End If

    'If execution reaches this point, the cache will be fully rebuilt...
    Dim fileNum As Long, syncID As String, b() As Byte
    #If Mac Then 'Variables for manual decoding of UTF-8, UTF-32 and ANSI
        Dim j As Long, k As Long, m As Long, ansi() As Byte, sAnsi As String
        Dim utf16() As Byte, sUtf16 As String, utf32() As Byte
        Dim utf8() As Byte, sUtf8 As String, numBytesOfCodePoint As Long
        Dim codepoint As Long, lowSurrogate As Long, highSurrogate As Long
    #End If

    lastCacheUpdate = Now()
    #If Mac Then 'Prepare building syncIDtoSyncDir dictionary. This involves
        'reading the ".849C9593-D756-4E56-8D6E-42412F2A707B" files inside the
        'subdirs of "~/Library/CloudStorage/", list of files and access required
        Dim coll As Collection: Set coll = New Collection
        dirName = Dir(cloudStoragePath, vbDirectory)
        Do Until dirName = vbNullString
            If dirName Like "OneDrive*" Then
                cloudStoragePathExists = True
                vDir = cloudStoragePath & dirName & ps
                vFile = cloudStoragePath & dirName & ps & syncIDFileName
                coll.Add Item:=vDir
                requiredFiles.Add Item:=vDir 'For pooling file access requests
                requiredFiles.Add Item:=vFile
            End If
            dirName = Dir(, vbDirectory)
        Loop

        'Pool access request for these files and the OneDrive/settings files
        If locToWebColl Is Nothing Then
            Dim vFiles As Variant
            If requiredFiles.count > 0 Then
                ReDim vFiles(1 To requiredFiles.count)
               For i = 1 To UBound(vFiles): vFiles(i) = requiredFiles(i): Next i
                If Not GrantAccessToMultipleFiles(vFiles) Then _
                    Err.Raise vbErrPermissionDenied, methodName
            End If
        End If

        'More access might be required if some folders inside cloudStoragePath
        'don't contain the hidden file ".849C9593-D756-4E56-8D6E-42412F2A707B".
        'In that case, access to their first level subfolders is also required.
        If cloudStoragePathExists Then
            For i = coll.count To 1 Step -1
                Dim fAttr As Long: fAttr = 0
                On Error Resume Next
                fAttr = GetAttr(coll(i) & syncIDFileName)
                Dim IsFile As Boolean: IsFile = False
                If Err.Number = 0 Then IsFile = Not CBool(fAttr And vbDirectory)
                On Error GoTo 0
                If Not IsFile Then 'hidden file does not exist
                'Dir(path, vbHidden) is unreliable and doesn't work on some Macs
                'If Dir(coll(i) & syncIDFileName, vbHidden) = vbNullString Then
                    dirName = Dir(coll(i), vbDirectory)
                    Do Until dirName = vbNullString
                        If Not dirName Like ".Trash*" And dirName <> "Icon" Then
                            coll.Add coll(i) & dirName & ps
                            coll.Add coll(i) & dirName & ps & syncIDFileName, _
                                     coll(i) & dirName & ps  '<- key for removal
                        End If
                        dirName = Dir(, vbDirectory)
                    Loop          'Remove the
                    coll.Remove i 'folder if it doesn't contain the hidden file.
                End If
            Next i
            If coll.count > 0 Then
                ReDim arrDirs(1 To coll.count)
                For i = 1 To coll.count: arrDirs(i) = coll(i): Next i
                If Not GrantAccessToMultipleFiles(arrDirs) Then _
                    Err.Raise vbErrPermissionDenied, methodName
            End If
            'Remove all files from coll (not the folders!): Reminder:
            On Error Resume Next 'coll(coll(i)) = coll(i) & syncIDFileName
            For i = coll.count To 1 Step -1
                coll.Remove coll(i)
            Next i
            On Error GoTo 0

            'Write syncIDtoSyncDir collection
            Dim syncIDtoSyncDir As Collection
            Set syncIDtoSyncDir = New Collection
            For Each vDir In coll
                fAttr = 0
                On Error Resume Next
                fAttr = GetAttr(vDir & syncIDFileName)
                IsFile = False
                If Err.Number = 0 Then IsFile = Not CBool(fAttr And vbDirectory)
                On Error GoTo 0
                If IsFile Then 'hidden file exists
                'Dir(path, vbHidden) is unreliable and doesn't work on some Macs
                'If Dir(vDir & syncIDFileName, vbHidden) <> vbNullString Then
                    fileNum = FreeFile(): s = "": vFile = vDir & syncIDFileName
                    'Somehow reading these files with "Open" doesn't always work
                    Dim readSucceeded As Boolean: readSucceeded = False
                    On Error GoTo ReadFailed
                    Open vFile For Binary Access Read As #fileNum
                        ReDim b(0 To LOF(fileNum)): Get fileNum, , b: s = b
                        readSucceeded = True
ReadFailed:             On Error GoTo -1
                    Close #fileNum: fileNum = 0
                    On Error GoTo 0
                    If readSucceeded Then
                        'Debug.Print "Used open statement to read file: " & _
                                    vDir & syncIDFileName
                        ansi = s 'If Open was used: Decode ANSI string manually:
                        If LenB(s) > 0 Then
                            ReDim utf16(0 To LenB(s) * 2 - 1): k = 0
                            For j = LBound(ansi) To UBound(ansi)
                                utf16(k) = ansi(j): k = k + 2
                            Next j
                            s = utf16
                        Else: s = vbNullString
                        End If
                    Else 'Reading the file with "Open" failed with an error. Try
                        'using AppleScript. Also avoids the manual transcoding.
                        'Somehow ApplScript fails too, sometimes. Seems whenever
                        '"Open" works, AppleScript fails and vice versa (?!?!)
                        vFile = MacScript("return path to startup disk as " & _
                                    "string") & Replace(Mid$(vFile, 2), ps, ":")
                        s = MacScript("return read file """ & _
                                      vFile & """ as string")
                       'Debug.Print "Used Apple Script to read file: " & vFile
                    End If
                    If InStr(1, s, """guid"" : """, vbBinaryCompare) Then
                        s = Split(s, """guid"" : """)(1)
                        syncID = Left$(s, InStr(1, s, """", 0) - 1)
                        syncIDtoSyncDir.Add key:=syncID, _
                             Item:=VBA.Array(syncID, Left$(vDir, Len(vDir) - 1))
                    Else
                        Debug.Print "Warning, empty syncIDFile encountered!"
                    End If
                End If
            Next vDir
        End If
        'Now all access requests have succeeded
        If Not accessRequestInfoMsgShown Then SaveSetting _
            "GetLocalPath", "AccessRequestInfoMsg", "Displayed", "True"
    #End If

    'Declare all variables that will be used in the loop over OneDrive settings
    Dim line As Variant, parts() As String, N As Long, libNr As String
    Dim tag As String, mainMount As String, relPath As String, email As String
    Dim parentID As String, folderID As String, folderName As String
    Dim folderIdPattern As String, folderType As String, keyExists As Boolean
    Dim siteID As String, libID As String, webID As String, lnkID As String
    Dim mainSyncID As String, syncFind As String, mainSyncFind As String
    'The following are "constants" and needed for reading the .dat files:
    Dim sig1 As String:       sig1 = ChrB$(2)
    Dim sig2 As String * 4:   MidB$(sig2, 1) = ChrB$(1)
    Dim vbNullByte As String: vbNullByte = ChrB$(0)
    #If Mac Then
        Const sig3 As String = vbNullChar & vbNullChar
    #Else 'Windows
        Const sig3 As String = vbNullChar
    #End If

    'Writing locToWebColl using .ini and .dat files in the OneDrive settings:
    'Here, a Scripting.Dictionary would be nice but it is not available on Mac!
    Dim lastAccountUpdates As Collection, lastAccountUpdate As Date
    Set lastAccountUpdates = New Collection
    Set locToWebColl = New Collection
    For Each vDir In oneDriveSettDirs 'One folder per logged in OD account
        dirName = Mid$(vDir, InStrRev(vDir, ps, Len(vDir) - 1, 0) + 1)
        dirName = Left$(dirName, Len(dirName) - 1)

        'Read global.ini to get cid
        If Dir(vDir & "global.ini", vbNormal) = "" Then GoTo NextFolder
        fileNum = FreeFile()
        Open vDir & "global.ini" For Binary Access Read As #fileNum
            ReDim b(0 To LOF(fileNum)): Get fileNum, , b
        Close #fileNum: fileNum = 0
        #If Mac Then 'On Mac, the OneDrive settings files use UTF-8 encoding
            sUtf8 = b: GoSub DecodeUTF8
            b = sUtf16 'b = StrConv(b, vbUnicode) <- UNRELIABLE
        #End If
        For Each line In Split(b, vbNewLine)
            If line Like "cid = *" Then cid = Mid$(line, 7): Exit For
        Next line

        If cid = vbNullString Then GoTo NextFolder
        If (Dir(vDir & cid & ".ini") = vbNullString Or _
            Dir(vDir & cid & ".dat") = vbNullString) Then GoTo NextFolder
        If dirName Like "Business#" Then
            folderIdPattern = Replace(Space$(32), " ", "[a-f0-9]")
        ElseIf dirName = "Personal" Then
            folderIdPattern = Replace(Space$(12), " ", "[A-F0-9]") & "*!###*"
        End If

        'Get email for business accounts
        '(only necessary to let user choose preferredMountPointOwner)
        FileName = Dir(clpPath, vbNormal)
        Do Until FileName = vbNullString
            i = InStrRev(FileName, cid, , vbTextCompare)
            If i > 1 And cid <> vbNullString Then _
                email = LCase$(Left$(FileName, i - 2)): Exit Do
            FileName = Dir
        Loop

        #If Mac Then
            On Error Resume Next
            lastAccountUpdate = lastAccountUpdates(dirName)
            keyExists = (Err.Number = 0)
            On Error GoTo 0
            If keyExists Then
                If FileDateTime(vDir & cid & ".ini") < lastAccountUpdate Then
                    GoTo NextFolder
                Else
                    For i = locToWebColl.count To 1 Step -1
                        If locToWebColl(i)(5) = dirName Then
                            locToWebColl.Remove i
                        End If
                    Next i
                    lastAccountUpdates.Remove dirName
                    lastAccountUpdates.Add key:=dirName, _
                                         Item:=FileDateTime(vDir & cid & ".ini")
                End If
            Else
                lastAccountUpdates.Add key:=dirName, _
                                      Item:=FileDateTime(vDir & cid & ".ini")
            End If
        #End If

        'Read all the ClientPloicy*.ini files:
        Dim cliPolColl As Collection: Set cliPolColl = New Collection
        FileName = Dir(vDir, vbNormal)
        Do Until FileName = vbNullString
            If FileName Like "ClientPolicy*.ini" Then
                fileNum = FreeFile()
                Open vDir & FileName For Binary Access Read As #fileNum
                    ReDim b(0 To LOF(fileNum)): Get fileNum, , b
                Close #fileNum: fileNum = 0
                #If Mac Then 'On Mac, OneDrive settings files use UTF-8 encoding
                    sUtf8 = b: GoSub DecodeUTF8
                    b = sUtf16 'StrConv(b, vbUnicode)UNRELIABLE
                #End If
                cliPolColl.Add key:=FileName, Item:=New Collection
                For Each line In Split(b, vbNewLine)
                    If InStr(1, line, " = ", vbBinaryCompare) Then
                        tag = Left$(line, InStr(1, line, " = ", 0) - 1)
                        s = Mid$(line, InStr(1, line, " = ", 0) + 3)
                        Select Case tag
                        Case "DavUrlNamespace"
                            cliPolColl(FileName).Add key:=tag, Item:=s
                        Case "SiteID", "IrmLibraryId", "WebID" 'Only used for
                            s = Replace(LCase$(s), "-", "") 'backup method later
                            If Len(s) > 3 Then s = Mid$(s, 2, Len(s) - 2)
                            cliPolColl(FileName).Add key:=tag, Item:=s
                        End Select
                    End If
                Next line
            End If
            FileName = Dir
        Loop

        'Read cid.dat file
        Const chunkOverlap          As Long = 1000
        Const maxDirName            As Long = 255
        Dim buffSize As Long: buffSize = -1 'Buffer uninitialized
Try:    On Error GoTo Catch
        Dim odFolders As Collection: Set odFolders = New Collection
        Dim lastChunkEndPos As Long: lastChunkEndPos = 1
        Dim lastFileUpdate As Date:  lastFileUpdate = FileDateTime(vDir & _
                                                                   cid & ".dat")
        i = 0 'i = current reading pos.
        Do
            'Ensure file is not changed while reading it
            If FileDateTime(vDir & cid & ".dat") > lastFileUpdate Then GoTo Try
            fileNum = FreeFile
            Open vDir & cid & ".dat" For Binary Access Read As #fileNum
                Dim lenDatFile As Long: lenDatFile = LOF(fileNum)
                If buffSize = -1 Then buffSize = lenDatFile 'Initialize buffer
                'Overallocate a bit so read chunks overlap to recognize all dirs
                ReDim b(0 To buffSize + chunkOverlap)
                Get fileNum, lastChunkEndPos, b: s = b
                Dim SIZE As Long: SIZE = LenB(s)
            Close #fileNum: fileNum = 0
            lastChunkEndPos = lastChunkEndPos + buffSize

            For vItem = 16 To 8 Step -8
                i = InStrB(vItem + 1, s, sig2, 0) 'Sarch pattern in cid.dat
                Do While i > vItem And i < SIZE - 168 'and confirm with another
                    If StrComp(MidB$(s, i - vItem, 1), sig1, 0) = 0 Then 'one
                        i = i + 8: N = InStrB(i, s, vbNullByte, 0) - i
                        If N < 0 Then N = 0           'i:Start pos, n: Length
                        If N > 39 Then N = 39
                        #If Mac Then 'StrConv doesn't work reliably on Mac ->
                            ansi = MidB$(s, i, N) 'Decode ANSI string manually:
                            j = UBound(ansi) - LBound(ansi) + 1
                            If j > 0 Then
                                ReDim utf16(0 To j * 2 - 1): k = 0
                                For j = LBound(ansi) To UBound(ansi)
                                    utf16(k) = ansi(j): k = k + 2
                                Next j
                                folderID = utf16
                            Else: folderID = vbNullString
                            End If
                        #Else 'Windows
                            folderID = StrConv(MidB$(s, i, N), vbUnicode)
                        #End If
                        i = i + 39: N = InStrB(i, s, vbNullByte, 0) - i
                        If N < 0 Then N = 0
                        If N > 39 Then N = 39
                        #If Mac Then 'StrConv doesn't work reliably on Mac ->
                            ansi = MidB$(s, i, N) 'Decode ANSI string manually:
                            j = UBound(ansi) - LBound(ansi) + 1
                            If j > 0 Then
                                ReDim utf16(0 To j * 2 - 1): k = 0
                                For j = LBound(ansi) To UBound(ansi)
                                    utf16(k) = ansi(j): k = k + 2
                                Next j
                                parentID = utf16
                            Else: parentID = vbNullString
                            End If
                        #Else 'Windows
                            parentID = StrConv(MidB$(s, i, N), vbUnicode)
                        #End If
                        i = i + 121
                        N = InStr(-Int(-(i - 1) / 2) + 1, s, sig3) * 2 - i - 1
                        If N > maxDirName * 2 Then N = maxDirName * 2
                        If N < 0 Then N = 0
                        If folderID Like folderIdPattern _
                        And parentID Like folderIdPattern Then
                            #If Mac Then 'Encoding of folder names is UTF-32-LE
                                Do While N Mod 4 > 0
                                    If N > maxDirName * 4 Then Exit Do
                                    N = InStr(-Int(-(i + N) / 2) + 1, s, sig3) _
                                        * 2 - i - 1
                                Loop
                                If N > maxDirName * 4 Then N = maxDirName * 4
                                utf32 = MidB$(s, i, N)
                                'UTF-32 can only be converted manually to UTF-16
                                ReDim utf16(LBound(utf32) To UBound(utf32))
                                j = LBound(utf32): k = LBound(utf32)
                                Do While j < UBound(utf32)
                                    If utf32(j + 2) + utf32(j + 3) = 0 Then
                                        utf16(k) = utf32(j)
                                        utf16(k + 1) = utf32(j + 1)
                                        k = k + 2
                                    Else
                                        If utf32(j + 3) <> 0 Then Err.Raise _
                                            vbErrInvalidFormatInResourceFile, _
                                            methodName
                                        codepoint = utf32(j + 2) * &H10000 + _
                                                    utf32(j + 1) * &H100& + _
                                                    utf32(j)
                                        m = codepoint - &H10000
                                        highSurrogate = &HD800& Or (m \ &H400&)
                                        lowSurrogate = &HDC00& Or (m And &H3FF)
                                        utf16(k) = highSurrogate And &HFF&
                                        utf16(k + 1) = highSurrogate \ &H100&
                                        utf16(k + 2) = lowSurrogate And &HFF&
                                        utf16(k + 3) = lowSurrogate \ &H100&
                                        k = k + 4
                                    End If
                                    j = j + 4
                                Loop
                                If k > LBound(utf16) Then
                                    ReDim Preserve utf16(LBound(utf16) To k - 1)
                                    folderName = utf16
                                Else: folderName = vbNullString
                                End If
                            #Else 'On Windows encoding is UTF-16-LE
                                folderName = MidB$(s, i, N)
                            #End If
                            'VBA.Array() instead of just Array() is used in this
                            'function because it ignores Option Base 1
                            odFolders.Add VBA.Array(parentID, folderName), _
                                          folderID
                        End If
                    End If
                    i = InStrB(i + 1, s, sig2, 0) 'Find next sig2 in cid.dat
                Loop
                If odFolders.count > 0 Then Exit For
            Next vItem
        Loop Until lastChunkEndPos >= lenDatFile _
                Or buffSize >= lenDatFile
        GoTo Continue
Catch:
        Select Case Err.Number
        Case vbErrKeyAlreadyExists
            'This can happen at chunk boundries, folder might get added twice:
            odFolders.Remove folderID 'Make sure the folder gets added new again
            Resume 'to avoid folderNames truncated by chunk ends
        Case Is <> vbErrOutOfMemory: Err.Raise Err, methodName
        End Select
        If buffSize > &HFFFFF Then buffSize = buffSize / 2: Resume Try
        Err.Raise Err, methodName 'Raise error if less than 1 MB RAM available
Continue:
        On Error GoTo 0

        'Read cid.ini file
        fileNum = FreeFile()
        Open vDir & cid & ".ini" For Binary Access Read As #fileNum
            ReDim b(0 To LOF(fileNum)): Get fileNum, , b
        Close #fileNum: fileNum = 0
        #If Mac Then 'On Mac, the OneDrive settings files use UTF-8 encoding
            sUtf8 = b: GoSub DecodeUTF8:
            b = sUtf16 'b = StrConv(b, vbUnicode) <- UNRELIABLE
        #End If
        Select Case True
        Case dirName Like "Business#" 'Settings files for a business OD account
        'Max 9 Business OneDrive accounts can be signed in at a time.
           Dim libNrToWebColl As Collection: Set libNrToWebColl = New Collection
            mainMount = vbNullString
            For Each line In Split(b, vbNewLine)
                webRoot = "": locRoot = "": parts = Split(line, """")
                Select Case Left$(line, InStr(1, line, " = ", 0) - 1)
                Case "libraryScope" 'One line per synchronized library
                    locRoot = parts(9)
                    syncFind = locRoot: syncID = Split(parts(10), " ")(2)
                    libNr = Split(line, " ")(2)
                    folderType = parts(3): parts = Split(parts(8), " ")
                    siteID = parts(1): webID = parts(2): libID = parts(3)
                    If mainMount = vbNullString And folderType = "ODB" Then
                        mainMount = locRoot: FileName = "ClientPolicy.ini"
                        mainSyncID = syncID: mainSyncFind = syncFind
                    Else: FileName = "ClientPolicy_" & libID & siteID & ".ini"
                    End If
                    On Error Resume Next 'On error try backup method...
                    webRoot = cliPolColl(FileName)("DavUrlNamespace")
                    On Error GoTo 0
                    If webRoot = "" Then 'Backup method to find webRoot:
                        For Each vItem In cliPolColl
                            If vItem("SiteID") = siteID _
                            And vItem("WebID") = webID _
                            And vItem("IrmLibraryId") = libID Then
                                webRoot = vItem("DavUrlNamespace"): Exit For
                            End If
                        Next vItem
                    End If
                    If webRoot = vbNullString Then Err.Raise vbErrFileNotFound _
                                                           , methodName
                    libNrToWebColl.Add VBA.Array(libNr, webRoot), libNr
                    If Not locRoot = vbNullString Then _
                        locToWebColl.Add VBA.Array(locRoot, webRoot, email, _
                                        syncID, syncFind, dirName), key:=locRoot
                Case "libraryFolder" 'One line per synchronized library folder
                    libNr = Split(line, " ")(3)
                    locRoot = parts(1): syncFind = locRoot
                    syncID = Split(parts(4), " ")(1)
                    s = vbNullString: parentID = Left$(Split(line, " ")(4), 32)
                    Do  'If not synced at the bottom dir of the library:
                        '   -> add folders below mount point to webRoot
                        On Error Resume Next: odFolders parentID
                        keyExists = (Err.Number = 0): On Error GoTo 0
                        If Not keyExists Then Exit Do
                        s = odFolders(parentID)(1) & "/" & s
                        parentID = odFolders(parentID)(0)
                    Loop
                    webRoot = libNrToWebColl(libNr)(1) & s
                    locToWebColl.Add VBA.Array(locRoot, webRoot, email, _
                                             syncID, syncFind, dirName), locRoot
                Case "AddedScope" 'One line per folder added as link to personal
                    relPath = parts(5): If relPath = " " Then relPath = ""  'lib
                    parts = Split(parts(4), " "): siteID = parts(1)
                    webID = parts(2): libID = parts(3): lnkID = parts(4)
                    FileName = "ClientPolicy_" & libID & siteID & lnkID & ".ini"
                    On Error Resume Next 'On error try backup method...
                    webRoot = cliPolColl(FileName)("DavUrlNamespace") & relPath
                    On Error GoTo 0
                    If webRoot = "" Then 'Backup method to find webRoot:
                        For Each vItem In cliPolColl
                            If vItem("SiteID") = siteID _
                            And vItem("WebID") = webID _
                            And vItem("IrmLibraryId") = libID Then
                                webRoot = vItem("DavUrlNamespace") & relPath
                                Exit For
                            End If
                        Next vItem
                    End If
                    If webRoot = vbNullString Then Err.Raise vbErrFileNotFound _
                                                           , methodName
                    s = vbNullString: parentID = Left$(Split(line, " ")(3), 32)
                    Do 'If link is not at the bottom of the personal library:
                        On Error Resume Next: odFolders parentID
                        keyExists = (Err.Number = 0): On Error GoTo 0
                        If Not keyExists Then Exit Do       'add folders below
                        s = odFolders(parentID)(1) & ps & s 'mount point to
                        parentID = odFolders(parentID)(0)   'locRoot
                    Loop
                    locRoot = mainMount & ps & s
                    locToWebColl.Add VBA.Array(locRoot, webRoot, email, _
                                     mainSyncID, mainSyncFind, dirName), locRoot
                Case Else: Exit For
                End Select
            Next line
        Case dirName = "Personal" 'Settings files for a personal OD account
        'Only one Personal OneDrive account can be signed in at a time.
            For Each line In Split(b, vbNewLine) 'Loop should exit at first line
                If line Like "library = *" Then
                    parts = Split(line, """"): locRoot = parts(3)
                    syncFind = locRoot: syncID = Split(parts(4), " ")(2)
                    Exit For
                End If
            Next line
            On Error Resume Next 'This file may be missing if the personal OD
            webRoot = cliPolColl("ClientPolicy.ini")("DavUrlNamespace") 'account
            On Error GoTo 0                  'was logged out of the OneDrive app
            If locRoot = "" Or webRoot = "" Or cid = "" Then GoTo NextFolder
            locToWebColl.Add VBA.Array(locRoot, webRoot & "/" & cid, email, _
                                       syncID, syncFind, dirName), key:=locRoot
            If Dir(vDir & "GroupFolders.ini") = "" Then GoTo NextFolder
            'Read GroupFolders.ini file
            cid = vbNullString: fileNum = FreeFile()
            Open vDir & "GroupFolders.ini" For Binary Access Read As #fileNum
                ReDim b(0 To LOF(fileNum)): Get fileNum, , b
            Close #fileNum: fileNum = 0
            #If Mac Then 'On Mac, the OneDrive settings files use UTF-8 encoding
                sUtf8 = b: GoSub DecodeUTF8
                b = sUtf16 'StrConv(b, vbUnicode) is UNRELIABLE
            #End If 'Two lines per synced folder from other peoples personal ODs
            For Each line In Split(b, vbNewLine)
                If line Like "*_BaseUri = *" And cid = vbNullString Then
                    cid = LCase$(Mid$(line, InStrRev(line, "/", , 0) + 1, _
                       InStrRev(line, "!", , 0) - InStrRev(line, "/", , 0) - 1))
                    folderID = Left$(line, InStr(line, "_") - 1)
                ElseIf cid <> vbNullString Then
                    locToWebColl.Add VBA.Array(locRoot & ps & odFolders( _
                                     folderID)(1), webRoot & "/" & cid & "/" & _
                                     Mid$(line, Len(folderID) + 9), email, _
                                     syncID, syncFind, dirName), _
                                key:=locRoot & ps & odFolders(folderID)(1)
                    cid = vbNullString: folderID = vbNullString
                End If
            Next line
        End Select
NextFolder:
        cid = vbNullString: s = vbNullString: email = vbNullString
    Next vDir

    'Clean the finished "dictionary" up, remove trailing "\" and "/"
    Dim tmpColl As Collection: Set tmpColl = New Collection
    For Each vItem In locToWebColl
        locRoot = vItem(0): webRoot = vItem(1): syncFind = vItem(4)
        If Right$(webRoot, 1) = "/" Then _
            webRoot = Left$(webRoot, Len(webRoot) - 1)
        If Right$(locRoot, 1) = ps Then _
            locRoot = Left$(locRoot, Len(locRoot) - 1)
        If Right$(syncFind, 1) = ps Then _
            syncFind = Left$(syncFind, Len(syncFind) - 1)
        tmpColl.Add VBA.Array(locRoot, webRoot, vItem(2), _
                              vItem(3), syncFind), locRoot
    Next vItem
    Set locToWebColl = tmpColl

    #If Mac Then 'deal with syncIDs
        If cloudStoragePathExists Then
            Set tmpColl = New Collection
            For Each vItem In locToWebColl
                locRoot = vItem(0): syncID = vItem(3): syncFind = vItem(4)
                locRoot = Replace(locRoot, syncFind, _
                                           syncIDtoSyncDir(syncID)(1), , 1)
                tmpColl.Add VBA.Array(locRoot, vItem(1), vItem(2)), locRoot
            Next vItem
            Set locToWebColl = tmpColl
        End If
    #End If

    GetLocalPath = GetLocalPath(path, returnAll, pmpo, False): Exit Function
    Exit Function
DecodeUTF8:         'StrConv doesn't work reliably, therefore UTF-8 must
    #If Mac Then    'be transcoded to UTF-16 manually (yes, this is insane)
        utf8 = sUtf8
        ReDim utf16(0 To (UBound(utf8) - LBound(utf8) + 1) * 2)
        i = LBound(utf8): k = 0
        Do While i <= UBound(utf8) 'Loop through the UTF-8 byte array
            'Determine the number of bytes in the current UTF-8 codepoint
            numBytesOfCodePoint = 1
            If utf8(i) And &H80 Then
                If utf8(i) And &H20 Then
                    If utf8(i) And &H10 Then
                        numBytesOfCodePoint = 4
                    Else: numBytesOfCodePoint = 3: End If
                Else: numBytesOfCodePoint = 2: End If
            End If
            If i + numBytesOfCodePoint - 1 > UBound(utf8) Then _
                Err.Raise vbErrInvalidFormatInResourceFile, methodName
            'Calculate the Unicode codepoint value from the UTF-8 bytes
            If numBytesOfCodePoint = 1 Then
                codepoint = utf8(i)
            Else: codepoint = utf8(i) And (2 ^ (7 - numBytesOfCodePoint) - 1)
                For j = 1 To numBytesOfCodePoint - 1
                    codepoint = (codepoint * 64) + (utf8(i + j) And &H3F)
                Next j
            End If
            'Convert the Unicode codepoint to UTF-16LE bytes
            If codepoint < &H10000 Then
                utf16(k) = codepoint And &HFF&
                utf16(k + 1) = codepoint \ &H100&
                k = k + 2
            Else 'Codepoint must be encoded as surrogate pair
                m = codepoint - &H10000
                highSurrogate = &HD800& Or (m \ &H400&)
                lowSurrogate = &HDC00& Or (m And &H3FF)
                utf16(k) = highSurrogate And &HFF&
                utf16(k + 1) = highSurrogate \ &H100&
                utf16(k + 2) = lowSurrogate And &HFF&
                utf16(k + 3) = lowSurrogate \ &H100&
                k = k + 4
            End If
            i = i + numBytesOfCodePoint 'Move to the next UTF-8 codepoint
        Loop
        If k > 0 Then
            ReDim Preserve utf16(0 To k - 1)
            sUtf16 = utf16
        Else: sUtf16 = ""
        End If
        Return 'Jump back to the statement after last encountered GoSub
    #End If
End Function

