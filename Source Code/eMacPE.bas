Attribute VB_Name = "eMacPE"
'''''''''''''''''''''''''''''''''
' WebTV IPE (In-place Edit) 4.0 '
'                               '
' By: Eric MacDonald            '
' Date: April 24, 2005          '
'                               '
' This is a patcher tool        '
' for any SuperViewer template  '
'''''''''''''''''''''''''''''''''

Option Explicit


Public Type IMAGE_DOS_HEADER
   Magic    As Integer
   cblp     As Integer
   cp       As Integer
   crlc     As Integer
   cparhdr  As Integer
   minalloc As Integer
   maxalloc As Integer
   ss       As Integer
   sp       As Integer
   csum     As Integer
   ip       As Integer
   cs       As Integer
   lfarlc   As Integer
   ovno     As Integer
   res(3)   As Integer
   oemid    As Integer
   oeminfo  As Integer
   res2(9)  As Integer
   lfanew      As Long
End Type

Public Type IMAGE_DATA_DIRECTORY
   DataRVA     As Long
   DataSize    As Long
End Type

Public Type IMAGE_FILE_HEADER
   Machine              As Integer
   NumberOfSections     As Integer
   TimeDateStamp        As Long
   PointerToSymbolTable As Long
   NumberOfSymbols      As Long
   SizeOfOtionalHeader  As Integer
   Characteristics      As Integer
End Type

Public Type IMAGE_OPTIONAL_HEADER
   Magic             As Integer
   MajorLinkVer      As Byte
   MinorLinkVer      As Byte
   CodeSize          As Long
   InitDataSize      As Long
   unInitDataSize    As Long
   EntryPoint        As Long
   CodeBase          As Long
   DataBase          As Long
   ImageBase         As Long
   SectionAlignment  As Long
   FileAlignment     As Long
   MajorOSVer        As Integer
   MinorOSVer        As Integer
   MajorImageVer     As Integer
   MinorImageVer     As Integer
   MajorSSVer        As Integer
   MinorSSVer        As Integer
   Win32Ver          As Long
   ImageSize         As Long
   HeaderSize        As Long
   Checksum          As Long
   Subsystem         As Integer
   DLLChars          As Integer
   StackRes          As Long
   StackCommit       As Long
   HeapReserve       As Long
   HeapCommit        As Long
   LoaderFlags       As Long
   RVAsAndSizes      As Long
   DataEntries(15)   As IMAGE_DATA_DIRECTORY
End Type

Public Type IMAGE_SECTION_HEADER
   sectionName(7)    As Byte
   Address           As Long
   VirtualAddress    As Long
   SizeOfData        As Long
   PData             As Long
   PReloc            As Long
   PLineNums         As Long
   RelocCount        As Integer
   LineCount         As Integer
   Characteristics   As Long
End Type

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Public Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Public Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long

Private SectionAlignment   As Long
Private FileAlignment      As Long
Private ResSectionRVA      As Long
Private ResSectionOffset   As Long
Private DestinationString() As Byte



Public Function updateSection(fileHandle As Long, sectionName As String, sectionRepl As String, exeImage() As Byte, sizeOfArry As Long)
    Dim lngBytesRead As Long
    Dim crazyMZHeader As IMAGE_DOS_HEADER
    Dim crazyPEHeader1 As IMAGE_FILE_HEADER
    Dim crazyPEHeader2 As IMAGE_OPTIONAL_HEADER
    Dim crazyPESections() As IMAGE_SECTION_HEADER
    Dim sectionNameE As String
    Dim i As Integer
    Dim chgOff As Long
    Dim sizTemp As Long
    Dim sizATemp As Long
    Dim sizATemp2 As Long
    Dim tempExe2 As String
    Dim tempExe() As Byte
    Dim VIEWER2 As Integer
    Dim sizDif As Long
    Dim whereAt As Long

    
    ' DOS HEADER
    SetFilePointer fileHandle, ByVal 0, 0, 0
    ReadFile fileHandle, crazyMZHeader, ByVal Len(crazyMZHeader), lngBytesRead, ByVal 0&
    
    ' PE Header
    SetFilePointer fileHandle, ByVal crazyMZHeader.lfanew + 4, 0, 0
    ReadFile fileHandle, crazyPEHeader1, ByVal Len(crazyPEHeader1), lngBytesRead, ByVal 0&
    ReadFile fileHandle, crazyPEHeader2, ByVal Len(crazyPEHeader2), lngBytesRead, ByVal 0&
    
    ' --Section header--
    ReDim crazyPESections(crazyPEHeader1.NumberOfSections - 1) As IMAGE_SECTION_HEADER
    ReadFile fileHandle, crazyPESections(0), ByVal Len(crazyPESections(0)) * crazyPEHeader1.NumberOfSections, lngBytesRead, ByVal 0&
    
    sizATemp2 = 0
    chgOff = 0
    For i = 0 To UBound(crazyPESections)
        sectionNameE = StrConv(crazyPESections(i).sectionName, vbUnicode)
        sectionNameE = Left(sectionNameE, InStr(sectionNameE, Chr(0)) - 1)
        
        If sectionNameE = sectionName Then
            
            crazyPESections(i).Address = Len(sectionRepl)
            
            sizTemp = crazyPESections(i).Address + (crazyPEHeader2.FileAlignment - (crazyPESections(i).Address Mod crazyPEHeader2.FileAlignment))
            
            sizDif = sizTemp - crazyPESections(i).SizeOfData
            crazyPESections(i).SizeOfData = sizTemp
            
            sectionRepl = pack(sectionRepl, crazyPESections(i).SizeOfData)
            
            sizATemp = crazyPESections(i).PData + Len(sectionRepl)
            
            crazyPEHeader2.InitDataSize = crazyPEHeader2.InitDataSize + sizDif
            
            ReDim tempExe(0 To sizATemp) As Byte
            
            sizATemp2 = crazyPESections(i).VirtualAddress + crazyPESections(i).Address
            
            crazyPEHeader2.ImageSize = sizATemp2 + (crazyPEHeader2.SectionAlignment - (sizATemp2 Mod crazyPEHeader2.SectionAlignment))
            crazyPEHeader2.DataEntries(2).DataSize = crazyPESections(i).SizeOfData
            tempExe2 = StrConv(exeImage, vbUnicode)
            tempExe2 = Left(tempExe2, crazyPESections(i).PData) & sectionRepl
            
            ReDim DestinationString(0 To sizATemp) As Byte
            Call CopyMemory(DestinationString(0), ByVal tempExe2, Len(tempExe2))
            
            whereAt = crazyMZHeader.lfanew + 4
            Call CopyMemory(DestinationString(whereAt), ByVal VarPtr(crazyPEHeader1), Len(crazyPEHeader1))
            
            whereAt = whereAt + Len(crazyPEHeader1)
            Call CopyMemory(DestinationString(whereAt), ByVal VarPtr(crazyPEHeader2), Len(crazyPEHeader2))
            
            whereAt = whereAt + Len(crazyPEHeader2)
            Call CopyMemory(DestinationString(whereAt), ByVal VarPtr(crazyPESections(0)), Len(crazyPESections(0)) * crazyPEHeader1.NumberOfSections)
        End If
    Next i
    
End Function

Public Function pack(block As String, Size As Long) As String
    Dim i As Long
    
    For i = Len(block) + 1 To Size
        block = block & Chr(0)
    Next i
    
    pack = block
End Function


Public Function writeHeader(tempDict As Dictionary, mode As Integer, backup As Boolean) As Integer
    Dim hashItem
    Dim hashItem2
    Dim block As String
    Dim DestSize As Long
    Dim Size As Long
    Dim Offset As Long
    Dim OPENFILE As String
    Dim fileSiz As Long
    Dim readByes As Long
    Dim crazyMZHeader As IMAGE_DOS_HEADER
    Dim progressTot As Long
    Dim progressIn As Long
    Dim VIEWER As Long
    Dim VIEWER2 As Long
    
    progressTot = 0
    
    Err = 0

    Select Case mode
        Case 1:
            progressTot = frmMain.blockVars.count
        Case 2:
            progressTot = frmMain.editCodes.count
        Case 3:
            progressTot = frmMain.editCodes.count + frmMain.blockVars.count
    End Select
    
    progressTot = progressTot + 1
    
    OPENFILE = App.Path & "\" & frmMain.pathto
    
    VIEWER = CreateFile(OPENFILE, ByVal &H80000000, 0, ByVal 0&, 3, 0, ByVal 0)

    If Err <> 0 Then
        MsgBox "writeHeader(): I couldn't open the file '" & OPENFILE & "' I had an Error: " & Err, vbCritical, "Something's fucked up"
        writeHeader = 0
    Else
        frmProgress.Show
        
        fileSiz = GetFileSize(VIEWER, 0)
        ReDim DestinationString(0 To fileSiz)
    
        frmProgress.Label1.Caption = "Reading viewer data"
        
        ReadFile VIEWER, DestinationString(0), fileSiz, readByes, ByVal 0&

        If backup = True Then
            frmProgress.Label1.Caption = "Backing up data..."
            
            VIEWER2 = FreeFile
            Open OPENFILE & "BACK" & ".exe" For Binary Access Write As #VIEWER2
            Put #VIEWER2, , DestinationString
            Close #VIEWER2
        End If
    If mode = 1 Or mode = 3 Then
    For Each hashItem In frmMain.blockVars
        If frmMain.blockVars(hashItem).Exists("headers") <> False Then
            block = ""
            For Each hashItem2 In frmMain.blockVars(hashItem)("headers")
                If tempDict.Exists(hashItem2) <> False Then
                    block = block & hashItem2 & ": " & tempDict(hashItem2) & vbCrLf
                End If
            Next hashItem2
            
            If frmMain.blockVars(hashItem)("write-end") = "yes" Then
                block = block & hashItem
            End If
            
            Size = CLng(frmMain.blockVars(hashItem)("block-size"))
            Offset = CLng("&H" & frmMain.blockVars(hashItem)("block-offset"))
            
            frmProgress.Label1.Caption = "Writing " & hashItem & " at " & Offset
            block = pack(block, Size)
            
            Call CopyMemory(DestinationString(Offset), ByVal block, Len(block))

        End If
        
        progressIn = progressIn + 1
        frmProgress.Shape2.Width = (progressIn * 3255) / progressTot
    
    Next hashItem
    
    End If
    
    If mode = 2 Or mode = 3 Then
        For Each hashItem In frmMain.editCodes
        
        frmProgress.Label1.Caption = "Writing enumerative data at " & hashItem
            
            If Left(hashItem, 4) = "sec_" Then
                Call updateSection(VIEWER, Mid(hashItem, 5), frmMain.editCodes(hashItem), DestinationString, fileSiz)
            Else
                Offset = CLng("&H" & hashItem)
                block = frmMain.editCodes(hashItem)
                Size = Len(block)
                
                Call CopyMemory(DestinationString(Offset), ByVal block, Size)
                
                progressIn = progressIn + 1
                frmProgress.Shape2.Width = (progressIn * 3255) / progressTot
            
            End If
        Next hashItem
    End If
    
    frmProgress.Label1.Caption = "Closing read viewer"
    
    CloseHandle (VIEWER)
    
    frmProgress.Label1.Caption = "Writing to viewer"


    VIEWER = FreeFile
    Kill OPENFILE
    Open OPENFILE For Binary Access Write As #VIEWER
    Put #VIEWER, , DestinationString
    Close #VIEWER
    
    progressIn = progressIn + 1
    frmProgress.Shape2.Width = (progressIn * 3255) / progressTot
    
    frmProgress.Hide
    
    writeHeader = 1
    
    End If
    

End Function

