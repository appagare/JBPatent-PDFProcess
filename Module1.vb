Imports System.IO
Module PDFProcess

    Sub Main()

        'ordinal
        '1 = method

        'if "v" (viewall)
        '2 = format
        '3 = AllFileName 
        '4 Final As boolean 0/1, 
        '5=SrcPath As String,
        '6 = DstPath As String 
        '7= Admin As Boolean 0/1, 
        '8 WaterMarkText As String - spaces in string
        '9 Option1 As String, 
        '10 = Option2 As String
        '11 = LicenseKey As String

        'if "c" cleanup temp folder
        '2 = DstPath / TempFolder
        '3 = days to keep

        Dim args() As String = System.Environment.GetCommandLineArgs()

        'For i As Integer = 0 To args.Length - 1
        '    Console.WriteLine("Arg: " & i & " is " & args(i))
        'Next

        If args.Length > 1 Then
            Select Case args(1)
                Case "v" 'viewall
                    If args.Length > 10 Then
                        System.Environment.Exit(ViewAll(args(2), args(3), IIf(args(4) = "1" OrElse LCase(args(4)) = "true", True, False), args(5), args(6), IIf(args(7) = "1" OrElse LCase(args(7)) = "true", True, False), args(8), args(9), args(10), args(11)))
                    Else
                        System.Environment.Exit(-2)
                    End If
                Case "c" 'clean up temp folder
                    If args.Length > 2 Then
                        CleanUpTempFolder(args(2), args(3), "")
                        System.Environment.Exit(0)
                    Else
                        System.Environment.Exit(-2)
                    End If
                Case Else
                    System.Environment.Exit(-1)
            End Select
        End If


    End Sub

    Private Function ViewAll(ByVal FileFormat As String, ByVal AllFileName As String,
ByVal Final As Boolean, ByVal SrcPath As String,
ByVal DstPath As String, ByVal Admin As Boolean, ByVal WaterMarkText As String, ByVal Option1 As String, ByVal Option2 As String, ByVal LicenseKey As String) As Integer

        'possible performance enhancement - make all final copies when using Formalizes

        On Error GoTo Err_Handler
        Dim DebugError As Integer = -100
        FileFormat = Trim(Replace(FileFormat, "|", " "))
        AllFileName = Trim(Replace(AllFileName, "|", " "))
        SrcPath = Trim(Replace(SrcPath, "|", " "))
        DstPath = Trim(Replace(DstPath, "|", " "))
        WaterMarkText = Trim(Replace(WaterMarkText, "|", " "))
        Option1 = Trim(Replace(Option1, "|", " "))
        Option2 = Trim(Replace(Option2, "|", " "))
        LicenseKey = Trim(Replace(LicenseKey, "|", " "))

        'currently - always regenerate drafts, keep formal for 1hr
        Dim HoursToKeepFinal As Integer = 1

        If IsNumeric(Option1) = True Then
            HoursToKeepFinal = CType(Option1, Integer)
        End If

        'Console.WriteLine(LicenseKey)
        'Console.ReadKey()

        If LicenseKey <> "" Then
            DebugError = -98
            PDFTech.PDFDocument.License = LicenseKey
        End If
        DebugError = -97
        Dim options As New PDFTech.PDFCreationOptions()
        If WaterMarkText <> "" AndAlso LCase(FileFormat) = "dra" Then
            DebugError = -96
            options.Watermark.SetText(WaterMarkText)
        End If

        DebugError = -95
        Dim obj As PDFTech.PDFDocument
        Dim basefile As String = ""
        Dim x As Integer = 0
        Dim y As Integer = 0
        'Dim NumPages1 As Integer = 0
        'Dim NumPages2 As Integer = 0
        Dim FirstPage As Boolean = True
        Dim format2 As String = ""

        DebugError = -94
        Dim o As New DirectoryInfo(SrcPath)
        Dim FileList As FileInfo()
        Dim FileDetail As FileInfo
        DebugError = -93
        FileList = o.GetFiles

        DebugError = -92
        If Admin = False Then
            'from main
            x = 0
            DebugError = -91
            'sets format2 to either PCT, EPO, or format
            If FileFormat = "pct" Then
                For Each FileDetail In FileList
                    If InStr(1, SrcPath & FileDetail.Name, "PCT", 1) > 0 Then
                        x = 1
                    End If
                Next
                If x = 1 Then
                    format2 = "pct"
                Else
                    format2 = "epo"
                End If
            Else
                format2 = FileFormat
            End If
            basefile = ""
            y = 0
            DebugError = -90

            'rule 
            '- keep DRA for x hrs since it would take time for JBL to generate new drafts
            '- keep non-DRA versions since they should not be out of date unless user wants to re-do them which runs it through the dra process again
            '- so, if dra, delete other versions usa, ua4, epo, pct, sff

            If LCase(format2) = "dra" Then
                On Error Resume Next 'proceed no matter what
                'if this is a draft request, delete any non-draft versions (they shouldn't exist unless redoing a previously formalized drawing)
                File.Delete(DstPath & "usa" & Right(AllFileName, Len(AllFileName) - 3))
                File.Delete(DstPath & "ua4" & Right(AllFileName, Len(AllFileName) - 3))
                File.Delete(DstPath & "epo" & Right(AllFileName, Len(AllFileName) - 3))
                File.Delete(DstPath & "pct" & Right(AllFileName, Len(AllFileName) - 3))
                File.Delete(DstPath & "sff" & Right(AllFileName, Len(AllFileName) - 3))
            End If

            If File.Exists(DstPath & AllFileName) = True Then
                'If format2 <> "dra" Then
                '    Dim ThisFile As New FileInfo(DstPath & AllFileName)
                '    If DateDiff(DateInterval.Hour, ThisFile.CreationTime, Now) > HoursToKeepFinal Then
                '        'if this is non-stale final, use it 
                '        Return 0
                '    End If
                'End If
                On Error Resume Next 'proceed no matter what
                File.Delete(DstPath & AllFileName) 'if here, we are deleting/rebuilding a stale draft
            End If
            On Error GoTo Err_Handler 'always resume normal error handling

            obj = New PDFTech.PDFDocument(DstPath & AllFileName, options) 'create new ALL file, no matter what
            DebugError = -87
            'not a bac (normal merge)
            For Each FileDetail In FileList
                If InStr(1, SrcPath & FileDetail.Name, format2, 1) > 0 Then
                    DebugError = -86
                    obj.LoadPdf(SrcPath & FileDetail.Name, "")
                End If
            Next
            DebugError = -85
        Else
            'from Admin - not presently used
            DebugError = -84
            obj = New PDFTech.PDFDocument(DstPath & AllFileName, options) 'create new ALL file, no matter what

            If FileFormat = "pct" Then
                format2 = "epo"
            Else
                format2 = FileFormat
            End If
            basefile = ""
            DebugError = -80
            For Each FileDetail In FileList
                If InStr(1, SrcPath & FileDetail.Name, format2, 1) > 0 Then
                    DebugError = -79
                    obj.LoadPdf(SrcPath & FileDetail.Name, "")
                End If
            Next
            DebugError = -78
        End If

        DebugError = -77
        obj.Pages.Delete(obj.Pages(0)) 'delete the first, blank page?
        obj.Save() 'save whatever we created

        Return 0
        Exit Function
Err_Handler:

        Return DebugError

    End Function

    Private Sub CleanUpTempFolder(ByVal DstPath As String, ByVal DaysToKeep As Integer, ByVal ExceptionMask As String)
        'On Error Resume Next
        Dim o As New DirectoryInfo(DstPath)
        Dim FileList As FileInfo()
        Dim FileDetail As FileInfo
        FileList = o.GetFiles
        For Each FileDetail In FileList
            'if date1 Later than date2, datediff is neg
            If DateDiff(DateInterval.Day, FileDetail.LastAccessTime, Now) > DaysToKeep AndAlso UCase(FileDetail.Name) <> UCase(ExceptionMask) Then
                'older than X days and not a current file
                FileDetail.Delete()
            End If
        Next
    End Sub

End Module
