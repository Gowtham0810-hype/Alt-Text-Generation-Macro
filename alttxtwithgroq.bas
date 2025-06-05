Attribute VB_Name = "Module1"
''
' VBA-JSON v2.3.1
' (c) Tim Hall - https://github.com/VBA-tools/VBA-JSON
'
' JSON Converter for VBA
'
' Errors:
' 10001 - JSON parse error
'


'============================================
' MODULE LEVEL DECLARATIONS (must be at top)
'============================================

' Windows API Declarations
Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As LongPtr
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function CopyImage Lib "user32" (ByVal handle As LongPtr, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As LongPtr
Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As LongPtr, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Private Declare PtrSafe Function OleLoadPicture Lib "olepro32" (pStream As Any, ByVal lSize As Long, ByVal fRunmode As Long, riid As Any, ppvObj As Any) As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' GDI+ Declarations
Private Declare PtrSafe Function GdiplusStartup Lib "gdiplus" (token As Long, inputbuf As Any, Optional ByVal outputbuf As LongPtr = 0) As Long
Private Declare PtrSafe Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As Long
Private Declare PtrSafe Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hbm As LongPtr, ByVal hpal As LongPtr, ByRef Bitmap As LongPtr) As Long
Private Declare PtrSafe Function GdipDisposeImage Lib "gdiplus" (ByVal Image As LongPtr) As Long
Private Declare PtrSafe Function GdipSaveImageToFile Lib "gdiplus" (ByVal Image As LongPtr, ByVal filename As LongPtr, ByRef clsidEncoder As Any, ByRef encoderParams As Any) As Long
Private Declare PtrSafe Function CLSIDFromString Lib "ole32" (ByVal str As LongPtr, id As Any) As Long

' Clipboard Formats
Private Const CF_BITMAP As Long = 2
Private Const CF_ENHMETAFILE As Long = 14
Private Const IMAGE_BITMAP As Long = 0
Private Const LR_COPYRETURNORG As Long = &H4

Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As LongPtr
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

''
' VBA-JSON v2.3.1
' (c) Tim Hall - https://github.com/VBA-tools/VBA-JSON

'============================================
' MAIN FUNCTIONALITY
'============================================

Sub GenerateAltTextForImagesWithGroq()
    Dim doc As Document
    Dim inlineShape As inlineShape
    Dim shp As Shape
    Dim groqApiKey As String
    Dim groqModel As String
    Dim base64Image As String
    Dim altText As String
    Dim processedCount As Long
    Dim skippedCount As Long
    Dim failedCount As Long
    Dim totalImages As Long
    Dim startTime As Double
    
    ' Initialize timing
    startTime = Timer
    
    ' Get API credentials
    groqApiKey = "your groq api key"
    groqModel = "meta-llama/llama-4-scout-17b-16e-instruct"
    
    Set doc = ActiveDocument
    totalImages = CountImagesInDocument(doc)
    
    ' Initialize counters
    processedCount = 0
    skippedCount = 0
    failedCount = 0
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Process InlineShapes
    For Each inlineShape In doc.InlineShapes
        If inlineShape.Type = wdInlineShapePicture Then
            ' Skip if already has alt text
            If Len(inlineShape.AlternativeText) > 0 Then
                skippedCount = skippedCount + 1
                GoTo NextInlineShape
            End If
            
            ' Get image as Base64
            base64Image = GetImageBase64Direct(inlineShape)
            
            If base64Image <> "" Then
                altText = CallGroqVisionAPI(base64Image, groqApiKey, groqModel)
                
                If altText <> "" Then
                    inlineShape.AlternativeText = altText
                    processedCount = processedCount + 1
                Else
                    failedCount = failedCount + 1
                End If
            Else
                failedCount = failedCount + 1
            End If
        End If
NextInlineShape:
    Next inlineShape
    
    ' Process Shapes
    For Each shp In doc.Shapes
        If shp.Type = msoPicture Then
            ' Skip if already has alt text
            If Len(shp.AlternativeText) > 0 Then
                skippedCount = skippedCount + 1
                GoTo NextShape
            End If
            
            ' Get image as Base64
            base64Image = GetImageBase64Direct(shp)
            
            If base64Image <> "" Then
                altText = CallGroqVisionAPI(base64Image, groqApiKey, groqModel)
                
                If altText <> "" Then
                    shp.AlternativeText = altText
                    processedCount = processedCount + 1
                Else
                    failedCount = failedCount + 1
                End If
            Else
                failedCount = failedCount + 1
            End If
        End If
NextShape:
    Next shp
    
    ' Clean up
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    ' Show results
    Dim timeElapsed As Double
    timeElapsed = Round(Timer - startTime, 2)
    
    MsgBox "Alt text generation complete!" & vbCrLf & _
            "Total Images: " & totalImages & vbCrLf & _
            "Processed: " & processedCount & vbCrLf & _
            "Skipped (had alt text): " & skippedCount & vbCrLf & _
            "Failed: " & failedCount & vbCrLf & _
            "Time taken: " & timeElapsed & " seconds", _
            vbInformation, "Results"
End Sub

'============================================
' IMAGE PROCESSING FUNCTIONS
'============================================

Function GetImageBase64Direct(imgObject As Object) As String
    Dim tempFile As String
    Dim fso As Object
    Dim success As Boolean
    
    On Error GoTo ErrorHandler
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    tempFile = Environ("TEMP") & "\" & fso.GetTempName & ".png"
    
    Debug.Print "the type name of the image: " & TypeName(imgObject)
    ' Copy image to clipboard
    If TypeName(imgObject) = "InlineShape" Then
        imgObject.Range.Select
        Selection.Copy
        DoEvents: Sleep 100
        Selection.Copy
        DoEvents: Sleep 200
    Else
        imgObject.Copy
    End If
    
    ' Wait for clipboard
    DoEvents
    Sleep 200
    
    ' Save clipboard image to file
    success = SaveClipboardImageToFile(tempFile)
    
    If success Then
        GetImageBase64Direct = EncodeFileToBase64(tempFile)
    End If
    
    ' Clean up
    If fso.FileExists(tempFile) Then fso.DeleteFile tempFile
    
    Exit Function
    
ErrorHandler:
    Debug.Print "Error in GetImageBase64Direct: " & Err.Description
    If Not fso Is Nothing Then
        If fso.FileExists(tempFile) Then fso.DeleteFile tempFile
    End If
    GetImageBase64Direct = ""
End Function

Function SaveClipboardImageToFile(filePath As String) As Boolean
    Dim hBitmap As LongPtr
    Dim hCopy As LongPtr
    Dim gdiToken As Long
    
    Dim bitmapHandle As LongPtr
    Dim clsidEncoder(0 To 15) As Byte
    Dim hr As Long
    
    On Error GoTo ErrorHandler
    Dim gdiInput As GdiplusStartupInput
    ' Initialize GDI+
    gdiInput.GdiplusVersion = 1
    If GdiplusStartup(gdiToken, gdiInput) <> 0 Then
        Exit Function
    End If
    
    ' Open clipboard
    If OpenClipboard(0) = 0 Then Exit Function
    
    
    
    ' Get bitmap from clipboard
    hBitmap = GetClipboardData(CF_BITMAP)
    If hBitmap = 0 Then
        CloseClipboard
        Exit Function
    End If
    
    ' Make copy of bitmap
    hCopy = CopyImage(hBitmap, IMAGE_BITMAP, 0, 0, LR_COPYRETURNORG)
    CloseClipboard
    
    If hCopy = 0 Then Exit Function
    
    ' Create GDI+ bitmap from HBITMAP
    If GdipCreateBitmapFromHBITMAP(hCopy, 0, bitmapHandle) <> 0 Then
        DeleteObject hCopy
        GdiplusShutdown gdiToken
        Exit Function
    End If
    
    ' Get PNG encoder CLSID
    If CLSIDFromString(StrPtr("{557CF406-1A04-11D3-9A73-0000F81EF32E}"), clsidEncoder(0)) <> 0 Then
        GdipDisposeImage bitmapHandle
        DeleteObject hCopy
        GdiplusShutdown gdiToken
        Exit Function
    End If
    
    ' Save to file
    hr = GdipSaveImageToFile(bitmapHandle, StrPtr(filePath), clsidEncoder(0), ByVal 0&)
    
    ' Clean up
    GdipDisposeImage bitmapHandle
    DeleteObject hCopy
    GdiplusShutdown gdiToken
    
    SaveClipboardImageToFile = (hr = 0)
    Exit Function
    
ErrorHandler:
    Debug.Print "Error in SaveClipboardImageToFile: " & Err.Description
    SaveClipboardImageToFile = False
End Function

'============================================
' API COMMUNICATION FUNCTIONS
'============================================
Function CallGroqVisionAPI(base64Image As String, apiKey As String, modelName As String) As String
    Dim objHTTP As Object
    Dim jsonRequest As String, jsonResponse As String
    Dim parsedResponse As Object
    Dim altTextContent As String
    Dim json As Object ' Add JSON parser object
    Dim choice As Object
    Dim content As String
    Dim startPos As Long
    Dim endPos As Long
    
    
    On Error GoTo ErrorHandler
    
    ' Initialize JSON parser (alternative if JsonConverter fails)
    Set json = CreateObject("Scripting.Dictionary")
    Set objHTTP = CreateObject("MSXML2.XMLHTTP")
    
    ' Construct the JSON request
    jsonRequest = "{""messages"": [{""role"": ""user"", ""content"": [{""type"": ""text"", ""text"": ""Describe this image concisely for alt text.just return the visual description nothing else. ""}, {""type"": ""image_url"", ""image_url"": {""url"": ""data:image/png;base64," & base64Image & """, ""detail"": ""low""}}]}], ""model"": """ & modelName & """}"
    
    With objHTTP
        .Open "POST", "https://api.groq.com/openai/v1/chat/completions", False
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "Authorization", "Bearer " & apiKey
        .send jsonRequest
        
        If .Status = 200 Then
            jsonResponse = .responseText
            ' Debug.Print "Raw JSON Response: " & jsonResponse
            
            startPos = InStr(jsonResponse, """content"":""")
            If startPos > 0 Then
                startPos = startPos + Len("""content"":""")
                endPos = InStr(startPos, jsonResponse, """}")
                If endPos > 0 Then
                    content = Mid(jsonResponse, startPos, endPos - startPos)
                    content = Replace(content, "\""", """") ' Unescape quotes
                    Debug.Print "? Extracted content: " & content
                End If
            End If
            
            
            altTextContent = content
            
            
            
            ' Clean and return the result
            altTextContent = Trim(Replace(altTextContent, vbLf, " "))
            altTextContent = Left(altTextContent, 150)
            CallGroqVisionAPI = altTextContent
        Else
            Debug.Print "API Error: " & .Status & " - " & .StatusText
            Debug.Print "Response: " & .responseText
            CallGroqVisionAPI = ""
        End If
    End With
    
    Exit Function
    
ErrorHandler:
    Debug.Print "Error in CallGroqVisionAPI: " & Err.Description & " (Line " & Erl & ")"
    CallGroqVisionAPI = ""
End Function

'============================================
' HELPER FUNCTIONS
'============================================

Function CountImagesInDocument(doc As Document) As Long
    Dim count As Long
    Dim shp As Shape
    Dim ils As inlineShape
    
    count = 0
    
    For Each ils In doc.InlineShapes
        If ils.Type = wdInlineShapePicture Then count = count + 1
    Next ils
    
    For Each shp In doc.Shapes
        If shp.Type = msoPicture Then count = count + 1
    Next shp
    
    CountImagesInDocument = count
End Function

Function EncodeFileToBase64(filePath As String) As String
    Dim objStream As Object
    Dim objXML As Object
    Dim objNode As Object
    Dim base64String As String
    
    On Error GoTo ErrorHandler
    
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Type = 1 ' adTypeBinary
    objStream.Open
    objStream.LoadFromFile filePath
    
    Set objXML = CreateObject("MSXML2.DOMDocument")
    Set objNode = objXML.createElement("b64")
    objNode.dataType = "bin.base64"
    objNode.nodeTypedValue = objStream.Read
    base64String = objNode.Text
    
    base64String = Replace(base64String, vbCrLf, "")
    base64String = Replace(base64String, vbLf, "")
    base64String = Replace(base64String, vbCr, "")
    base64String = Replace(base64String, Chr(10), "") ' Remove \n
    base64String = Replace(base64String, Chr(13), "") ' Remove \r
    
    EncodeFileToBase64 = base64String
    
    objStream.Close
    Exit Function
    
ErrorHandler:
    Debug.Print "Error in EncodeFileToBase64: " & Err.Description
    EncodeFileToBase64 = ""
    If Not objStream Is Nothing Then objStream.Close
End Function

' Removed UpdateProgress function as it relies on frmProgress
' Sub UpdateProgress(current As Long, total As Long, Optional message As String = "")
'     On Error Resume Next ' In case form isn't loaded
'     With frmProgress
'         .lblProgress.Caption = current & " of " & total & " (" & Format(current / total, "0%") & ")"
'         If message <> "" Then .lblStatus.Caption = message
'         .Repaint
'     End With
'     DoEvents
' End Sub

'============================================
' TYPE DEFINITIONS
'============================================





