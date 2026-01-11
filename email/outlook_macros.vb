'~~> Set a reference to Microsoft Internet Controls
'~~> Set a reference to Microsoft Scripting Runtime

'~~> The GetWindow function retrieves the handle of a window that has
'~~> the specified relationship (Z order or owner) to the specified window.
Private Declare PtrSafe Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long

'~~> The GetForegroundWindow function returns the handle of the foreground
'~~> window (the window with which the user is currently working).
Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As Long

' for the outlook folder "archivo"
' sources: https://chat.openai.com/c/437bc0a0-2621-42d6-81c6-dbb861f41843
Public ArchivoFolder As Outlook.folder
Public ExitLoop As Boolean

Private Sub Application_Startup()

    ' for the outlook folder "archivo"
    Call FoldersFindArchive
    
End Sub

Function GetURL() As String

Dim sw As SHDocVw.ShellWindows
Dim objIE As SHDocVw.InternetExplorer
Dim topHwnd As Long, nextHwnd As Long
Dim sURL As String, hwnds As String

    On Error GoTo errHandler
    
    Set sw = New SHDocVw.ShellWindows

    '~~> Check the number of IE Windows Opened
    '~~> If more than 1
    hwnds = "|"
    
    If sw.Count > 1 Then
        '~~> Create a string of hwnds of all IE windows
        For Each objIE In sw
            '~~> si es una instancia de TF7 no lo considera, si no sigue. la instancia se encuentra por el "path"
            If objIE.path = "c:\tf7\bin\" Then
            Else: hwnds = hwnds & objIE.hwnd & "|"
            End If
        Next

        '~~> Get handle of handle of the foreground window
        nextHwnd = GetForegroundWindow

        '~~> Check for the 1st IE window after foreground window
        Do While nextHwnd > 0
            nextHwnd = GetWindow(nextHwnd, 2&)
            If InStr(hwnds, "|" & nextHwnd & "|") > 0 Then
                topHwnd = nextHwnd
                Exit Do
            End If
        Loop

        '~~> Get the URL from the relevant IE window
        
        For Each objIE In sw
            '~~> como da un error de automatización cuando está abierto el TF7, en la segunda llamada a "sw", en vez de
            '~~> fijar la condición igual a topHwnd, la fijo como distinta, de forma que cuando falla no haga nada
            If objIE.hwnd <> topHwnd Then
            Else
            sURL = objIE.LocationURL
            Exit For
            End If
        Next

    '~~> If only 1 was found
    Else
        For Each objIE In sw
            sURL = objIE.LocationURL
        Next
    End If
    
    ' sustituye los caracteres del formato http por formato explorador de windows
    sURL = Replace(sURL, "/", "\")
    
    ' sustituye el formato de ruta con servidor por el L:\
    sURL = Replace(sURL, "file:\\S5557D16\depa16\", "L:\")
    sURL = Replace(sURL, "file:\\S5557d16\depa16\", "L:\")
    sURL = Replace(sURL, "file:\\S5557D16\DEPA16\", "L:\")
    sURL = Replace(sURL, "file:\\S5557d16\DEPA16\", "L:\")
    sURL = Replace(sURL, "file:\10.116.253.164\Depa16\", "L:\")
    
    ' sustituye el formato de ruta con servidor por el Y:\ con la función "environ"
    ' hay diferentes formatos según como llegue del explorador..
    sURL = Replace(sURL, "file:\\Svcphd01\users01\EXTERNOS\" & UCase(Environ("UserName")), "Y:")
    sURL = Replace(sURL, "file:\\svcphd01\users01\EXTERNOS\" & UCase(Environ("UserName")), "Y:")
        
    ' sustituye el formato de ruta con disco\
    sURL = Replace(sURL, "file:\\\", "")
    
    ' sustituye los caracteres del formato http por formato explorador de windows
    sURL = Replace(sURL, "%20", " ")
    sURL = Replace(sURL, "%E1", "á")
    sURL = Replace(sURL, "%E9", "é")
    sURL = Replace(sURL, "%ED", "í")
    sURL = Replace(sURL, "%F3", "ó")
    sURL = Replace(sURL, "%FA", "ú")
    sURL = Replace(sURL, "%E7", "ç")
    sURL = Replace(sURL, "%F1", "ñ")
    sURL = Replace(sURL, "%C1", "Á")
    sURL = Replace(sURL, "%C9", "É")
    sURL = Replace(sURL, "%CD", "Í")
    sURL = Replace(sURL, "%D3", "Ó")
    sURL = Replace(sURL, "%DA", "Ú")
    sURL = Replace(sURL, "%C7", "Ç")
    sURL = Replace(sURL, "%D1", "Ñ")
        
    ' añade corchete final
    sURL = sURL & "\"

    ' si se ha seleccionado una unidad raíz, aparecen dos corchetes en vez de uno
    sURL = Replace(sURL, "\\", "\")

    ' devuelve la variable como resultado de la función
    GetURL = sURL

    ' libera objetos y memoria
    Set sw = Nothing: Set objIE = Nothing
    
errHandler:

    ' se ignoran conscientemente estos errores, pues se piensa que pueden ser provocados por el TF7. Existe un riesgo.
    Select Case Err.Number
    ' error de automatización
    Case -2147467259: Resume Next
    ' llamada a miembro o a procedimiento no válido
    Case 5: Resume Next
    Case Else: Err.Raise Err.Number
    End Select
    
End Function

Function GetSequentialNumber(SourceFolderName As String) As String

Dim fso As Scripting.FileSystemObject
Dim SourceFolder As Scripting.folder
Dim FileItem As Scripting.File
Dim contador As Integer
Dim n As Integer
Dim str As String
Dim nFicheros As Integer

    ' inicializa el contador
    contador = 0
    
    ' lists information about the files in SourceFolder ' example: ListFilesInFolder "C:\FolderName\", True Dim FSO As Scripting.FileSystemObject Dim SourceFolder As Scripting.Folder, SubFolder As Scripting.Folder Dim FileItem As Scripting.File Dim r As Long
    Set fso = New Scripting.FileSystemObject
    Set SourceFolder = fso.GetFolder(SourceFolderName)
       
    For Each FileItem In SourceFolder.Files
    
        ' busca de cada fichero las primeras tres letras, y las pasa a número
        str = FileItem.Name
        If Val(Left(str, 3)) > 0 Then n = Val(Left(str, 3))
        ' si consigue pasarlas a numérico si es mayor que el contador, le asigna el valor encontrado
        If n > contador Then contador = n
        
    Next FileItem
    
    ' suma uno al contador para que se grabe el fichero en el próximo secuencial
    contador = contador + 1
    str = Right("000" & contador, 3)
    
    ' devuelve el resultado de la función
    GetSequentialNumber = str
    
    ' libera memoria y objetos
    Set fso = Nothing
    Set FileItem = Nothing
    Set SourceFolder = Nothing
    
End Function

Public Sub DeleteAttachments()

Dim objOL As Outlook.Application
Dim objMsg As Outlook.MailItem 'Object
Dim objAttachments As Outlook.Attachments
Dim objSelection As Outlook.Selection
Dim i As Long
Dim lngCount As Long
Dim Counter As Long

    ' Instantiate an Outlook Application object.
    Set objOL = CreateObject("Outlook.Application")

    ' Get the collection of selected objects.
    Set objSelection = objOL.ActiveExplorer.Selection

    ' Check each selected item for attachments.
    Counter = 1
    For Each objMsg In objSelection
    
        Set objAttachments = objMsg.Attachments
        lngCount = objAttachments.Count
    
        If lngCount > 0 Then
    
        ' Use a count down loop for removing items
        ' from a collection. Otherwise, the loop counter gets
        ' confused and only every other item is removed.
    
        For i = lngCount To 1 Step -1
        
            ' Delete the attachment.
            Select Case Right(objAttachments.Item(i).filename, 3)
            
            Case "pdf": objAttachments.Item(i).Delete
            Case "zip": objAttachments.Item(i).Delete
            Case "msg": objAttachments.Item(i).Delete
            Case "xls": objAttachments.Item(i).Delete
            Case "lsx": objAttachments.Item(i).Delete
            Case "lsm": objAttachments.Item(i).Delete
            Case "lsb": objAttachments.Item(i).Delete
            Case "doc": objAttachments.Item(i).Delete
            Case "ocx": objAttachments.Item(i).Delete
            Case "ppt": objAttachments.Item(i).Delete
            Case "ptx": objAttachments.Item(i).Delete
            
            Case Else
            
            End Select
        
            Counter = Counter + 1
    
        Next i
    
        End If
        
        objMsg.Save

    Next

ExitSub:

    ' libera memoria y objetos
    Set objAttachments = Nothing
    Set objMsg = Nothing
    Set objSelection = Nothing
    Set objOL = Nothing

End Sub

Sub FoldersFindArchive()

    Dim ns As Outlook.NameSpace
    Dim mainFolder As Outlook.folder
    Dim folder As Outlook.folder
    
    Set ns = Application.GetNamespace("MAPI")
    
    ' Loop through all folders (accounts) and find the Gmail account
    For Each folder In ns.Folders
        If InStr(1, folder.Name, strNamespace, vbTextCompare) > 0 Then
            Set mainFolder = folder
            Exit For
        End If
    Next folder
    
    If Not mainFolder Is Nothing Then
        ' Call the function to loop through all folders in the mailbox
        FoldersProcess mainFolder
    Else
        MsgBox "desired account not found."
    End If
    
End Sub

Sub FoldersProcess(CurrentFolder As Outlook.folder)
    Dim i As Integer
    Dim subFolder As Outlook.folder
    
    ' Check if folder name ends with "Archivo" and not "Archivos"
    If InStr(1, CurrentFolder.Name, "Archivo", vbTextCompare) > 0 And Not ExitLoop Then
        Set ArchivoFolder = CurrentFolder
        Debug.Print "match: " & CurrentFolder.FolderPath
        ExitLoop = True
    End If
    
    If ExitLoop Then Exit Sub 'if we've found the folder, exit
    
    ' Loop through all subfolders of the current folder
    For i = 1 To CurrentFolder.Folders.Count
        Set subFolder = CurrentFolder.Folders(i)
        ' Call the function recursively to get all levels of subfolders
        FoldersProcess subFolder
    Next i
    
End Sub


' Function to remove illegal characters from a string
Function ReplaceIllegalChars(ByVal strIn As String) As String

    ' List of illegal characters for Windows filenames: < > : " | ? * \ /
    ' In VBA, double quote must be escaped as "" inside a string literal
    Dim strChars As String: strChars = "<>:""|?*\/"
    Dim intIndex As Integer

    ' Loop through each character in the string and replace it if it is an illegal character
    For intIndex = 1 To Len(strChars)
        strIn = Replace(strIn, Mid(strChars, intIndex, 1), " ")
    Next

    ' Clean up multiple spaces and trim
    strIn = Trim(strIn)
    Do While InStr(strIn, "  ") > 0
        strIn = Replace(strIn, "  ", " ")
    Loop

    ' Return the cleaned string
    ReplaceIllegalChars = strIn

End Function

' Function to verify and correct the file path if it's too long
' source > https://chat.openai.com/c/4ac8607e-db30-4bb6-89ae-0ac87d8bb52f
Function VerifyAndShortenPath(ByVal path As String, ByVal filename As String) As String

    Const MaxPathLength As Integer = 260  ' Maximum file path length in Windows
    Dim TotalLength As Integer
    Dim AllowedLength As Integer
    
    TotalLength = Len(path) + Len(filename)
    
    If TotalLength > MaxPathLength Then
        ' Calculate how many characters we are allowed for filename
        AllowedLength = MaxPathLength - Len(path)
        
        ' Shorten the filename while preserving the extension
        filename = Left(filename, AllowedLength - 5) & Right(filename, 4)
    End If
    
    VerifyAndShortenPath = filename
    
End Function

' Function to archive and manage attachments in selected emails
Function ArchiveVB(ByVal putInput As Boolean, preference As Integer)
    
    ' Declare variables for the main Outlook application, the email messages, attachments, and selection of emails
    Dim objOL As Outlook.Application
    Dim objMsg As Outlook.MailItem
    Dim objAttachments As Outlook.Attachments
    Dim objSelection As Outlook.Selection
    Dim i As Long
    Dim lngCount As Long
    Dim Counter As Long
    Dim ruta As String
    Dim secuencial As String
    Dim nombreCorreo As String
    Dim correctedFilename As String
    
    ' Call the custom function GetURL() to get the file path for saving attachments
    ruta = GetURL()
    
    ' Check if putInput is False, then get a sequential number for the attachment file name
    ' Otherwise, prompt for a sequential number via an input box
    If putInput = False Then
        secuencial = GetSequentialNumber(ruta)
    Else
        secuencial = InputBox("incorporar secuencial", "secuencial")
        
        ' If the input box is cancelled or left blank, exit the function
        If secuencial = Null Or secuencial = "" Then Exit Function
    End If
    
    ' Create a new instance of the Outlook application and get the currently selected emails
    Set objOL = CreateObject("Outlook.Application")
    Set objSelection = objOL.ActiveExplorer.Selection
    
    ' Loop through each selected email
    For Each objMsg In objSelection
    
        ' Get the attachments of the email and the count of attachments
        Set objAttachments = objMsg.Attachments
        lngCount = objAttachments.Count
        Counter = 1
        
        ' If there are any attachments
        If lngCount > 0 Then
        
            ' Loop through each attachment
            For i = lngCount To 1 Step -1
            
                ' Save the attachment if the file type is one of the specified types
                Select Case LCase(Right(objAttachments.Item(i).filename, 3))
                    Case "pdf", "zip", "msg", "xls", "lsx", "lsm", "lsb", "doc", "ocx", "ppt", "ptx"
                        correctedFilename = VerifyAndShortenPath(ruta, secuencial & " - " & objAttachments.Item(i).filename)
                        objAttachments.Item(i).SaveAsFile (ruta & "\" & correctedFilename)
                End Select
                Counter = Counter + 1
            Next i
        End If
        
        ' Clean up the subject line of the email to create the name for the saved email file
        nombreCorreo = ReplaceIllegalChars(objMsg.Subject)
        nombreCorreo = secuencial & " - " & nombreCorreo & ".msg"
        
        ' Verify and correct the path if it's too long
        nombreCorreo = VerifyAndShortenPath(ruta, nombreCorreo)
        
        ' Save the email to the specified file path
        objMsg.SaveAs ruta & nombreCorreo, olMSG
        
        ' Perform an action based on the preference parameter
        Select Case preference
            Case 1
                ' Delete the email
                objMsg.Delete
            Case 2
                ' Call a custom function to delete the attachments, then save the email
                Call DeleteAttachments
                objMsg.Close olSave
            Case 3
                ' Discard the email without saving changes
                objMsg.Close olDiscard
            Case 4
                ' Move the email to the "Archivo" folder
                ' Call the custom function to find the "Archivo" folder if it hasn't been found yet
                If ArchivoFolder Is Nothing Then Call FoldersFindArchive
                objMsg.Move ArchivoFolder
        End Select
    Next objMsg
    
    ' Clean up and release the Outlook objects
    Set objAttachments = Nothing
    Set objMsg = Nothing
    Set objSelection = Nothing
    Set objOL = Nothing
    
End Function

' Function to archive only in the "Archivo"
Sub ArchiveVBonly()
    
    ' Declare variables for the main Outlook application, the email messages, attachments, and selection of emails
    Dim objOL As Outlook.Application
    Dim objMsg As Outlook.MailItem
    Dim objAttachments As Outlook.Attachments
    Dim objSelection As Outlook.Selection
        
    ' Create a new instance of the Outlook application and get the currently selected emails
    Set objOL = CreateObject("Outlook.Application")
    Set objSelection = objOL.ActiveExplorer.Selection
    
    ' Loop through each selected email
    For Each objMsg In objSelection
    
        If ArchivoFolder Is Nothing Then Call FoldersFindArchive
        objMsg.Move ArchivoFolder

    Next objMsg
    
    ' Clean up and release the Outlook objects
    Set objAttachments = Nothing
    Set objMsg = Nothing
    Set objSelection = Nothing
    Set objOL = Nothing
    
End Sub

Sub ArchiveVBinput()

' con input y archivando el correo tal cual
ArchiveVB True, 3

End Sub

Sub ArchiveVBdefault()

' con secuencial automático y archivando el correo tal cual
ArchiveVB False, 3

End Sub

Sub Link()

Dim objOL As Outlook.Application
Dim objNS As Outlook.NameSpace
Dim objDoc As Word.Document
Dim objSel As Word.Selection
Dim ruta As String
  
    On Error GoTo error:
       
    Set objOL = Application
    If objOL.ActiveInspector.EditorType = olEditorWord Then
    
        ' use WordEditor
        Set objDoc = objOL.ActiveInspector.WordEditor
        Set objNS = objOL.Session
        ' set current selection
        Set objSel = objDoc.Windows(1).Selection
            
        ' check selection
        Debug.Print objSel.Text
        
        ' recupera la ruta con la función anterior
        ruta = GetURL()

        ' añade el link
        objDoc.Hyperlinks.Add objSel.Range, ruta, , "link", ruta, ""

error:
        If Err.Number = 5824 Then Resume Next

    End If
    
    ' libera memoria y objetos
    Set objOL = Nothing
    Set objNS = Nothing
    Set objDoc = Nothing
    Set objSel = Nothing
        
End Sub

Public Sub CloseNoSave()

Dim objOL As Outlook.Application
Dim objMsg As Outlook.MailItem 'Object
Dim objWindow As Object

    ' Instantiate an Outlook Application object.
    Set objOL = CreateObject("Outlook.Application")

    ' Get the collection of selected objects.
    Set objWindow = objOL.ActiveWindow

    ' close
    objWindow.Close olDiscard

ExitSub:

    ' libera memoria y objetos
    Set objMsg = Nothing
    Set objWindow = Nothing
    Set objOL = Nothing

End Sub

Public Sub CloseSave()

Dim objOL As Outlook.Application
Dim objMsg As Outlook.MailItem 'Object
Dim objWindow As Object

    ' Instantiate an Outlook Application object.
    Set objOL = CreateObject("Outlook.Application")

    ' Get the collection of selected objects.
    Set objWindow = objOL.ActiveWindow

    ' close
    objWindow.Close olSave

ExitSub:

    ' libera memoria y objetos
    Set objMsg = Nothing
    Set objWindow = Nothing
    Set objOL = Nothing

End Sub

Sub ColorTextGreen()

Dim objOL As Outlook.Application
Dim objNS As Outlook.NameSpace
Dim objDoc As Word.Document
Dim objSel As Word.Selection
  
    On Error GoTo error:
       
    Set objOL = Application
    If objOL.ActiveInspector.EditorType = olEditorWord Then
    
        ' use WordEditor
        Set objDoc = objOL.ActiveInspector.WordEditor
        Set objNS = objOL.Session
        ' set current selection
        Set objSel = objDoc.Windows(1).Selection
            
        ' check selection
        ' Debug.Print objSel.Text

        ' cambia el color
        With objSel.Font
            .Name = "Calibri"
            .Size = 11
            .Bold = True
            .Italic = False
            .Underline = wdUnderlineNone
            .UnderlineColor = wdColorAutomatic
            .Strikethrough = False
            .DoubleStrikeThrough = False
            .Outline = False
            .Emboss = False
            .Shadow = False
            .Hidden = False
            .Smallcaps = False
            .Allcaps = False
            .Color = 5287936
            .Engrave = False
            .Superscript = False
            .Subscript = False
            .Spacing = 0
            .Scaling = 100
            .Position = 0
            .Kerning = 0
            .Animation = wdAnimationNone
        End With

error:
        If Err.Number = 5824 Then Resume Next

    End If
    
    ' libera memoria y objetos
    Set objOL = Nothing
    Set objNS = Nothing
    Set objDoc = Nothing
    Set objSel = Nothing
    
End Sub

Sub ColorTextBlue()

Dim objOL As Outlook.Application
Dim objNS As Outlook.NameSpace
Dim objDoc As Word.Document
Dim objSel As Word.Selection
  
    On Error GoTo error:
       
    Set objOL = Application
    If objOL.ActiveInspector.EditorType = olEditorWord Then
    
        ' use WordEditor
        Set objDoc = objOL.ActiveInspector.WordEditor
        Set objNS = objOL.Session
        ' set current selection
        Set objSel = objDoc.Windows(1).Selection
            
        ' check selection
        ' Debug.Print objSel.Text

        ' cambia el color
        With objSel.Font
            .Name = "Calibri"
            .Size = 11
            .Bold = True
            .Italic = False
            .Underline = wdUnderlineNone
            .UnderlineColor = wdColorAutomatic
            .Strikethrough = False
            .DoubleStrikeThrough = False
            .Outline = False
            .Emboss = False
            .Shadow = False
            .Hidden = False
            .Smallcaps = False
            .Allcaps = False
            .Color = 12611584
            .Engrave = False
            .Superscript = False
            .Subscript = False
            .Spacing = 0
            .Scaling = 100
            .Position = 0
            .Kerning = 0
            .Animation = wdAnimationNone
        End With

error:
        If Err.Number = 5824 Then Resume Next

    End If
    
    ' libera memoria y objetos
    Set objOL = Nothing
    Set objNS = Nothing
    Set objDoc = Nothing
    Set objSel = Nothing
    
End Sub

Sub ColorTextBlack()
    
Dim objOL As Outlook.Application
Dim objNS As Outlook.NameSpace
Dim objDoc As Word.Document
Dim objSel As Word.Selection
  
    On Error GoTo error:
       
    Set objOL = Application
    If objOL.ActiveInspector.EditorType = olEditorWord Then
    
        ' use WordEditor
        Set objDoc = objOL.ActiveInspector.WordEditor
        Set objNS = objOL.Session
        ' set current selection
        Set objSel = objDoc.Windows(1).Selection
            
        ' check selection
        ' Debug.Print objSel.Text

        ' cambia el color
        With objSel.Font
            .Name = "Calibri"
            .Size = 11
            .Bold = False
            .Italic = False
            .Underline = wdUnderlineNone
            .UnderlineColor = wdColorAutomatic
            .Strikethrough = False
            .DoubleStrikeThrough = False
            .Outline = False
            .Emboss = False
            .Shadow = False
            .Hidden = False
            .Smallcaps = False
            .Allcaps = False
            .Color = -553582593
            .Engrave = False
            .Superscript = False
            .Subscript = False
            .Spacing = 0
            .Scaling = 100
            .Position = 0
            .Kerning = 0
            .Animation = wdAnimationNone
        End With

error:
        If Err.Number = 5824 Then Resume Next

    End If
    
    ' libera memoria y objetos
    Set objOL = Nothing
    Set objNS = Nothing
    Set objDoc = Nothing
    Set objSel = Nothing
        
End Sub

Sub LangCat()
    
Dim objOL As Outlook.Application
Dim objNS As Outlook.NameSpace
Dim objDoc As Word.Document
Dim objSel As Word.Selection
   
Set objOL = Application
If objOL.ActiveInspector.EditorType = olEditorWord Then

    ' use WordEditor
    Set objDoc = objOL.ActiveInspector.WordEditor
    Set objNS = objOL.Session
    ' set current selection
    Set objSel = objDoc.Windows(1).Selection

' idioma catalán
objSel.LanguageID = wdCatalan
objSel.NoProofing = False

End If

' libera memoria y objetos
Set objOL = Nothing
Set objNS = Nothing
Set objDoc = Nothing
Set objSel = Nothing

End Sub

Sub LangCas()
    
Dim objOL As Outlook.Application
Dim objNS As Outlook.NameSpace
Dim objDoc As Word.Document
Dim objSel As Word.Selection
   
Set objOL = Application
If objOL.ActiveInspector.EditorType = olEditorWord Then

    ' use WordEditor
    Set objDoc = objOL.ActiveInspector.WordEditor
    Set objNS = objOL.Session
    ' set current selection
    Set objSel = objDoc.Windows(1).Selection

' idioma castellano
objSel.LanguageID = wdSpanishModernSort
objSel.NoProofing = False

End If

' libera memoria y objetos
Set objOL = Nothing
Set objNS = Nothing
Set objDoc = Nothing
Set objSel = Nothing
    
End Sub

Sub Send8h_default()

Dim obj As Object
Dim Mail As Outlook.MailItem
Dim WkDay As String
Dim MinNow As Integer
Dim SecNow As Integer
Dim SendHour As Integer
Dim SendDate As Date
Dim SendNow As Boolean
    
'Set Variables
SendDate = Now()
SendHour = Hour(Now)
MinNow = Minute(Now)
SecNow = Second(Now)
WkDay = Weekday(Now)
SendNow = True

' sábado y domingo, se envía lunes a las 8h

Select Case WkDay

    Case 1 ' domingo, un día más
        SendDate = DateAdd("d", 1, SendDate)
        SendHour = 8 - SendHour
        SendDate = DateAdd("h", SendHour, SendDate)
        SendDate = DateAdd("n", -MinNow, SendDate)
        SendDate = DateAdd("s", -SecNow, SendDate)
    
    Case 7 ' sábado, dos días más
        SendDate = DateAdd("d", 2, SendDate)
        SendHour = 8 - SendHour
        SendDate = DateAdd("h", SendHour, SendDate)
        SendDate = DateAdd("n", -MinNow, SendDate)
        SendDate = DateAdd("s", -SecNow, SendDate)
                
    Case Else
    
    Select Case SendHour
    
        Case Is < 8: SendHour = 8 - SendHour
        Case Is < 13: SendHour = 13 - SendHour
        Case Is < 16: SendHour = 16 - SendHour
        Case Is < 19: SendHour = 19 - SendHour
        Case Is > 19: SendHour = 32 - SendHour    'Send a 8 am next day
    
    End Select
    
    SendDate = DateAdd("h", SendHour, SendDate)
    SendDate = DateAdd("n", -MinNow, SendDate)
    SendDate = DateAdd("s", -SecNow, SendDate)

End Select

SendNow = False

'Send the Email
Set obj = Application.ActiveInspector.CurrentItem
If TypeOf obj Is Outlook.MailItem Then
  Set Mail = obj
  'Check if we need to delay delivery
  If SendNow = False Then
    
    Mail.DeferredDeliveryTime = SendDate
  End If
  Mail.Send
End If
  
End Sub

Sub Send8hSel()

Dim objOL As Outlook.Application
Dim objMsg As Outlook.MailItem 'Object
Dim objSelection As Outlook.Selection

Dim WkDay As String
Dim MinNow As Integer
Dim SecNow As Integer
Dim SendHour As Integer
Dim SendDate As Date
Dim SendNow As Boolean

'Set Variables
SendDate = Now()
SendHour = Hour(Now)
MinNow = Minute(Now)
SecNow = Second(Now)
WkDay = Weekday(Now)
SendNow = True

' sábado y domingo, se envía lunes a las 8h

Select Case WkDay

    Case 1 ' domingo, un día más
        SendDate = DateAdd("d", 1, SendDate)
        SendHour = 8 - SendHour
        SendDate = DateAdd("h", SendHour, SendDate)
        SendDate = DateAdd("n", -MinNow, SendDate)
        SendDate = DateAdd("s", -SecNow, SendDate)
    
    Case 7 ' sábado, dos días más
        SendDate = DateAdd("d", 2, SendDate)
        SendHour = 8 - SendHour
        SendDate = DateAdd("h", SendHour, SendDate)
        SendDate = DateAdd("n", -MinNow, SendDate)
        SendDate = DateAdd("s", -SecNow, SendDate)
                
    Case Else
    
    Select Case SendHour
    
        Case Is < 8: SendHour = 8 - SendHour
        Case Is < 13: SendHour = 13 - SendHour
        Case Is < 16: SendHour = 16 - SendHour
        Case Is < 19: SendHour = 19 - SendHour
        Case Is > 19: SendHour = 32 - SendHour    'Send a 8 am next day
    
    End Select
    
    SendDate = DateAdd("h", SendHour, SendDate)
    SendDate = DateAdd("n", -MinNow, SendDate)
    SendDate = DateAdd("s", -SecNow, SendDate)

End Select

SendNow = False

    ' Instantiate an Outlook Application object.
    Set objOL = CreateObject("Outlook.Application")

    ' Get the collection of selected objects.
    Set objSelection = objOL.ActiveExplorer.Selection

    
    'Check if we need to delay delivery
    If SendNow = False Then
      
    For Each objMsg In objSelection
    
        objMsg.DeferredDeliveryTime = SendDate
        objMsg.Send
    
    Next
    
    End If

ExitSub:

    ' libera memoria y objetos
    Set objMsg = Nothing
    Set objSelection = Nothing
    Set objOL = Nothing

End Sub

Function RunPythonScript(PythonExePath As String, ScriptPath As String)

    ' Declare variables
    Dim objShell As Object
    Dim strCmd As String

    ' Create the command string
    strCmd = """" & PythonExePath & """ """ & ScriptPath & """"

    ' Create a shell object
    Set objShell = CreateObject("WScript.Shell")

    ' Run the Python script
    objShell.Run strCmd, 1, True

    ' Clean up
    Set objShell = Nothing

End Function

Sub ArchivePy()
RunPythonScript pythonExe, pythonArchive
End Sub

Sub ClassifyPy()
RunPythonScript pythonExe, pythonClassify
End Sub

Sub SortAddressesDefault()

' credits to
' https://techniclee.wordpress.com/2010/08/26/sorting-addressees/

' variable definition

Dim colNames As New Collection, _
    intCounter As Integer, _
    intIndex As Integer, _
    olkItem As Object, _
    olkRecipient As Outlook.Recipient, _
    olkAddressee As Outlook.Recipient, _
    varName As Variant
        
' create an instance for the current item. has to be an outlook item (may be an email or an appointment)

Set olkItem = Application.ActiveInspector.CurrentItem

' takes out all the recipients and puts them into an array > ColNames (collection)
' then removes them one by one from the email

For intCounter = olkItem.Recipients.Count To 1 Step -1
    Set olkRecipient = olkItem.Recipients.Item(intCounter)
    If colNames.Count > 0 Then
        intIndex = 1
        For Each varName In colNames
            If intIndex = colNames.Count Then
                If LCase(olkRecipient.Name) > LCase(varName) Then
                    colNames.Add olkRecipient, LCase(olkRecipient.Name), , intIndex
                    Exit For
                Else
                    colNames.Add olkRecipient, LCase(olkRecipient.Name), intIndex
                    Exit For
                End If
            End If
            If LCase(olkRecipient.Name) < LCase(varName) Then
                colNames.Add olkRecipient, LCase(olkRecipient.Name), intIndex
                Exit For
            End If
            intIndex = intIndex + 1
        Next
    Else
        colNames.Add olkRecipient, LCase(olkRecipient.Name)
    End If
    olkItem.Recipients.Remove intCounter
Next

' resolves the name and put the recipient back into the email
' TO, CC and BCC are respected

For Each olkAddressee In colNames
    Set olkRecipient = Session.CreateRecipient(olkAddressee.Name)
    olkRecipient.Resolve
    If olkRecipient.Resolved Then
        Set olkRecipient = olkItem.Recipients.Add(olkAddressee.Name)
    Else
        Set olkRecipient = olkItem.Recipients.Add(olkAddressee.Address)
    End If
    olkRecipient.Type = olkAddressee.Type
Next

' last check before sending, resolves address. no duplicates are allowed

olkItem.Recipients.ResolveAll

' clears temporary data

Set olkRecipient = Nothing
Set colNames = Nothing

End Sub

Sub SortAddressesBCC()

' credits to
' https://techniclee.wordpress.com/2010/08/26/sorting-addressees/

' variable definition

Dim colNames As New Collection, _
    intCounter As Integer, _
    intIndex As Integer, _
    olkItem As Outlook.MailItem, _
    olkRecipient As Outlook.Recipient, _
    olkAddressee As Outlook.Recipient, _
    varName As Variant, _
    strBCC As String
   
' set the BCC (custom)
strBCC = "rferraro@caixabank.com"

' create an instance for the current item. has to be an outlook item (may be an email or an appointment)

Set olkItem = Application.ActiveInspector.CurrentItem

' takes out all the recipients and puts them into an array > ColNames (collection)
' then removes them one by one from the email

For intCounter = olkItem.Recipients.Count To 1 Step -1
    Set olkRecipient = olkItem.Recipients.Item(intCounter)
    If colNames.Count > 0 Then
        intIndex = 1
        For Each varName In colNames
            If intIndex = colNames.Count Then
                If LCase(olkRecipient.Name) > LCase(varName) Then
                    colNames.Add olkRecipient, LCase(olkRecipient.Name), , intIndex
                    Exit For
                Else
                    colNames.Add olkRecipient, LCase(olkRecipient.Name), intIndex
                    Exit For
                End If
            End If
            If LCase(olkRecipient.Name) < LCase(varName) Then
                colNames.Add olkRecipient, LCase(olkRecipient.Name), intIndex
                Exit For
            End If
            intIndex = intIndex + 1
        Next
    Else
        colNames.Add olkRecipient, LCase(olkRecipient.Name)
    End If
    olkItem.Recipients.Remove intCounter
Next

' resolves the name and put the recipient back into the email
' TO, CC and BCC are respected

For Each olkAddressee In colNames
    Set olkRecipient = Session.CreateRecipient(olkAddressee.Name)
    olkRecipient.Resolve
    If olkRecipient.Resolved Then
        Set olkRecipient = olkItem.Recipients.Add(olkAddressee.Name)
    Else
        Set olkRecipient = olkItem.Recipients.Add(olkAddressee.Address)
    End If
    olkRecipient.Type = olkAddressee.Type
Next

' set BCC (custom)

olkItem.BCC = strBCC

' last check before sending, resolves address. no duplicates are allowed

olkItem.Recipients.ResolveAll

' clears temporary data

Set olkRecipient = Nothing
Set colNames = Nothing

End Sub

Sub ResizeImage_execute(ByVal imgHeight As Single, ByVal imgWidth As Single)

    ' Get the current open Inspector object in Outlook
    Dim Inspector As Outlook.Inspector
    Set Inspector = Application.ActiveInspector
    
    ' Get the Word Editor object for the current email
    Dim WordDoc As Word.Document
    Set WordDoc = Inspector.WordEditor
    
    ' Work with the Word Editor selection
    With WordDoc.Application.Selection
    
        If .InlineShapes.Count > 0 Then
        
            With .InlineShapes(1)
                ' Lock aspect ratio
                .LockAspectRatio = msoTrue
                ' Set height
                .Height = CentimetersToPoints(imgHeight)
                ' Set width
                .Width = CentimetersToPoints(imgWidth)
            End With
            Exit Sub
            
        Else
        
            With .ShapeRange
                ' Lock aspect ratio
                .LockAspectRatio = msoTrue
                ' Set height
                .Height = CentimetersToPoints(imgHeight)
                ' Set width
                .Width = CentimetersToPoints(imgWidth)
            End With
            
        End If
        
    End With
End Sub

Sub ResizeImage(ByVal imgHeight As Single, ByVal imgWidth As Single)

    ' Get the current open Inspector object in Outlook
    Dim Inspector As Outlook.Inspector
    Set Inspector = Application.ActiveInspector
    
    ' Get the Word Editor object for the current email
    Dim WordDoc As Word.Document
    Set WordDoc = Inspector.WordEditor
    
    ' Work with the Word Editor selection
    With WordDoc.Application.Selection
    
        ' Check if the current selection is an image object
        If .Type = wdSelectionInlineShape Or .Type = wdSelectionShape Then
        
            Call ResizeImage_execute(imgHeight, imgWidth)
            Exit Sub
        
        Else
        
            ' Paste whatever is in the clipboard
            .Paste
            
            ' Select the last pasted object
            .MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
            
            ' Check if the pasted object is an image
            If .Type = wdSelectionInlineShape Or .Type = wdSelectionShape Then
            
                ' Apply the code and resize the image
                Call ResizeImage_execute(imgHeight, imgWidth)
                Exit Sub
            
            End If
        
        End If
        
    End With

    ' If the pasted object is not an image, end the process

End Sub

Sub ResizeImage_18()
Call ResizeImage(18, 18)
End Sub

Sub ResizeImage_09()
Call ResizeImage(9, 9)
End Sub

Sub PasteAsValues()

 ' Get the current open Inspector object in Outlook
    Dim Inspector As Outlook.Inspector
    Set Inspector = Application.ActiveInspector
    
    ' Get the Word Editor object for the current email
    Dim WordDoc As Word.Document
    Set WordDoc = Inspector.WordEditor
    
    ' Work with the Word Editor selection
    With WordDoc.Application.Selection
  
        On Error Resume Next
        .PasteAndFormat (wdFormatPlainText)
       
    End With
    
End Sub
