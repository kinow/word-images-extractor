' Extract images from word documents.
'
' Author:  Bruno de Paula Kinoshiyta (http://www.kinoshita.eti.br)
' License: GNU/LGPL (http://www.gnu.org/licenses/lgpl.html)
' Based on: 
'
'  - http://www.source-code.biz/snippets/vbscript/1.htm
'	 VBScript code structure. Very clean and structured.
'	 by Christian d'Heureuse
'
'  - http://www.robvanderwoude.com/vbstech_automation_word.php#SaveAsHTML
'    Sub that saves a word document as html.
'	 by Rob van der Woude
'
' Changes:
' 18.feb.2009 Had the idea after having to use some old graphics from a .doc
' 26.feb.2009 Added file dialog
' 01.mar.2009 Script structure modified after seeing d'Heureuse code
'

Option Explicit

Main

Sub Main
	
	Dim dialog: Set dialog = CreateObject("UserAccounts.CommonDialog")
	dialog.Filter = "Word Documents|*.doc"
	dialog.FilterIndex = 2
	Dim status: status = dialog.ShowOpen
	If status = 0 Then
		WScript.Echo "Script Error: Please select a file!"
		WScript.Quit
	End If 

	Doc2HTML ( dialog.FileName )

End Sub

Sub Doc2HTML( input )

	Const wdFormatDocument                    =  0
    Const wdFormatDocument97                  =  0
    Const wdFormatDocumentDefault             = 16
    Const wdFormatDOSText                     =  4
    Const wdFormatDOSTextLineBreaks           =  5
    Const wdFormatEncodedText                 =  7
    Const wdFormatFilteredHTML                = 10
    Const wdFormatFlatXML                     = 19
    Const wdFormatFlatXMLMacroEnabled         = 20
    Const wdFormatFlatXMLTemplate             = 21
    Const wdFormatFlatXMLTemplateMacroEnabled = 22
    Const wdFormatHTML                        =  8
    Const wdFormatPDF                         = 17
    Const wdFormatRTF                         =  6
    Const wdFormatTemplate                    =  1
    Const wdFormatTemplate97                  =  1
    Const wdFormatText                        =  2
    Const wdFormatTextLineBreaks              =  3
    Const wdFormatUnicodeText                 =  7
    Const wdFormatWebArchive                  =  9
    Const wdFormatXML                         = 11
    Const wdFormatXMLDocument                 = 12
    Const wdFormatXMLDocumentMacroEnabled     = 13
    Const wdFormatXMLTemplate                 = 14
    Const wdFormatXMLTemplateMacroEnabled     = 15
    Const wdFormatXPS                         = 18

	Dim fso: Set fso = CreateObject( "Scripting.FileSystemObject" )
	Dim word: Set word = CreateObject( "Word.Application" )

    With word
        .Visible = False
        Dim inputPath
		Dim inputFile
        If fso.FileExists( input ) Then
            Set inputFile = fso.GetFile( input )
            inputPath = inputFile.Path
        Else
            WScript.Echo "FILE OPEN ERROR: The file does not exist" & vbCrLf
            .Quit
            Exit Sub
        End If

		Dim parent: Set parent = inputFile.ParentFolder
		Dim htmlFile: htmlFile = fso.GetBaseName( inputFile ) & ".html"
		Dim html: html = fso.BuildPath( parent, htmlFile )
		
        ' Open the Word document
        .Documents.Open inputPath

        ' Make the opened file the active document
		Dim doc:  Set doc = .ActiveDocument

        ' Save as HTML
        doc.SaveAs html, wdFormatFilteredHTML

        ' Close the active document
        doc.Close
		
		Dim imagesFolder: imagesFolder = fso.GetBaseName( inputFile ) & "_arquivos"
		Dim images: images = fso.BuildPath ( parent, imagesFolder )

		Dim currentFolder

		If fso.FolderExists( images ) Then

			currentFolder = inputFile.ParentFolder
		
			Dim Folder: Set Folder = fso.GetFolder( images )
			Dim Files: Set Files = Folder.Files
			Dim File
			For Each File in Files
				fso.GetFile ( File ).Copy currentFolder & "\" & fso.GetFileName( File ), True
			Next

			fso.DeleteFolder images

		End If	
		
		fso.DeleteFile html
		
        .Quit

    End With
	
End Sub 