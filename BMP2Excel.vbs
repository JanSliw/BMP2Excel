
	originalFilePath = WScript.Arguments(0)

	Set bt = New BMPTranslator
	
	bt.translate(originalFilePath)
	
	Set bt = Nothing
	
	MsgBox "Your file is ready"

Class BITMAPFILEHEADER
	Public bfType
	Public bfSize
	Public bfReserved1
	Public bfReserved2
	Public bfOffBits
	
	Public Function toString()
		toString = "BITMAPFILEHEADER " & vbNewLine & _
					"bfType: " & bfType & vbNewLine & _
					"bfSize: " & bfSize & vbNewLine & _
					"bfReserved1: " & bfReserved1 & vbNewLine & _
					"bfReserved2: " & bfReserved2 & vbNewLine & _
					"bfOffBits: " & bfOffBits  
	End Function
	
End Class

Class BITMAPINFOHEADER
	Public biSize
	Public biWidth
	Public biHeight
	Public biPlanes
	Public biBitCount
	Public biCompresion
	Public biSizeImage
	Public biXPelsPerMeter
	Public biYPelsPerMeter
	Public biClrUsed
	Public biClrImportant
	Public biRestOfHeaderData
	
	Public Function toString()
		toString = "BITMAPINFOHEADER " & vbNewLine & _
					"biSize: " & vbNewLine & _
					"biWidth: " & vbNewLine & _
					"biHeight: " & vbNewLine & _
					"biPlanes: " & vbNewLine & _
					"biBitCount: " & vbNewLine & _
					"biCompresion: " & vbNewLine & _
					"biSizeImage: " & vbNewLine & _
					"biXPelsPerMeter: " & vbNewLine & _
					"biYPelsPerMeter: " & vbNewLine & _
					"biClrUsed: " & vbNewLine & _
					"biClrImportant: " & vbNewLine & _
					"biRestOfHeaderData: " & biRestOfHeaderData
	End Function
	
End Class

Class BMPTranslator

	Private fso
	Private bfr
	Private wc
	Private originalFilePath
	Private outputFilePath
	Private tempFolderPath
	
	Private bfh
	Private bih
	Private outputTempFile
	Private cellsTempFile
	Private emptyExcelFile
	
	Private Sub Class_Initialize()
		Set fso = CreateObject("Scripting.FileSystemObject") 
		Set bfr = New BinaryFileReader
		Set wc = New WorkbookController
		Set bfh = New BITMAPFILEHEADER
		Set bih = New BITMAPINFOHEADER
	End Sub
	
	Private Sub Class_Terminate()
		Set bih = Nothing
		Set bfh = Nothing
		Set wc = Nothing
		Set bfr = Nothing
		Set fso = Nothing
	End Sub
	
	Public Sub translate(path)
		setFilePath(path)
		If fileIsBmp() Then readMetaData()
		createTempFolderAndFiles()
		translateImage()
		optimizeOutputFile()
		cleanUp()
		openOutputFile()
	End Sub
	
	Private Sub setFilePath(path)
		originalFilePath = path
		If pathExists() Then
			bfr.loadFile(originalFilePath)
		Else
			WScript.Echo "Given path does not exists:\n" & originalFilePath
		End If
	End Sub
	
	Private Function pathExists()
		pathExists = Len(originalFilePath) <> 0 And fso.FileExists(originalFilePath)
	End Function
	
	Private Sub readMetaData()
		readBitMapFileHeader()
		readBitMapInfoHeader()
	End Sub
	
	Private Sub readBitMapFileHeader()
		bfr.setLittleEndian
	
		bfh.bfType = bfr.readBytesAsHex(0, 2)
		
		If bfh.bfType <> "4D42" Then
			MsgBox bfh.toString()
			Err.Raise 10004, , "Loaded file is not a bitmap"
		End If
		
		bfh.bfSize = bfr.readBytesAsDec(2, 4)
		bfh.bfReserved1 = bfr.readBytesAsDec(6, 2)
		bfh.bfReserved2 = bfr.readBytesAsDec(8, 2)
		bfh.bfOffBits = bfr.readBytesAsDec(10, 4)		
	
		If bfr.getSize < bfh.bfOffBits Then
			MsgBox bfh.toString()
			Err.Raise 10005, , "Bitmap File Header stanity check failed"
		End If
		
	End Sub

	Private Sub readBitMapInfoHeader()
		bfr.setLittleEndian()
	
		bih.biSize = bfr.readBytesAsDec(14, 4)
		
		If bih.biSize < 40 Then
			MsgBox bfh.toString() & vbNewLine & vbNewLine & _
				bih.toString()
			Err.Raise 10004, , "Loaded bitmap file type is not supported"
		End If
		
		bih.biWidth = bfr.readBytesAsDec(18, 4)
		bih.biHeight = bfr.readBytesAsDec(22, 4)
		bih.biPlanes = bfr.readBytesAsDec(26, 2)
		bih.biBitCount = bfr.readBytesAsDec(28, 2)
		
		If bih.biBitCount <> 24 Then
			MsgBox bfh.toString() & vbNewLine & vbNewLine & _
				bih.toString()
			Err.Raise 10004, , "Loaded bitmap file type is not supported"
		End If		
		
		bih.biCompresion = bfr.readBytesAsDec(30, 4)
		
		If bih.biCompresion <> 0 Then
			MsgBox bfh.toString() & vbNewLine & vbNewLine & _
				bih.toString()
			Err.Raise 10004, , "Loaded bitmap file type is not supported"
		End If		
		
		bih.biSizeImage = bfr.readBytesAsDec(34, 4)
		bih.biXPelsPerMeter = bfr.readBytesAsDec(38, 4)
		bih.biYPelsPerMeter = bfr.readBytesAsDec(42, 4)
		bih.biClrUsed = bfr.readBytesAsDec(46, 4)
		bih.biClrImportant = bfr.readBytesAsDec(50, 4)
		
		If bih.biSize = 40 Then
			bih.biRestOfHeaderData = ""
		Else	
			bih.biRestOfHeaderData = bfr.readBytesAsHex(54, bih.biSize - 40)				
		End If
	
	End Sub

	Private Function fileIsBmp()
		bfr.setBigEndian()
		fileIsBmp = bfr.readBytesAsString(0, 2) = "BM"
	End Function
	
	Private Sub createTempFolderAndFiles()
		createTempFolder()
		createTempFiles()
	End Sub
	
	Private Sub createTempFolder()
		tempFolderPath = fso.GetFile(originalFilePath).ParentFolder.Path & "\BMPTranslatorTemp"
		If fso.FolderExists(tempFolderPath) Then
			fso.GetFolder(tempFolderPath).Delete
		End If
		fso.CreateFolder(tempFolderPath)
	End Sub
	
	Private Sub createTempFiles()
		Set outputTempFile = New FileReaderWriter
		Set cellsTempFile = New FileReaderWriter
		cellsTempFile.openForWriting(tempFolderPath & "\cells.xml")
		outputTempFile.openForWriting(tempFolderPath & "\temp.xml")
		wc.PrepareNewXmlWorkbook(tempFolderPath & "\empty.xml")
		wc.close()
		Set emptyExcelFile = New FileReaderWriter
		emptyExcelFile.openForReading(wc.getFullPath)
	End Sub
	
	Private Sub translateImage()

		createXmlWorkbookHeader()
		suffixLength = (4 -((bih.biWidth * 3) Mod 4)) Mod 4
		position = bfh.bfOffBits
		bfr.setLittleEndian()
		appendHeaderToCellsTempFile()
		
		For h = bih.biHeight To 1 Step -1
		
			cellsTempFile.writeLine("   <Row ss:AutoFitHeight=""0"" ss:Height=""7.5"">")

			For w = 1 To bih.biWidth
				rgbHex = bfr.readBytesAsHex(position, 3)
				position = position + 3
				styleId = h*bih.biWidth + w + 100
				appendStyleElementToOutputTempFile styleId, rgbHex
				styleId = (bih.biHeight - h + 1)*bih.biWidth + w + 100
				appendCellElementToCellsTempFile styleId
			Next
			
			cellsTempFile.writeLine("   </Row>")
			position = position + suffixLength
			
		Next
		
		cellsTempFile.writeLine("  </Table>")
		cellsTempFile.closeFile()
		finishAppendngToOutputTempFile()
	
	End Sub
	
	Private Sub createXmlWorkbookHeader()
		outputTempFile.writeLine(emptyExcelFile.readLinesUntil("</Style>", True))
	End Sub

	Private Sub appendHeaderToCellsTempFile()
		Dim header
		header = "  <Table ss:ExpandedColumnCount=""" & bih.biWidth & """" & _
						" ss:ExpandedRowCount=""" & bih.biHeight & """" & _
						" x:FullColumns=""1"" x:FullRows=""1"" ss:DefaultRowHeight=""15"">" & _
				 "   <Column ss:AutoFitWidth=""0"" ss:Width=""7.5"" ss:Span=""" & bih.biWidth - 1 & """/>"
		cellsTempFile.writeLine(header)
	End Sub

	Private Sub appendStyleElementToOutputTempFile(ByVal styleId, ByVal rgbHex)
		Dim style
		style = "  <Style ss:ID=""s" & styleId & """>" & vbNewLine & _
				"   <Interior ss:Color=""#" & rgbHex & """ ss:Pattern=""Solid""/>" & vbNewLine & _
				"  </Style>"
		outputTempFile.writeLine(style)
	End Sub 

	Private Sub appendCellElementToCellsTempFile(styleId)
		Dim cell
		cell = "    <Cell ss:StyleID=""s" & styleId & """/>"
		cellsTempFile.writeLine(cell)
	End Sub
		
	Private Sub finishAppendngToOutputTempFile()
		outputTempFile.writeLine(emptyExcelFile.readLinesUntil("<Table ", False))
		cellsTempFile.openForReading(cellsTempFile.getFullPath())
		outputTempFile.writeLine(cellsTempFile.ReadAll())
		dumpFunctionOutput = emptyExcelFile.readLinesUntil("</Table>", False)
		outputTempFile.writeLine(emptyExcelFile.readLinesUntil("</PageSetup>", True))
		outputTempFile.writeLine("   <Zoom>10</Zoom>")		
		outputTempFile.writeLine(emptyExcelFile.readLinesAfter("<Selected/>", True))
	End Sub
	
	Private Sub optimizeOutputFile()
		outputFilePath = fso.GetFile(outputTempFile.getFullPath).ParentFolder.ParentFolder.Path & "\picture.xlsx"
		wc.Open(outputTempFile.getFullPath)
		If fso.FileExists(outputFilePath) Then fso.GetFile(outputFilePath).Delete()
		wc.SaveAs(outputFilePath)
	End Sub
	
	Private Sub cleanUp()
		emptyExcelFile.closeFile()
		cellsTempFile.closeFile()
		outputTempFile.closeFile()
		Set emptyExcelFile = Nothing
		Set cellsTempFile = Nothing
		Set outputTempFile = Nothing
		fso.GetFolder(tempFolderPath).Delete()
	End Sub
	
	Public Sub openOutputFile()
		wc.Open(outputFilePath)
		wc.showAppWindow()
	End Sub 
	
End Class

Class BinaryFileReader 

	Private adodbStream
	Private pByteOrder
	Private pBytes
	Private pStart
	Private pFinish
	Private pDirection
	Private pSize
	
	Public Sub setBigEndian()
		pByteOrder = 1
		pStart = 1
		pDirection = 1
	End Sub
	
	Public Sub setLittleEndian()
		pByteOrder = 2		
		pFinish = 1
		pDirection = -1
	End Sub 
	
	Public Property Get getSize()
		getSize = pSize
	End Property
	
	Private Sub Class_Initialize()
		Set adodbStream = CreateObject("ADODB.Stream")
		adodbStream.Type = 1
		adodbStream.Open
	End Sub
	
	Private Sub Class_Terminate()
		Set adodbStream = Nothing
	End Sub
	
	Public Function loadFile(path)
		adodbStream.LoadFromFile path
		pSize = adodbStream.Size
	End Function
	
	Public Function readBytesAsString(position, numberOfBytes)
		readBytes position, numberOfBytes
		readBytesAsString = changeBytesToString()
	End Function
	
	Public Function readBytesAsHex(position, numberOfBytes)
		readBytes position, numberOfBytes
		readBytesAsHex = changeBytesToHex()
	End Function
	
	Public Function readBytesAsDec(position, numberOfBytes)
		readBytes position, numberOfBytes
		readBytesAsDec = changeBytesToDec()
	End Function
	
	Private Function changeBytesToString()		
		For i = pStart To pFinish Step pDirection
			ascii = AscB(MidB(pBytes, i, 1))
			val = Chr(ascii)
			changeBytesToString = changeBytesToString & val
		Next
	
	End Function
	
	Private Function changeBytesToHex()
	
		For i = pStart To pFinish Step pDirection
			ascii = AscB(MidB(pBytes, i, 1))
			val = Right("00" & Hex(ascii), 2)
			changeBytesToHex = changeBytesToHex & val
		Next

	End Function
	
	Private Function changeBytesToDec()

		bytesInHex = changeBytesToHex		
		chunksCount = Int(Len(bytesInHex)/6)
		If Len(bytesInHex)/6 > chunksCount Then 
			remainder = Len(bytesInHex) - 6*chunksCount
			chunksCount = chunksCount + 1
		Else
			remainder = 6
		End If
		
		For i = 1 To chunksCount
			
			If i = chunksCount Then
				chunk = CDbl("&H" & Left(Right(bytesInHex, i * 6), remainder))
				power = CDbl(CDbl(2)^CDbl(4*6*(i - 1)+reminder))
			Else
				chunk = CDbl("&H" & Left(Right(bytesInHex, i * 6), 6))
				power = CDbl(CDbl(2)^CDbl(4*6*(i - 1)))
			End If
			
			val = CDbl(val) + CDbl(chunk*power)
				
		Next
		
		changeBytesToDec = val
		
	End Function
	
	Private Sub readBytes(position, numberOfBytes)
		' Begining of file starts at position = 0
		If position > pSize Then Err.Raise 10002, , "Given position (" & position & "} is greater than size of file (" & pSize & ")"
		If position + numberOfBytes > pSize Then Err.Raise 10003, , "Given position (" &  position & ") and numberOfBytes (" & numberOfBytes & ") exceeds size of file {" & pSize & ")"
		adodbStream.Position = position
		pBytes = adodbStream.Read(numberOfBytes)
		setStartOrFinish()
	End Sub
	
	Private Sub setStartOrFinish()
		If pByteOrder = 1 Then
			pFinish = LenB(pBytes)
		ElseIf pByteOrder = 2 Then
			pStart = LenB(pBytes)
		End If
	End Sub
	
End Class	


Class FileReaderWriter

	Private fso
	Private textStream
	Private IOMode
	Private pPath

	Public Sub openForReading(path)
		IOMode = 1
		pPath = path
		openFileIfExists()
	End Sub

	Public Sub openForWriting(path)
		IOMode = 2
		pPath = path
		openFile()
	End Sub

	Public Sub openForAppending(path)
		IOMode = 8
		pPath = path
		openFileIfExists()
	End Sub
	
	Public Sub writeLine(textWritten)
		textStream.WriteLine(textWritten)
	End Sub
	
	Public Function readLinesUntil(textSearched, includeFoundLine)
		Dim line
		
		Do Until textStream.AtEndOfStream
			line = textStream.ReadLine
			If InStr(line, textSearched) <> 0 Then
				If includeFoundLine Then
					readLinesUntil = readLinesUntil & line
				End If
				Exit Do
			Else
				readLinesUntil = readLinesUntil & line & vbNewLine
			End If
		Loop 
		
	End Function 
	
	Public Function readLinesAfter(textSearched, includeFoundLine)
		Dim line
		Dim isAfter

		isAfter = False
		
		Do Until textStream.AtEndOfStream
			line = textStream.ReadLine
			If isAfter Then
				readLinesAfter = readLinesAfter & line
				If textStream.AtEndOfStream Then
					Exit Do
				Else
					readLinesAfter = readLinesAfter & vbNewLine
				End If
			End If
			If InStr(line, textSearched) <> 0 Then
				isAfter = True
				If includeFoundLine Then
					readLinesAfter = line & vbNewLine
				End If
			End If
		Loop 
		
	End Function 

	Public Sub closeFile()
		textStream.Close
		Set textStream = Nothing
	End Sub

	Public Function readAll()
		readAll = textStream.ReadAll()
	End Function
	
	Public Function getFullPath()
		getFullPath = pPath
	End Function
	
	Private Sub openFileIfExists()
		If fso.FileExists(pPath) Then
			openFile()
		Else
			Err.Raise 10001, , "Given path does not exists: " & path
		End If
	End Sub
	
	Private Sub openFile()
		Set textStream = fso.OpenTextFile(pPath, IOMode, True)
	End Sub
	
	Private Sub Class_Initialize()
		Set fso = CreateObject("Scripting.FileSystemObject")
	End Sub
	
	Private Sub Class_Terminate()
		Set textStream = Nothing
		Set fso = Nothing
	End Sub

End Class



Class WorkbookController
	
	Private application
	Private workbook
	Private filePath
	Private fso
	
	Public Function getFullPath()
		getFullPath = filePath
	End Function
	
	Private Sub Class_Initialize()
		Set application = CreateObject("Excel.Application")
		Set fso = CreateObject("Scripting.FileSystemObject") 
	End Sub

	Public Sub prepareNewXmlWorkbook(path)
		filePath = path
		Create()
		SaveAsXml()
	End Sub
	
	Private Sub create()
		Set workbook = application.workbooks.Add()
	End Sub
		
	Private Sub saveAsXml()
		If fso.FileExists(filePath) Then fso.DeleteFile(filePath) 
		workbook.SaveAs filePath, 46 ' 46 = xlXMLSpreadsheet
	End Sub
	
	Public Sub open(path)
		Set workbook = application.workbooks.Open(path)
	End Sub
	
	Public Sub close()
		workbook.Close()
	End Sub
	
	Public Sub saveAs(path)
		workbook.SaveAs path, 51 '51 = xlOpenXMLWorkbook
	End Sub
	
	Public Sub showAppWindow()
		application.Visible = True
	End Sub
	
	Private Sub Class_Terminate()
		Set workbook = Nothing
		Set application = Nothing
		Set fso = Nothing
	End Sub
	
End Class