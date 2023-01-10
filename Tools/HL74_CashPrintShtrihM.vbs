'#include "HL74_Const.avb"

Option Explicit
'	prmPrinter
'	0	- 	Msc
'	1 	-	кассир
'	2	-	пароль кассира
'	3 	-	прочее, если нужно

'	prmData для печати чеков
'	0 	-	операция
'	1	-	номер проводки
'	2	-	тип платежа

'	prmData для печати отчетов 
'	0	-	вид отчета
'	1	-	тип отчета (сокращенный или развернутый)
'	2	-	выводить полный период
'	3	-	начало периода
'	4	-	конец периода
'	5	-	начальный номер
'	6	-	конечный номер

Const MSG_ERROR_PAYMENT = "Ошибка оплаты фискального чека"
Const MSG_ERROR_PRINTROW = "Ошибка печати строки фискального чека"
Const MSG_ERROR_NO_REPORT = "Отчет данным ЭККА не поддерживается"
Const MSG_ERROR_PRN_OPEN = "Ошибка доступа к принтеру"
Const MSG_ERROR_NO_COM = "Ошибка создания OLE/COM объекта"
Const MSG_ERROR_PRINT_OUT = "Ошибка печати чека выемки денег"
Const MSG_ERROR_PRINT_IN = "Ошибка печати чека внесения денег"
Const MSG_ERROR_OPEN_DRAWER = "Ошибка открытия денежного ящика"
Const MSG_ERROR_RESET_CHECK = "Ошибка аннулирования чека"
Const MSG_ERROR_OPEN_CHECK = "Ошибка открытия чека"

Const MSG_NO_SUPPPORT = "Функция для данного принтера не поддерживается"
Const MSG_BOX_CAPTION = "Печать фискального чека"

Const prmEntCashPrnName = "Наименование для ЭККА"

Const PORT_NO = "1"
Const BAUD_RATE = 115200

main

Sub main()
	Dim oArg, cmd, filename

	Set oArg = Wscript.Arguments

	If oArg.Count = 0 Then
		MsgBox "Using: HL74_CashPrintShtrihM.vbs <cmd> [<filename>]"
	Else
		cmd = oArg.Item(0)

		If oArg.Count = 1 Then
			filename = oArg.Item(1)
		Else
			filename = ""
		End If

		RunCommand cmd, filename
	End If
End Sub
'---
'
'---
Sub RunCommand(cmd, filename)
	Dim f, fso, Buffer

	If filename <> "" Then
		Set fso = CreateObject("Scripting.FileSystemObject")

		If fso.isfileexists(filename) Then
			Set f = f.OpenTextFile(filename)
			Buffer = f.ReadAll
			f.Close
		Else
			MsgBox "Файл не найден (" & filename & ")", 48, "Внимание !"
		End If

		Select Case LCase(cmd)
			Case "printcheck"
				'ART_NO;ART_NAME;PRICE;SUM
				PrintCash_DoPrintCheck Buffer
			Case "printretcheck"
				'ART_NO;ART_NAME;PRICE;SUM
				PrintCash_DoPrintCheckRet Buffer
			Case "printcheckin"
				'SUM
				PrintCash_PrintCheckIn Buffer
			Case "printcheckout"
				'SUM
				PrintCash_PrintCheckOut Buffer
			Case "printxreport"
				PrintCash_PrintReports 1
			Case "printzreport"
				PrintCash_PrintReports 2
			Case "opencashdrawer"
				PrintCash_OpenCashDrawer
			Case "resetcheck"
				PrintCash_ResetCheck
		End Select
	End If
End Sub
'---
'
'---
Sub PrintCash_DoPrintCheck(Buffer)
	Dim CashPrn

	Set CashPrn = PrintCash_Open(PORT_NO, BOUD_RATE)

	If Not CashPrn Is Nothing Then
		PrintCash_PrintCheck CashPrn, Buffer
		PrintCash_Close CashPrn
		Set CashPrn = Nothing
	End If
End Sub
'---
'
'---
Sub PrintCash_DoPrintCheckRet(Buffer)
	Dim CashPrn

	Set CashPrn = PrintCash_Open(PORT_NO, BOUD_RATE)

	If Not CashPrn Is Nothing Then
		PrintCash_PrintCheckRet CashPrn, Buffer
		PrintCash_Close CashPrn
		Set CashPrn = Nothing
	End If
End Sub
'---
'
'---
Sub SetPaymentType(CashPrn, PaymentType, PaymentSum)
	Select Case PaymentType
		Case 1
			CashPrn.Summ1 = PaymentSum
		Case 2
			CashPrn.Summ2 = PaymentSum
		Case 3
			CashPrn.Summ3 = PaymentSum
		Case 4
			CashPrn.Summ4 = PaymentSum
	End Select

End Sub 
'---
'
'---
Function PrintCash_Open(PortNo, BoudRate)
	Dim CashPrn, Passw

	Passw = "0000"	

	On Error Resume Next

	Set PrintCash_Open = Nothing
	Set CashPrn = CreateObject("Addin.DRvFR")

	If Err.Number = 0 Then
		If CashPrn.Connect2(CLng(Passw), CLng(PortNo), CLng(BoudRate), 5000) = 0 Then
			Set PrintCash_Open = CashPrn
		Else
			MsgBox MSG_ERROR_PRN_OPEN & Chr(13) + Chr(10) & CashPrn.ResultCodeDescription, 16, MSG_BOX_CAPTION
		End If
	Else
		MsgBox MSG_ERROR_NO_COM & Chr(13) + Chr(10) & Err.description & " (" & Err.Number & ")", 16, MSG_BOX_CAPTION
	End If

End Function
'---
'
'---
Function PrintCash_PrintCheckRet(CashPrn, Buffer)
	Dim aBuffer, i, aLine, Res

	PrintCash_PrintCheckRet = True
	CashPrn.CheckType = 2	' возврат продажи

	If CashPrn.OpenRetCheck() = 0 Then
		aBuffer = Split(Buffer, Chr(13) + Chr(10))
		
		For i = 0 To UBound(aBuffer)
			aLine = Split(aBuffer(i), ";")

			CashPrn.StringForPrinting = aLine(0) & " " & aLine(1)
			CashPrn.Quantity = CDbl(aLine(2))
			CashPrn.Price = CDbl(aLine(3))

			Res = CashPrn.ReturnSale()

			If Res <> 0 Then 
				MsgBox MSG_ERROR_PRINTROW & Chr(13) + Chr(10) & CashPrn.ResultCodeDescription, 16, MSG_BOX_CAPTION
				If CashPrn.ResultCode <> 0 Then CashPrn.CancelCheck
				PrintCash_PrintCheckRet = False
				Exit Function
			End If
		Next
	Else
		MsgBox MSG_ERROR_OPEN_CHECK & Chr(13) + Chr(10) & CashPrn.ResultCodeDescription, 16, MSG_BOX_CAPTION		
	End If

End Function

'---
'
'---
Function PrintCash_PrintCheck(CashPrn)
	Dim aBuffer, i, aLine, Res

	PrintCash_PrintCheck = True
	CashPrn.CheckType = 0	' возврат продажи

	If CashPrn.OpenRetCheck() = 0 Then
		aBuffer = Split(Buffer, Chr(13) + Chr(10))
		
		For i = 0 To UBound(aBuffer)
			aLine = Split(aBuffer(i), ";")

			CashPrn.StringForPrinting = aLine(0) & " " & aLine(1)
			CashPrn.Quantity = CDbl(aLine(2))
			CashPrn.Price = CDbl(aLine(3))

			Res = CashPrn.Sale()

			If Res <> 0 Then 
				MsgBox MSG_ERROR_PRINTROW & Chr(13) + Chr(10) & CashPrn.ResultCodeDescription, 16, MSG_BOX_CAPTION
				PrintCash_PrintCheck = False
				If CashPrn.ResultCode <> 0 Then CashPrn.CancelCheck
				Exit Function
			End If
		Next
	Else
		MsgBox MSG_ERROR_OPEN_CHECK & Chr(13) + Chr(10) & CashPrn.ResultCodeDescription, 16, MSG_BOX_CAPTION		
	End If

End Function
'---
'
'---
Function PrintCash_Close(CashPrn)
	PrintCash_Close = CashPrn.ClosePort
End Function
'---
'
'	prmData для печати чеков
'	0	-	вид отчета
'	1	-	тип отчета (сокращенный или развернутый)
'	2	-	выводить полный период
'	3	-	начало периода
'	4	-	конец периода
'	5	-	начальный номер
'	6	-	конечный номер
'
'---
Sub PrintCash_PrintReports(prm)
	Dim IsErr, Passw
	Dim sDate, eDate, sN, eN, RepMode

	Passw = "00000"

	Set CashPrn = PrintCash_Open(PORT_NO, BOUD_RATE)

	If Not CashPrn Is Nothing Then
		
		Select Case prm
			Case 1	'	x-report
				IsErr = CashPrn.PrintXReport
	
			Case 2	'	z-report
				IsErr = CashPrn.PrintZReport
	
			Case 3	'	report by art
				MsgBox MSG_NO_SUPPPORT, 64, MSG_BOX_CAPTION
	
			Case 4	'	report by disount
				MsgBox MSG_NO_SUPPPORT, 64, MSG_BOX_CAPTION
	
			Case 5	'	report by Date
				MsgBox MSG_NO_SUPPPORT, 64, MSG_BOX_CAPTION
	
			Case 6	'	report by no
				MsgBox MSG_NO_SUPPPORT, 64, MSG_BOX_CAPTION

			Case Else
				MsgBox MSG_NO_SUPPPORT, 64, MSG_BOX_CAPTION
	
		End Select 
	End If

	Set CashPrn = Nothing
End Sub
'---
'
'---
Sub PrintCash_PrintNullCheck()
	MsgBox MSG_NO_SUPPPORT, 64, MSG_BOX_CAPTION
End Sub
'---
'
'---
Sub PrintCash_PrintCheckOut(Buffer)
	Set CashPrn = PrintCash_Open(PORT_NO, BOUD_RATE)

	If Not CashPrn Is Nothing Then
		If Not CashPrn.PrintCheckOut(CDbl(Buffer)) Then
			MsgBox MSG_ERROR_PRINT_IN & Chr(13) & Chr(10) & CashPrn.ErrorMsg, 16, MSG_BOX_CAPTION
		End If
	End If

	Set CashPrn = Nothing

End Sub
'---
'
'---
Sub PrintCash_PrintCheckIn(buffer)
	Set CashPrn = PrintCash_Open(PORT_NO, BOUD_RATE)

	If Not CashPrn Is Nothing Then
		If Not CashPrn.PrintCheckIn(CDbl(buffer)) Then
			MsgBox MSG_ERROR_PRINT_OUT & Chr(13) & Chr(10) & CashPrn.ErrorMsg, 16, MSG_BOX_CAPTION
		End If
	End If

	Set CashPrn = Nothing

End Sub
'---
'
'---
Sub PrintCash_OpenCashDrawer()
	Set CashPrn = PrintCash_Open(PORT_NO, BOUD_RATE)

	If Not CashPrn Is Nothing Then
		If Not CashPrn.OpenCashDrawer Then
			MsgBox MSG_ERROR_OPEN_DRAWER & Chr(13) & Chr(10) & CashPrn.ErrorMsg, 16, MSG_BOX_CAPTION
		End If
	End If

	Set CashPrn = Nothing
End Sub
'---
'
'---
Function PrintCash_ResetCheck()
	Set CashPrn = PrintCash_Open(PORT_NO, BOUD_RATE)

	If Not CashPrn Is Nothing Then
		If Not CashPrn.ResetCheck() Then
			MsgBox MSG_ERROR_RESET_CHECK & Chr(13) & Chr(10) & CashPrn.ErrorMsg, 13, MSG_BOX_CAPTION
		End If
	End If

	Set CashPrn = Nothing

End Function
'---
'
'---

