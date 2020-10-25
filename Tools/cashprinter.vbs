Option Explicit

Const MSC_BOX_CAPTION = "Работа с ЭККА"

Dim m119F

stop
Set m119F = CreateObject("CashPrints.CashPrint")

If Not m119F.open(1, 9600, 0, "0000") Then
	MsgBox "Ошибка открытия порта принтера" & vbNewLine & m119F.GetErrorDescription & "( #" & m119F.GetErrorNo & ")", vbCritical, MSC_BOX_CAPTION
Else
	m119F.close
End If

Set m119F = Nothing