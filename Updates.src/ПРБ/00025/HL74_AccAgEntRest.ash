аЯрЁБс                >  ўџ	                               ўџџџ       џџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџ = dPKind(Command - 200)
			prKindID = PriceKind.ID
			prKindName = PriceKind.Name
		End If
	End If

	NeedRecalc = True
	RefreshToolbarButon
	recalc
End Sub
'---
'
'---
Sub ShtBook_OnRequery
	MakeReport
End Sub
'---
'
'---
Sub ShtBook_OnPeriodClick(Buttons)
	If workarea.period.browse Then MakeReport
End Sub
'---
'
'---
Sub ShtBook_OnBarClick(Command)
 	ExecBarTag Command
	Recalc
End Sub
'---
'
'---
Sub SelectAccount
	Dim AccIDNew

	AccIDNew = workarea.browse(acAccount, AccID,,, "Тћсх№шђх ёїхђ")

	If AccIDNew <> 0 Then 
		AccID = AccIDNew
		AccName = Workarea.Account(AccID).DrawText(3)
		NeedRecalc = True
		RefreshToolbarButon
	End If
End Sub
'---
'
'---
Sub SelectAgent
	Dim AgIDNew

	AgIDNew = workarea.browse(acAgent, AgID,,, "Тћсх№шђх ьхёђю ѕ№рэхэшџ")

	If AgIDNew <> 0 Then 
		AgID = AgIDNew
		AgNAme = workarea.Agent(AgID).Name		
		NeedRecalc = True
		RefreshToolbarButon
	End If
End Sub
'---
'
'---
Sub SelectEnt
	Dim EntIDNew

	EntIDNew = workaR o o t   E n t r y                                               џџџџџџџџ                               #Ми         C o n t e n t s                                                  џџџџ   џџџџ                                       C       S u m m a r y I n f o r m a t i o n                           (  џџџџџџџџџџџџ                                        e                                                                          џџџџџџџџџџџџ                                                      ўџџџ§џџџўџџџўџџџ      	   
                                                                       !   "   #   $   %   ўџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџ               ўџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџўџ   рђљOhЋ +'Гй   рђљOhЋ +'Гй0   5        x                   Ќ      И      Ф      а   	   м   
   ш      є                       (     у        HL74_AccAgEntRest.ash                                                           7   @   ,Оз   @   аМи@   Р9qCж      Ръіхэђ 7.4                                                                                                                                                                               џ2%'#include "HL74_Common.avb"
'#include "HL74_Const.avb"
'#include "HL74_ADO.avb"

Option Explicit

Const START_ROW = 10
Const MAX_COLUMNS = 15

Dim dPrice, dPKind
Dim prListID, 	prKindID, PriceKind
Dim AccID, AgID, EntID
Dim AgName, AccName, EntName, prListName, prKindName
'---
'
'---
Sub SaveReportSettings
	Dim aPrm, key

	aPrm = Array(AccID, AgID, prListID, prKindID, EntID)
	Key = Workarea.Map("RepID") & ":" & workarea.dbuser
	workarea.userparam(Key) = Join(aPrm, ",")

End Sub
'---
'
'---
Sub LoadReportSettings
	Dim prm, key, aPrm

	Key = Workarea.Map("RepID") & ":" & workarea.dbuser
	prm = workarea.userparam(Key)

	If IsNull(prm) Then
		AccID = workarea.GetAccID("281")
		AgID = 0
		prListID = workarea.DefPriceList 
		prKindID = 0
		EntID = 0

		With WorkArea.Site
			Select Case .Kind
				Case acAgent
					AgID = iif(.ID <> 0, .ID, AgID)
				Case acAccount
					AccID = .ID
				Case acEntity
					EntID = .ID
			End Select
		End With
	Else
		If prm = "" Then prm = ",,,,"
		aPrm = Split(prm, ",")

		AccID = str2long(aPrm(0))
		AgID = str2long(aPrm(1))
		prListID = str2long(aPrm(2)) 
		prKindID = str2long(aPrm(3))
		EntID = str2long(aPrm(4))

		If AccID <> 0 Then AccID = iif(IsNull(workarea.account(AccID)), workarea.GetAccID("281"), AccID)
		If AgID <> 0 Then AgID = iif(IsNull(workarea.agent(AgID)), 0, AgID)
		If prListID <> 0 Then prListID = iif(IsNull(workarea.PriceList(prListID)), workarea.DefPriceList, prListID)
		If prKindID <> 0 Then prKindID = iif(IsNull(workarea.PriceKind(prKindID)), 0, prKindID)
		If EntID <> 0 Then EntID = iif(IsNull(workarea.entity(EntID)), 0, EntID)
	End If

	If prKindID = 0 And AgID <> 0 Then prKindID = com_getparamvalue(WorkArea.Agent(AgID), prmAgDefPrice, 0)
	If AgID <> 0 Then AgName = workarea.Agent(AgID).Name Else AgName = "<Тёх ъю№№хёяюэфхэђћ>"
	AccName	 = workarea.account(AccID).DrawText(3)
	If EntID <> 0 Then EntName = workarea.entity(EntID).Name Else EntName = "<Тхёќ ђютр№>"
	prListName = workarea.pricelist(prListID).Name
	If prKindID <> 0 Then prKindName = workarea.PriceKind(prKindID).Name Else prKindName = "<гїхђэћх іхэћ>"
End Sub
'---
'
'---
Sub RefreshToolbarButon
	With ToolBar
		.ItemByCommand(1).Caption = "Я№рщё ышёђ: " & prListName
		.ItemByCommand(2).Caption = "Тшф іхэћ: " & prKindName
		.ItemByCommand(3).Caption = "бїхђ: " & AccName
		.ItemByCommand(4).Caption = "Ьхёђю ѕ№рэхэшџ: " & AgName
		.ItemByCommand(5).Caption = "Яряър ђютр№р: " & EntName
		.Refresh
	End With

	recalc
End Sub
'---
'
'---
Sub MakeReport
	Dim aData, i, aPrm, CurrAgID, RowAg, RowNo
	Dim WC, Mtr

	Set WC = WaitCursor
	
	Sheet1.Rows = START_ROW
	Sheet1.Columns = MAX_COLUMNS
	Sheet1.Range(4, START_ROW, MAX_COLUMNS, START_ROW).Value = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)

	PeriodText = workarea.period.title

	aPrm = Array(AccID, AgID, EntID, workarea.period.start, workarea.period.end, prListID, prKindID, -1, -1, -1, workarea.mycompany.id)
	NeedRecalc = False

	If Query("HL74_AccAgEntRest", aPrm, aData) Then
		CurrAgID = -1
		Set Mtr = Meter
		Mtr.Open "Чряюыэџў ђрсышіѓ ...", 0, UBound(aData, 2)

		If prKindID <> 0 Then 
			Set PriceKind = Workarea.PriceKind(prKindID)
		Else
			Set PriceKind = Nothing
		End If

		For i = 0 To UBound(aData, 2)
			Mtr.Pos = i
			If aData(0, i) <> CurrAgID Then
				If CurrAgID <> -1 Then Sheet1.Range(1, RowAg + 1, Sheet1.Columns, Sheet1.Rows).Stripe

				CurrAgID = aData(0, i)
				RowAg = Sheet1.InsertRow2
				Sheet1.Cell(RowAg, 1).Value = aData(2, i)	' a.ag_name
				Sheet1.Range(1, RowAg, 3, RowAg).MergeCells = True
				With Sheet1.Range(1, RowAg, Sheet1.Columns, RowAg)
					.Font.Bold = True
					.BackColor = StdBackColor(2)
				End With
			End If

			RowNo = Sheet1.InsertRow2

			Sheet1.Cell(RowNo, 1).Value = aData(3, i)		' e.ent_name
			Sheet1.Cell(RowNo, 2).Value = aData(4, i)		' e.ent_cat
			Sheet1.Cell(RowNo, 3).Value = aData(5, i)		' e.ent_art

			CalcSumRow aData, RowNo, i, 5, 6, 8
			CalcSumRow aData, RowNo, i, 8, 11, 9
			CalcSumRow aData, RowNo, i, 11, 14, 12
			CalcSumRow aData, RowNo, i, 14, 17, 15

			Sheet1.Cell(RowNo, 4).Value = aData(6, i)		' StartQty
			Sheet1.Cell(RowNo, 6).Value = aData(8, i)		' StartSum

			CalcTotals RowAg, 4, 6, aData(6, i), aData(8, i)

			Sheet1.Cell(RowNo, 7).Value = aData(9, i)		' DBQty
			Sheet1.Cell(RowNo, 9).Value = aData(11, i)		' DBSum

			CalcTotals RowAg, 7, 9, aData(9, i), aData(11, i)

			Sheet1.Cell(RowNo, 10).Value = aData(12, i)		' CrQty
			Sheet1.Cell(RowNo, 12).Value = aData(14, i)		' CrSum

			CalcTotals RowAg, 10, 12, aData(12, i), aData(14, i)

			Sheet1.Cell(RowNo, 13).Value = aData(15, i)		' EndQty
			Sheet1.Cell(RowNo, 15).Value = aData(17, i)		' EndSum

			CalcTotals RowAg, 13, 15, aData(15, i), aData(17, i)
		Next

		Sheet1.Range(1, RowAg + 1, Sheet1.Columns, Sheet1.Rows).Stripe
		Mtr.Close
		FormatSheet
	
	End If

	Set WC = Nothing
	recalc
End Sub
'---
'
'---
Sub CalcSumRow(aData, RowNo, Row, PriceCol, QtyCol, SumCol)
	If PriceKind Is Nothing Then
		If aData(QtyCol, Row) <> 0 Then 
			Sheet1.Cell(RowNo, PriceCol).Value = aData(SumCol, Row) / aData(QtyCol, Row) 
		Else 
			Sheet1.Cell(RowNo, PriceCol).Value = 0
		End If
	Else
		Sheet1.Cell(RowNo, PriceCol).Value = PriceKind.GetEntPrice(aData(1, Row), Workarea.period.End, prListID)
		aData(SumCol, Row) = Sheet1.Cell(RowNo, PriceCol).Value * aData(QtyCol, Row)
	End If
End Sub
'---
'
'---
Sub FormatSheet
	Sheet1.Range(1, START_ROW + 1, 1, Sheet1.Rows).Multiline = True
	Sheet1.Range(4, START_ROW, Sheet1.Columns, Sheet1.Rows).Alignment = acRight
	Sheet1.Range(2, START_ROW - 1, Sheet1.Columns, Sheet1.Rows).AutoFit 2
	Sheet1.Range(1, START_ROW, Sheet1.Columns, Sheet1.Rows).SetBorder acBrdGrid, , RGB(125, 158, 192)
	Sheet1.AutoFit 1
	Sheet1.FixRows START_ROW
End Sub
'---
'
'---
Sub CalcTotals(RowNo, xQty, xSum, Qty, Sum)
	Sheet1.Cell(RowNo, xQty).Value = Sheet1.Cell(RowNo, xQty).Value + Qty
	Sheet1.Cell(RowNo, xSum).Value = Sheet1.Cell(RowNo, xSum).Value + Sum
	Sheet1.Cell(START_ROW, xQty).Value = Sheet1.Cell(START_ROW, xQty).Value + Qty
	Sheet1.Cell(START_ROW, xSum).Value = Sheet1.Cell(START_ROW, xSum).Value + Sum
End Sub
'---
'
'---
Sub ShtBook_OnLoad
	Set dPrice = CreateLibObject("Map")
	Set dPKind = CreateLibObject("Map")

	Sheet1.DisplayGrid = False
	Sheet1.DisplayHeadings = False

	LoadReportSettings
	RefreshToolbarButon

	LoadPrices
	enablebuttons 1

	MakeReport
		
End Sub
'---
'
'---
Sub LoadPrices
	Dim i, TBtn, MenuPopup

	MenuPopup = "1:<Эх ѓърчрэ>:100"

	With workarea.pricelists
		For i = 1 To .count
			With .Item(i)
				MenuPopup = MenuPopup & "|1:" & .Name & ":1" & com_strzero(i, 2)
				If prListID = .ID Then prListName = .Name
			End With

			Set dPrice.Item(i) = .Item(i)
		Next
	End With

	LoadPKinds prListID

	Set TBtn = ToolBar.ItemByCommand(1)
	TBtn.PopUp = MenuPopup
End Sub
'---
'
'---
Sub LoadPKinds(prListID)
	Dim i, TBtn, MenuPopup

	MenuPopup = "1:<гїхђэрџ іхэр>:200"

	If prListID > 0 Then
		With workarea.PriceList(prListID)
			With .PriceKinds
				For i = 1 To .count
					With .Item(i)
						MenuPopup = MenuPopup & "|1:" & .Name & ":2" & com_strzero(i, 2)
						If prKindID = .ID Or (prKindID = 0 And i = 1) Then prKindName = .Name
					End With

					Set dPKind.Item(i) = .Item(i)
				Next
			End With
		End With
	End If

	Set TBtn = ToolBar.ItemByCommand(2)
	TBtn.PopUp = MenuPopup
End Sub
'---
'
'---
Sub ShtBook_OnPopup(Command)
	Dim pList

	If Command < 200 Then
		' я№рщё ышёђћ
		If Command <> 100 Then
			Set pList = dPrice(Command - 100)
			prListID = pList.ID
			prListName = pList.Name
			LoadPKinds prListID
		Else
			Set PriceKind = Nothing
			prKindID = 0
			prListID = 0
			prListName = "<Эх ѓърчрэ>"
			prKindName = "<гїхђэрџ іхэр>"
		End If

	ElseIf Command = 301 Then
		EntID = 0
		EntName = "<Тхёќ ђютр№>"
	Else
		' тшф іхэћ
		If Command = 200 Then
			Set PriceKind = Nothing
			prKindID = 0
			prKindName = "<гїхђэрџ іхэр>"
		Else
			Set PriceKindrea.browse(acEntity, EntID,, 4 + 8, "Тћсх№шђх яряъѓ ђютр№р")

	If EntIDNew <> 0 Then 
		EntID = EntIDNew	
		EntName = Workarea.Entity(EntID).Name
		NeedRecalc = True
		RefreshToolbarButon
	End If
End Sub
'---
'
'---
Sub ShtBook_CanClose(ByRef Cancel)
	SaveReportSettings
End Sub
'---
'
'---
<     ш                Ь Arial Йа[wz   pяџ     єяџ -єu|kТhs      џџџџ                 џџ 
 CSheetPageSheet1Ышёђ 1   
      61    7                                       ^   X   :   :   O   @   #   U   @   #   U   :   #   O   
                               
 џџ  CRow	   џџ  CCell6 *=iif(NeedRecalc, "ваХСгХвбп ЯХаХбзХв", "")   џџ
                           џџџџ                         џџџџ                         џџџџ                         џџџџ                         џџџџ                         џџџџ                         џџџџ                        $ џџџџ   $   2865=85  B>20@0  70                   =Workarea.Period.Title$   џџ
                         $   џџ                       $   џџ                       $   џџ                       $   џџ                       $   џџ                       $   џџ                       $   џџ                       $   џџ                       $   џџ                       $   џџ                       $   џџ                       $   џџ                       $   џџ                        $ џџџџ      ?>  AG5BC                   =AccName$   џџ
                         $   џџ                       $   џџ                       $   џџ                       $   џџ                       $   џџ                       $   џџ                       $   џџ                       $   џџ                       $   џџ                       $   џџ                       $   џџ                       $   џџ                       $   џџ                        $ џџџџ      :>@@5A?>=45=BC                   =AgName$   џџ
                         $   џџ                       $   џџ                       $   џџ                       $   џџ                       $   џџ                       $   џџ                       $   џџ                       $   џџ                       $   џџ                       $   џџ                       $   џџ                       $   џџ                       $   џџ                        $ џџџџ      ?0?:5  #                   =EntName$   џџ
                         $   џџ                       $   џџ                       $   џџ                       $   џџ                       $   џџ                       $   џџ                       $   џџ                       $   џџ                       $   џџ                       $   џџ                       $   џџ                       $   џџ                       $   џџ                        $ џџџџ      @09A  ;8AB                   =prListName$   џџ
                         $   џџ                       $   џџ                       $   џџ                       $   џџ                       $   џџ                       $   џџ                       $   џџ                       $   џџ                       $   џџ                       $   џџ                       $   џџ                       $   џџ                       $   џџ                        $ џџџџ      84  F5=K                   =prKindName$   џџ
                         $   џџ                       $   џџ                       $   џџ                       $   џџ                       $   џџ                       $   џџ                       $   џџ                       $   џџ                       $   џџ                       $   џџ                       $   џџ                       $   џџ                       $   џџ                        5 џџ  b   08<5=>20=85  :>@@5A?>=45=B0  /   08<5=>20=85  B>20@0                      џџ                          џџ                        5 џџ  "   AB0B>:  =0  =0G0;>                    5 џџ                           5 џџ                           5 џџ     0:C?:0                    5 џџ                           5 џџ                           5 џџ     @>4068                    5 џџ                           5 џџ                           5 џџ     AB0B>:                    5 џџ                           5 џџ                            5 џџ       08<5=>20=85                    % џџ       0B0;.   !                    % џџ       @B8:C;                    5 џџ       >;- 2>                    5 џџ       &5=0                    5 џџ    
   !C<<0                    5 џџ       >;- 2>                    5 џџ       &5=0                    5 џџ    
   !C<<0                    5 џџ       >;- 2>                    5 џџ       &5=0                    5 џџ    
   !C<<0                    5 џџ       >;- 2>                    5 џџ       &5=0                    5 џџ    
   !C<<0                     $        B>3>:                     $                           $                           &       ррG                        &                             &       жR9                       &       Р)Х                        &                             &       Н)']Й                      &       р$сЙ                        &                             &       ЖЃP
                      &       РС                        &                             &       ­Ѕr                         џџџџ                           џџџџ                       d   d   d   d    <     ш                Ь Arial /Y|ъ~  цЄ   8хЄxяџ  <     ш                Ь Arial        Мџ эi|ъцЄ    <     ш                Ь Arial aГѓш8ИX;И       ь  ўџџџ <     ш                Ь Arial /Y|ъ~  цЄ   8хЄxяџ  <     ш                Ь Arial        Мџ эi|ъцЄ    <     ш                Ь Arial aГѓш8ИX;И       ь  ўџџџ    2   2       <     ш         џ   Ь Arial xяџ xяџ     p јНџ VY|8хЄ <     ш         џ   Ь Arial xяџ xяџ     p јНџ VY|8хЄ <     ш           џ   Ь Arial xяџ xяџ     p јНџ VY|8хЄ <     ш            џ   Ь Arial xяџ xяџ     p јНџ VY|8хЄ <     ш            џ­[ Ь Arial xяџ xяџ     p јНџ VY|8хЄ  <     ш               Ь Arial xяџ xяџ     p јНџ VY|8хЄ    ццњ    }Р    }Р    }Р    }Р  џџџ    }Р    }Р    }Р    }Р  рџџ    }Р    }Р    }Р    }Р  џџџ                                  рџџ                                  џџр                                  џџр    }Р    }Р    }Р    }Р   џџџџ   }Р    }Р    }Р    }Р  џџ< 
 CBarButtonO      Я№рщё ышёђ: Я№рщё-ышёђ ѓіхэърТћсх№шђх я№рщё ышёђ   B1:<Эх ѓърчрэ>:100|1:Я№рщё фыџ ъышхэђют:101|1:Я№рщё-ышёђ ѓіхэър:102(ShowPopUp ToolBar.ItemByCommand(1).PopUp                #      Тшф іхэћ: <гїхђэћх іхэћ>Тћсх№шђх тшф іхэћ   1:<гїхђэрџ іхэр>:200(ShowPopUp ToolBar.ItemByCommand(2).PopUp                      бїхђ: 281 вютр№ћ эр ёъырфх     SelectAccount                $      $Ьхёђю ѕ№рэхэшџ: <Тёх ъю№№хёяюэфхэђћ>     SelectAgent                %      Яряър ђютр№р: <Тхёќ ђютр№>    1:Тхёќ ђютр№:301	SelectEnt                                                                                                                              