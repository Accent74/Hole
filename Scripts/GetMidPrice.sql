USE [wolmart_work]
GO
/****** Object:  StoredProcedure [dbo].[ST7_GetMidEntPrice]    Script Date: 09.06.2016 21:03:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO






ALTER PROCEDURE [dbo].[ST7_GetMidEntPrice]
      @AgID Int, @AccID Int, @EntID Int, @DocID Int, @OnDate DateTime, @MC int

AS
SET NOCOUNT ON

Select	Sum(Case 
				When ACC_DB = @AccID And J_AG1 = @AgID Then J_QTY 
				Else 0 
			End - 
			case 
				when ACC_CR = @AccID and J_AG2 = @AgID and j_DATE < @OnDate then J_QTY
				else 0
			end),
        Sum(Case 
				When ACC_DB = @AccID And J_AG1 = @AgID Then J_SUM 
				Else 0 
			End - 
			case
				when ACC_CR = @AccID and J_AG2 = @AgID and j_DATE < @OnDate then J_SUM
				else 0
			end),
        Sum(Case 
				When ACC_DB = @AccID And J_AG1 = @AgID Then J_QTY
				Else 0 
			End - 
			case
				when ACC_CR = @AccID and J_AG2 = @AgID then J_QTY
				else 0
			end),
        Sum(Case 
				When ACC_DB = @AccID And J_AG1 = @AgID Then J_SUM 
				Else 0 
			End - 
			case
				when ACC_CR = @AccID and J_AG2 = @AgID then J_SUM
				else 0
			end)


FROM JOURNAL
Where (ACC_DB = @AccID And J_AG1 = @AgID Or ACC_CR = @AccID And J_AG2 = @AgID) And 
J_DONE = 2 And J_DATE <= @OnDate And J_ENT = @EntID And DOC_ID <> @DocID And MC_ID = @MC;





