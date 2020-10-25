USE [Avantage]
GO

/****** Object:  StoredProcedure [dbo].[ST7_GetMidEntPrice_Retail]    Script Date: 26.05.2016 14:04:06 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO






CREATE PROCEDURE [dbo].[ST7_GetMidEntPrice_Retail]
      @AgID Int, @AccID Int, @EntID Int, @DocID Int, @OnDate DateTime, @MC int

AS
SET NOCOUNT ON

Select Sum(Case When ACC_DB = @AccID And J_AG1 = @AgID Then J_QTY Else 0 End),
           Sum(Case When ACC_DB = @AccID And J_AG1 = @AgID Then J_SUM Else 0 End),
           Sum((Case When ACC_DB = @AccID And J_AG1 = @AgID Then J_SUM Else 0 End) * Case When IsNull(J_QTY, 0) = 0 Then 1 Else 0 End)
FROM JOURNAL
Where (ACC_DB = @AccID And J_AG1 = @AgID Or ACC_CR = @AccID And J_AG2 = @AgID) And 
J_DONE = 2 And J_DATE <= @OnDate And J_ENT = @EntID And DOC_ID <> @DocID And MC_ID = @MC;





GO

/****** Object:  StoredProcedure [dbo].[ST7_GetMidEntPrice]    Script Date: 26.05.2016 14:04:07 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO






CREATE PROCEDURE [dbo].[ST7_GetMidEntPrice]
      @AgID Int, @AccID Int, @EntID Int, @DocID Int, @OnDate DateTime, @MC int

AS
SET NOCOUNT ON

Select Sum(Case When ACC_DB = @AccID And J_AG1 = @AgID Then J_QTY Else 0 End),
           Sum(Case When ACC_DB = @AccID And J_AG1 = @AgID Then J_SUM Else 0 End)
FROM JOURNAL
Where (ACC_DB = @AccID And J_AG1 = @AgID Or ACC_CR = @AccID And J_AG2 = @AgID) And 
J_DONE = 2 And J_DATE <= @OnDate And J_ENT = @EntID And DOC_ID <> @DocID And MC_ID = @MC;





GO

/****** Object:  StoredProcedure [dbo].[ST74_CalcEntRest]    Script Date: 26.05.2016 14:04:07 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO



CREATE PROCEDURE [dbo].[ST74_CalcEntRest]
   @AgID Int, 
   @AccID Int, 
   @DocID Int, 
   @OnDate DateTime, 
   @MC int

AS
SET NOCOUNT ON

---
--- Создаем временную таблицу для выборки объектов в разрезе серий
---
CREATE TABLE #Ent(SerID int, EntID int, RestQty Money, RestSum Money)

--вставляем дебет (с плюсами)
INSERT INTO #Ent (SerID, EntID, RestQty, RestSum)

SELECT SER_ID, J_ent, 
	sum(CASE 
		WHEN (ACC_DB=@AccID And J_AG1=@AgID ) THEN IsNull(J_QTY, 0)
	ELSE 0
	END),

	sum(CASE 
		WHEN (ACC_DB=@AccID And J_AG1=@AgID) THEN J_SUM
	ELSE 0
	END)

FROM JOURNAL
WHERE JOURNAL.ACC_DB=@AccID AND JOURNAL.J_AG1=@AgID
		and JOURNAL.J_DONE=2
		AND (doc_id<>@DocID) 
		AND (JOURNAL.J_DATE<=@OnDate) 
		and JOURNAL.MC_ID = @MC
group by SER_ID, J_ent
--вставляем кредит (с минусами)
INSERT INTO #Ent (SerID, EntID, RestQty, RestSum)

SELECT SER_ID, J_ent, 
	sum(CASE 
		WHEN (ACC_CR=@AccID And J_AG2=@AgID ) THEN IsNull(-J_QTY, 0)
	ELSE 0
	END),

	sum(CASE 
		WHEN (ACC_CR=@AccID And J_AG2=@AgID) THEN -J_SUM
	ELSE 0
	END)
	
FROM JOURNAL
WHERE	JOURNAL.ACC_CR=@AccID AND JOURNAL.J_AG2=@AgID
		and JOURNAL.J_DONE=2
		AND (doc_id<>@DocID) 
		AND (JOURNAL.J_DATE<=@OnDate) 
		and JOURNAL.MC_ID = @MC
group by SER_ID, J_ent
--Окончательная выборка
SELECT	0, 
		EntID, 
		Sum(RestQty) AS RestQty, 
		Sum (RestSum) AS RestSum
FROM #Ent 
group by EntID
HAVING Sum(RestQty)<>0 or Sum(RestSum)<>0 

drop table #Ent




GO

