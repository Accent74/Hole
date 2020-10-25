/****** Object:  StoredProcedure [dbo].[HL74_RestEntSeries]    Script Date: 23.02.2016 21:29:13 ******/
DROP PROCEDURE [dbo].[HL74_RestEntSeries]
GO

/****** Object:  StoredProcedure [dbo].[st7_oddments_of_the_goods_price]    Script Date: 23.02.2016 21:29:13 ******/
DROP PROCEDURE [dbo].[st7_oddments_of_the_goods_price]
GO

/****** Object:  StoredProcedure [dbo].[st7_oddments_of_the_goods]    Script Date: 23.02.2016 21:29:13 ******/
DROP PROCEDURE [dbo].[st7_oddments_of_the_goods]
GO

/****** Object:  StoredProcedure [dbo].[HL74_custombind_goods_oddments_price]    Script Date: 23.02.2016 21:29:13 ******/
DROP PROCEDURE [dbo].[HL74_custombind_goods_oddments_price]
GO

/****** Object:  StoredProcedure [dbo].[HL74_custombind_goods_oddments_price]    Script Date: 23.02.2016 21:29:13 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


CREATE PROCEDURE [dbo].[HL74_custombind_goods_oddments_price]
		  @EntName nvarchar(255), @AgID Int, @AccID Int, @DocID Int, @OnDate DateTime, @MC int, @pID int, @pKindID int

	AS
	SET NOCOUNT ON

	Declare @str nvarchar(255)

	Set @str = N'%' + @EntName + N'%' 

	---
	--- Создаем временную таблицу для выборки объектов в разрезе серий
	---
	CREATE TABLE #Ent(EntityID Int, Qty Money, Suma Money)

	INSERT INTO #Ent (EntityID,  Qty, Suma)
	select ent_id, 0, 0
	from entities 
	where ENTITIES.ENT_TYPE > 1000 and upper( ENTITIES.ENT_NAME) like upper(@str)

	--вставляем дебет (с плюсами)
	INSERT INTO #Ent (EntityID,  Qty, Suma)

	SELECT JOURNAL.J_ENT, 
		CASE 
			WHEN (ACC_DB=@AccID And J_AG1=@AgID) THEN IsNull(J_QTY, 0)
		ELSE 0
		END,

		CASE 
			WHEN (ACC_DB=@AccID And J_AG1=@AgID) THEN J_SUM
		ELSE 0
		END

	FROM JOURNAL
		left join entities as e on e.ent_id = journal.J_ENT 
	WHERE (JOURNAL.ACC_DB=@AccID) AND (JOURNAL.J_AG1=@AgID)  AND (JOURNAL.J_DONE=2) and JOURNAL.MC_ID = @MC
		 AND (JOURNAL.DOC_ID<>@DocID) AND (JOURNAL.J_DATE<@OnDate)
		 and E.ENT_TYPE > 1000 and upper( E.ENT_NAME) like upper(@str)

	--вставляем кредит (с минусами)
	INSERT INTO #Ent (EntityID,  Qty, Suma)

	SELECT JOURNAL.J_ENT, 
		CASE 
			WHEN (ACC_CR=@AccID And J_AG2=@AgID) THEN IsNull(-J_QTY, 0)
		ELSE 0
		END,

		CASE 
			WHEN (ACC_CR=@AccID And J_AG2=@AgID) THEN -J_SUM
		ELSE 0
		END

	FROM JOURNAL 
		left join entities as e on e.ent_id = journal.J_ENT 
	WHERE (JOURNAL.ACC_CR=@AccID) AND (JOURNAL.J_AG2=@AgID)  AND (JOURNAL.J_DONE=2) and JOURNAL.MC_ID = @MC
		 AND (JOURNAL.DOC_ID<>@DocID) AND (JOURNAL.J_DATE<@OnDate) 
		 and E.ENT_TYPE > 1000 and upper( E.ENT_NAME) like upper(@str)
	
	--Окончательная выборка
	SELECT top 21 ENTITIES.ENT_NAME AS EntityName, 
			(select top 1 p.prc_value from prc_contents as p where p.ent_id = EntityID and p.prl_id = @pID and p.prc_id = @pKindID and p.prc_date <= @OnDate order by p.prc_date desc) as EntityPrice,
			Sum(Qty) AS Qty, 
			Sum (Suma) AS Sum, #Ent.EntityID as EntID

	FROM #Ent LEFT JOIN  ENTITIES ON #Ent.EntityID= ENTITIES.ENT_ID
	GROUP BY #Ent.EntityID,  ENTITIES.ENT_NAME,  ENTITIES.ENT_TYPE
	---HAVING ENTITIES.ENT_TYPE > 1000 and upper( ENTITIES.ENT_NAME) like upper(@str)
	ORDER BY  Sum (Suma) desc, ENTITIES.ENT_NAME

GO

/****** Object:  StoredProcedure [dbo].[st7_oddments_of_the_goods]    Script Date: 23.02.2016 21:29:13 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[st7_oddments_of_the_goods]
      @AgID Int, @AccID Int, @DocID Int, @OnDate DateTime, @EntName nvarchar(255), @MC int

AS
SET NOCOUNT ON

Declare @str nvarchar(255)

Set @str = N'%' + @EntName + N'%' 

---
--- Создаем временную таблицу для выборки объектов в разрезе серий
---
CREATE TABLE #Ent(EntityID Int, Qty Money, Suma Money)

--вставляем дебет (с плюсами)
INSERT INTO #Ent (EntityID,  Qty, Suma)

SELECT JOURNAL.J_ENT, 
	CASE 
		WHEN (ACC_DB=@AccID And J_AG1=@AgID) THEN IsNull(J_QTY, 0)
	ELSE 0
	END,

	CASE 
		WHEN (ACC_DB=@AccID And J_AG1=@AgID) THEN J_SUM
	ELSE 0
	END

FROM JOURNAL 
WHERE (JOURNAL.ACC_DB=@AccID) AND (JOURNAL.J_AG1=@AgID)  AND (JOURNAL.J_DONE=2) and JOURNAL.MC_ID = @MC
	 AND (JOURNAL.DOC_ID<>@DocID) AND (JOURNAL.J_DATE<=@OnDate)

--вставляем кредит (с минусами)
INSERT INTO #Ent (EntityID,  Qty, Suma)

SELECT JOURNAL.J_ENT, 
	CASE 
		WHEN (ACC_CR=@AccID And J_AG2=@AgID) THEN IsNull(-J_QTY, 0)
	ELSE 0
	END,

	CASE 
		WHEN (ACC_CR=@AccID And J_AG2=@AgID) THEN -J_SUM
	ELSE 0
	END

FROM JOURNAL 
WHERE (JOURNAL.ACC_CR=@AccID) AND (JOURNAL.J_AG2=@AgID)  AND (JOURNAL.J_DONE=2) and JOURNAL.MC_ID = @MC
	 AND (JOURNAL.DOC_ID<>@DocID) AND (JOURNAL.J_DATE<=@OnDate) 
	 

--Окончательная выборка
SELECT ENTITIES.ENT_NAME AS EntityName, 
		Sum(Qty) AS Qty, Sum (Suma) AS Sum, #Ent.EntityID as EntID

FROM #Ent LEFT JOIN  ENTITIES ON #Ent.EntityID= ENTITIES.ENT_ID
where ENTITIES.ENT_TYPE > 1000 and upper(@str) like upper( ENTITIES.ENT_NAME)
GROUP BY #Ent.EntityID,  ENTITIES.ENT_NAME,  ENTITIES.ENT_TYPE
HAVING Sum(Qty)<>0 
ORDER BY  ENTITIES.ENT_NAME;
GO

/****** Object:  StoredProcedure [dbo].[st7_oddments_of_the_goods_price]    Script Date: 23.02.2016 21:29:13 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[st7_oddments_of_the_goods_price]
		  @EntName nvarchar(255), @AgID Int, @AccID Int, @DocID Int, @OnDate DateTime, @MC int, @pID int, @pKindID int

	AS
	SET NOCOUNT ON

	Declare @str nvarchar(255)

	Set @str = N'%' + @EntName + N'%' 

	---
	--- Создаем временную таблицу для выборки объектов в разрезе серий
	---
	CREATE TABLE #Ent(EntityID Int, Qty Money, Suma Money)

	--вставляем дебет (с плюсами)
	INSERT INTO #Ent (EntityID,  Qty, Suma)

	SELECT JOURNAL.J_ENT, 
		CASE 
			WHEN (ACC_DB=@AccID And J_AG1=@AgID) THEN IsNull(J_QTY, 0)
		ELSE 0
		END,

		CASE 
			WHEN (ACC_DB=@AccID And J_AG1=@AgID) THEN J_SUM
		ELSE 0
		END

	FROM JOURNAL 
	WHERE (JOURNAL.ACC_DB=@AccID) AND (JOURNAL.J_AG1=@AgID)  AND (JOURNAL.J_DONE=2) and JOURNAL.MC_ID = @MC
		 AND (JOURNAL.DOC_ID<>@DocID) AND (JOURNAL.J_DATE<@OnDate)

	--вставляем кредит (с минусами)
	INSERT INTO #Ent (EntityID,  Qty, Suma)

	SELECT JOURNAL.J_ENT, 
		CASE 
			WHEN (ACC_CR=@AccID And J_AG2=@AgID) THEN IsNull(-J_QTY, 0)
		ELSE 0
		END,

		CASE 
			WHEN (ACC_CR=@AccID And J_AG2=@AgID) THEN -J_SUM
		ELSE 0
		END

	FROM JOURNAL 
	WHERE (JOURNAL.ACC_CR=@AccID) AND (JOURNAL.J_AG2=@AgID)  AND (JOURNAL.J_DONE=2) and JOURNAL.MC_ID = @MC
		 AND (JOURNAL.DOC_ID<>@DocID) AND (JOURNAL.J_DATE<=@OnDate) 

	--Окончательная выборка
	SELECT ENTITIES.ENT_NAME AS EntityName, 
			(select top 1 p.prc_value from prc_contents as p where p.ent_id = EntityID and p.prl_id = @pID and p.prc_id = @pKindID and p.prc_date <= @OnDate order by p.prc_date desc) as EntityPrice,
			Sum(Qty) AS Qty, Sum (Suma) AS Sum, #Ent.EntityID as EntID

	FROM #Ent LEFT JOIN  ENTITIES ON #Ent.EntityID= ENTITIES.ENT_ID
	GROUP BY #Ent.EntityID,  ENTITIES.ENT_NAME,  ENTITIES.ENT_TYPE
	HAVING Sum(Qty)<>0 and  ENTITIES.ENT_TYPE > 1000 and upper( ENTITIES.ENT_NAME) like upper(@str)
	ORDER BY  ENTITIES.ENT_NAME
GO

/****** Object:  StoredProcedure [dbo].[HL74_RestEntSeries]    Script Date: 23.02.2016 21:29:13 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

create procedure [dbo].[HL74_RestEntSeries]
	@AgID as int, 
	@AccID as int, 
	@EntID as int, 
	@DocID as int, 
	@DocDate as DateTime, 
	@MC as int
as
	set nocount on 
	
	SELECT JOURNAL.Ser_ID,
			Sum(
			case 
				when ACC_DB=@AccID And J_AG1=@AgID then J_QTY
				when ACC_CR=@AccID And J_AG2=@AgID then -J_QTY
				else 0
			end),
			Sum(
			case 
				when ACC_DB=@AccID And J_AG1=@AgID then J_SUM
				when ACC_CR=@AccID And J_AG2=@AgID then -J_SUM
				else 0
			end)

	FROM ENTITIES INNER JOIN JOURNAL ON ENTITIES.ENT_ID = JOURNAL.J_ENT
	WHERE ((JOURNAL.ACC_DB=@AccID AND JOURNAL.J_AG1=@AgID) or (JOURNAL.ACC_CR=@AccID AND JOURNAL.J_AG2=@AgID))
		AND JOURNAL.J_DONE=2 
		AND JOURNAL.DOC_ID<>@DocID 
		AND JOURNAL.J_DATE<=@DocDate 
		AND JOURNAL.MC_ID=@MC 
		AND JOURNAL.J_ENT=@EntID
	GROUP BY JOURNAL.Ser_ID
	

GO


