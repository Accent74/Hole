/****** Object:  StoredProcedure [dbo].[st7_oddments_of_the_goods]    Script Date: 16.05.2019 14:22:21 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[HL74_custombind_goods_oddments_twice_price]
		  @EntName nvarchar(255), 
		  @AgID Int, 
		  @AccID Int, 
		  @DocID Int, 
		  @OnDate DateTime, 
		  @MC int, 
		  @pID int, 
		  @pKindID int

	AS
	SET NOCOUNT ON

	Declare @str nvarchar(255)

	Set @str = N'%' + @EntName + N'%' 

	---
	--- Создаем временную таблицу для выборки объектов 
	---
	CREATE TABLE #Ent(EntID Int, EntQty Money, EntSum Money)

	--- добавляем все ОУ, подходящие по имени
	INSERT INTO #Ent (EntID,  EntQty, EntSum)
	select ent_id, 0, 0
	from entities 
	where ENTITIES.ENT_TYPE > 1000 and upper( ENTITIES.ENT_NAME) like upper(@str)

	--вставляем дебет (с плюсами)
	INSERT INTO #Ent (EntID,  EntQty, EntSum)

	SELECT JOURNAL.J_ENT, 
		Sum(CASE 
			WHEN (ACC_DB=@AccID And J_AG1=@AgID) THEN IsNull(J_QTY, 0)
		ELSE 0
		END),

		Sum(CASE 
			WHEN (ACC_DB=@AccID And J_AG1=@AgID) THEN J_SUM
		ELSE 0
		END)

	FROM JOURNAL
		left join entities as e on e.ent_id = journal.J_ENT 
	WHERE (JOURNAL.ACC_DB=@AccID) 
			AND (JOURNAL.J_AG1=@AgID)  
			AND (JOURNAL.J_DONE=2) 
			and JOURNAL.MC_ID = @MC
			AND (JOURNAL.DOC_ID<>@DocID) 
			AND (JOURNAL.J_DATE < @OnDate)
			and j_ent in (select EntID from #Ent)
	group by JOURNAL.J_ENT
	--вставляем кредит (с минусами)
	INSERT INTO #Ent (EntID,  EntQty, EntSum)

	SELECT JOURNAL.J_ENT, 
		Sum(CASE 
			WHEN (ACC_CR=@AccID And J_AG2=@AgID) THEN IsNull(-J_QTY, 0)
		ELSE 0
		END),

		Sum(CASE 
			WHEN (ACC_CR=@AccID And J_AG2=@AgID) THEN -J_SUM
		ELSE 0
		END)

	FROM JOURNAL 
		left join entities as e on e.ent_id = journal.J_ENT 
	WHERE (JOURNAL.ACC_CR=@AccID) 
			AND (JOURNAL.J_AG2=@AgID)  
			AND (JOURNAL.J_DONE=2) 
			and JOURNAL.MC_ID = @MC
			AND (JOURNAL.DOC_ID<>@DocID) 
			AND (JOURNAL.J_DATE<@OnDate) 
			and j_ent in (select EntID from #Ent)
	group by JOURNAL.J_ENT

	--Окончательная выборка
	SELECT top 21 ENTITIES.ENT_NAME, 
			Sum(EntQty),
			case 
				when Sum(EntQty) <> 0 then Sum(EntSum) / Sum(EntQty)
				else 0
			end,
			(select top 1 p.prc_value 
				from prc_contents as p 
				where p.ent_id = EntID 
					and p.prl_id = @pID 
					and p.prc_id = @pKindID 
					and p.prc_date <= @OnDate 
				order by p.prc_date desc) as EntityPrice,
			Sum (EntSum) AS Sum, 
			#Ent.EntID as EntID

	FROM #Ent LEFT JOIN  ENTITIES ON #Ent.EntID= ENTITIES.ENT_ID
	GROUP BY #Ent.EntID,  ENTITIES.ENT_NAME,  ENTITIES.ENT_TYPE
	HAVING Sum (EntSum) <> 0
	ORDER BY  Sum (EntSum) desc, ENTITIES.ENT_NAME

GRANT EXECUTE ON HL74_custombind_goods_oddments_twice_price TO ap_public
GO

/****** Object:  StoredProcedure [dbo].[HL74_custombind_goods_oddments_price]    Script Date: 16.05.2019 14:23:59 ******/
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
		 AND (JOURNAL.DOC_ID<>@DocID) AND (JOURNAL.J_DATE < @OnDate)
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

CREATE PROCEDURE [dbo].[st7_FindEntityByFirstChar] 
      @EntName nvarchar(255)
AS
SET NOCOUNT ON

select ent_id 
into #disable
from entities 
where ent_id in (select id from ent_tree where @EntRoot in (p0, p1, p2, p3, p4, p5, p6, p7))

declare @str nvarchar(255)
declare @ent_item nvarchar(255), @ent_index int

set @EntName = upper(@EntName)
create table #Ents(ent_id int, ent_name nvarchar(255))

declare ent_find cursor for
	select items, iindex from dbo.f_Split(@EntName, '%')

open ent_find 
FETCH NEXT FROM ent_find INTO @ent_item, @ent_index

WHILE @@FETCH_STATUS = 0
	BEGIN

		Select @str = '%' + @ent_item + '%'

		if @ent_index = 0
			insert into #Ents(ent_id, ent_name)
			select ent_id, ent_name
			from entities
			where upper(ent_name) like @str
				and ent_type > 1000
--				and ent_id not in (select ent_id from #disable)
			order by ent_name, ent_art desc
		else
			delete from #Ents
			where not upper(ent_name) like @str
			
		FETCH NEXT FROM ent_find INTO @ent_item, @ent_index
	end
	
CLOSE ent_find
DEALLOCATE ent_find

Select TOP 21 ent_name AS EntName
FROM #Ents 
ORDER BY ENT_NAME
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