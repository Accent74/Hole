USE [FoodsnGoods]
GO

/****** Object:  StoredProcedure [dbo].[_Export_ClientsForManagers2]    Script Date: 26.11.2017 20:35:32 ******/
SET ANSI_NULLS OFF
GO

SET QUOTED_IDENTIFIER ON
GO

--
-- Процедура расчета остатков
--
CREATE PROCEDURE [dbo].[_Export_ClientsForManagers2]
      @ManagerID Int
AS
SET NOCOUNT ON
SET DATEFORMAT dmy

declare @MscFldID int = 6368

select m.MSC_LNG1 as AgID
into #clients
from MISC m
where msc_id in (select id from MISC_TREE where @MscFldID in (p0, p1, p2, p3, p4, p5, p6, p7))

SELECT	AGENTS.AG_ID, 
		AG_NAME, 
		AG_ADDRESS,
		AG_GUID
from AGENTS 
	inner join AG_PARAMS on AGENTS.AG_ID = AG_PARAMS.AG_ID
	inner join AG_PARAM_NAMES on AG_PARAMS.PRM_ID = AG_PARAM_NAMES.PRM_ID

where (AG_TYPE = 1 or AG_TYPE = 4) 
	and PRM_NAME = 'Менеджер' 
	and PRM_LONG = @ManagerID
	and AGENTS.ag_id in (select AgID from #clients)
order by AGENTS.AG_NAME

drop table #clients

GO

/****** Object:  StoredProcedure [dbo].[_Export_RestsForManagers]    Script Date: 26.11.2017 20:35:32 ******/
SET ANSI_NULLS OFF
GO

SET QUOTED_IDENTIFIER ON
GO

--
-- Процедура расчета остатков
--
ALTER PROCEDURE [dbo].[_Export_RestsForManagers]
      @RootAccID Int, @OnDate DateTime
AS
SET NOCOUNT ON
SET DATEFORMAT dmy

---
--- Создаем временные таблицы
---
CREATE TABLE #Accs(AccID Int)
INSERT INTO #Accs(AccID)
SELECT DISTINCT ID
FROM ACC_TREE with (nolock)
WHERE (ID = @RootAccID) Or (P0=@RootAccID) or (P1=@RootAccID) or (P2=@RootAccID) or (P3=@RootAccID) or (P4=@RootAccID) or (P5=@RootAccID) or (P6=@RootAccID) or (P7=@RootAccID)
CREATE TABLE #Leaves(EntID Int, AccID Int, Qty Money)
--Формируем остатки
INSERT INTO #Leaves(EntID, AccID, Qty)
   SELECT J_ENT, AccID, Sum(QTY) FROM
   (
   SELECT J_ENT, ACC_DB As AccID, J_QTY As Qty
   FROM JOURNAL with (nolock) INNER JOIN #Accs ON JOURNAL.ACC_DB = #Accs.AccID
   WHERE (J_DONE=2) AND (J_DATE <= @OnDate)
   UNION ALL
   SELECT J_ENT, ACC_CR As AccID, -J_QTY As Qty
   FROM JOURNAL with (nolock) INNER JOIN #Accs ON JOURNAL.ACC_CR = #Accs.AccID
   WHERE (J_DONE=2) AND (J_DATE <= @OnDate)
   ) As #Turns
   GROUP BY J_ENT, AccID
   HAVING Sum(QTY) <> 0   

SELECT  e.ENT_ID, 
		e.ENT_GUID, 
		e.ENT_TYPE, 
		e.ENT_NAME, 
		e.ENT_TAG, 
		e.ENT_CAT, 
		UN_SHORT, 
		#Leaves.ACCID, 
		#Leaves.Qty,
	(select top 1 prc_value from PRC_CONTENTS with (nolock) where PRC_CONTENTS.ENT_ID = #Leaves.EntID and PRC_ID = 5 and PRL_ID = 8 and PRC_DATE <= @OnDate order by PRC_DATE desc) As Price,
		ep.ent_id,
		ep.ent_name
FROM #Leaves
	INNER JOIN ENTITIES e ON #Leaves.ENTID = E.ENT_ID
	left join UNITS on E.UN_ID = UNITS.UN_ID
	left join ent_tree on id = #Leaves.EntID
	left join entities ep on ep.ent_id = ent_tree.P0
order by ep.ent_name, ep.ent_id, e.ent_name, e.ent_id 

drop table #Leaves
drop table #Accs

GO


