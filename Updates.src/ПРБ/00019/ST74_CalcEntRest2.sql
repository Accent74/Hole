CREATE PROCEDURE [dbo].[ST74_CalcEntRest2]
		   @EntID Int,
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
		SELECT	SerID, EntID, 
				Sum(RestQty) AS RestQty, 
				Sum (RestSum) AS RestSum
		FROM #Ent e, ENT_TREE et
		WHERE et.ID = e.EntID AND @EntID in (p0,p1,p2,p3,p4,p5,p6,p7)
		group by SerID, EntID
		HAVING Sum(RestQty)<>0 or Sum(RestSum)<>0
		order by SerID, EntID

		drop table #Ent

GRANT EXECUTE ON [dbo].[ST74_CalcEntRest2] TO [ap_public]
