USE [walmart_work]
GO

/****** Object:  StoredProcedure [dbo].[ST74_CRReport]    Script Date: 22.11.2017 17:11:37 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

alter procedure [dbo].[ST74_CRReport] 
	@AccID int,
	@dEnd datetime,
	@mc int
as
	set nocount on

	--- результирующая таблица
	create table #report (AgID int, TotalSum money, DocID int, DocDate datetime, DocNo nvarchar(50), DocSum money, PayDate datetime)

	--- выбрали кредиторку. только задолженности по покупателям
	select	j_ag1 as AgID, 
			sum(j_sum) as TotalSum
	into #JO6
	from journal j
	with (nolock) 
	where
		j_done = 2 and MC_ID = @mc 
		and j_date < @dEnd
		and acc_db = @AccID
	group by j_ag1
	
	insert into #JO6 (AgID, TotalSum)
	select	j_ag2, 
			-sum(j_sum)
	from 	journal j
	with (nolock) 
	where
		j_done = 2 and MC_ID = @mc 
		and j_date < @dEnd
		and acc_cr = @AccID
	group by j_ag2

	--- курсор по все записям дебиторки
	declare @AgID int, @AgSum Money
	declare DocList cursor for

		select #JO6.AgID, 
			-Sum(#JO6.TotalSum)
		from #JO6 with (nolock) 
		Group by #JO6.AgID
		Having Sum(#JO6.TotalSum) < 0

	open DocList

	fetch next From DocList into @AgID, @AgSum

	while @@FETCH_STATUS = 0
		begin
			declare @DocID int, 
					@DocDate datetime, 
					@DocNo nvarchar(50),
					@DocSum money,
					@PayDate datetime,
					@PaySum money = 0,
					@TotalSum money = 0

			--- выбрали все документы по покупателю до даты отчета
			declare AgDocList cursor for
				select j.doc_id, 
						doc_date, 
						doc_no, 
						sum(j.j_sum), 
						isnull(doc_pd3, doc_date) 
				from journal j
					left join documents d on d.doc_id = j.doc_id
				where j_ag1 = @AgID
					and acc_db = @AccID 
					and acc_cr <> @AccID
					and j_done = 2
					and j.mc_id = @mc 
					and d.doc_date < @dEnd 
				group by j.doc_id, doc_date, doc_no, isnull(doc_pd3, doc_date)
				having sum(j.j_sum) > 0
				order by isnull(doc_pd3, doc_date) desc, doc_no desc

			open AgDocList
			fetch next from AgDocList
			into @DocID, @DocDate, @DocNo, @DocSum, @PayDate

			--- макс сумма задолженности (для расчетов)
			set @TotalSum = @AgSum

			insert into #report (AgID, DocID, TotalSum) values (@AgID, 0, @AgSum)

			--- по всем документам от конца пока не выбрали всю сумму
			while @@FETCH_STATUS = 0 and @TotalSum > 0
				begin
					--- если остаток суммы < суммы по документу, то делаем сумму задолженности == остатку
					if @TotalSum < @DocSum
						set @DocSum = @TotalSum

					--- уменьшили сумму задолженности на сумму документа
					set @TotalSum = @TotalSum - @DocSum
					
					insert into #report (AgID, TotalSum, DocID, DocDate, DocNo, DocSum, PayDate) values
									(@AgID, 0, @DocID, @DocDate, @DocNo, @DocSum, @PayDate)
					fetch next from AgDocList
					into @DocID, @DocDate, @DocNo, @DocSum, @PayDate
				end

			close AgDocList
			deallocate AgDocList

			fetch next From DocList into @AgID, @AgSum
		end

	close DocList
	deallocate DocList

	select AgID, 
		Ag_Code, 
		Ag_NAME, 
		TotalSum, 
		DocID, 
		DocDate, 
		DocNo, 
		DocSum, 
		PayDate
	from #report
		left join AGENTS on ag_id = AgID
	order by Ag_Name, AgID, DocDate, DocID, DocNo, PayDate

	drop table #report, #JO6

GO

/****** Object:  StoredProcedure [dbo].[ST74_CRReportCurs]    Script Date: 22.11.2017 17:11:37 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

alter procedure [dbo].[ST74_CRReportCurs] 
	@AccID int,
	@CursID int,
	@dEnd datetime,
	@mc int
as
	set nocount on

	--- результирующая таблица
	create table #report (AgID int, TotalSum money, DocID int, DocDate datetime, DocNo nvarchar(50), DocSum money, PayDate datetime)

	--- выбрали кредиторку. только задолженности по поставщикам
	select	j_ag1 as AgID, 
			sum(JC_SUM) as TotalSum
	into #JO6
	from journal j with (nolock) 
		left join JRN_CRC jc on jc.J_ID = j.J_ID
	where
		j_done = 2 and MC_ID = @mc 
		and j_date < @dEnd
		and acc_db = @AccID
		and jc.CRC_ID = @CursID
	group by j_ag1
	
	insert into #JO6 (AgID, TotalSum)
	select	j_ag2, 
			-sum(JC_SUM)
	from 	journal j with (nolock) 
		left join JRN_CRC jc on jc.J_ID = j.J_ID
	where
		j_done = 2 and MC_ID = @mc 
		and j_date < @dEnd
		and acc_cr = @AccID
		and jc.CRC_ID = @CursID
	group by j_ag2

	--- курсор по все записям дебиторки
	declare @AgID int, @AgSum Money
	declare DocList cursor for

		select #JO6.AgID, 
			-Sum(#JO6.TotalSum)
		from #JO6 with (nolock) 
		Group by #JO6.AgID
		Having Sum(#JO6.TotalSum) < 0

	open DocList

	fetch next From DocList into @AgID, @AgSum

	while @@FETCH_STATUS = 0
		begin
			declare @DocID int, 
					@DocDate datetime, 
					@DocNo nvarchar(50),
					@DocSum money,
					@PayDate datetime,
					@PaySum money = 0,
					@TotalSum money = 0

			--- выбрали все документы по покупателю до даты отчета
			declare AgDocList cursor for
				select j.doc_id, 
						doc_date, 
						doc_no, 
						sum(jc.jc_sum), 
						isnull(doc_pd3, doc_date) 
				from journal j with (nolock)
					left join documents d on d.doc_id = j.doc_id
					left join JRN_CRC jc on jc.J_ID = j.J_ID
				where j_ag1 = @AgID
					and acc_db = @AccID 
					and acc_cr <> @AccID
					and j_done = 2
					and j.mc_id = @mc 
					and d.doc_date < @dEnd 
					and jc.CRC_ID = @CursID
				group by j.doc_id, doc_date, doc_no, isnull(doc_pd3, doc_date)
				having sum(jc.jc_sum) > 0
				order by isnull(doc_pd3, doc_date) desc, doc_no desc

			open AgDocList
			fetch next from AgDocList
			into @DocID, @DocDate, @DocNo, @DocSum, @PayDate

			--- макс сумма задолженности (для расчетов)
			set @TotalSum = @AgSum

			insert into #report (AgID, DocID, TotalSum) values (@AgID, 0, @AgSum)

			--- по всем документам от конца пока не выбрали всю сумму
			while @@FETCH_STATUS = 0 and @TotalSum > 0
				begin
					--- если остаток суммы < суммы по документу, то делаем сумму задолженности == остатку
					if @TotalSum < @DocSum
						set @DocSum = @TotalSum

					--- уменьшили сумму задолженности на сумму документа
					set @TotalSum = @TotalSum - @DocSum
					
					insert into #report (AgID, TotalSum, DocID, DocDate, DocNo, DocSum, PayDate) values
									(@AgID, 0, @DocID, @DocDate, @DocNo, @DocSum, @PayDate)
					fetch next from AgDocList
					into @DocID, @DocDate, @DocNo, @DocSum, @PayDate
				end

			close AgDocList
			deallocate AgDocList

			fetch next From DocList into @AgID, @AgSum
		end

	close DocList
	deallocate DocList

	select AgID, Ag_Code, Ag_NAME, TotalSum, DocID, DocDate, DocNo, DocSum, PayDate
	from #report
		left join AGENTS on ag_id = AgID
	order by Ag_Name, AgID, DocDate, DocID, DocNo, PayDate

	drop table #report, #JO6

GO

/****** Object:  StoredProcedure [dbo].[ST74_DBReport]    Script Date: 22.11.2017 17:11:37 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

alter procedure [dbo].[ST74_DBReport] 
	@AccID int,
	@dEnd datetime,
	@mc int
as
	set nocount on

	--- результирующая таблица
	create table #report (AgID int, TotalSum money, DocID int, DocDate datetime, DocNo nvarchar(50), DocSum money, PayDate datetime)

	--- выбрали кредиторку. только задолженности по покупателям
	select	j_ag1 as AgID, 
			sum(j_sum) as TotalSum
	into #JO6
	from journal j
	with (nolock) 
	where
		j_done = 2 and MC_ID = @mc 
		and j_date < @dEnd
		and acc_db = @AccID
	group by j_ag1
	
	insert into #JO6 (AgID, TotalSum)
	select	j_ag2, 
			-sum(j_sum)
	from 	journal j
	with (nolock) 
	where
		j_done = 2 and MC_ID = @mc 
		and j_date < @dEnd
		and acc_cr = @AccID
	group by j_ag2

	--- курсор по все записям дебиторки
	declare @AgID int, @AgSum Money
	declare DocList cursor for

		select #JO6.AgID, 
			Sum(#JO6.TotalSum)
		from #JO6 with (nolock) 
		Group by #JO6.AgID
		Having Sum(#JO6.TotalSum) > 0

	open DocList

	fetch next From DocList into @AgID, @AgSum

	while @@FETCH_STATUS = 0
		begin
			declare @DocID int, 
					@DocDate datetime, 
					@DocNo nvarchar(50),
					@DocSum money,
					@PayDate datetime,
					@PaySum money = 0,
					@TotalSum money = 0

			--- выбрали все документы по покупателю до даты отчета
			declare AgDocList cursor for
				select j.doc_id, 
						doc_date, 
						doc_no, 
						sum(j.j_sum), 
						isnull(doc_pd3, doc_date) 
				from journal j
					left join documents d on d.doc_id = j.doc_id
				where j_ag1 = @AgID
					and acc_db = @AccID 
					and acc_cr <> @AccID
					and j_done = 2
					and j.mc_id = @mc 
					and d.doc_date < @dEnd 
				group by j.doc_id, doc_date, doc_no, isnull(doc_pd3, doc_date)
				having sum(j.j_sum) > 0
				order by isnull(doc_pd3, doc_date) desc, doc_no desc

			open AgDocList
			fetch next from AgDocList
			into @DocID, @DocDate, @DocNo, @DocSum, @PayDate

			--- макс сумма задолженности (для расчетов)
			set @TotalSum = @AgSum

			insert into #report (AgID, DocID, TotalSum) values (@AgID, 0, @AgSum)

			--- по всем документам от конца пока не выбрали всю сумму
			while @@FETCH_STATUS = 0 and @TotalSum > 0
				begin
					--- если остаток суммы < суммы по документу, то делаем сумму задолженности == остатку
					if @TotalSum < @DocSum
						set @DocSum = @TotalSum

					--- уменьшили сумму задолженности на сумму документа
					set @TotalSum = @TotalSum - @DocSum
					
					insert into #report (AgID, TotalSum, DocID, DocDate, DocNo, DocSum, PayDate) values
									(@AgID, 0, @DocID, @DocDate, @DocNo, @DocSum, @PayDate)
					fetch next from AgDocList
					into @DocID, @DocDate, @DocNo, @DocSum, @PayDate
				end

			close AgDocList
			deallocate AgDocList

			fetch next From DocList into @AgID, @AgSum
		end

	close DocList
	deallocate DocList

	select AgID, Ag_Code, Ag_NAME, TotalSum, DocID, DocDate, DocNo, DocSum, PayDate
	from #report
		left join AGENTS on ag_id = AgID
	order by Ag_Name, AgID, DocDate, DocID, DocNo, PayDate

	drop table #report, #JO6

GO

/****** Object:  StoredProcedure [dbo].[ST74_DBReportCurs]    Script Date: 22.11.2017 17:11:37 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

alter procedure [dbo].[ST74_DBReportCurs] 
	@AccID int,
	@CursID int,
	@dEnd datetime,
	@mc int
as
	set nocount on

	--- результирующая таблица
	create table #report (AgID int, TotalSum money, DocID int, DocDate datetime, DocNo nvarchar(50), DocSum money, PayDate datetime)

	--- выбрали кредиторку. только задолженности по поставщикам
	select	j_ag1 as AgID, 
			sum(JC_SUM) as TotalSum
	into #JO6
	from journal j with (nolock) 
		left join JRN_CRC jc on jc.J_ID = j.J_ID
	where
		j_done = 2 and MC_ID = @mc 
		and j_date < @dEnd
		and acc_db = @AccID
		and jc.CRC_ID = @CursID
	group by j_ag1
	
	insert into #JO6 (AgID, TotalSum)
	select	j_ag2, 
			-sum(JC_SUM)
	from 	journal j with (nolock) 
		left join JRN_CRC jc on jc.J_ID = j.J_ID
	where
		j_done = 2 and MC_ID = @mc 
		and j_date < @dEnd
		and acc_cr = @AccID
		and jc.CRC_ID = @CursID
	group by j_ag2

	--- курсор по все записям дебиторки
	declare @AgID int, @AgSum Money
	declare DocList cursor for

		select #JO6.AgID, 
			Sum(#JO6.TotalSum)
		from #JO6 with (nolock) 
		Group by #JO6.AgID
		Having Sum(#JO6.TotalSum) > 0

	open DocList

	fetch next From DocList into @AgID, @AgSum

	while @@FETCH_STATUS = 0
		begin
			declare @DocID int, 
					@DocDate datetime, 
					@DocNo nvarchar(50),
					@DocSum money,
					@PayDate datetime,
					@PaySum money = 0,
					@TotalSum money = 0

			--- выбрали все документы по покупателю до даты отчета
			declare AgDocList cursor for
				select j.doc_id, 
						doc_date, 
						doc_no, 
						sum(jc.jc_sum), 
						isnull(doc_pd3, doc_date) 
				from journal j with (nolock)
					left join documents d on d.doc_id = j.doc_id
					left join JRN_CRC jc on jc.J_ID = j.J_ID
				where j_ag1 = @AgID
					and acc_db = @AccID 
					and acc_cr <> @AccID
					and j_done = 2
					and j.mc_id = @mc 
					and d.doc_date < @dEnd 
					and jc.CRC_ID = @CursID
				group by j.doc_id, doc_date, doc_no, isnull(doc_pd3, doc_date)
				having sum(jc.jc_sum) > 0
				order by isnull(doc_pd3, doc_date) desc, doc_no desc

			open AgDocList
			fetch next from AgDocList
			into @DocID, @DocDate, @DocNo, @DocSum, @PayDate

			--- макс сумма задолженности (для расчетов)
			set @TotalSum = @AgSum

			insert into #report (AgID, DocID, TotalSum) values (@AgID, 0, @AgSum)

			--- по всем документам от конца пока не выбрали всю сумму
			while @@FETCH_STATUS = 0 and @TotalSum > 0
				begin
					--- если остаток суммы < суммы по документу, то делаем сумму задолженности == остатку
					if @TotalSum < @DocSum
						set @DocSum = @TotalSum

					--- уменьшили сумму задолженности на сумму документа
					set @TotalSum = @TotalSum - @DocSum
					
					insert into #report (AgID, TotalSum, DocID, DocDate, DocNo, DocSum, PayDate) values
									(@AgID, 0, @DocID, @DocDate, @DocNo, @DocSum, @PayDate)
					fetch next from AgDocList
					into @DocID, @DocDate, @DocNo, @DocSum, @PayDate
				end

			close AgDocList
			deallocate AgDocList

			fetch next From DocList into @AgID, @AgSum
		end

	close DocList
	deallocate DocList

	select AgID, Ag_Code, Ag_NAME, TotalSum, DocID, DocDate, DocNo, DocSum, PayDate
	from #report
		left join AGENTS on ag_id = AgID
	order by Ag_Name, AgID, DocDate, DocID, DocNo, PayDate

	drop table #report, #JO6

GO


