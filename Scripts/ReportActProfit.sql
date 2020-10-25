USE [walmart_work]
GO
/****** Object:  StoredProcedure [dbo].[HL74_ReportProfitAct]    Script Date: 25.04.2020 9:45:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
ALTER procedure [dbo].[HL74_ReportProfitAct]
	@acc36 int = 220,
	@acc70 int = 333,
	@acc28 int = 194,
	@MscNo int = 30,
	@dStart datetime = '20200501',
	@dEnd datetime = '20200601',
	@mc int = 1
as
begin
	set nocount on ;

	select distinct id into #Acc36 from acc_tree where @acc36 in (id, p0, p1, p2, p3, p4, p5, p6, p7)
	select distinct id into #Acc70 from acc_tree where @acc70 in (id, p0, p1, p2, p3, p4, p5, p6, p7)
	select distinct id into #Acc28 from acc_tree where @acc28 in (id, p0, p1, p2, p3, p4, p5, p6, p7)

	select
		j.j_ag1 AgID,
		j.j_ent EntID,
		j.doc_id DocID,
		sum(j.j_qty) EntQty,
		sum(j.j_sum) EntSum,
		isnull(jm.msc_id, 0) MscID,
		CONVERT(money, 0) EntStorageSum,
		CONVERT(money, 0) EntSeriesSum,
		(select sum(j2.j_sum) 
			from journal j2
			where doc_id = j.doc_Id
				and j2.acc_db = 865 
				and j2.j_ent = j.j_ent) SumSellCost
	into #res
	from
		journal j
			left join JRN_MISC jm on jm.J_ID = j.j_id and jm.msc_no = @MscNo
	where
		j.j_done = 2
		and j.mc_id = @mc
		and j.j_date >= @dStart and j.j_date < @dEnd
		and j.acc_Db in (select id from #Acc36)
		and j.acc_cr in (select id from #Acc70)
	group by
		j.j_ag1,
		j.j_ent,
		j.doc_id,
		isnull(jm.msc_id, 0)

	update #res
	set EntStorageSum = (select sum(j_sum) 
							from journal 
							where 
								doc_id = DocID 
								and j_Ent = EntID
								and acc_cr in (select id from #Acc28)),
		EntSEriesSum = (select	sum(s.ser_cy1 * j.j_qty)
							from journal j
								left join series s on s.SER_ID = j.SER_ID
							where 
								doc_id = DocID 
								and j_Ent = EntID
								and j.acc_cr in (select id from #Acc28)
								and j.DOC_ID in (select DocID from #res))

	--- Корреспондент, ОУ, Документ, Кол-во, Цена прихода,	Доп. Затраты, Сумма прихода, Цена реализации, Договор
	
	select	isnull(AgID, 0),
			isnull(EntID, 0),
			isnull(Ag_Name, '< Не указан >'),
			isnull(Ent_NAme, '< Не указан >'),
			Doc_Date,
			Doc_No,
			EntQty,
			case
				when EntQty > 0 then EntSeriesSum / EntQty
				else 0
			end,
			EntStorageSum - EntSeriesSum,
			EntStorageSum, 
			case 
				when EntQty > 0 then EntSum / EntQty
				else 0
			end,
			EntSum,
			SumSellCost,
			EntSum - EntStorageSum,
			case 
				when EntSum <> 0 then (1 - EntStorageSum / EntSum) * 100
				else 0
			end,
			msc_name,
			DocID
	from #res
		left join DOCUMENTS on doc_id = DocID
		left join agents on ag_id = AgID
		left join entities on ent_id = EntID
		left join misc on msc_id = MscID
	order by
		isnull(Ag_Name, '< Не указан >'),
		isnull(AgID, 0),
		isnull(Ent_NAme, '< Не указан >'),
		isnull(EntID, 0),
		Doc_Date,
		doc_no

	drop table #Acc36, #Acc70, #Acc28, #res

end

