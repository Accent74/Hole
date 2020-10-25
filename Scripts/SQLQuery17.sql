declare @AccBID int = 220,
		@AgBID int = 1396,
		@AgSID int = 1,
		@AccSID int= 741,
		@OnDate datetime = '20190101',
		@mc int = 1


declare @EntID int = 10643

select	j_date,
		doc_id,
		sum(case 
			when (j.acc_db = @AccBID and j.j_ag1 = @AgBID) then j.j_qty
			else 0
		end -
		case 
			when (j.acc_cr = @AccBID and j.j_ag2 = @AgBID) then j.j_qty
			else 0
		end) eQty,
		sum(case 
			when (j.acc_db = @AccBID and j.j_ag1 = @AgBID) then j.j_sum
			else 0
		end -
		case 
			when (j.acc_cr = @AccBID and j.j_ag2 = @AgBID) then j.j_sum
			else 0
		end) eSum
into #res
from journal j
where
		j.J_DONE = 2
	and j.MC_ID = @mc
	and j.j_date <= @OnDate
	and ((j.acc_db = @AccBID and j.j_ag1 = @AgBID) or (j.acc_cr = @AccBID and j.j_ag2 = @AgBID))
	and j.j_ent = @EntID
group by 
	j.J_DATE,
	j.doc_id
order by 
	j.J_DATE desc


select	ser_id,
		j_date,
		sum(case 
			when (j.acc_cr = @AccSID and j.j_ag2 = @AgSID) then j.j_qty
			else 0
		end -
		case 
			when (j.acc_db = @AccSID and j.j_ag1 = @AgSID) then j.j_qty
			else 0
		end) eQty,
		sum(case 
			when (j.acc_cr = @AccSID and j.j_ag2 = @AgSID) then j.j_sum
			else 0
		end -
		case 
			when (j.acc_db = @AccSID and j.j_ag1 = @AgSID) then j.j_sum
			else 0
		end) eSum
from journal j
where
		j.J_DONE = 2
	and j.MC_ID = @mc
	and j.j_date <= @OnDate
	and ((j.acc_cr = @AccSID and j.j_ag2 = @AgSID) or (j.acc_db = @AccSID and j.j_ag1 = @AgSID))
	and j.j_ent = @EntID
	and j.DOC_ID in (select #res.doc_id from #res)
group by 
	ser_id,
	j_date
order by j_date desc


	
select * from #res
drop table #res