create procedure HL74_CalcFirstEntSeriesCostPrice 
	@Acc28 int,
	@Acc63 int,
	@serID int,
	@OnDate datetime,
	@mc int

as
	set nocount on

	select id
	into #acc28
	from acc_tree
	where @Acc28 in (id, p0, p1, p2, p3, p4, p5, p6, p7)

	select id
	into #acc63
	from acc_tree
	where @acc63 in (id, p0, p1, p2, p3, p4, p5, p6, p7)

	select sum(j_sum), sum(j_qty)
	from journal
	where 
		j_done = 2
		and mc_id = @mc
		and acc_db in (select id from #acc28)
		and acc_cr in (select id from #acc63)
		and j_date < @OnDate 
		and ser_id = @serID

drop table #acc28
drop table #acc63