USE [StroyCity]
GO

/****** Object:  StoredProcedure [dbo].[cipher_add_ent_bar_code]    Script Date: 05/31/2013 05:37:54 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[cipher_add_ent_bar_code]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[cipher_add_ent_bar_code]
GO

/****** Object:  StoredProcedure [dbo].[cipher_ent_all]    Script Date: 05/31/2013 05:37:54 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[cipher_ent_all]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[cipher_ent_all]
GO

/****** Object:  StoredProcedure [dbo].[cipher_ent_bar_code_list]    Script Date: 05/31/2013 05:37:54 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[cipher_ent_bar_code_list]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[cipher_ent_bar_code_list]
GO

/****** Object:  StoredProcedure [dbo].[cipher_ent_rest]    Script Date: 05/31/2013 05:37:54 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[cipher_ent_rest]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[cipher_ent_rest]
GO

/****** Object:  StoredProcedure [dbo].[cipher_find_ent_by_bar_code]    Script Date: 05/31/2013 05:37:54 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[cipher_find_ent_by_bar_code]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[cipher_find_ent_by_bar_code]
GO

USE [StroyCity]
GO

/****** Object:  StoredProcedure [dbo].[cipher_add_ent_bar_code]    Script Date: 05/31/2013 05:37:54 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


CREATE procedure [dbo].[cipher_add_ent_bar_code]
	@EntID int, @BarCode nchar(13)
as
	begin
		set nocount on ;
		declare @find int
		set @find = (select count(*) from entities_bar_codes where ent_bar = @BarCode)
		
		if @find = 0
			begin
				insert into entities_bar_codes
				select @EntID, @BarCode
			
				select 0
			end
		else
			select 1
	end

GO

/****** Object:  StoredProcedure [dbo].[cipher_ent_all]    Script Date: 05/31/2013 05:37:54 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


CREATE procedure [dbo].[cipher_ent_all]
	@acc_root_id int,
	@ent_root_id int,
	@dEnd datetime,
	@MC int
as
	begin
		set nocount on ;
		
		create table #ent(ID int)
	
		if @ent_root_id = 0
			insert into #ent
			select e.ent_id as id
			from entities as e
				left join entities_bar_codes on entities_bar_codes.ent_id = e.ent_id
			where	ent_type > 1000
					and isnull(entities_bar_codes.ent_bar, '0') <> '0'
			group by e.ent_id
		else
			insert into #ent
			select id
			from ent_tree
				left join entities on entities.ent_id = id
				left join entities_bar_codes on entities_bar_codes.ent_id = id
			where	(p0 = @ent_root_id or
					p1 = @ent_root_id or
					p2 = @ent_root_id or
					p3 = @ent_root_id or
					p4 = @ent_root_id or
					p5 = @ent_root_id or
					p6 = @ent_root_id or
					p7 = @ent_root_id)
					and ent_type > 1000
					and isnull(entities_bar_codes.ent_bar, '0') <> '0'
			group by id

		select id
		into #acc
		from acc_tree
		where	p0 = @acc_root_id or
				p1 = @acc_root_id or
				p2 = @acc_root_id or
				p3 = @acc_root_id or
				p4 = @acc_root_id or
				p5 = @acc_root_id or
				p6 = @acc_root_id or
				p7 = @acc_root_id
		group by id
	
		insert into #acc
		select @acc_root_id

		select	j_ent as e_id,
				sum(j_qty) as e_qty,
				sum(j_sum) as e_sum
		into	#rest
		from	journal
		where	j_done = 2 and j_date < @dEnd and mc_id = @mc
				and acc_db in (select id from #acc where id = acc_db)
				and j_ent in (select id from #ent where id = j_ent)
				and isnull(j_ent, 0) <> 0
		group by j_ent

		insert into #rest
		select id, 0 , 0
		from #ent
		where id not in (select e_id 
							from #rest 
							where e_id = id)
			
		insert into #rest(e_id, e_qty, e_sum)
		select	j_ent as e_id,
				-sum(j_qty) as e_qty,
				-sum(j_sum) as e_sum
		from	journal
		where	j_done = 2 and j_date < @dEnd and mc_id = @mc
				and acc_cr in (select id from #acc where id = acc_cr)
				and j_ent in (select id from #ent where id = j_ent)
				and isnull(j_ent, 0) <> 0
		group by j_ent
			
		select	e_id,
				isnull(entities_bar_codes.ent_bar, '0'),
				isnull(ent_art, ''),
				ent_name,
				isnull(un_short, ''),
				sum(e_qty),
				Sum(e_sum)
		from	#rest
				left join entities_bar_codes on entities_bar_codes.ent_id = e_id
				left join entities on entities.ent_id = e_id
				left join units on entities.un_id = units.un_id
		where	isnull(entities_bar_codes.ent_bar, '0') <> '0'
		group by	e_id,
					entities_bar_codes.ent_bar,
					isnull(ent_art, ''),
					ent_name,
					isnull(un_short, '')
		order by	e_id 

		drop table #rest
		drop table #acc
		drop table #ent
	end
	

GO

/****** Object:  StoredProcedure [dbo].[cipher_ent_bar_code_list]    Script Date: 05/31/2013 05:37:54 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE procedure [dbo].[cipher_ent_bar_code_list]
	@entid int
as
begin
	set nocount on ;

	select ENT_BAR
	from ENTITIES_BAR_CODES
	where ENT_ID = @EntID
end
	

GO

/****** Object:  StoredProcedure [dbo].[cipher_ent_rest]    Script Date: 05/31/2013 05:37:54 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


CREATE procedure [dbo].[cipher_ent_rest]
	@acc_root_id int,
	@pList1 int,
	@pList2 int,
	@pKind1 int,
	@pKind2 int,
	@dEnd datetime,
	@MC int
as
	begin
		set nocount on ;
		
		select id
		into #acc
		from acc_tree
		where	p0 = @acc_root_id or
				p1 = @acc_root_id or
				p2 = @acc_root_id or
				p3 = @acc_root_id or
				p4 = @acc_root_id or
				p5 = @acc_root_id or
				p6 = @acc_root_id or
				p7 = @acc_root_id
		group by id
		
		insert into #acc
		select @acc_root_id

		select	j_ent as e_id,
				sum(j_qty) as e_qty,
				sum(j_sum) as e_sum
		into	#rest
		from	journal
		where	j_done = 2 and j_date < @dEnd and mc_id = @mc
				and acc_db in (select id from #acc where id = acc_db)
				and isnull(j_ent, 0) <> 0
		group by j_ent
				
		insert into #rest(e_id, e_qty, e_sum)
		select	j_ent as e_id,
				-sum(j_qty) as e_qty,
				-sum(j_sum) as e_sum
		from	journal
		where	j_done = 2 and j_date < @dEnd and mc_id = @mc
				and acc_cr in (select id from #acc where id = acc_cr)
				and isnull(j_ent, 0) <> 0
		group by j_ent
				
		select	e_id,
				isnull(entities_bar_codes.ent_bar, '0'),
				isnull(ent_art, ''),
				ent_name,
				isnull(un_short, ''),
				isnull(ent_memo, ''),
				sum(e_qty),
				isnull((select top 1 prc_value 
					from prc_contents 
					where ent_id = e_id 
							and prc_id = @pKind1
							and prl_id = @pList1
							and prc_date <= @dEnd
					order by prc_date desc), 0),
				isnull((select top 1 prc_value 
					from prc_contents 
					where ent_id = e_id 
							and prc_id = @pKind2
							and prl_id = @pList2
							and prc_date <= @dEnd
					order by prc_date desc), 0),
				Sum(e_sum)
					 
		from	#rest
				left join entities_bar_codes on entities_bar_codes.ent_id = e_id
				left join entities on entities.ent_id = e_id
				left join units on entities.un_id = units.un_id
		where	isnull(entities_bar_codes.ent_bar, '0') <> '0'
		group by	e_id,
					entities_bar_codes.ent_bar,
					isnull(ent_art, ''),
					isnull(ent_memo, ''),
					ent_name,
					isnull(un_short, '')
		having sum(e_qty) > 0
		order by	e_id 

		drop table #rest
		drop table #acc
	end
	

GO

/****** Object:  StoredProcedure [dbo].[cipher_find_ent_by_bar_code]    Script Date: 05/31/2013 05:37:54 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


CREATE procedure [dbo].[cipher_find_ent_by_bar_code]
	@BarCode nchar(13)
as
	set nocount on ;
	select ent_id from entities_bar_codes where ent_bar = @BarCode


GO

