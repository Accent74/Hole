---
--- модифицированные скрипты для работы с системным протоколом 
---

ALTER Procedure [dbo].[ap_param_deletename]
@et int,
@id int
as
set nocount on
begin tran
declare @tbl nvarchar(64)
declare @sql nvarchar(255)
declare @prm nvarchar(255)
 
exec apx_et_paramname @et, @tbl OUT
if @tbl is null
  begin
    rollback tran
    raiserror (N'Недопустимый тип параметра',16,-1)
    return 0
  end
  
select @sql = N'DELETE FROM ' + 
	      @tbl + 
	      N' WHERE PRM_ID = @id'

select @prm = N'@id int'

execute sp_executesql @sql, @prm, @id

declare @itemName nvarchar(1)
set @itemName = (case
					when @tbl = 'fld_param_names' then 'Z'
					when @tbl = 'acc_param_names' then 'U'
					when @tbl = 'ag_param_names' then 'N'
					when @tbl = 'ent_param_names' then 'Y'
					when @tbl = 'msc_param_names' then 'I'
					when @tbl = 'bind_param_names' then 'V'
					when @tbl = 'tml_param_names' then 'P'
					when @tbl = 'rcp_param_names' then 'C'
					when @tbl = 'db_param_names' then 'S'
					when @tbl = 'ser_param_names' then 'H'
					when @tbl = 'frm_param_names' then 'Q'
					when @tbl = 'doc_param_names' then 'K'
					when @tbl = 'jrn_param_names' then 'J'
					when @tbl = 'rpe_param_names' then 'W'
				end)
exec ap_log_add @itemName,N'D',@id


commit tran



GO

/****** Object:  StoredProcedure [dbo].[ap_param_rename]    Script Date: 23.01.2019 14:40:17 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO


---------------------------------------------------------------------
ALTER Procedure [dbo].[ap_param_rename]
@et int,
@id int,
@name nvarchar(256)
as
set nocount on
begin tran
declare @tbl nvarchar(64)
declare @sql nvarchar(255)
declare @prm nvarchar(255)
 
exec apx_et_paramname @et, @tbl OUT
if @tbl is null
  begin
    rollback tran
    raiserror (N'Недопустимый тип элемента',16,-1)
    return 0
  end
  
select @sql = N'UPDATE ' + 
	      @tbl + 
              N' SET PRM_NAME = @name' +
	      N' WHERE PRM_ID = @id'

select @prm = N'@id int, @name nvarchar(256)'

execute sp_executesql @sql, @prm, @id, @name

declare @itemName nvarchar(1)
set @itemName = (case
					when @tbl = 'fld_param_names' then 'Z'
					when @tbl = 'acc_param_names' then 'U'
					when @tbl = 'ag_param_names' then 'N'
					when @tbl = 'ent_param_names' then 'Y'
					when @tbl = 'msc_param_names' then 'I'
					when @tbl = 'bind_param_names' then 'V'
					when @tbl = 'tml_param_names' then 'P'
					when @tbl = 'rcp_param_names' then 'C'
					when @tbl = 'db_param_names' then 'S'
					when @tbl = 'ser_param_names' then 'H'
					when @tbl = 'frm_param_names' then 'Q'
					when @tbl = 'doc_param_names' then 'K'
					when @tbl = 'jrn_param_names' then 'J'
					when @tbl = 'rpe_param_names' then 'W'
				end)
exec ap_log_add @itemName,N'N',@id

commit tran


GO

/****** Object:  StoredProcedure [dbo].[ap_param_addname]    Script Date: 23.01.2019 14:40:17 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-----------------------------------------
ALTER Procedure [dbo].[ap_param_addname]
@et int,
@name nvarchar(256),
@type smallint,
@node smallint,
@ref int,
@refid int,
@descr nvarchar(256) = null
as
set nocount on
begin tran
if @ref = 0
  select @ref = null
if @refid = 0
  select @refid = null
declare @tbl nvarchar(64)
declare @sql nvarchar(255)
declare @rv int
declare @prm nvarchar(255)
 
exec apx_et_paramname @et, @tbl OUT
if @tbl is null
  begin
    rollback tran
    raiserror (N'Недопустимый тип параметра',16,-1)
    return 0
  end
  
select @sql = N'INSERT INTO ' + 
	      @tbl + 
	      N' (PRM_NAME,PRM_TYPE,NODE_TYPE,PRM_REF,PRM_REFID,PRM_DESCR) VALUES (@name,@type,@node,@ref,@refid,@descr)'

select @prm = N'@name nvarchar(255), @type smallint, @node smallint, @ref int, @refid int, @descr nvarchar(256)'

execute sp_executesql @sql, @prm, @name, @type, @node, @ref, @refid, @descr

select @rv = ident_current(@tbl)

declare @itemName nvarchar(1)
set @itemName = (case
					when @tbl = 'fld_param_names' then 'Z'
					when @tbl = 'acc_param_names' then 'U'
					when @tbl = 'ag_param_names' then 'N'
					when @tbl = 'ent_param_names' then 'Y'
					when @tbl = 'msc_param_names' then 'I'
					when @tbl = 'bind_param_names' then 'V'
					when @tbl = 'tml_param_names' then 'P'
					when @tbl = 'rcp_param_names' then 'C'
					when @tbl = 'db_param_names' then 'S'
					when @tbl = 'ser_param_names' then 'H'
					when @tbl = 'frm_param_names' then 'Q'
					when @tbl = 'doc_param_names' then 'K'
					when @tbl = 'jrn_param_names' then 'J'
					when @tbl = 'rpe_param_names' then 'W'
				end)
exec ap_log_add @itemName,N'I',@rv

commit tran

return @rv
GO

/****** Object:  StoredProcedure [dbo].[ap_template_rename]    Script Date: 23.01.2019 14:40:17 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO


---------------------------------------------------------------------
ALTER Procedure [dbo].[ap_template_rename] 
@id int, @name nvarchar(256)
as
set nocount on
begin tran
update TEMPLATES set TML_NAME=@name where TML_ID=@id
exec ap_log_add N'T',N'N',@id
commit tran



GO

/****** Object:  StoredProcedure [dbo].[ap_misc_rename]    Script Date: 23.01.2019 14:40:17 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO


---------------------------------------------------------------------
ALTER Procedure [dbo].[ap_misc_rename] 
@id int, @name nvarchar(256)
as
set nocount on
begin tran
update MISC set MSC_NAME=@name where MSC_ID=@id
exec ap_log_add N'M',N'N',@id
commit tran



GO

/****** Object:  StoredProcedure [dbo].[ap_folder_rename]    Script Date: 23.01.2019 14:40:17 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO


---------------------------------------------------------------------
-- переименование папки
---------------------------------------------------------------------
ALTER Procedure [dbo].[ap_folder_rename] 
@id int, @name nvarchar(256)
as
set nocount on
begin tran
update FOLDERS set FLD_NAME=@name where FLD_ID=@id
exec ap_log_add N'F',N'N',@id
commit tran



GO

/****** Object:  StoredProcedure [dbo].[ap_entity_rename]    Script Date: 23.01.2019 14:40:17 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO


----------------------------------------
ALTER Procedure [dbo].[ap_entity_rename] 
@id int, 
@name nvarchar(256)
as
set nocount on
begin tran
update ENTITIES set ENT_NAME=@name where ENT_ID=@id
exec ap_log_add N'E',N'N',@id
commit tran


GO

/****** Object:  StoredProcedure [dbo].[ap_binder_rename]    Script Date: 23.01.2019 14:40:17 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO


---------------------------------------------------------------------
ALTER Procedure [dbo].[ap_binder_rename] 
@id int, @name nvarchar(256)
as
set nocount on
begin tran
update BINDERS set BIND_NAME=@name where BIND_ID=@id
exec ap_log_add N'B',N'N',@id

commit tran



GO

/****** Object:  StoredProcedure [dbo].[ap_agent_rename]    Script Date: 23.01.2019 14:40:17 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO


----------------------------------------
ALTER Procedure [dbo].[ap_agent_rename] 
@id int, @name nvarchar(256)
as
set nocount on
begin tran
update AGENTS set AG_NAME=@name where AG_ID=@id
exec ap_log_add N'G',N'N',@id
commit tran


GO

/****** Object:  StoredProcedure [dbo].[ap_account_rename]    Script Date: 23.01.2019 14:40:17 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO


---------------------------------------------------------------------
ALTER Procedure [dbo].[ap_account_rename] 
@id int, @name nvarchar(256)
as
set nocount on
begin tran
update ACCOUNTS set ACC_NAME=@name where ACC_ID=@id
exec ap_log_add N'A',N'N',@id
commit tran


GO

/****** Object:  StoredProcedure [dbo].[ap_template_savescript]    Script Date: 23.01.2019 14:40:17 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO


---------------------------------------------------------------------
ALTER Procedure [dbo].[ap_template_savescript]
@id int,
@text ntext,
@form int,
@saveform smallint
as
set nocount on
begin tran
if @saveform = 1 -- save form id
  update TEMPLATES set TML_SCRIPT=@text, FRM_ID=@form where TML_ID=@id
else
  update TEMPLATES set TML_SCRIPT=@text where TML_ID=@id  
commit tran

exec ap_log_add N'T',N'U',@id
GO

/****** Object:  StoredProcedure [dbo].[ap_template_save]    Script Date: 23.01.2019 14:40:17 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

---------------------------------------------------------------------
ALTER Procedure [dbo].[ap_template_save]
@id int,
@name nvarchar(256),
@tag nvarchar(51) = null,
@memo nvarchar(256) = null,
@frm int = null,
@hidden bit = 0
as
set nocount on

begin transaction
update TEMPLATES set 
  TML_NAME=@name, TML_TAG=@tag, TML_MEMO=@memo, FRM_ID=@frm, TML_HIDDEN=@hidden
where TML_ID=@id
exec ap_log_add N'T',N'U',@id
commit 
GO

/****** Object:  StoredProcedure [dbo].[ap_mscattr_save]    Script Date: 23.01.2019 14:40:17 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO


----------------------------------------
ALTER Procedure [dbo].[ap_mscattr_save]
@no int,
@vis int = 0,
@s1 nvarchar(51) = null,
@s2 nvarchar(51) = null,
@s3 nvarchar(51) = null,
@l1 nvarchar(51) = null,
@l2 nvarchar(51) = null,
@l3 nvarchar(51) = null,
@t1 int = null,
@t2 int = null,
@t3 int = null,
@c1 nvarchar(51) = null,
@c2 nvarchar(51) = null,
@c3 nvarchar(51) = null,
@d1 nvarchar(51) = null,
@d2 nvarchar(51) = null,
@d3 nvarchar(51) = null,
@reset int = 0,
@flags int = 0
as
set nocount on
begin tran
update MISC_ATTR set
  MSC_VISIBLE=@vis,MS1_NAME=@s1, MS2_NAME=@s2, MS3_NAME=@s3,
  ML1_NAME=@l1, ML2_NAME=@l2, ML3_NAME=@l3, ML1_TYPE=@t1, ML2_TYPE=@t2, ML3_TYPE=@t3,
  MC1_NAME=@c1, MC2_NAME=@c2, MC3_NAME=@c3, MD1_NAME=@d1, MD2_NAME=@d2, MD3_NAME=@d3,
  MSC_FLAGS=@flags
where MSC_NO=@no
if @reset & 0x00000001 <> 0
  begin
     -- сбросить значения LONG1
     update MISC set MSC_LNG1 = null where MSC_NO=@no
  end
if @reset & 0x00000002 <> 0
  begin
     -- сбросить значения LONG2
     update MISC set MSC_LNG2 = null where MSC_NO=@no
  end
if @reset & 0x00000004 <> 0
  begin
     -- сбросить значения LONG3
     update MISC set MSC_LNG3 = null where MSC_NO=@no
  end

 commit tran


GO

/****** Object:  StoredProcedure [dbo].[ap_misc_save]    Script Date: 23.01.2019 14:40:17 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO


----------------------------------------
ALTER Procedure [dbo].[ap_misc_save]
@id int,
@name nvarchar(256),
@tag nvarchar(51) = null,
@memo nvarchar(256) = null,
@s1 nvarchar(256) = null,
@s2 nvarchar(256) = null,
@s3 nvarchar(256) = null,
@l1 int = null,
@l2 int = null,
@l3 int = null,
@c1 money = null,
@c2 money = null,
@c3 money = null,
@d1 datetime = null,
@d2 datetime = null,
@d3 datetime = null
as
set nocount on
begin tran
update MISC set 
  MSC_NAME=@name,MSC_TAG=@tag,MSC_MEMO=@memo,
  MSC_STR1=@s1,MSC_STR2=@s2,MSC_STR3=@s3, 
  MSC_LNG1=@l1,MSC_LNG2=@l2,MSC_LNG3=@l3, 
  MSC_CY1=@c1, MSC_CY2=@c2, MSC_CY3=@c3, 
  MSC_DT1=@d1, MSC_DT2=@d2, MSC_DT3=@d3 
where MSC_ID=@id

exec ap_log_add N'M',N'U',@id

commit tran


GO

/****** Object:  StoredProcedure [dbo].[ap_folder_save]    Script Date: 23.01.2019 14:40:17 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO


---------------------------------------------------------------------
ALTER Procedure [dbo].[ap_folder_save]
@id int,
@name nvarchar(256),
@tag nvarchar(256),
@memo nvarchar(256),
@tml int = 0,
@frm int = 0
as
set nocount on
begin transaction
if @tml = 0
  select @tml = null
if @frm = 0
  select @frm = null
update FOLDERS set FLD_NAME=@name, FLD_TAG=@tag, FLD_MEMO=@memo, TML_ID=@tml, FRM_ID=@frm
where FLD_ID=@id

exec ap_log_add N'F',N'U',@id
commit transaction


GO

/****** Object:  StoredProcedure [dbo].[ap_binder_save]    Script Date: 23.01.2019 14:40:17 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO


--------------------------------------------
ALTER Procedure [dbo].[ap_binder_save]
@id int,
@name nvarchar(256),
@tag nvarchar(51) = null,
@memo nvarchar(256) = null

as

set nocount on

begin tran
update BINDERS set BIND_NAME=@name,BIND_TAG=@tag,BIND_MEMO=@memo
where BIND_ID=@id
exec ap_log_add N'B',N'U',@id
commit tran


GO

/****** Object:  StoredProcedure [dbo].[ap_template_createnew]    Script Date: 23.01.2019 14:40:17 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-----------------------------------------
ALTER Procedure [dbo].[ap_template_createnew] 
@name nvarchar(255),
@type smallint,
@parent int,
@guid uniqueidentifier = null
as

set nocount on

begin tran

declare @id int
declare @p0 int, @p1 int, @p2 int, @p3 int, @p4 int, @p5 int, @p6 int, @sc bit

if @guid is null
begin
  insert into TEMPLATES (TML_NAME,TML_TYPE) VALUES (@name,@type)
  select @id = ident_current('TEMPLATES')
end
else
begin
  insert into TEMPLATES (TML_NAME,TML_TYPE,TML_GUID) VALUES (@name,@type,@guid)
  select @id = ident_current('TEMPLATES')
end

select P0,P1,P2,P3,P4,P5,P6,P7,SHORTCUT into #tmptmltree
from TML_TREE where ID=@parent

if 0 = (select count(*) from #tmptmltree)
   insert into #tmptmltree (P0,P1,P2,P3,P4,P5,P6,P7,SHORTCUT) values (@parent,0,0,0,0,0,0,0,0)

declare #tmpcrs cursor local fast_forward read_only for
  select P0,P1,P2,P3,P4,P5,P6,SHORTCUT from #tmptmltree

if 0 <> (select count(*) from #tmptmltree where P7 <> 0)
   begin
      rollback tran
      raiserror ('Слишком большой уровень вложенности',16,-1)
      return 0
   end   

open #tmpcrs
fetch next from #tmpcrs into @p0,@p1,@p2,@p3,@p4,@p5,@p6,@sc
while @@fetch_status = 0
  begin
     -- insert into tml_tree
     insert into TML_TREE (ID,P0,P1,P2,P3,P4,P5,P6,P7,SHORTCUT) values (@id,@parent,@p0,@p1,@p2,@p3,@p4,@p5,@p6,@sc)
     fetch next from #tmpcrs into @p0,@p1,@p2,@p3,@p4,@p5,@p6,@sc
  end
close #tmpcrs
deallocate #tmpcrs

exec ap_log_add N'T',N'I',@id
commit 
return @id
GO

/****** Object:  StoredProcedure [dbo].[ap_misc_createnew]    Script Date: 23.01.2019 14:40:17 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-----------------------------------------
ALTER Procedure [dbo].[ap_misc_createnew] 
@no int, -- номер аналитики
@name nvarchar(255),
@type smallint,
@parent int,
@guid uniqueidentifier = null
as

set nocount on

begin tran

declare @id int
declare @p0 int, @p1 int, @p2 int, @p3 int, @p4 int, @p5 int, @p6 int, @sc bit

if @type = -1
  begin
     -- добавим запись в MISC_ATTR    
    insert into MISC_ATTR (MSC_VISIBLE) values (0)
    select @no = ident_current('MISC_ATTR')
  end

if @guid is null
begin
  insert into MISC (MSC_NO,MSC_NAME,MSC_TYPE) VALUES (@no, @name,@type)
  select @id = ident_current('MISC')
end
else
begin
  insert into MISC (MSC_NO,MSC_NAME,MSC_TYPE,MSC_GUID) VALUES (@no, @name,@type,@guid)
  select @id = ident_current('MISC')
end

select P0,P1,P2,P3,P4,P5,P6,P7,SHORTCUT into #tmpmsctree
from MISC_TREE where ID=@parent

if 0 = (select count(*) from #tmpmsctree)
   insert into #tmpmsctree (P0,P1,P2,P3,P4,P5,P6,P7,SHORTCUT) values (@parent,0,0,0,0,0,0,0,0)

declare #tmpcrs cursor local fast_forward read_only for
  select P0,P1,P2,P3,P4,P5,P6,SHORTCUT from #tmpmsctree

if exists(select * from #tmpmsctree where P7 <> 0)
   begin
      rollback tran
      raiserror ('Слишком большой уровень вложенности',16,-1)
      return 0
   end   

open #tmpcrs
fetch next from #tmpcrs into @p0,@p1,@p2,@p3,@p4,@p5,@p6,@sc
while @@fetch_status = 0
  begin
     -- insert into msc_tree
     insert into MISC_TREE (ID,P0,P1,P2,P3,P4,P5,P6,P7,SHORTCUT) values (@id,@parent,@p0,@p1,@p2,@p3,@p4,@p5,@p6,@sc)
     fetch next from #tmpcrs into @p0,@p1,@p2,@p3,@p4,@p5,@p6,@sc
  end
close #tmpcrs
deallocate #tmpcrs

exec ap_log_add N'M',N'I',@id

commit 
return @id
GO

/****** Object:  StoredProcedure [dbo].[ap_folder_createnew]    Script Date: 23.01.2019 14:40:17 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-----------------------------------------------
ALTER Procedure [dbo].[ap_folder_createnew] 
@name nvarchar(255),
@type smallint,
@parent int,
@guid uniqueidentifier = null
as

set nocount on

begin tran

declare @id int
declare @p0 int, @p1 int, @p2 int, @p3 int, @p4 int, @p5 int, @p6 int, @sc bit

if @guid is null
begin
  insert into FOLDERS (FLD_NAME) VALUES (@name)
  select @id = ident_current('FOLDERS')
end 
else
begin
  insert into FOLDERS (FLD_NAME,FLD_GUID) VALUES (@name,@guid)
  select @id = ident_current('FOLDERS')
end

select P0,P1,P2,P3,P4,P5,P6,P7,SHORTCUT into #tmpfldtree
from FLD_TREE where ID=@parent

if 0 = (select count(*) from #tmpfldtree)
   insert into #tmpfldtree (P0,P1,P2,P3,P4,P5,P6,P7,SHORTCUT) values (@parent,0,0,0,0,0,0,0,0)

declare #tmpcrs cursor local fast_forward read_only for
  select P0,P1,P2,P3,P4,P5,P6,SHORTCUT from #tmpfldtree

if 0 <> (select count(*) from #tmpfldtree where P7 <> 0)
   begin
      rollback tran
      raiserror ('Слишком большой уровень вложенности',16,-1)
      return 0
   end   

open #tmpcrs
fetch next from #tmpcrs into @p0,@p1,@p2,@p3,@p4,@p5,@p6,@sc
while @@fetch_status = 0
  begin
     -- insert into fld_tree
     insert into FLD_TREE (ID,P0,P1,P2,P3,P4,P5,P6,P7,SHORTCUT) values (@id,@parent,@p0,@p1,@p2,@p3,@p4,@p5,@p6,@sc)
     fetch next from #tmpcrs into @p0,@p1,@p2,@p3,@p4,@p5,@p6,@sc
  end
close #tmpcrs
deallocate #tmpcrs

exec ap_log_add N'F',N'I',@id
commit 
return @id
GO

/****** Object:  StoredProcedure [dbo].[ap_binder_createnew]    Script Date: 23.01.2019 14:40:17 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-----------------------------------------
ALTER Procedure [dbo].[ap_binder_createnew] 
@name nvarchar(255),
@type smallint,
@parent int,
@guid uniqueidentifier = null
as

set nocount on

begin tran

declare @id int
declare @p0 int, @p1 int, @p2 int, @p3 int, @p4 int, @p5 int, @p6 int, @sc bit

if @guid is null
begin
  insert into BINDERS (BIND_NAME,BIND_TYPE) VALUES (@name,@type)
  select @id = ident_current('BINDERS')
end
else
begin
  insert into BINDERS (BIND_NAME,BIND_TYPE,BIND_GUID) VALUES (@name,@type,@guid)
  select @id = ident_current('BINDERS')
end

select P0,P1,P2,P3,P4,P5,P6,P7,SHORTCUT into #tmpbndtree
from BIND_TREE where ID=@parent

if 0 = (select count(*) from #tmpbndtree)
   insert into #tmpbndtree (P0,P1,P2,P3,P4,P5,P6,P7,SHORTCUT) values (@parent,0,0,0,0,0,0,0,0)

declare #tmpcrs cursor local fast_forward read_only for
  select P0,P1,P2,P3,P4,P5,P6,SHORTCUT from #tmpbndtree

if 0 <> (select count(*) from #tmpbndtree where P7 <> 0)
   begin
      rollback tran
      raiserror ('Слишком большой уровень вложенности',16,-1)
      return 0
   end   

open #tmpcrs
fetch next from #tmpcrs into @p0,@p1,@p2,@p3,@p4,@p5,@p6,@sc
while @@fetch_status = 0
  begin
     -- insert into bind_tree
     insert into BIND_TREE (ID,P0,P1,P2,P3,P4,P5,P6,P7,SHORTCUT) values (@id,@parent,@p0,@p1,@p2,@p3,@p4,@p5,@p6,@sc)
     fetch next from #tmpcrs into @p0,@p1,@p2,@p3,@p4,@p5,@p6,@sc
  end
close #tmpcrs
deallocate #tmpcrs

exec ap_log_add N'B',N'I',@id
commit 
return @id
GO

/****** Object:  StoredProcedure [dbo].[ap_account_createnew]    Script Date: 23.01.2019 14:40:17 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-----------------------------------------
ALTER Procedure [dbo].[ap_account_createnew] 
@code nvarchar(51),
@name nvarchar(256),
@type smallint,
@parent int,
@guid uniqueidentifier = null
as

set nocount on

begin tran

declare @id int
declare @p0 int, @p1 int, @p2 int, @p3 int, @p4 int, @p5 int, @p6 int, @sc bit

declare @plid int
declare @plcode nvarchar(51)

select @plid = 0
select @plid = ACC_PLAN, @plcode = ACC_PL_CODE from ACCOUNTS where ACC_ID=@parent

if @plid = 0  
  begin
      rollback tran
      raiserror ('Ошибка при вставке счета',16,-1)
      return
  end

if @guid is null
begin
  insert into ACCOUNTS (ACC_NAME,ACC_TYPE,ACC_CODE,ACC_PLAN,ACC_PL_CODE,ACC_S_TYPE,ACC_T_TYPE) 
     values (@name,@type,@code,@plid,@plcode,0,0)
  select @id = ident_current('ACCOUNTS')
end
else
begin
  insert into ACCOUNTS (ACC_NAME,ACC_TYPE,ACC_CODE,ACC_PLAN,ACC_PL_CODE,ACC_S_TYPE,ACC_T_TYPE,ACC_GUID) 
     values (@name,@type,@code,@plid,@plcode,0,0,@guid)
  select @id = ident_current('ACCOUNTS')
end

select P0,P1,P2,P3,P4,P5,P6,P7,SHORTCUT into #tmpagtree
from ACC_TREE where ID=@parent

if 0 = (select count(*) from #tmpagtree)
   insert into #tmpagtree (P0,P1,P2,P3,P4,P5,P6,P7,SHORTCUT) values (@parent,0,0,0,0,0,0,0,0)

declare #tmpcrs cursor local fast_forward read_only for
  select P0,P1,P2,P3,P4,P5,P6,SHORTCUT from #tmpagtree

if 0 <> (select count(*) from #tmpagtree where P7 <> 0)
   begin
      rollback tran
      raiserror ('Слишком большой уровень вложенности',16,-1)
      return
   end   

open #tmpcrs
fetch next from #tmpcrs into @p0,@p1,@p2,@p3,@p4,@p5,@p6,@sc
while @@fetch_status = 0
  begin
     -- insert into ag_tree
     insert into ACC_TREE (ID,P0,P1,P2,P3,P4,P5,P6,P7,SHORTCUT) values (@id,@parent,@p0,@p1,@p2,@p3,@p4,@p5,@p6,@sc)
     fetch next from #tmpcrs into @p0,@p1,@p2,@p3,@p4,@p5,@p6,@sc
  end
close #tmpcrs
deallocate #tmpcrs

exec ap_log_add N'A',N'I',@id
commit tran

select @id, @plid, @plcode
GO

/****** Object:  StoredProcedure [dbo].[ap_accplan_createnew]    Script Date: 23.01.2019 14:40:17 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-----------------------------------------
ALTER Procedure [dbo].[ap_accplan_createnew]
@code nvarchar(51),
@name nvarchar(256),
@guid uniqueidentifier = null
as
set nocount on

declare @id int

begin tran
if @guid is null
begin
  insert into ACCOUNTS (ACC_PLAN,ACC_PL_CODE,ACC_TYPE,ACC_NAME,ACC_CODE) values
     (0,@code,-1,@name,@code)
  select @id = ident_current('ACCOUNTS')
end
else
begin
  insert into ACCOUNTS (ACC_PLAN,ACC_PL_CODE,ACC_TYPE,ACC_NAME,ACC_CODE,ACC_GUID) values
    (0,@code,-1,@name,@code,@guid)
  select @id = ident_current('ACCOUNTS')
end
update ACCOUNTS set ACC_PLAN=@id where ACC_ID=@id
-- в дереве всегда в корне
insert into ACC_TREE (ID)values (@id)
-- остальные значения по умолчанию

exec ap_log_add N'A',N'I',@id
commit tran
return @id
GO

/****** Object:  StoredProcedure [dbo].[ap_account_save]    Script Date: 23.01.2019 14:40:17 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO


--------------------------------------------
ALTER Procedure [dbo].[ap_account_save]
@id int,
@type smallint,
@name nvarchar(256),
@tag nvarchar(51),
@memo nvarchar(256),
@code nvarchar(51),
@tt smallint,
@st smallint
as
set nocount on
begin tran
update ACCOUNTS set ACC_NAME=@name, ACC_TAG=@tag,
  ACC_MEMO=@memo, ACC_CODE=@code, ACC_TYPE=@type, ACC_T_TYPE=@tt, ACC_S_TYPE=@st 
where ACC_ID=@id
if @type = -1 -- план счетов
  begin
    -- обновить коды для всех подчиненных счетов
    update ACCOUNTS set ACC_PL_CODE=@code where ACC_PLAN=@id
  end

  exec ap_log_add N'A',N'U',@id
commit tran


GO

/****** Object:  StoredProcedure [dbo].[ap_agent_settype]    Script Date: 23.01.2019 14:40:17 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO


----------------------------------------
ALTER Procedure [dbo].[ap_agent_settype]
@id int,
@type smallint
as
set nocount on

begin transaction
update AGENTS set AG_TYPE=@type where AG_ID=@id
exec ap_log_add N'G',N'T',@id
commit 

GO

/****** Object:  StoredProcedure [dbo].[ap_agent_save]    Script Date: 23.01.2019 14:40:17 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO


----------------------------------------
ALTER Procedure [dbo].[ap_agent_save]
@id int,
@name nvarchar(256),
@tag nvarchar(51),
@memo nvarchar(256),
@code nvarchar(51) = null,
@addr nvarchar(256) = null,
@phone nvarchar(51) = null,
@vat nvarchar(51) = null,
@reg nvarchar(51) = null,
@tab nvarchar(51) = null,
@pasp nvarchar(256) = null,
@date1 datetime = null,
@date2 datetime = null,
@cnt nvarchar(256) = null,
@country nvarchar(51) = null,
@email nvarchar(256) = null,
@www nvarchar(256) = null
as
set nocount on
begin transaction
update AGENTS set 
  AG_NAME=@name, AG_TAG=@tag, AG_MEMO=@memo, AG_CODE=@code, AG_ADDRESS=@addr,
  AG_PHONE=@phone, AG_VATNO=@vat, AG_REGNO=@reg, AG_TABNO=@tab, AG_PASSPORT=@pasp,
  AG_DATE1=@date1, AG_DATE2=@date2, AG_CONTACT=@cnt, AG_COUNTRY=@country,
  AG_EMAIL=@email, AG_WWW=@www
where AG_ID=@id

exec ap_log_add N'G',N'U',@id
commit transaction


GO

/****** Object:  StoredProcedure [dbo].[ap_agent_createnew]    Script Date: 23.01.2019 14:40:17 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-----------------------------------------
ALTER Procedure [dbo].[ap_agent_createnew] 
@name nvarchar(255),
@type smallint,
@parent int,
@guid uniqueidentifier = null
as
set nocount on
begin tran

declare @id int
declare @p0 int, @p1 int, @p2 int, @p3 int, @p4 int, @p5 int, @p6 int, @sc bit

if @guid is null
begin
  insert into AGENTS (AG_NAME,AG_TYPE) VALUES (@name,@type)
  select @id = ident_current('AGENTS')
end
else
begin
  insert into AGENTS (AG_NAME,AG_TYPE,AG_GUID) VALUES (@name,@type,@guid)
  select @id = ident_current('AGENTS')
end 

select P0,P1,P2,P3,P4,P5,P6,P7,SHORTCUT into #tmpagtree
from AG_TREE where ID=@parent

if 0 = (select count(*) from #tmpagtree)
   insert into #tmpagtree (P0,P1,P2,P3,P4,P5,P6,P7,SHORTCUT) values (@parent,0,0,0,0,0,0,0,0)

declare #tmpcrs cursor local fast_forward read_only for
  select P0,P1,P2,P3,P4,P5,P6,SHORTCUT from #tmpagtree

if 0 <> (select count(*) from #tmpagtree where P7 <> 0)
   begin
      rollback tran
      raiserror ('Слишком большой уровень вложенности',16,-1)
      return 0
   end   

open #tmpcrs
fetch next from #tmpcrs into @p0,@p1,@p2,@p3,@p4,@p5,@p6,@sc
while @@fetch_status = 0
  begin
     -- insert into ag_tree
     insert into AG_TREE (ID,P0,P1,P2,P3,P4,P5,P6,P7,SHORTCUT) values (@id,@parent,@p0,@p1,@p2,@p3,@p4,@p5,@p6,@sc)
     fetch next from #tmpcrs into @p0,@p1,@p2,@p3,@p4,@p5,@p6,@sc
  end
close #tmpcrs
deallocate #tmpcrs

exec ap_log_add N'G',N'I',@id
commit 
return @id
GO

/****** Object:  StoredProcedure [dbo].[ap_entity_settype]    Script Date: 23.01.2019 14:40:17 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO


----------------------------------------
ALTER Procedure [dbo].[ap_entity_settype]
@id int,
@type smallint
as
set nocount on

begin transaction
update ENTITIES set ENT_TYPE=@type where ENT_ID=@id
exec ap_log_add N'E',N'T',@id
commit

GO

/****** Object:  StoredProcedure [dbo].[ap_entity_save]    Script Date: 23.01.2019 14:40:17 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO


----------------------------------------
ALTER Procedure [dbo].[ap_entity_save]
@id int,
@name nvarchar(255),
@tag nvarchar(255) = null,
@memo nvarchar(255) = null,
@cat nvarchar(51) = null,
@nom nvarchar(51) = null,
@art nvarchar(51) = null,
@bar nvarchar(51) = null,
@code nvarchar(51) = null,
@un int = null,
@unself smallint = null,
@acc int = null,
@acself smallint = null
as
set nocount on
begin tran
update ENTITIES set 
  ENT_NAME=@name, ENT_TAG=@tag, ENT_MEMO=@memo,
  ENT_CAT=@cat, ENT_NOM=@nom, ENT_ART=@art, ENT_BAR=@bar, ENT_CODE=@code, 
  UN_ID=@un, ACC_ID=@acc, UN_SELF=@unself, ACC_SELF=@acself
where 
  ENT_ID=@id

exec ap_log_add N'E',N'U',@id

commit tran


GO

/****** Object:  StoredProcedure [dbo].[ap_entity_createnew]    Script Date: 23.01.2019 14:40:17 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-----------------------------------------
ALTER Procedure [dbo].[ap_entity_createnew] 
@name nvarchar(255),
@type smallint,
@parent int,
@guid uniqueidentifier = null
as

set nocount on

begin tran

declare @id int
declare @p0 int, @p1 int, @p2 int, @p3 int, @p4 int, @p5 int, @p6 int, @sc bit

if @guid is null
begin
  insert into ENTITIES (ENT_NAME,ENT_TYPE) VALUES (@name,@type)
  select @id = ident_current('ENTITIES')
end
else
begin
  insert into ENTITIES (ENT_NAME,ENT_TYPE,ENT_GUID) VALUES (@name,@type,@guid)
  select @id = ident_current('ENTITIES')
end 

select P0,P1,P2,P3,P4,P5,P6,P7,SHORTCUT into #tmpenttree
from ENT_TREE where ID=@parent

if 0 = (select count(*) from #tmpenttree)
   insert into #tmpenttree (P0,P1,P2,P3,P4,P5,P6,P7,SHORTCUT) values (@parent,0,0,0,0,0,0,0,0)

declare #tmpcrs cursor local fast_forward read_only for
  select P0,P1,P2,P3,P4,P5,P6,SHORTCUT from #tmpenttree

if 0 <> (select count(*) from #tmpenttree where P7 <> 0)
   begin
      rollback tran
      raiserror ('Слишком большой уровень вложенности',16,-1)
      return 0
   end   

open #tmpcrs
fetch next from #tmpcrs into @p0,@p1,@p2,@p3,@p4,@p5,@p6,@sc
while @@fetch_status = 0
  begin
     -- insert into ent_tree
     insert into ENT_TREE (ID,P0,P1,P2,P3,P4,P5,P6,P7,SHORTCUT) values (@id,@parent,@p0,@p1,@p2,@p3,@p4,@p5,@p6,@sc)
     fetch next from #tmpcrs into @p0,@p1,@p2,@p3,@p4,@p5,@p6,@sc
  end
close #tmpcrs
deallocate #tmpcrs


exec ap_log_add N'E',N'I',@id
commit 
return @id
GO


