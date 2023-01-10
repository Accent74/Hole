GO

ALTER procedure [dbo].[HL74_TaxInDocs]
		@Acc631ID int = 298, 
		@StartDate DateTime = '20220701', 
		@EndDate DateTime = '20220801', 
		@MCID int = 1, 
		@TaxRepID int = 8
As
begin
	set nocount on ;

	drop table if exists #prm
	select 
		FormID = p.FRM_ID
		,FormName = isnull(p.prm_string, '')
	into #prm
	from FRM_PARAMS as p
		left join FRM_PARAM_NAMES as pn on pn.PRM_ID = p.PRM_ID
	where 
		pn.PRM_NAME like 'Наименование для отчета'


	SELECT 
      AG_CODE, 
      AG_NAME, 
      J_DATE, 
      DOC_NO, 
      SUM(J_SUM),
      isnull(FormName, '')
	FROM JRN_TAX AS JT 
      INNER JOIN JOURNAL AS J ON J.J_ID = JT.J_ID
      INNER JOIN DOCUMENTS AS D ON D.DOC_ID = J.DOC_ID
      INNER JOIN AGENTS AS A ON A.AG_ID = J.J_AG2
      left join #prm as f on f.FormID = d.FRM_ID
	WHERE JT.TX_ID = @TaxRepID 
			And JT_ADDR1 = 'Д.А.' 
			And J.MC_ID = @MCID 
			And ACC_CR in (select id  from acc_tree where @Acc631ID in (id, p0, p1, p2, p3, p4, p5, p6, p7))
			And J_DATE >= @StartDate 
			And J_DATE < @EndDate And  J_DONE = 2
	GROUP BY 
      AG_CODE,
	  AG_TYPE, 
      AG_NAME, 
      J_DATE, 
      DOC_NO,
      isnull(FormName, '')
	order by 
    case	
		when CHARINDEX('ФЛ-П', AG_NAME) = 0 and 
			not AG_NAME like 'Конечный потребитель' and
			AG_TYPE < 3 then 1
		when CHARINDEX('ФЛ-П', AG_NAME) > 0 then 10
		else 100
	end,
	isnull(AG_CODE, '999999999'),
	AG_NAME
end

go
--declare
ALTER procedure [dbo].[HL74_TaxOutDocs]
		@Acc361ID int = 219, 
		@StartDate DateTime = '20220701', 
		@EndDate DateTime = '20220801', 
		@MCID int = 1, 
		@TaxRepID int = 8

as
begin
   set nocount on ;

	drop table if exists #prm
	select 
		FormID = p.FRM_ID
		,FormName = isnull(p.prm_string, '')
	into #prm
	from FRM_PARAMS as p
		left join FRM_PARAM_NAMES as pn on pn.PRM_ID = p.PRM_ID
	where 
		pn.PRM_NAME like 'Наименование для отчета'

   SELECT 
      AG_CODE, 
      AG_NAME, 
      J_DATE, 
      DOC_NO, 
      SUM(J_SUM),
      isnull(FormName, '')
	FROM JRN_TAX AS JT 
      INNER JOIN JOURNAL AS J ON J.J_ID = JT.J_ID
      INNER JOIN DOCUMENTS AS D ON D.DOC_ID = J.DOC_ID
      INNER JOIN AGENTS AS A ON A.AG_ID = J.J_AG1
      left join #prm as f on f.FormID = d.FRM_ID

   WHERE 
	   JT.TX_ID = @TaxRepID 
	   And 
	   JT_ADDR1 = 'Д.А.' 
	   And 
	   J.MC_ID = @MCID 
	   And ACC_DB in (select id  from acc_tree where @Acc361ID in (id, p0, p1, p2, p3, p4, p5, p6, p7))
	   And J_DATE >= @StartDate 
	   And J_DATE < @EndDate 
	   And  J_DONE = 2
   GROUP BY 
      AG_CODE, 
	  AG_TYPE, 
      AG_NAME, 
      J_DATE, 
      DOC_NO,
      isnull(FormName, '')
   order by 
    case	
		when CHARINDEX('ФЛ-П', AG_NAME) = 0 and 
			not AG_NAME like 'Конечный потребитель' and
			AG_TYPE < 3 then 1
		when CHARINDEX('ФЛ-П', AG_NAME) > 0 then 10
		else 100
	end,

	isnull(AG_CODE, '999999999'),
	AG_NAME
end		

go



