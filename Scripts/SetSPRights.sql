
declare c_t cursor for
SELECT  o.name AS StoredProcedureName ,
        o.[Type]
FROM    sys.objects o
WHERE   o.[Type] = 'P' -- Stored Procedures 
		and DB_NAME()='walmart_work2'

declare @t_name nvarchar(255), @sql nvarchar(255)

open c_t

FETCH NEXT FROM c_t 
INTO @t_name

WHILE @@FETCH_STATUS = 0  
BEGIN  
	set @SQL = 'GRANT INSERT ON [dbo].[' + @t_name + '] TO [ap_public] '
	exec sp_executesql @SQL

	set @SQL = 'GRANT ALTER ON [dbo].[' + @t_name + '] TO [ap_public] '
	exec sp_executesql @SQL

	set @SQL = 'GRANT UPDATE ON [dbo].[' + @t_name + '] TO [ap_public] '
	exec sp_executesql @SQL

	set @SQL = 'GRANT DELETE ON [dbo].[' + @t_name + '] TO [ap_public] '
	exec sp_executesql @SQL
	
	set @SQL = 'GRANT SELECT ON [dbo].[' + @t_name + '] TO [ap_public] '
	exec sp_executesql @SQL

	FETCH NEXT FROM c_t 
	INTO @t_name
end

close c_t
deallocate c_t




