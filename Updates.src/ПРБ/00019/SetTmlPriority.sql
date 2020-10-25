
/****** Object:  StoredProcedure [dbo].[ST7_SelectDocsForTMLPriority2]    Script Date: 11.01.2018 18:31:22 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE 	[dbo].[ST7_SelectDocsForTMLPriority2]
	@dStart DateTime, @dEnd DateTime, @pStart DateTime, @pEnd DateTime, @TmlPrmID int, @MC int
AS
SET NOCOUNT ON

Select DOCUMENTS.DOC_ID
FROM DOCUMENTS INNER Join TML_PARAMS On DOCUMENTS.TML_ID = TML_PARAMS.TML_ID
WHERE (((TML_PARAMS.PRM_LONG)<>0) And ((DOCUMENTS.DOC_DONE) < 100) 
	And ((DOCUMENTS.DOC_DATE) between @dStart and @dEnd) 
	And ((DOCUMENTS.DOC_DATE) between @pStart and @pEnd) 
	And ((TML_PARAMS.PRM_ID)=@TmlPrmID)) and MC_ID = @MC
ORDER BY documents.DOC_DATE, TML_PARAMS.PRM_LONG




GO

/****** Object:  StoredProcedure [dbo].[ST7_ClearDocDoneForTMLPriority2]    Script Date: 11.01.2018 18:31:22 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


CREATE PROCEDURE [dbo].[ST7_ClearDocDoneForTMLPriority2] 
	@dStart DateTime, @dEnd DateTime, @pStart DateTime, @pEnd DateTime, @TmlPrmID int, @MC int

AS
SET NOCOUNT ON

select doc_id 
into #docs
from DOCUMENTS
Where DOCUMENTS.TML_ID in (Select TML_ID FROM TML_PARAMS Where PRM_ID = @TmlPrmID And PRM_LONG <> 0)
		And (DOC_DATE between @dStart and @dEnd )
		And (DOC_DATE between @pStart and @pEnd )
		And DOC_DONE <100 and MC_ID = @MC


BEGIN TRY
    BEGIN TRAN
		update documents 
		set doc_done = 0
		where doc_id in (select doc_id from #docs)

		update journal
		set j_done = 0
		where doc_id in (select doc_id from #docs);
    COMMIT TRAN
END TRY

BEGIN CATCH
    ROLLBACK TRAN
END CATCH






GO


