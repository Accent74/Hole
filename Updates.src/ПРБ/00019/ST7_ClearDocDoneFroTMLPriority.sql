
ALTER PROCEDURE [dbo].[ST7_ClearDocDoneForTMLPriority] 
	@dStart DateTime, @dEnd DateTime, @TmlPrmID int, @MC int

AS
SET NOCOUNT ON

select doc_id 
into #docs
from DOCUMENTS
Where DOCUMENTS.TML_ID in (Select TML_ID FROM TML_PARAMS Where PRM_ID = @TmlPrmID And PRM_LONG <> 0)
		And (DOC_DATE between @dStart and @dEnd )
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





