USE [TPMDB]
GO
/****** Object:  StoredProcedure [dbo].[TPIBK_getRecipestepdownload]    Script Date: 2/16/2023 12:05:16 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
ALTER PROCEDURE [dbo].[TPIBK_getRecipestepdownload]



	@RID			as nvarchar(50),
	@Ver			as nvarchar(50),	
	@Train			as nvarchar(50)

AS

SET NOCOUNT ON

	SELECT * 
	FROM v_recipestepdownload 
	WHERE rid=@RID And [Version]=@Ver And ((RecipeTrainName= @Train) or (RecipeTrainName is null) ) 
	order by step
SET NOCOUNT OFF
