USE [TPMDB]
GO
/****** Object:  StoredProcedure [dbo].[TPIBK_ValidateRecipe1]    Script Date: 2/16/2023 12:09:15 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
ALTER PROCEDURE [dbo].[TPIBK_ValidateRecipe1]

/*****************************************************************************************************************/
-- ---------------------------------------------------------------------------------------------------------------
-- Object		Stored Procedure TPIBK_CopyRecipe
-- ---------------------------------------------------------------------------------------------------------------
-- Author		Carey Warren 
-- Created		2014-02-04
-- Description	Copy TPI BK Recipe
-- ---------------------------------------------------------------------------------------------------------------
-- History
--	[Version]	[Name]				[Date]		[Comments]
-- ---------------------------------------------------------------------------------------------------------------

	@RID		as nvarchar(30),
	@Version	as nvarchar(10),
	@ResultMsg	as int OUTPUT
		
AS

Declare @KeyTable TABLE (oKey INT ,nKey INT)

Declare
	@ID				as int,
	@AllocCount		as int,
	@DeallocCount	as int,
	@Count			as int,
	@Result			as int = 1

SET NOCOUNT ON

	/* Validate Alloc/Dealloc */
	DECLARE InsEquip CURSOR FOR
	SELECT DISTINCT ProcessClass_ID
	FROM TPIBK_RecipeBatchData
	WHERE Recipe_RID = @RID AND Recipe_Version = @Version 

	OPEN InsEquip
	-- Select first record
	FETCH NEXT FROM InsEquip
	INTO @ID
	WHILE @@FETCH_STATUS = 0 AND @Result = 1
	BEGIN

		SELECT TPIBK_StepType_ID 
		FROM TPIBK_RecipeBatchData
		WHERE ProcessClass_ID = @ID AND TPIBK_StepType_ID = 8
		SET @AllocCount = @@ROWCOUNT
	
		SELECT DISTINCT TPIBK_StepType_ID 
		FROM TPIBK_RecipeBatchData
		WHERE ProcessClass_ID = @ID AND TPIBK_StepType_ID = 9
		SET @DeallocCount = @@ROWCOUNT

		IF @AllocCount <> @DeallocCount 
			SET @Result = 0

		--Move to next record
		FETCH NEXT FROM InsEquip
		INTO @ID
	END
	CLOSE InsEquip
	DEALLOCATE InsEquip	

	/* Check step phase exists */
	IF @Result = 1 
		SELECT TPIBK_RecipeBatchData.TPIBK_StepType_ID, TPIBK_StepType.ID
		FROM TPIBK_RecipeBatchData 
		LEFT OUTER JOIN TPIBK_StepType ON TPIBK_RecipeBatchData.TPIBK_StepType_ID = TPIBK_StepType.ID
		WHERE TPIBK_RecipeBatchData.Recipe_RID = @RID AND TPIBK_RecipeBatchData.Recipe_Version = @Version AND TPIBK_StepType.ID IS NULL
		SET @Count = @@ROWCOUNT
		IF @Count > 0 
			SET @Result = 0
			
	/* Check step type exists */
	IF @Result = 1 
		SELECT TPIBK_RecipeBatchData.ProcessClassPhase_ID, ProcessClassPhase.ID
		FROM TPIBK_RecipeBatchData 
		LEFT OUTER JOIN ProcessClassPhase ON TPIBK_RecipeBatchData.ProcessClassPhase_ID = ProcessClassPhase.ID
		WHERE TPIBK_RecipeBatchData.Recipe_RID = @RID AND TPIBK_RecipeBatchData.Recipe_Version = @Version AND ProcessClassPhase.ID IS NULL
		SET @Count = @@ROWCOUNT
		IF @Count > 0 
			SET @Result = 0
			
	/* Check step material exists */
	IF @Result = 1 
		SELECT TPIBK_RecipeBatchData.Material_ID, Material.ID
		FROM TPIBK_RecipeBatchData 
		LEFT OUTER JOIN Material ON TPIBK_RecipeBatchData.Material_ID = Material.ID
		WHERE TPIBK_RecipeBatchData.Recipe_RID = @RID AND TPIBK_RecipeBatchData.Recipe_Version = @Version AND Material.ID IS NULL
		SET @Count = @@ROWCOUNT
		IF @Count > 0 
			SET @Result = 0
	
	-- Return validation result
	SET @ResultMsg = @Result
	--return @ResultMsg
SET NOCOUNT OFF
