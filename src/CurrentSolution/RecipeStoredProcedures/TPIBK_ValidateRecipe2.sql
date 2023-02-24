USE [TPMDB]
GO
/****** Object:  StoredProcedure [dbo].[TPIBK_ValidateRecipe2]    Script Date: 2/16/2023 12:09:47 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
ALTER PROCEDURE [dbo].[TPIBK_ValidateRecipe2]

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
	@ReturnMsg	as nvarchar(100) OUTPUT
		
AS

Declare
	@ID									as int,
	@Recipe_RID							as nvarchar(25),
	@Recipe_Version						as nvarchar(10),
	@TPIBK_StepType_ID					as int,
	@ProcessClassPhase_ID				as int,
	@Step								as smallint,
	@UserString							as nvarchar(100),
	@RecipeEquipmentTransition_Data_ID	as int,
	@NextStep							as int,
	@Allocation_Type_ID					as int,
	@LateBinding						as smallint,
	@Material_ID						as int,
	@ProcessClass_ID					as int,	

	@PrevProcessClass_ID				as int = 0,
	@Allocated							as bit = 0,
	@ValidateMsg						as nvarchar(100) = 'Valid',
	@Result			as int = 1

SET NOCOUNT ON

	/* Validate Alloc/Dealloc */
	DECLARE ValidateAlloc CURSOR FOR
	SELECT ID,Recipe_RID,Recipe_Version,TPIBK_StepType_ID,ProcessClassPhase_ID,Step,UserString,RecipeEquipmentTransition_Data_ID,
			NextStep,Allocation_Type_ID,LateBinding,Material_ID,ProcessClass_ID
	FROM TPIBK_RecipeBatchData
	WHERE Recipe_RID = @RID AND Recipe_Version = @Version 
	ORDER BY ProcessClass_ID, Step
  
	OPEN ValidateAlloc
	-- Select the first record
	FETCH NEXT FROM ValidateAlloc
	INTO @ID, @Recipe_RID, @Recipe_Version, @TPIBK_StepType_ID, @ProcessClassPhase_ID, @Step, @UserString, 
			@RecipeEquipmentTransition_Data_ID, @NextStep, @Allocation_Type_ID, @LateBinding, @Material_ID, @ProcessClass_ID
	WHILE @@FETCH_STATUS = 0 AND @Result = 1
	BEGIN
		--Check the previous unit was deallocated
		IF @ProcessClass_ID <> @PrevProcessClass_ID AND @Allocated = 1
		BEGIN
			SET @Result = 0
			SET @ValidateMsg = 'Process class not deallocated after step ' + CONVERT(VARCHAR, @Step)
		END

		--Update unit
		IF @ProcessClass_ID <> @PrevProcessClass_ID	SET @PrevProcessClass_ID = @ProcessClass_ID

		--Check the new unit is allocated	
		IF @Allocated = 0 AND @TPIBK_StepType_ID <> 8
		BEGIN
			SET @Result = 0
			SET @ValidateMsg = 'Step ' + CONVERT(VARCHAR, @Step) + 'process class not allocated'
		END
		
		--Update allocation status for unit
		IF @TPIBK_StepType_ID = 8 SET @Allocated = 1
		ELSE IF @TPIBK_StepType_ID = 9 SET @Allocated = 0

		--Check step type exists
		IF @Result = 1
			SELECT TPIBK_StepType.ID
			FROM TPIBK_StepType
			WHERE TPIBK_StepType.ID = @TPIBK_StepType_ID
			IF @@ROWCOUNT = 0 
			BEGIN
				SET @Result = 0			
				SET @ValidateMsg = 'Step ' + CONVERT(VARCHAR, @Step) + ' - invalid type'
			END

		--Check step phase exists
		IF @Result = 1
			SELECT ProcessClassPhase.ID
			FROM ProcessClassPhase
			WHERE ProcessClassPhase.ID = @ProcessClassPhase_ID  
			IF @@ROWCOUNT = 0 		
			BEGIN
				SET @Result = 0
				SET @ValidateMsg = 'Step ' + CONVERT(VARCHAR, @Step) + ' - phase not defined'
			END
	
		--Check step material exists 
		IF @Result = 1
			SELECT Material.ID
			FROM Material
			WHERE Material.ID = @Material_ID
			IF @@ROWCOUNT = 0 			
			BEGIN
				SET @Result = 0
				SET @ValidateMsg = 'Step ' + CONVERT(VARCHAR, @Step) + ' - material #' + CONVERT(VARCHAR, @Material_ID) + ' not defined'
			END
		
		--Move to next record
		FETCH NEXT FROM ValidateAlloc
		INTO @ID, @Recipe_RID, @Recipe_Version, @TPIBK_StepType_ID, @ProcessClassPhase_ID, @Step, @UserString, 
			@RecipeEquipmentTransition_Data_ID, @NextStep, @Allocation_Type_ID, @LateBinding, @Material_ID, @ProcessClass_ID
	END
	CLOSE ValidateAlloc
	DEALLOCATE ValidateAlloc	

	--Check the last unit is deallocated
	IF @Allocated = 1	
		BEGIN
			SET @Result = 0
			SET @ValidateMsg = 'Process class not deallocated after ' + CONVERT(VARCHAR, @Step)
		END

	-- Return validation result
	SET @ReturnMsg = @ValidateMsg
	RETURN @Result
