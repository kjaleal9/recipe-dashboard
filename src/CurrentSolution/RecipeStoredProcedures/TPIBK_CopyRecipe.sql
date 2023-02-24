USE [TPMDB]
GO
/****** Object:  StoredProcedure [dbo].[TPIBK_CopyRecipe]    Script Date: 2/16/2023 12:03:16 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
ALTER PROCEDURE [dbo].[TPIBK_CopyRecipe]

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

	@oRID		as nvarchar(25),
	@oVersion	as nvarchar(10),
	@nRID		as nvarchar(30),
	@nVersion	as nvarchar(10)

AS

Declare @KeyTable TABLE (oKey INT ,nKey INT)

Declare
	@InsEquipID		as int,
	@InsBatchID		as int,
	@InsStepID		as int,
	@NewEquipKey	as int,
	@oEquipKey		as int,
	@NewBatchKey	as int

SET NOCOUNT ON

DECLARE InsEquip CURSOR FOR
	SELECT ID
	FROM RecipeEquipmentRequirement
	WHERE Recipe_RID = @oRID AND Recipe_Version = @oVersion 

	OPEN InsEquip
	-- Select first record
	FETCH NEXT FROM InsEquip
	INTO @InsEquipID
	WHILE @@FETCH_STATUS = 0
	BEGIN
		/* Copy data in RecipeEquipmentRequirement */
		INSERT INTO RecipeEquipmentRequirement
		SELECT EquipmentType, @nRID AS Recipe_RID, @nVersion AS Recipe_Version, ProcessClass_Name, Equipment_Name, LateBinding, IsMainBatchUnit,EqIdx1
		FROM RecipeEquipmentRequirement
		WHERE ID = @InsEquipID	

		--SELECT SCOPE_IDENTITY() AS NewID
		set @NewEquipKey = @@IDENTITY
		INSERT INTO @KeyTable (oKey, nKey) VALUES(@InsEquipID,@NewEquipKey)

		--Move to next record
		FETCH NEXT FROM InsEquip
		INTO @InsEquipID
	END
	CLOSE InsEquip
	DEALLOCATE InsEquip	

	/* Copy data in TPIBK_RecipeBatchData */
	DECLARE InsBatch CURSOR FOR
		SELECT ID, ProcessClass_ID
		from TPIBK_RecipeBatchData
		where Recipe_RID = @oRID AND Recipe_Version = @OVersion
	
		OPEN InsBatch
		-- Select first record
	    FETCH NEXT FROM InsBatch
		INTO @InsBatchID, @oEquipKey
		WHILE @@FETCH_STATUS = 0
		BEGIN
		
			SET @NewEquipKey = (Select nKey from @KeyTable where oKey = @oEquipKey)
		
			INSERT INTO TPIBK_RecipeBatchData
			SELECT @nRID AS Recipe_RID, @nVersion AS Recipe_Version,TPIBK_StepType_ID,ProcessClassPhase_ID,Step,UserString,
			RecipeEquipmentTransition_Data_ID,NextStep,Allocation_Type_ID,LateBinding,Material_ID,@NewEquipKey as ProcessClass_ID
			FROM TPIBK_RecipeBatchData
			WHERE ID = @InsBatchID

			--SELECT SCOPE_IDENTITY() AS NewID
			SET @NewBatchKey = @@IDENTITY

			/* Copy data in TPIBK_RecipeStepData */
			DECLARE InsStep CURSOR FOR
				SELECT ID
				FROM TPIBK_RecipeStepData
				WHERE TPIBK_RecipeBatchData_ID = @InsBatchID

			OPEN InsStep
			-- Select first record
			FETCH NEXT FROM InsStep
			into @InsStepID
			WHILE @@FETCH_STATUS = 0
			BEGIN		
				INSERT INTO TPIBK_RecipeStepData
				SELECT TPIBK_RecipeParameterData_ID, @NewBatchKey AS TPIBK_RecipeBatchData_ID, Value,CustomEU --[CustomEU] as EU
				FROM TPIBK_RecipeStepData
				WHERE ID = @InsStepID	

				--Move to next record
			    FETCH NEXT FROM InsStep
				INTO @InsStepID	
			END
			CLOSE InsStep
			DEALLOCATE InsStep

		--Move to next record
        FETCH NEXT FROM InsBatch
        INTO @InsBatchID, @oEquipKey
	END
	CLOSE InsBatch
	DEALLOCATE InsBatch

	Declare
	@nID		as int,
	@nStep		as int,
	@nNextStep		as int,
	@oID		as int,
	@oStep		as int,
	@oNextStep		as int
	/*Update the Id for jump steps */
	--find all new jump steps
	DECLARE InsBatch_Jump CURSOR FOR
		SELECT ID, Step,NextStep
		from TPIBK_RecipeBatchData
		where (Recipe_RID = @nRID) AND (Recipe_Version = @nVersion) and (not Nextstep is null) and ([TPIBK_StepType_ID]=10)
	
		OPEN InsBatch_Jump
		-- Select first record
	    FETCH NEXT FROM InsBatch_Jump
		INTO @nID,@nStep,@oNextstep
		WHILE @@FETCH_STATUS = 0
		BEGIN

			--find the step number from the old step
			set @oStep=(select top 1 step from TPIBK_RecipeBatchData where id=@oNextstep)
			--find the new id of the step
			set @nNextStep=(select top 1 id from TPIBK_RecipeBatchData where (Recipe_RID = @nRID) AND (Recipe_Version = @nVersion) and (step =@oStep))
			--fix the record
			update TPIBK_RecipeBatchData set NextStep=@nNextStep where id=@nID
			--Move to next record
			FETCH NEXT FROM InsBatch_Jump
			INTO @nID,@nStep,@oNextstep
		END

	CLOSE InsBatch_Jump
	DEALLOCATE InsBatch_Jump


SET NOCOUNT OFF
