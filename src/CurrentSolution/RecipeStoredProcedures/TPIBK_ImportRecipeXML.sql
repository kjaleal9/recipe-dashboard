USE [TPMDB]
GO
/****** Object:  StoredProcedure [dbo].[TPIBK_ImportRecipeXML]    Script Date: 2/16/2023 12:06:46 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

ALTER PROCEDURE [dbo].[TPIBK_ImportRecipeXML]

	@xml			as xml

AS

SET NOCOUNT ON
declare @result int
set @result=1

declare  @Recipe table(
	[RID] [nvarchar](25) NOT NULL,
	[Version] [nvarchar](10) NOT NULL,
	[RecipeType] [nvarchar](10) NOT NULL,
	[NbrOfExecutions] [int] NULL,
	[VersionDate] [datetime] NULL,
	[Description] [nvarchar](60) NULL,
	[EffectiveDate] [datetime] NULL,
	[ExpirationDate] [datetime] NULL,
	[ProductID] [nvarchar](50) NOT NULL,
	[BatchSizeNominal] [float] NULL,
	[BatchSizeMin] [float] NULL,
	[BatchSizeMax] [float] NULL,
	[Status] [nvarchar](10) NULL,
	[UseBatchKernel] [bit] NULL,
	[CurrentElementID] [int] NULL,
	[RecipeData] [image] NULL,
	[RunMode] [int] NOT NULL,
	[IsPackagingRecipeType] [bit] NOT NULL)
declare  @RecipeEquipmentRequirement table(
	[ID] [int] NOT NULL,
	[EquipmentType] [nvarchar](10) NOT NULL,
	[Recipe_RID] [nvarchar](25) NOT NULL,
	[Recipe_Version] [nvarchar](10) NOT NULL,
	[ProcessClass_Name] [nvarchar](30) NULL,
	[Equipment_Name] [nvarchar](20) NULL,
	[LateBinding] [int] NOT NULL,
	[IsMainBatchUnit] [bit] NOT NULL,
	[EqIdx1] [int] NULL)
declare @TPIBK_RecipeBatchData table(
	[ID] [int] NOT NULL,
	[Recipe_RID] [nvarchar](25) NOT NULL,
	[Recipe_Version] [nvarchar](10) NOT NULL,
	[TPIBK_StepType_ID] [int] NOT NULL,
	[ProcessClassPhase_ID] [int] NULL,
	[Step] [smallint] NULL,
	[UserString] [nvarchar](100) NULL,
	[RecipeEquipmentTransition_Data_ID] [int] NULL,
	[NextStep] [int] NULL,
	[Allocation_Type_ID] [int] NULL,
	[LateBinding] [smallint] NULL,
	[Material_ID] [int] NULL,
	[ProcessClass_ID] [int] NULL)
declare @TPIBK_RecipeStepData TABLE(
	[ID] [int] NOT NULL,
	[TPIBK_RecipeParameterData_ID] [int] NOT NULL,
	[TPIBK_RecipeBatchData_ID] [int] NOT NULL,
	[Value] [real] NOT NULL,
	[CustomEU] [nvarchar](50) NULL)
--Load data from xml into temp tables--------------------------------------------------------------------------
DECLARE @docHandle int  
declare @x as xml
DECLARE readxml CURSOR FOR

	select @xml
	
	OPEN readxml
	-- Select first record
	FETCH NEXT FROM readxml
	INTO @x
	WHILE @@FETCH_STATUS = 0
	BEGIN
		
		--print cast(@x as nvarchar(max))
		EXEC sp_xml_preparedocument @docHandle OUTPUT, @x
		INSERT INTO @Recipe
           ([RID]
           ,[Version]
           ,[RecipeType]
           ,[NbrOfExecutions]
           ,[VersionDate]
           ,[Description]
           ,[EffectiveDate]
           ,[ExpirationDate]
           ,[ProductID]
           ,[BatchSizeNominal]
           ,[BatchSizeMin]
           ,[BatchSizeMax]
           ,[Status]
           ,[UseBatchKernel]
           ,[CurrentElementID]
           ,[RecipeData]
           ,[RunMode]
           ,[IsPackagingRecipeType])
		SELECT *   FROM OPENXML(@docHandle, N'Recipes/Recipe')   
		with
		(
		RID[nvarchar](25) '@RID'
		,Version[nvarchar](10) '@Version'
	  ,RecipeType[nvarchar](10) '@RecipeType'
	  ,NbrOfExecutions[int] '@NbrOfExecutions'
	  ,VersionDate[datetime] '@VersionDate'
	  ,Description[nvarchar](60) '@Description'
	  ,EffectiveDate[datetime] '@EffectiveDate'
	  ,ExpirationDate[datetime] '@ExpirationDate'
	  ,ProductID[nvarchar](50) '@ProductID'
	  ,BatchSizeNominal[float] '@BatchSizeNominal'
	  ,BatchSizeMin[float] '@BatchSizeMin'
	  ,BatchSizeMax[float] '@BatchSizeMax'
	  ,Status[nvarchar](10) '@Status'
	  ,UseBatchKernel[bit] '@UseBatchKernel'
	  ,CurrentElementID[int] '@CurrentElementID'
	  ,RecipeData[image] '@RecipeData'
	  ,RunMode[int] '@RunMode'
	  ,IsPackagingRecipeType[bit] '@IsPackagingRecipeType'
		)
		
		INSERT INTO @RecipeEquipmentRequirement
           (ID,[EquipmentType]
           ,[Recipe_RID]
           ,[Recipe_Version]
           ,[ProcessClass_Name]
           ,[Equipment_Name]
           ,[LateBinding]
           ,[IsMainBatchUnit]
           ,[EqIdx1])
		
		SELECT *   FROM OPENXML(@docHandle, N'Recipes/Recipe/RecipeEquipmentRequirement_Rows/RecipeEquipmentRequirement')   
		with
		(
		ID[int] '@ID'
      ,EquipmentType[nvarchar](10) '@EquipmentType'
      ,Recipe_RID[nvarchar](25)  '@Recipe_RID'
      ,Recipe_Version[nvarchar](10) '@Recipe_Version'
      ,ProcessClass_Name[nvarchar](30) '@ProcessClass_Name'
      ,Equipment_Name[nvarchar](20) '@Equipment_Name'
      ,LateBinding[int] '@LateBinding'
      ,IsMainBatchUnit[bit] '@IsMainBatchUnit'
      ,EqIdx1[int] '@EqIdx1'
		)
	
		INSERT INTO @TPIBK_RecipeBatchData
           (ID,[Recipe_RID]
           ,[Recipe_Version]
           ,[TPIBK_StepType_ID]
           ,[ProcessClassPhase_ID]
           ,[Step]
           ,[UserString]
           ,[RecipeEquipmentTransition_Data_ID]
           ,[NextStep]
           ,[Allocation_Type_ID]
           ,[LateBinding]
           ,[Material_ID]
           ,[ProcessClass_ID])
           
		SELECT *   FROM OPENXML(@docHandle, N'Recipes/Recipe/TPIBK_RecipeBatchData_Rows/TPIBK_RecipeBatchData')   
		with
		(
		ID[int] '@ID'
      ,Recipe_RID[nvarchar](25)  '@Recipe_RID'
      ,Recipe_Version[nvarchar](10) '@Recipe_Version'
	  ,TPIBK_StepType_ID[int] '@TPIBK_StepType_ID'
      ,ProcessClassPhase_ID[int] '@ProcessClassPhase_ID'
      ,Step[smallint] '@Step'
      ,UserString[nvarchar](100) '@UserString'
      ,RecipeEquipmentTransition_Data_ID[int] '@RecipeEquipmentTransition_Data_ID'
      ,NextStep[int] '@NextStep'
      ,Allocation_Type_ID[int] '@Allocation_Type_ID'
      ,LateBinding[smallint] '@LateBinding'
      ,Material_ID[int] '@Material_ID'
      ,ProcessClass_ID[int] '@ProcessClass_ID'
		)
		INSERT INTO @TPIBK_RecipeStepData
           (ID,[TPIBK_RecipeParameterData_ID]
           ,[TPIBK_RecipeBatchData_ID]
           ,[Value]
           ,[CustomEU])
		
		SELECT *   FROM OPENXML(@docHandle, N'Recipes/Recipe/TPIBK_RecipeBatchData_Rows/TPIBK_RecipeBatchData/TPIBK_RecipeStepData_Rows/TPIBK_RecipeStepData')   
		with
		(
		ID[int] '@ID'
       ,TPIBK_RecipeParameterData_ID[int] '@TPIBK_RecipeParameterData_ID'
		  ,TPIBK_RecipeBatchData_ID[int] '@TPIBK_RecipeBatchData_ID'
		  ,Value[real] '@Value'
		  ,CustomEU[nvarchar](50) '@CustomEU'
		)
		EXEC sp_xml_removedocument @docHandle;  
		
		--Move to next record
		FETCH NEXT FROM readxml
		INTO @x
	END
	CLOSE readxml
	DEALLOCATE readxml	
--validate data in temp files--------------------------------------------------------------------------
	declare @version as int
	declare @RID as nvarchar(25)
	set @rid=(select top 1 rid from @Recipe)
	--create a new version number for the import.  if non exist default 1
	set @version=(select coalesce(Max(CONVERT(INT,[Version])),0)+1 from recipe  where RecipeType='Master' and RID= @RID)
	--change to now date and registered
	update @Recipe set VersionDate=GETDATE(), Status='Registered', Version=@version
	--update the rest of the temp tables with new version
	update @RecipeEquipmentRequirement set Recipe_Version=@version		
	update @TPIBK_RecipeBatchData set Recipe_Version=@version
	--check product id--------------------------------------
	if not exists(select name from Material where name in(select productid from @Recipe))
	begin
		set @result=2
	end
	--check process class name--------------------------------
	declare @n as nvarchar(30)
	DECLARE checkxml CURSOR FOR

	select processclass_name from @RecipeEquipmentRequirement
	
	OPEN checkxml
	-- Select first record
	FETCH NEXT FROM checkxml
	INTO @n
	WHILE @@FETCH_STATUS = 0 and @result=1
	BEGIN
	if not exists(select name from ProcessClass where name =@n)
	begin
		set @result=3
	end
	--Move to next record
		FETCH NEXT FROM checkxml
		INTO @n
	END
	CLOSE checkxml
	DEALLOCATE checkxml	
	--recipe batch data values--------------------------------
	declare @steptype int
	declare @pcp int
	declare @ret int
	declare @matid int
	declare @pc int
	
	DECLARE checkxml1 CURSOR FOR

	select TPIBK_StepType_ID,ProcessClassPhase_ID,RecipeEquipmentTransition_Data_ID,Material_ID,ProcessClass_ID from @TPIBK_RecipeBatchData
	
	OPEN checkxml1
	-- Select first record
	FETCH NEXT FROM checkxml1
	INTO @steptype,@pcp,@ret,@matid,@pc
	WHILE @@FETCH_STATUS = 0 and @result=1
	BEGIN
	if not exists(select id from TPIBK_StepType where id =@steptype)
	begin
		set @result=4
	end
	if not exists(select id from ProcessClassPhase where id =@pcp) and not @pcp is null
	begin
		set @result=5
	end
	if not exists(select id from RecipeEquipmentTransition_Data where RecipeEquipmentTransition_ID =@ret) and not @ret is null
	begin
		set @result=6
	end
	if not exists(select id from Material where id =@matid) and not @matid is null
	begin
		set @result=7
	end
	--Move to next record
		FETCH NEXT FROM checkxml1
		INTO @steptype,@pcp,@ret,@matid,@pc
	END
	CLOSE checkxml1
	DEALLOCATE checkxml1	
	--check recipe step data--------------------------------
	declare @s as int
	DECLARE checkxml2 CURSOR FOR

	select TPIBK_RecipeParameterData_ID from @TPIBK_RecipeStepData
	
	OPEN checkxml2
	-- Select first record
	FETCH NEXT FROM checkxml2
	INTO @s
	WHILE @@FETCH_STATUS = 0 and @result=1
	BEGIN
	if not exists(select id from TPIBK_RecipeParameterData where id =@s)
	begin
		set @result=9
	end
	--Move to next record
		FETCH NEXT FROM checkxml2
		INTO @s
	END
	CLOSE checkxml2
	DEALLOCATE checkxml2	
	
	
	--no erros then do import
	If @result=1
	Begin
		Insert into Recipe select * from @Recipe
		--Insert into RecipeEquipmentRequirement select * from @RecipeEquipmentRequirement
		Declare @KeyTable TABLE (oKey INT ,nKey INT)

		Declare
			@InsEquipID		as int,
			@InsBatchID		as int,
			@InsStepID		as int,
			@NewEquipKey	as int,
			@oEquipKey		as int,
			@NewBatchKey	as int

		DECLARE InsEquip CURSOR FOR
		SELECT ID
		FROM @RecipeEquipmentRequirement

		OPEN InsEquip
		-- Select first record
		FETCH NEXT FROM InsEquip
		INTO @InsEquipID
		WHILE @@FETCH_STATUS = 0
		BEGIN
			/* Copy data in RecipeEquipmentRequirement */
			INSERT INTO RecipeEquipmentRequirement
			SELECT EquipmentType, Recipe_RID, Recipe_Version, ProcessClass_Name, Equipment_Name, LateBinding, IsMainBatchUnit,EqIdx1
			FROM @RecipeEquipmentRequirement
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
		
		DECLARE InsBatch CURSOR FOR
			SELECT ID, ProcessClass_ID
			from @TPIBK_RecipeBatchData
	
		OPEN InsBatch
		-- Select first record
	    FETCH NEXT FROM InsBatch
		INTO @InsBatchID, @oEquipKey
		WHILE @@FETCH_STATUS = 0
		BEGIN
		
			SET @NewEquipKey = (Select nKey from @KeyTable where oKey = @oEquipKey)
		
			INSERT INTO TPIBK_RecipeBatchData
			SELECT Recipe_RID, Recipe_Version,TPIBK_StepType_ID,ProcessClassPhase_ID,Step,UserString,
			RecipeEquipmentTransition_Data_ID,NextStep,Allocation_Type_ID,LateBinding,Material_ID,@NewEquipKey as ProcessClass_ID
			FROM @TPIBK_RecipeBatchData
			WHERE ID = @InsBatchID

			--SELECT SCOPE_IDENTITY() AS NewID
			SET @NewBatchKey = @@IDENTITY

			/* Copy data in TPIBK_RecipeStepData */
			DECLARE InsStep CURSOR FOR
				SELECT ID
				FROM @TPIBK_RecipeStepData
				WHERE TPIBK_RecipeBatchData_ID = @InsBatchID

			OPEN InsStep
			-- Select first record
			FETCH NEXT FROM InsStep
			into @InsStepID
			WHILE @@FETCH_STATUS = 0
			BEGIN		
				INSERT INTO TPIBK_RecipeStepData
				SELECT TPIBK_RecipeParameterData_ID, @NewBatchKey AS TPIBK_RecipeBatchData_ID, Value,CustomEU --[CustomEU] as EU
				FROM @TPIBK_RecipeStepData
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
	
	
		print 'good'
	END
	--debug
	select * from @Recipe
	select * from @RecipeEquipmentRequirement
	select * from @TPIBK_RecipeBatchData
	Select * from @TPIBK_RecipeStepData
	return( @result)

SET NOCOUNT OFF


