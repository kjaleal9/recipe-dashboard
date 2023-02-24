USE [TPMDB]
GO
/****** Object:  StoredProcedure [dbo].[TPIBK_getRecipeXML]    Script Date: 2/16/2023 12:05:33 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


ALTER PROCEDURE [dbo].[TPIBK_getRecipeXML]



	@RID			as nvarchar(50),
	@Ver			as nvarchar(50)

AS

SET NOCOUNT ON
DECLARE @XML XML

SET @XML =(SELECT [RID] as "@RID"
      ,[Version] as "@Version"
      ,[RecipeType] as "@RecipeType"
      ,[NbrOfExecutions] as "@NbrOfExecutions"
      ,[VersionDate] as "@VersionDate"
      ,[Description] as "@Description"
      ,[EffectiveDate] as "@EffectiveDate"
      ,[ExpirationDate] as "@ExpirationDate"
      ,[ProductID] as "@ProductID"
      ,[BatchSizeNominal] as "@BatchSizeNominal"
      ,[BatchSizeMin] as "@BatchSizeMin"
      ,[BatchSizeMax] as "@BatchSizeMax"
      ,[Status] as "@Status"
      ,[UseBatchKernel] as "@UseBatchKernel"
      ,[CurrentElementID] as "@CurrentElementID"
      ,'' as "@RecipeData"
      ,[RunMode] as "@RunMode"
      ,[IsPackagingRecipeType] as "@IsPackagingRecipeType"
      
      ,(SELECT [ID] as "@ID"
      ,[EquipmentType] as "@EquipmentType"
      ,[Recipe_RID] as  "@Recipe_RID"
      ,[Recipe_Version] as "@Recipe_Version"
      ,[ProcessClass_Name] as "@ProcessClass_Name"
      ,[Equipment_Name] as "@Equipment_Name"
      ,[LateBinding] as "@LateBinding"
      ,[IsMainBatchUnit] as "@IsMainBatchUnit"
      ,[EqIdx1] as "@EqIdx1"
		FROM RecipeEquipmentRequirement re
			where (re.Recipe_RID=r.rid) and (re.Recipe_Version=r.Version)				
			for xml path('RecipeEquipmentRequirement') ,type) as 'RecipeEquipmentRequirement_Rows'
			
	,(SELECT [ID] as "@ID"
      ,[Recipe_RID] as  "@Recipe_RID"
      ,[Recipe_Version] as "@Recipe_Version"
      ,[TPIBK_StepType_ID] as "@TPIBK_StepType_ID"
      ,[ProcessClassPhase_ID] as "@ProcessClassPhase_ID"
      ,[Step] as "@Step"
      ,[UserString] as "@UserString"
      ,[RecipeEquipmentTransition_Data_ID] as "@RecipeEquipmentTransition_Data_ID"
      ,[NextStep] as "@NextStep"
      ,[Allocation_Type_ID] as "@Allocation_Type_ID"
      ,[LateBinding] as  "@LateBinding"
      ,[Material_ID] as "@Material_ID"
      ,[ProcessClass_ID] as "@ProcessClass_ID"
		  ,(SELECT [ID] as "@ID"
		  ,[TPIBK_RecipeParameterData_ID] as "@TPIBK_RecipeParameterData_ID"
		  ,[TPIBK_RecipeBatchData_ID] as "@TPIBK_RecipeBatchData_ID"
		  ,[Value] as "@Value"
		  ,[CustomEU] as "@CustomEU"
	  FROM [TPMDB].[dbo].[TPIBK_RecipeStepData] sd
				where (sd.TPIBK_RecipeBatchData_ID=bd.id)			
				for xml path('TPIBK_RecipeStepData') ,type) as 'TPIBK_RecipeStepData_Rows'
  FROM [TPMDB].[dbo].[TPIBK_RecipeBatchData] bd
			where (bd.Recipe_RID=r.rid) and (bd.Recipe_Version=r.Version)				
			for xml path('TPIBK_RecipeBatchData') ,type) as 'TPIBK_RecipeBatchData_Rows'
      
  FROM [TPMDB].[dbo].[Recipe] R	 	
	where RID=@RID and Version=@Ver				
		for xml path('Recipe'),root('Recipes')
		)
	select @XML
		

SET NOCOUNT OFF


