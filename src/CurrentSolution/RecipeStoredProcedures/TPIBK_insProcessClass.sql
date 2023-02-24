USE [TPMDB]
GO
/****** Object:  StoredProcedure [dbo].[TPIBK_insProcessClass]    Script Date: 2/16/2023 12:07:28 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
ALTER PROCEDURE [dbo].[TPIBK_insProcessClass] 

	@RID	as nvarchar(50),
	@Ver	as nvarchar(50),
	@PC		as nvarchar(50),
	@EN		as nvarchar(50)
AS

DECLARE

	@Index as nvarchar (50)

	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

	-- Is the Process Class already existing?
	IF EXISTS(SELECT ID FROM RecipeEquipmentRequirement RER WHERE Recipe_RID = @RID And Recipe_Version = @Ver and ProcessClass_Name = @PC) 
		-- Yes - increment Process Class index
		BEGIN
			SET @Index = '_' + CAST((SELECT Max(EN) FROM (SELECT row_number() over (partition by ProcessClass_Name Order By RER.ID) as EN FROM RecipeEquipmentRequirement RER WHERE Recipe_RID = @RID And Recipe_Version = @Ver and ProcessClass_Name = @PC) t) + 1 As varchar)
		END
	ELSE
		-- No - Set Process Class index to 1
		BEGIN
			SET @Index = '' 
		END

	-- Insert record into table
	INSERT INTO RecipeEquipmentRequirement(EquipmentType , Recipe_RID , Recipe_Version , Equipment_Name , ProcessClass_Name , LateBinding , IsMainBatchUnit) values( 'Class', @RID, @Ver,@EN + @Index, @PC, 0, 0)



