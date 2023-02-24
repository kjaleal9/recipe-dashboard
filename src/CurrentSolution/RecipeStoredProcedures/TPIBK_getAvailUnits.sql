USE [TPMDB]
GO
/****** Object:  StoredProcedure [dbo].[TPIBK_getAvailUnits]    Script Date: 2/16/2023 12:04:30 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
ALTER PROCEDURE [dbo].[TPIBK_getAvailUnits]

	@RID	as nvarchar(50),
	@Ver	as nvarchar(50)

AS
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

	SELECT DISTINCT (PC.[ID]) ID,PC.[Name] Name, PC.[Description] Description 
	FROM (
		SELECT PC.[Name], Count(PC.[Name]) Qty
		FROM [Equipment] Eq
		Join [ProcessClass] PC On PC.ID = Eq.ProcessClass_ID
		GROUP BY  PC.[Name]
		EXCEPT
		SELECT [ProcessClass_Name] Name, Count(ProcessClass_Name) Qty
		FROM [v_RecipeEquipmentRequirement]
		WHERE Recipe_RID= @RID AND Recipe_Version=@Ver
		GROUP BY [ProcessClass_Name]
		) t
	Join [ProcessClass] PC On PC.Name = t.Name
	WHERE PC.[TypeBatchKernel] = 1
	ORDER BY Description

	SET NOCOUNT OFF



