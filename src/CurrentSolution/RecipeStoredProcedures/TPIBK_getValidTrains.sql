USE [TPMDB]
GO
/****** Object:  StoredProcedure [dbo].[TPIBK_getValidTrains]    Script Date: 2/16/2023 12:06:17 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


ALTER PROCEDURE [dbo].[TPIBK_getValidTrains]



	@RID			as nvarchar(50),
	@Ver			as nvarchar(50),	
	@BatchTank		as nvarchar(50)   --Kris Mod 05/17/2022 where RecipeTrain.Name like '%' + @BatchTank + '%'

AS

SET NOCOUNT ON
	declare @train			as nvarchar(50)
	Declare @temp TABLE (name nvarchar(50))	
	declare @pc as nvarchar(50)
	declare @pc_count as int
	declare @ok as int


	
--find out what each recipe needs and how many
	select * into #req from(
		SELECT        ProcessClass_Name, COUNT(*) AS NumRequired 
		FROM            RecipeEquipmentRequirement
		WHERE        (Recipe_RID = @RID) AND (Recipe_Version = @Ver)
		GROUP BY ProcessClass_Name
		) as req


--get all the trains
	DECLARE TPIBK_getValidTrains_Trains CURSOR FOR
	select distinct(Name) from RecipeTrain where RecipeTrain.Name like '%' + @BatchTank + '%' order by Name ASC

	OPEN TPIBK_getValidTrains_Trains
	-- Select first record
	FETCH NEXT FROM TPIBK_getValidTrains_Trains
	INTO @train
	WHILE @@FETCH_STATUS = 0
	BEGIN

		drop table if exists #trainpc

		--what does each train have for process class and how many
		select * into #trainpc from(
		SELECT        COUNT(*) AS NumUsed, ProcessClass.Name as ProcessClass_Name
		FROM            Equipment INNER JOIN
								 RecipeTrainEquipment ON Equipment.ID = RecipeTrainEquipment.Equipment_ID INNER JOIN
								 RecipeTrain ON RecipeTrainEquipment.RecipeTrain_ID = RecipeTrain.ID INNER JOIN
								 ProcessClass ON Equipment.ProcessClass_ID = ProcessClass.ID
		WHERE        (RecipeTrain.Name = @train)
		GROUP BY ProcessClass.Name
		) as trainpc

		set @ok= 1

		--check that the train has all the process class and quantities required
		DECLARE TPIBK_getValidTrains_Check CURSOR FOR
		select ProcessClass_Name, NumRequired from #req 

		OPEN TPIBK_getValidTrains_Check
		-- Select first record
		FETCH NEXT FROM TPIBK_getValidTrains_Check
		INTO @pc,@pc_count
		WHILE @@FETCH_STATUS = 0
		BEGIN

			if not exists( select * from #trainpc where  ProcessClass_Name=@pc and NumUsed=@pc_count)
			begin
				set @ok=0
			end

			--Move to next record
			FETCH NEXT FROM TPIBK_getValidTrains_Check
			INTO @pc,@pc_count
		END
		CLOSE TPIBK_getValidTrains_Check
		DEALLOCATE TPIBK_getValidTrains_Check	

		if @ok=1 
		begin
			insert into @temp(name) values(@train) 
		end



		--Move to next record
		FETCH NEXT FROM TPIBK_getValidTrains_Trains
		INTO @train
	END
	CLOSE TPIBK_getValidTrains_Trains
	DEALLOCATE TPIBK_getValidTrains_Trains	



	select * from @temp


SET NOCOUNT OFF

