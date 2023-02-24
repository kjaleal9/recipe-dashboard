SELECT RER."ID", "ProcessClass_Name", 
      CASE coalesce(RER."Equipment_Name",PC."Description") 
          WHEN '' THEN PC."Description"
          ELSE RER."Equipment_Name" 
      END as Equipment_Name, 
                    
      CASE 
          WHEN ROW_NUMBER() over(partition by "ProcessClass_Name" order by "ProcessClass_Name")<2 
          THEN RER."ProcessClass_Name" 
          ELSE RER."ProcessClass_Name"
          -- ||' #'|| ltrim(ROW_NUMBER() over(partition by ProcessClass_Name order by ProcessClass_Name)) 
      END as message, PC."Description" As Description
                    
FROM public."RecipeEquipmentRequirement" AS RER
JOIN public."ProcessClass" AS PC ON RER."ProcessClass_Name" = PC."Name" 
WHERE RER."Recipe_RID"= 'BatchTanksTest' AND RER."Recipe_Version" = 1
ORDER BY Description;