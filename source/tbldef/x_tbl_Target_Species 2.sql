CREATE TABLE [x_tbl_Target_Species 2] (
  [Tgt_Species_ID] AUTOINCREMENT,
  [Master_Plant_Code_FK] VARCHAR (20),
  [Park_Code] VARCHAR (4),
  [Target_Year] SHORT ,
  [Species_Name] VARCHAR (255),
  [LU_Code] VARCHAR (20),
  [Priority] SHORT ,
  [Transect_Only] BYTE ,
  [Target_Area_ID] SHORT ,
  [Comments] VARCHAR (255)
)
