CREATE TABLE [temp_Listbox_Recordset] (
  [Code] VARCHAR (255) CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [Species] VARCHAR (255),
  [LUCode] VARCHAR (255),
  [Transect_Only] SHORT ,
  [Target_Area_ID] SHORT 
)
