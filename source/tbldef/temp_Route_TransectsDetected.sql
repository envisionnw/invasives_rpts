CREATE TABLE [temp_Route_TransectsDetected] (
  [Unit_Code] VARCHAR (4),
  [Visit_Year] SHORT ,
  [Route] VARCHAR (255),
  [Area] VARCHAR (255),
  [PlantCode] VARCHAR (50),
  [IsDead] SHORT ,
  [TransectsDetected] LONG ,
   CONSTRAINT [PrimaryKey] PRIMARY KEY ([Unit_Code], [Visit_Year], [Route], [PlantCode], [IsDead])
)
