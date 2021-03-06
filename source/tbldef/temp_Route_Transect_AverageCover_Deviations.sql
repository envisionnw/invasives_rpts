CREATE TABLE [temp_Route_Transect_AverageCover_Deviations] (
  [Unit_Code] VARCHAR (4),
  [Visit_Year] SHORT ,
  [Route] VARCHAR (255),
  [Transect] SHORT ,
  [Area] VARCHAR (50),
  [E_Coord] DOUBLE ,
  [N_Coord] DOUBLE ,
  [PlantCode] VARCHAR (50),
  [Species] VARCHAR (255),
  [Master_Common_Name] VARCHAR (50),
  [IsDead] SHORT ,
  [TransectsSampled] DOUBLE ,
  [TotalCover] DOUBLE ,
  [TransectAverageCover] DOUBLE ,
  [RouteAverageCover] DOUBLE ,
  [Deviation] DOUBLE ,
  [DeviationSquared] DOUBLE 
)
