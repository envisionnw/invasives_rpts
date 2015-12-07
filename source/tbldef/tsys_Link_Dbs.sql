CREATE TABLE [tsys_Link_Dbs] (
  [Link_type] VARCHAR (255),
  [Link_db] VARCHAR (100),
  [Db_desc] VARCHAR (50),
  [Backups] BYTE ,
  [Is_ODBC] BYTE ,
  [Is_Network_db] BYTE ,
  [File_path] VARCHAR (255),
  [Server] VARCHAR (100),
  [New_db] VARCHAR (100),
  [New_path] VARCHAR (255),
  [New_server] VARCHAR (100),
  [Sort_order] SHORT ,
   CONSTRAINT [PrimaryKey] PRIMARY KEY ([Link_type], [Link_db])
)
