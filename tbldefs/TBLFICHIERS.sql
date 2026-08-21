CREATE TABLE [TBLFICHIERS] (
  [FicType] VARCHAR (8),
  [FicCode] VARCHAR (100),
  [FicValeur] VARCHAR (200),
  [FicValide] BIT,
   CONSTRAINT [PrimaryKey] PRIMARY KEY ([FicType], [FicCode])
)
