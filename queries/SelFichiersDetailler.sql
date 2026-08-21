SELECT
  TBLFICHIERS.FicType,
  Left([FicType], 3) AS FicType1,
  Mid([FicType], 4) AS FicType2,
  TBLFICHIERS.FicCode,
  Left(
    [FicCode],
    InStr([FicCode], "=") - 1
  ) AS FicCode1,
  Right(
    [FicCode],
    Len([FicCode]) - InStr([FicCode], "=")
  ) AS FicCode2,
  TBLFICHIERS.FicValeur,
  TBLFICHIERS.FicValide
FROM
  TBLFICHIERS;
