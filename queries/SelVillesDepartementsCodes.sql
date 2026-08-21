SELECT
  TBLVILLES.*,
  IIf(
    Left([VilCodePostal], 2) = "97",
    Left([VilCodePostal], 3),
    Left([VilCodePostal], 2)
  ) AS VilDepCode
FROM
  TBLVILLES;
