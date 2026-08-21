SELECT
  TBLPARAMETRES.ParType,
  Left([ParType], 3) AS ParType1,
  Right([ParType], 3) AS ParType2,
  TBLPARAMETRES.ParCode,
  Left(
    [ParCode],
    InStr([ParCode], "=") - 1
  ) AS ParCode1,
  Left(
    RemplaceChr([ParCode], [ParCode1] & "=", ""),
    InStr(
      RemplaceChr([ParCode], [ParCode1] & "=", ""),
      "="
    ) - 1
  ) AS ParCode2,
  Left(
    RemplaceChr(
      [ParCode], [ParCode1] & "=" & [ParCode2] & "=",
      ""
    ),
    InStr(
      RemplaceChr(
        [ParCode], [ParCode1] & "=" & [ParCode2] & "=",
        ""
      ),
      "="
    ) - 1
  ) AS ParCode3,
  Left(
    RemplaceChr(
      [ParCode], [ParCode1] & "=" & [ParCode2] & "=" & [ParCode3] & "=",
      ""
    ),
    InStr(
      RemplaceChr(
        [ParCode], [ParCode1] & "=" & [ParCode2] & "=" & [ParCode3] & "=",
        ""
      ),
      "="
    ) - 1
  ) AS ParCode4,
  RemplaceChr(
    [ParCode], [ParCode1] & "=" & [ParCode2] & "=" & [ParCode3] & "=" & [ParCode4] & "=",
    ""
  ) AS ParCode5,
  TBLPARAMETRES.ParValeur,
  TBLPARAMETRES.ParValide
FROM
  TBLPARAMETRES;
