SELECT
  cns_Country.*,
  cns_Country.GUID AS PrimaryGUID
FROM
  cns_Country
ORDER BY
  cns_Country.Iso2;
