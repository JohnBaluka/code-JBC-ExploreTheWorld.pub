SELECT
  MsOfficeEvents.Log,
  MsOfficeEvents.Category,
  MsOfficeEvents.Name
FROM
  MsOfficeEvents
WHERE
  (
    (
      (MsOfficeEvents.Excel)= True
    )
  )
ORDER BY
  MsOfficeEvents.Category,
  MsOfficeEvents.Name;
