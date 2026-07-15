SELECT
  MsOfficeEvents.Log,
  MsOfficeEvents.Category,
  MsOfficeEvents.Name
FROM
  MsOfficeEvents
WHERE
  (
    (
      (MsOfficeEvents.PowerPoint)= True
    )
  )
ORDER BY
  MsOfficeEvents.Category,
  MsOfficeEvents.Name;
