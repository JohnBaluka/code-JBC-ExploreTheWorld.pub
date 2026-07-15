SELECT
  MsOfficeEvents.Log,
  MsOfficeEvents.Category,
  MsOfficeEvents.Name
FROM
  MsOfficeEvents
WHERE
  (
    (
      (MsOfficeEvents.Word)= True
    )
  )
ORDER BY
  MsOfficeEvents.Category,
  MsOfficeEvents.Name;
