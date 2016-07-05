Select  [Resource].[ResourceID] as [Resource ID],
        [LaborActivity].[JobHead_JobNum] as [JobHead_JobNum],
        [LaborActivity].[JobHead_PartNum] as [JobHead_PartNum],
        [LaborActivity].[JobHead_ProdQty] as [JobHead_ProdQty],
        [LaborActivity].[JobHead_ReqDueDate] as [JobHead_ReqDueDate],
        [LaborActivity].[OpNum] as [OpNum],
        [LaborActivity].[EmpBasic_FirstName] as [EmpBasic_FirstName],
        [LaborActivity].[LaborDtl_ClockInDate] as [LaborDtl_ClockInDate],
        [LaborActivity].[Resource_Description] as [Resource_Description],
        [LaborActivity].[LaborDtl_ClockinTime]/24 as [LaborDtl_ClockinTime],
        [LaborActivity].[LaborDtl_LaborType] as [LaborDtl_LaborType],
        [LaborActivity].[JobOpDtl_ProdStandard] as [JobOpDtl_ProdStandard],
        [LaborActivity].[LaborHistory_TotalLabor] as [LaborHistory_TotalLabor]

FROM Erp.Resource as Resource
 FULL OUTER JOIN (SELECT
  [JobHead].[JobNum] as [JobHead_JobNum],
  [JobHead].[PartNum] as [JobHead_PartNum],
  [JobHead].[ProdQty] as [JobHead_ProdQty],
  [JobHead].[ReqDueDate] as [JobHead_ReqDueDate],
  [LaborDtl].[OprSeq] as [OpNum],
  [EmpBasic].[FirstName] as [EmpBasic_FirstName],
  [LaborDtl].[ClockInDate] as [LaborDtl_ClockInDate],
  [LaborDtl].[ResourceID] as [LaborDtl_ResourceID],
  [Resource].[Description] as [Resource_Description],
  [Resource].[ResourceType] as [ResourceType],
  [LaborDtl].[ClockinTime]/24 as [LaborDtl_ClockinTime],
  [LaborDtl].[LaborType] as [LaborDtl_LaborType],
  [JobOpDtl].[ProdStandard] as [JobOpDtl_ProdStandard],
  [LaborHistory].[TotalLabor] as [LaborHistory_TotalLabor]

FROM Erp.LaborDtl as LaborDtl
  INNER JOIN Erp.JobHead as JobHead
    ON   	LaborDtl.Company = JobHead.Company
    AND  	LaborDtl.JobNum = JobHead.JobNum

  INNER JOIN Erp.EmpBasic as EmpBasic
    ON   	LaborDtl.Company = EmpBasic.Company
    AND  	LaborDtl.EmployeeNum = EmpBasic.EmpID
  INNER JOIN Erp.Resource as Resource
    ON   	LaborDtl.Company = Resource.Company
    AND  	LaborDtl.ResourceID = Resource.ResourceID
  INNER JOIN Erp.JobOpDtl as JobOpDtl
    ON   	LaborDtl.Company = JobOpDtl.Company
    AND  	LaborDtl.JobNum = JobOpDtl.JobNum
    AND  	LaborDtl.OprSeq = JobOpDtl.OprSeq

  LEFT JOIN ( --LABOR HISTORY SUB QUERY
        SELECT
        l.JobNum,
        l.OprSeq ,
        SUM(l.LaborQty)  as TotalLabor,
        jh.PartNum

        FROM Erp.LaborDtl l
          INNER JOIN Erp.JobHead jh
            ON  l.Company = jh.Company
            AND l.JobNum = jh.JobNum

        GROUP BY l.JobNum, l.OprSeq, jh.PartNum
        HAVING COUNT(CASE ActiveTrans WHEN 1 THEN 1 END) >0
            ) AS LaborHistory

    ON LaborDtl.JobNum = LaborHistory.JobNum
    AND LaborDtl.OprSeq = LaborHistory.OprSeq
  --Active transactions with a clock in date of today
  WHERE LaborDtl.ActiveTrans = 1 AND convert(varchar(10), LaborDtl.ClockInDate, 102)
    = convert(varchar(10), getdate(), 102)) as LaborActivity

    ON Resource.ResourceID = LaborActivity.LaborDtl_ResourceID

  WHERE Resource.ResourceType = 'M'