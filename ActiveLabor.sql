Select  [Resource].[ResourceID] as [Resource ID],
        [LaborActivity].[JobHead_JobNum] as [JobNum],
        [LaborActivity].[JobHead_PartNum] as [PartNum],
        [LaborActivity].[JobHead_ProdQty] as [ProdQty],
        [LaborActivity].[JobHead_ReqDueDate] as [ReqDueDate],
        [LaborActivity].[OpNum] as [OpNum],
        [LaborActivity].[EmpBasic_FirstName] as [Employee],
        [LaborActivity].[LaborDtl_ClockInDate] as [ClockInDate],
        [LaborActivity].[Resource_Description] as [ResourceDescription],
        [LaborActivity].[LaborDtl_ClockinTime] as [ClockinTime],
        [LaborActivity].[LaborDtl_LaborType] as [LaborType],
        [LaborActivity].[JobOpDtl_ProdStandard] as [ProdStandard],
        [LaborActivity].[LaborHistory_TotalLabor] as [TotalLabor],
        [LaborActivity].[SetupTime] as [SetupTime],
        [LaborActivity].[ProdTime] as [ProdTime]

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
  [JobOpDtl].[EstSetHoursPerMch] as [SetupTime],
  [JobOpDtl].[EstProdHours] as [ProdTime],
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
     OR Resource.ResourceType = 'P'