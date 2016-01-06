Insert Into LVOBODocument (
lvobo_DocNum,
lvobo_DocDate,
lvobo_DocDueDate,
lvobo_PostingDate,
lvobo_DocReference,
lvobo_Remarks,
lvobo_ItemCode,
lvobo_Quantity,
lvobo_Price,
lvobo_Discount,
lvobo_Warehouse,
lvobo_Status)
Select T0.DocNum,GetDate(),GetDate(),GetDate(),T0.DocNum,T0.Comments
,T1.ItemCode,T1.Quantity,T1.Price,T1.DiscPrcnt,
Filler,0 As 'Status'
From
[WMS].[dbo].OWTQ T0 JOIN [WMS].[dbo].WTQ1 T1 On T0.DocEntry = T1.DocEntry

