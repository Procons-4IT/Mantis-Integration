Insert Into LVCBODocument (
lvcbo_DocNum,
lvcbo_CountDate,
lvcbo_DocReference,
lvcbo_ItemCode,
lvcbo_UOMEntry,
lvcbo_UOMName,
lvcbo_Quantity,
lvcbo_Warehouse,
lvcbo_Remarks,
lvcbo_Status)
Select T0.DocNum,GetDate(),T0.DocNum
,T1.ItemCode,T3.UomEntry,T3.UomCode,T1.Quantity
,Filler,T0.Comments,0 As 'Status'
From
[WMS].[dbo].OWTQ T0 JOIN [WMS].[dbo].WTQ1 T1 On T0.DocEntry = T1.DocEntry
JOIN [WMS].[dbo].OITM T2 On T2.ItemCode = T1.ItemCode
JOIN [WMS].[dbo].OUOM T3 On T2.UgpEntry = T3.UomEntry

Select ItemCode,ItemName,UgpEntry,PUomEntry,SUomEntry,IUomEntry,CntUnitMsr,INUomEntry From [WMS].[dbo].OITM

Select * From [WMS].[dbo].OUOM

