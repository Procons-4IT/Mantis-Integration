Insert Into LVSBODocumentHeader (lvsboh_DocNum,
lvsboh_CustomerCode,
lvsboh_DocDate,
lvsboh_DocDueDate,
lvsboh_PostingDate,
lvsboh_DocReference,
lvsboh_DocTotal,
lvsboh_Remarks,
lvsboh_FaxSent,
lvsboh_OrderType,
lvsboh_NoOfLines,
lvsboh_Warehouse,
lvsboh_Status)
Select Distinct Top 10 T0.DocNum,T0.CardCode,GetDate() As 'DD',GetDate() As 'DDD',GetDate() As 'PD',
DocNum,DocTotal,Comments
,'NO' As 'FS','SAMP' As 'OT',1 As 'NL',T1.WhsCode As 'WH',0 As 'Status'
From
[WMS].[dbo].ORDR T0 JOIN [WMS].[dbo].RDR1 T1 
ON T0.DocEntry = T1.DocEntry
Where T0.DocStatus = 'O'
And T1.ItemCode In
(
	Select ItemCode From [WMS].[dbo].OITM Where 
	ISNULL(ManSerNum,'N') = 'N'
	And 
	ISNULL(ManBtchNum,'N') = 'N'
) 

Insert Into LVSBODocumentLines (
lvsbol_DocNum,
lvsbol_LineNo,
lvsbol_ItemCode,
lvsbol_Quantity,
lvsbol_Price,
lvsbol_Discount,
lvsbol_LineTotal,
lvsbol_BaseType,
lvsbol_BaseRef,
lvsbol_BaseLine,
lvsbol_IsClosing,
lvsbol_Status
)
Select T0.DocNum,T1.LineNum,T1.ItemCode,T1.Quantity,T1.Price,T1.DiscPrcnt,T1.LineTotal,'15' As 'BT',
T0.DocEntry,T1.LineNum,'N' As 'NL',0 As 'Status'
From
[WMS].[dbo].ORDR T0 JOIN [WMS].[dbo].RDR1 T1 On T0.DocEntry = T1.DocEntry
JOIN LVSBODocumentHeader T2 On T0.DocNum = T2.lvsboh_DocNum
