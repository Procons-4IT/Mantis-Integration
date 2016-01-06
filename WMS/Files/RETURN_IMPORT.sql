Insert Into LVRBODocumentHeader (lvrboh_DocNum,
lvrboh_CustomerCode,
lvrboh_DocDate,
lvrboh_DocDueDate,
lvrboh_PostingDate,
lvrboh_DocReference,
lvrboh_DocTotal,
lvrboh_Remarks,
lvrboh_FaxSent,
lvrboh_OrderType,
lvrboh_NoOfLines,
lvrboh_Warehouse,
lvrboh_Status)
Select Distinct Top 10 T0.DocNum,T0.CardCode,GetDate() As 'DD',GetDate() As 'DDD',GetDate() As 'PD',
DocNum,DocTotal,Comments
,'NO' As 'FS','SAMP' As 'OT',1 As 'NL',T1.WhsCode As 'WH',0 As 'Status'
From
[WMS].[dbo].ODLN T0 JOIN [WMS].[dbo].DLN1 T1 
ON T0.DocEntry = T1.DocEntry
Where T0.DocStatus = 'O'
And T1.ItemCode In
(
	Select ItemCode From [WMS].[dbo].OITM Where 
	ISNULL(ManSerNum,'N') = 'N'
	And 
	ISNULL(ManBtchNum,'N') = 'N'
)


Insert Into LVRBODocumentLines (
lvrbol_DocNum,
lvrbol_LineNo,
lvrbol_ItemCode,
lvrbol_Quantity,
lvrbol_Price,
lvrbol_Discount,
lvrbol_LineTotal,
lvrbol_BaseType,
lvrbol_BaseRef,
lvrbol_BaseLine,
lvrbol_IsClosing,
lvrbol_Status
)
Select T0.DocNum,T1.LineNum,T1.ItemCode,T1.Quantity,T1.Price,T1.DiscPrcnt,T1.LineTotal,'15' As 'BT',
T0.DocEntry,T1.LineNum,'N' As 'NL',0 As 'Status'
From
[WMS].[dbo].ODLN T0 JOIN [WMS].[dbo].DLN1 T1 On T0.DocEntry = T1.DocEntry
JOIN LVRBODocumentHeader T2 On T0.DocNum = T2.lvrboh_DocNum
