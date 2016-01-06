Insert Into LVPBODocumentHeader (lvpboh_DocNum,
lvpboh_CustomerCode,
lvpboh_DocDate,
lvpboh_DocDueDate,
lvpboh_PostingDate,
lvpboh_DocReference,
lvpboh_DocTotal,
lvpboh_Remarks,
lvpboh_FaxSent,
lvpboh_OrderType,
lvpboh_NoOfLines,
lvpboh_Warehouse,
lvpboh_Status)
Select Distinct Top 10 T0.DocNum,T0.CardCode,GetDate() As 'DD',GetDate() As 'DDD',GetDate() As 'PD',
DocNum,DocTotal,Comments
,'NO' As 'FS','SAMP' As 'OT',1 As 'NL',T1.WhsCode As 'WH',0 As 'Status'
From
[WMS].[dbo].OPOR T0 JOIN [WMS].[dbo].POR1 T1 
ON T0.DocEntry = T1.DocEntry
Where T0.DocStatus = 'O'
And T1.ItemCode In
(
	Select ItemCode From [WMS].[dbo].OITM Where 
	ISNULL(ManSerNum,'N') = 'N'
	And 
	ISNULL(ManBtchNum,'N') = 'N'
)


Insert Into LVPBODocumentLines (
lvpbol_DocNum,
lvpbol_LineNo,
lvpbol_ItemCode,
lvpbol_Quantity,
lvpbol_Price,
lvpbol_Discount,
lvpbol_LineTotal,
lvpbol_BaseType,
lvpbol_BaseRef,
lvpbol_BaseLine,
lvpbol_IsClosing,
lvpbol_Status
)
Select T0.DocNum,T1.LineNum,T1.ItemCode,T1.Quantity,T1.Price,T1.DiscPrcnt,T1.LineTotal,'15' As 'BT',
T0.DocEntry,T1.LineNum,'N' As 'NL',0 As 'Status'
From
[WMS].[dbo].OPOR T0 JOIN [WMS].[dbo].POR1 T1 On T0.DocEntry = T1.DocEntry
JOIN LVPBODocumentHeader T2 On T0.DocNum = T2.lvpboh_DocNum
