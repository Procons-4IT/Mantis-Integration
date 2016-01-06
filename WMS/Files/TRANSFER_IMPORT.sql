Insert Into LVTBODocumentHeader (lvtboh_DocNum,
lvtboh_CustomerCode,
lvtboh_DocDate,
lvtboh_DocDueDate,
lvtboh_PostingDate,
lvtboh_DocReference,
lvtboh_Remarks,
lvtboh_FaxSent,
lvtboh_OrderType,
lvtboh_NoOfLines,
lvtboh_FWarehouse,
lvtboh_TWarehouse,
lvtboh_Status)
Select Top 5 DocNum,CardCode,GetDate() As 'DD',GetDate() As 'DDD',GetDate() As 'PD',
DocNum,Comments
,'NO' As 'FS','SAMP' As 'OT',1 As 'NL',Filler,ToWhsCode,0 As 'Status'
From
[WMS].[dbo].OWTQ Where DocStatus = 'O'
Order By DocEntry Desc

Insert Into LVTBODocumentLines (
lvtbol_DocNum,
lvtbol_LineNo,
lvtbol_ItemCode,
lvtbol_Quantity,
lvtbol_Price,
lvtbol_Discount,
lvtbol_LineTotal,
lvtbol_BaseType,
lvtbol_BaseRef,
lvtbol_BaseLine,
lvtbol_IsClosing,
lvtbol_Status
)
Select T0.DocNum,T1.LineNum,T1.ItemCode,T1.Quantity,T1.Price,T1.DiscPrcnt,T1.LineTotal,'15' As 'BT',
T0.DocEntry,T1.LineNum,'N' As 'NL',0 As 'Status'
From
[WMS].[dbo].OWTQ T0 JOIN [WMS].[dbo].WTQ1 T1 On T0.DocEntry = T1.DocEntry
JOIN LVTBODocumentHeader T2 On T0.DocNum = T2.lvtboh_DocNum
