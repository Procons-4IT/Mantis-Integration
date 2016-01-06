--(SALE ORDER)-->Out Bound
Select * From SBOLVDocumentHeader Where 1 = 2
Select * From SBOLVDocumentLines Where 1 = 2
--(SALE ORDER --> Delivery)--> In Bound
Select * From LVSBODocumentHeader Where 1 = 2
Select * From LVSBODocumentLines Where 1 = 2
--(SALE RETURN --> Delivery)--> In Bound
Select * From LVRBODocumentHeader Where 1 = 2
Select * From LVRBODocumentLines Where 1 = 2
--(A/R Reserve Invoice)--> Out Bound
Select * From SBORIDocumentHeader Where 1 = 2
Select * From SBORIDocumentLines Where 1 = 2
--(Purchase Orders)--> Out Bound
Select * From PBOLVDocumentHeader Where 1 = 2
Select * From PBOLVDocumentLines Where 1 = 2
--(Purchase Orders--> GRPO) --> In Bound
Select * From LVPBODocumentHeader Where 1 = 2
Select * From LVPBODocumentLines   Where 1 = 2
--(Inventory Request) --> Out Bound
Select * From TBOLVDocumentHeader Where 1 = 2
Select * From TBOLVDocumentLines Where 1 = 2
--(Inventory Request --> Inventory Trasfer)-In Bound
Select * From LVTBODocumentHeader Where 1 = 2
Select * From LVTBODocumentLines Where 1 = 2
--(Inventory Goods Receipt) --> In Bound
Select * From LVIBODocument Where 1 = 2
--(Inventory Goods Issue) --> In Bound
Select * From LVOBODocument Where 1 = 2
--(Inventory Count) --> In Bound
Select * From LVCBODocument Where 1 = 2

