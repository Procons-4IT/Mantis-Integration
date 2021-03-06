SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[LVPBODocumentHeader]') AND type in (N'U'))
BEGIN
CREATE TABLE [LVPBODocumentHeader](
	[lvpboh_DocNum] [varchar](20) NULL,
	[lvpboh_CustomerCode] [varchar](40) NULL,
	[lvpboh_DocDate] [date] NULL,
	[lvpboh_DocDueDate] [date] NULL,
	[lvpboh_PostingDate] [date] NULL,
	[lvpboh_DocReference] [varchar](100) NULL,
	[lvpboh_DocTotal] [numeric](19, 6) NULL,
	[lvpboh_Remarks] [varchar](255) NULL,
	[lvpboh_FaxSent] [varchar](10) NULL,
	[lvpboh_OrderType] [varchar](10) NULL,
	[lvpboh_NoOfLines] [numeric](12, 0) NULL,
	[lvpboh_Warehouse] [varchar](10) NULL,
	[lvpboh_Status] [int] NULL
) ON [PRIMARY]
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[LVPBODocumentLines]') AND type in (N'U'))
BEGIN
CREATE TABLE [LVPBODocumentLines](
	[lvpbol_DocNum] [varchar](20) NULL,
	[lvpbol_LineNo] [numeric](12, 0) NULL,
	[lvpbol_ItemCode] [varchar](40) NULL,
	[lvpbol_Quantity] [numeric](12, 0) NULL,
	[lvpbol_Price] [numeric](19, 6) NULL,
	[lvpbol_Discount] [numeric](19, 6) NULL,
	[lvpbol_LineTotal] [numeric](19, 6) NULL,
	[lvpbol_BaseType] [varchar](2) NULL,
	[lvpbol_BaseRef] [numeric](12, 0) NULL,
	[lvpbol_BaseLine] [numeric](12, 0) NULL,
	[lvpbol_IsClosing] [varchar](1) NULL,
	[lvpbol_Status] [int] NULL
) ON [PRIMARY]
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[LVRBODocumentHeader]') AND type in (N'U'))
BEGIN
CREATE TABLE [LVRBODocumentHeader](
	[lvrboh_DocNum] [varchar](20) NULL,
	[lvrboh_CustomerCode] [varchar](40) NULL,
	[lvrboh_DocDate] [date] NULL,
	[lvrboh_DocDueDate] [date] NULL,
	[lvrboh_PostingDate] [date] NULL,
	[lvrboh_DocReference] [varchar](100) NULL,
	[lvrboh_DocTotal] [numeric](19, 6) NULL,
	[lvrboh_Remarks] [varchar](255) NULL,
	[lvrboh_FaxSent] [varchar](10) NULL,
	[lvrboh_OrderType] [varchar](10) NULL,
	[lvrboh_NoOfLines] [numeric](12, 0) NULL,
	[lvrboh_Warehouse] [varchar](10) NULL,
	[lvrboh_Status] [int] NULL
) ON [PRIMARY]
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[LVRBODocumentLines]') AND type in (N'U'))
BEGIN
CREATE TABLE [LVRBODocumentLines](
	[lvrbol_DocNum] [varchar](20) NULL,
	[lvrbol_LineNo] [numeric](12, 0) NULL,
	[lvrbol_ItemCode] [varchar](40) NULL,
	[lvrbol_Quantity] [numeric](12, 0) NULL,
	[lvrbol_Price] [numeric](19, 6) NULL,
	[lvrbol_Discount] [numeric](19, 6) NULL,
	[lvrbol_LineTotal] [numeric](19, 6) NULL,
	[lvrbol_BaseType] [varchar](2) NULL,
	[lvrbol_BaseRef] [numeric](12, 0) NULL,
	[lvrbol_BaseLine] [numeric](12, 0) NULL,
	[lvrbol_IsClosing] [varchar](1) NULL,
	[lvrbol_Status] [int] NULL
) ON [PRIMARY]
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[LVSBODocumentHeader]') AND type in (N'U'))
BEGIN
CREATE TABLE [LVSBODocumentHeader](
	[lvsboh_DocNum] [varchar](20) NULL,
	[lvsboh_CustomerCode] [varchar](40) NULL,
	[lvsboh_DocDate] [date] NULL,
	[lvsboh_DocDueDate] [date] NULL,
	[lvsboh_PostingDate] [date] NULL,
	[lvsboh_DocReference] [varchar](100) NULL,
	[lvsboh_DocTotal] [numeric](19, 6) NULL,
	[lvsboh_Remarks] [varchar](255) NULL,
	[lvsboh_FaxSent] [varchar](10) NULL,
	[lvsboh_OrderType] [varchar](10) NULL,
	[lvsboh_NoOfLines] [numeric](12, 0) NULL,
	[lvsboh_Warehouse] [varchar](10) NULL,
	[lvsboh_Status] [int] NULL
) ON [PRIMARY]
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[LVSBODocumentLines]') AND type in (N'U'))
BEGIN
CREATE TABLE [LVSBODocumentLines](
	[lvsbol_DocNum] [varchar](20) NULL,
	[lvsbol_LineNo] [numeric](12, 0) NULL,
	[lvsbol_ItemCode] [varchar](40) NULL,
	[lvsbol_Quantity] [numeric](12, 0) NULL,
	[lvsbol_Price] [numeric](19, 6) NULL,
	[lvsbol_Discount] [numeric](19, 6) NULL,
	[lvsbol_LineTotal] [numeric](19, 6) NULL,
	[lvsbol_BaseType] [varchar](2) NULL,
	[lvsbol_BaseRef] [numeric](12, 0) NULL,
	[lvsbol_BaseLine] [numeric](12, 0) NULL,
	[lvsbol_IsClosing] [varchar](1) NULL,
	[lvsbol_Status] [int] NULL
) ON [PRIMARY]
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[LVSBODocumentLines]') AND type in (N'U'))
BEGIN
CREATE TABLE [LVSBODocumentLines](
	[lvsbol_DocNum] [varchar](20) NULL,
	[lvsbol_LineNo] [numeric](12, 0) NULL,
	[lvsbol_ItemCode] [varchar](40) NULL,
	[lvsbol_Quantity] [numeric](12, 0) NULL,
	[lvsbol_Price] [numeric](19, 6) NULL,
	[lvsbol_Discount] [numeric](19, 6) NULL,
	[lvsbol_LineTotal] [numeric](19, 6) NULL,
	[lvsbol_BaseType] [varchar](2) NULL,
	[lvsbol_BaseRef] [numeric](12, 0) NULL,
	[lvsbol_BaseLine] [numeric](12, 0) NULL,
	[lvsbol_IsClosing] [varchar](1) NULL,
	[lvsbol_Status] [int] NULL
) ON [PRIMARY]
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[LVSBODocumentHeader]') AND type in (N'U'))
BEGIN
CREATE TABLE [LVSBODocumentHeader](
	[lvsboh_DocNum] [varchar](20) NULL,
	[lvsboh_CustomerCode] [varchar](40) NULL,
	[lvsboh_DocDate] [date] NULL,
	[lvsboh_DocDueDate] [date] NULL,
	[lvsboh_PostingDate] [date] NULL,
	[lvsboh_DocReference] [varchar](100) NULL,
	[lvsboh_DocTotal] [numeric](19, 6) NULL,
	[lvsboh_Remarks] [varchar](255) NULL,
	[lvsboh_FaxSent] [varchar](10) NULL,
	[lvsboh_OrderType] [varchar](10) NULL,
	[lvsboh_NoOfLines] [numeric](12, 0) NULL,
	[lvsboh_Warehouse] [varchar](10) NULL,
	[lvsboh_Status] [int] NULL
) ON [PRIMARY]
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[LVRBODocumentLines]') AND type in (N'U'))
BEGIN
CREATE TABLE [LVRBODocumentLines](
	[lvrbol_DocNum] [varchar](20) NULL,
	[lvrbol_LineNo] [numeric](12, 0) NULL,
	[lvrbol_ItemCode] [varchar](40) NULL,
	[lvrbol_Quantity] [numeric](12, 0) NULL,
	[lvrbol_Price] [numeric](19, 6) NULL,
	[lvrbol_Discount] [numeric](19, 6) NULL,
	[lvrbol_LineTotal] [numeric](19, 6) NULL,
	[lvrbol_BaseType] [varchar](2) NULL,
	[lvrbol_BaseRef] [numeric](12, 0) NULL,
	[lvrbol_BaseLine] [numeric](12, 0) NULL,
	[lvrbol_IsClosing] [varchar](1) NULL,
	[lvrbol_Status] [int] NULL
) ON [PRIMARY]
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[LVRBODocumentHeader]') AND type in (N'U'))
BEGIN
CREATE TABLE [LVRBODocumentHeader](
	[lvrboh_DocNum] [varchar](20) NULL,
	[lvrboh_CustomerCode] [varchar](40) NULL,
	[lvrboh_DocDate] [date] NULL,
	[lvrboh_DocDueDate] [date] NULL,
	[lvrboh_PostingDate] [date] NULL,
	[lvrboh_DocReference] [varchar](100) NULL,
	[lvrboh_DocTotal] [numeric](19, 6) NULL,
	[lvrboh_Remarks] [varchar](255) NULL,
	[lvrboh_FaxSent] [varchar](10) NULL,
	[lvrboh_OrderType] [varchar](10) NULL,
	[lvrboh_NoOfLines] [numeric](12, 0) NULL,
	[lvrboh_Warehouse] [varchar](10) NULL,
	[lvrboh_Status] [int] NULL
) ON [PRIMARY]
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[LVPBODocumentLines]') AND type in (N'U'))
BEGIN
CREATE TABLE [LVPBODocumentLines](
	[lvpbol_DocNum] [varchar](20) NULL,
	[lvpbol_LineNo] [numeric](12, 0) NULL,
	[lvpbol_ItemCode] [varchar](40) NULL,
	[lvpbol_Quantity] [numeric](12, 0) NULL,
	[lvpbol_Price] [numeric](19, 6) NULL,
	[lvpbol_Discount] [numeric](19, 6) NULL,
	[lvpbol_LineTotal] [numeric](19, 6) NULL,
	[lvpbol_BaseType] [varchar](2) NULL,
	[lvpbol_BaseRef] [numeric](12, 0) NULL,
	[lvpbol_BaseLine] [numeric](12, 0) NULL,
	[lvpbol_IsClosing] [varchar](1) NULL,
	[lvpbol_Status] [int] NULL
) ON [PRIMARY]
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[LVPBODocumentHeader]') AND type in (N'U'))
BEGIN
CREATE TABLE [LVPBODocumentHeader](
	[lvpboh_DocNum] [varchar](20) NULL,
	[lvpboh_CustomerCode] [varchar](40) NULL,
	[lvpboh_DocDate] [date] NULL,
	[lvpboh_DocDueDate] [date] NULL,
	[lvpboh_PostingDate] [date] NULL,
	[lvpboh_DocReference] [varchar](100) NULL,
	[lvpboh_DocTotal] [numeric](19, 6) NULL,
	[lvpboh_Remarks] [varchar](255) NULL,
	[lvpboh_FaxSent] [varchar](10) NULL,
	[lvpboh_OrderType] [varchar](10) NULL,
	[lvpboh_NoOfLines] [numeric](12, 0) NULL,
	[lvpboh_Warehouse] [varchar](10) NULL,
	[lvpboh_Status] [int] NULL
) ON [PRIMARY]
END
GO
