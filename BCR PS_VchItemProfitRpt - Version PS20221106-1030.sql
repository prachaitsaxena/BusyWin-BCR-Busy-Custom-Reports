  SELECT TableItemsInTransaction.[VchCode]
 ,TableItemsInTransaction.[Date] As TIITDate
 ,LTRIM(RTRIM(TableItemsInTransaction.[VchNo])) As TIITVchNo
 ,TableItemsInTransaction.[MasterCode2] As TIITMCCode
 ,TableMaterialCenterMaster.[Name] As TMCMName
 ,TableItemsInTransaction.[CM6] As TIITSalesmanCode
 ,ISNULL(TableSalesmanMaster.[Name],'') As TSMName
 ,TableItemsInTransaction.[CM1] As TIITPartyMasterCode
 ,TablePartyMaster.[Name] As TPMName
 ,TablePartyMaster.[Alias] As TPMAlias
 ,TableAccountMasterDetailedInfo.[OF2] As TAMAIZoneCode
 ,ISNull(TableZoneMaster.[Name],'') As TZMName
 ,TableAccountMasterDetailedInfo.[OF3] As TAMAITerritoryCode
 ,ISNULL(TableTerritoryMaster.[Name],'') As TTMName
 ,TableAccountMasterDetailedInfo.[OF4] As TAMAICityCode
 ,ISNULL(TableCityMaster.[Name],'') As TCMName
 ,TableAccountMasterDetailedInfo.[PINCode] As PINCode
 ,TableAccountMasterDetailedInfo.[Contact] As Contact
 ,TableAccountMasterDetailedInfo.[WhatsAppNo] As WhatsAppNo
 ,TableAccountMasterDetailedInfo.[OF5] As TAMAITerritoryManagerCode
 ,ISNULL(TableTerritoryManagerMaster.[Name],'') As TTMMName
 ,TableAccountMasterDetailedInfo.[OF6] As TAMAIAccountManagerCode
 ,ISNULL(TableAccountManagerMaster.[Name],'') As TAMMName
 ,TableItemsInTransaction.[MasterCode1] As TIITItemCode
 ,TableItemMaster.[Name] As TIMName
 ,TableItemMaster.[Alias] As TIMAlias
 ,TableItemMaster.[ParentGrp] As TIMBusyParentGroupCode
 ,TableBusyGroupNameMaster.[Name] As TBPGNMName
 ,TableItemMasterDetailedInfo.[OF2] As TIMAIParentGroupCode
 ,ISNULL(TableParentGroupNameMaster.[Name],'') As TPGNMName
 ,TableItemMasterDetailedInfo.[OF3] As TIMAIGroupCode
 ,ISNULL(TableGroupNameMaster.[Name],'') As TGNMName
 ,TableItemMasterDetailedInfo.[OF4] As TIMAIManufacturingBrandCode
 ,ISNULL(TableManufacturingBrandMaster.[Name],'') As TMBMName
 ,ISNULL(TableItemMasterDetailedInfo.[OF5],'') As TIMAIConsumberBrandCode
 ,TableConsumerBrandMaster.[Name] As TCBMName
 ,LTRIM(RTRIM(ISNULL(TableItemSerialNo.[SerialNo],''))) As TISNSerialNo
 ,IIF(ISNULL(TableItemSerialNo.[SerialNo],1)='1',(0-TableItemsInTransaction.[Value1]),1) AS QtySold
 ,ISNULL((SELECT TOP 1 D1 FROM ItemSerialNo AS ItemSerialNoForPurchasePrice 
	WHERE ItemSerialNoForPurchasePrice.[ItemCode] = TableItemsInTransaction.[MasterCode1] 
	AND ItemSerialNoForPurchasePrice.[SerialNo] = TableItemSerialNo.[SerialNo]
	AND ItemSerialNoForPurchasePrice.[VchType] = '2' 
	Order By ItemSerialNoForPurchasePrice.[Date] ASC, ItemSerialNoForPurchasePrice.[VchCode] DESC),'') As BusyPurchasePrice
 ,ISNULL(TableItemSerialNo.[C3],'') As TISNOptionalFieldPurchasePrice
 ,TableItemsInTransaction.[D2] As TIITBusySalePricePerUnit
 ,(IIF(ISNULL(TableItemSerialNo.[SerialNo],1)='1',(0-TableItemsInTransaction.[Value1]),1) * TableItemsInTransaction.[D2])  As TIITBusySaleAmount
  FROM  [Tran2] As TableItemsInTransaction
 LEFT JOIN [Master1] As TableMaterialCenterMaster ON TableItemsInTransaction.[MasterCode2] = TableMaterialCenterMaster.[Code]
 LEFT JOIN [Master1] As TableSalesmanMaster ON TableItemsInTransaction.[CM6] = TableSalesmanMaster.[Code]
 LEFT JOIN [Master1] As TablePartyMaster ON TableItemsInTransaction.[CM1] = TablePartyMaster.[Code]
 LEFT JOIN [MasterAddressInfo] As TableAccountMasterDetailedInfo ON TableItemsInTransaction.[CM1] = TableAccountMasterDetailedInfo.[MasterCode] 
 LEFT JOIN [Master1] As TableZoneMaster ON TableZoneMaster.[Code] = TableAccountMasterDetailedInfo.[OF2]  
 LEFT JOIN [Master1] As TableTerritoryMaster ON TableTerritoryMaster.[Code] = TableAccountMasterDetailedInfo.[OF3]  
 LEFT JOIN [Master1] As TableCityMaster ON TableCityMaster.[Code] = TableAccountMasterDetailedInfo.[OF4]  
 LEFT JOIN [Master1] As TableTerritoryManagerMaster ON TableTerritoryManagerMaster.[Code] = TableAccountMasterDetailedInfo.[OF5]  
 LEFT JOIN [Master1] As TableAccountManagerMaster ON TableAccountManagerMaster.[Code] = TableAccountMasterDetailedInfo.[OF6]
 LEFT JOIN [Master1] As TableItemMaster ON TableItemsInTransaction.[MasterCode1] = TableItemMaster.[Code]
 LEFT JOIN [MasterAddressInfo] As TableItemMasterDetailedInfo ON TableItemsInTransaction.[MasterCode1] = TableItemMasterDetailedInfo.[MasterCode]
 LEFT JOIN [Master1] As TableBusyGroupNameMaster ON TableBusyGroupNameMaster.[Code] = TableItemMaster.[ParentGrp]
 LEFT JOIN [Master1] As TableParentGroupNameMaster ON TableParentGroupNameMaster.[Code] = TableItemMasterDetailedInfo.[OF2]  
 LEFT JOIN [Master1] As TableGroupNameMaster ON TableGroupNameMaster.[Code] = TableItemMasterDetailedInfo.[OF3]  
 LEFT JOIN [Master1] As TableManufacturingBrandMaster ON TableManufacturingBrandMaster.[Code] = TableItemMasterDetailedInfo.[OF4]  
 LEFT JOIN [Master1] As TableConsumerBrandMaster ON TableConsumerBrandMaster.[Code] = TableItemMasterDetailedInfo.[OF5]
 LEFT JOIN [ItemSerialNo] As TableItemSerialNo  ON TableItemsInTransaction.[VchCode] = TableItemSerialNo.[VchCode] AND TableItemsInTransaction.[MasterCode1] = TableItemSerialNo.[ItemCode]
 WHERE
  TableItemsInTransaction.[VchType] IN (3,9)
  AND TableItemsInTransaction.[RecType] = (2)
  And TableItemsInTransaction.[Date] >='11-01-2022' And TableItemsInTransaction.[Date] <='11-30-2022'
  Order By TableItemsInTransaction.[Date],TableItemsInTransaction.[VchCode],TableItemsInTransaction.[SrNo] ASC
  Go