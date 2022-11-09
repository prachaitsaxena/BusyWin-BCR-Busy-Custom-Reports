SELECT
TablePartyGroup.[Name] As TPGPartyGroup
,TableAccountMasterDetailedInfo.[OF2] As TAMAIZoneCode
,ISNull(TableZoneMaster.[Name],'') As TZMName
,TableAccountMasterDetailedInfo.[OF3] As TAMAITerritoryCode
,ISNULL(TableTerritoryMaster.[Name],'') As TTMName
,TableAccountMasterDetailedInfo.[OF4] As TAMAICityCode
,ISNULL(TableCityMaster.[Name],'') As TCMName
,TableAccountMasterDetailedInfo.[OF5] As TAMAITerritoryManagerCode
,ISNULL(TableTerritoryManagerMaster.[Name],'') As TTMMName
,TableAccountMasterDetailedInfo.[OF6] As TAMAIAccountManagerCode
,ISNULL(TableAccountManagerMaster.[Name],'') As TAMMName
,TablePartyMaster.[Code] As TPMCode
,TablePartyMaster.[Name] As TPMPartyName
,TablePartyMaster.[Alias] As TPMAPartyCode
,TablePartyMaster.[PrintName] As TPMPrintName
,TableAccountMasterDetailedInfo.[Contact] As TAMDIContact
,TableAccountMasterDetailedInfo.[Mobile] As TAMDIMobile
,TableAccountMasterDetailedInfo.[WhatsAppNo] As TAMDIWhatsAppNo
,TableAccountMasterDetailedInfo.[TelNo] As TAMDITelephoneNo
,TableAccountMasterDetailedInfo.[Email] As TAMDIEmail
,TableAccountMasterDetailedInfo.[OF7] As TAMDIPartyType
,ISNULL(TablePartyType.[Name],'') As TPTPartyType
,TableAccountMasterDetailedInfo.[OF8] As TAMDIPartyKYCStatus
,ISNULL(TablePartyKYCStatus.[Name],'') As TPKSTablePartyKYCStatus
,TableAccountMasterDetailedInfo.[Address1] As TAMDIAddressLine1
,TableAccountMasterDetailedInfo.[Address2] As TAMDIAddressLine2
,TableAccountMasterDetailedInfo.[Address3] As TAMDIAddressLine3
,TableAccountMasterDetailedInfo.[Address4] As TAMDIAddressLine4
,TableAccountMasterDetailedInfo.[Station] As TAMDIStation
,ISNULL(TableCityMasterAMDI.[Name],'') As TCMAMDICity
,TableStateMaster.[Name] As TCMStateName
,TableCountryMaster.[Name] As TCMCountryName
,TableAccountMasterDetailedInfo.[PINCode] As TAMDIPincode
,TableAccountMasterDetailedInfo.[ITPAN] As TAMDIITPAN
,TableAccountMasterDetailedInfo.[C3] As TAMDIAadharNumber
,TablePartyMaster.[CreatedBy] AS TPMCreatedBy
,TablePartyMaster.[CreationTime] AS TPMCreationTime
,TablePartyMaster.[ModifiedBy] AS TPMLastModifiedBy
,TablePartyMaster.[ModificationTime] AS TPMLastModificationTime
FROM [Master1] As TablePartyMaster
LEFT JOIN [Master1] As TablePartyGroup ON TablePartyGroup.[Code] = TablePartyMaster.[ParentGrp]
LEFT JOIN [MasterAddressInfo] As TableAccountMasterDetailedInfo ON TablePartyMaster.[Code] = TableAccountMasterDetailedInfo.[MasterCode]
LEFT JOIN [Master1] As TableZoneMaster ON TableZoneMaster.[Code] = TableAccountMasterDetailedInfo.[OF2]
LEFT JOIN [Master1] As TableTerritoryMaster ON TableTerritoryMaster.[Code] = TableAccountMasterDetailedInfo.[OF3]
LEFT JOIN [Master1] As TableCityMaster ON TableCityMaster.[Code] = TableAccountMasterDetailedInfo.[OF4]
LEFT JOIN [Master1] As TableTerritoryManagerMaster ON TableTerritoryManagerMaster.[Code] = TableAccountMasterDetailedInfo.[OF5]
LEFT JOIN [Master1] As TableAccountManagerMaster ON TableAccountManagerMaster.[Code] = TableAccountMasterDetailedInfo.[OF6]
LEFT JOIN [Master1] As TablePartyType ON TablePartyType.[Code] = TableAccountMasterDetailedInfo.[OF7]
LEFT JOIN [Master1] As TablePartyKYCStatus ON TablePartyKYCStatus.[Code] = TableAccountMasterDetailedInfo.[OF8]
LEFT JOIN [Master1] As TableCityMasterAMDI ON TableCityMasterAMDI.[Code] = TableAccountMasterDetailedInfo.[CityCodeLong]
LEFT JOIN [Master1] As TableStateMaster ON TableStateMaster.[Code] = TableAccountMasterDetailedInfo.[StateCodeLong]
LEFT JOIN [Master1] As TableCountryMaster ON TableCountryMaster.[Code] = TableAccountMasterDetailedInfo.[CountryCodeLong]
WHERE
TablePartyMaster.[MasterType] = 2
And TablePartyGroup.[Name] LIKE '%Sundry%'
Order By TablePartyGroup.[Name], TableZoneMaster.[Name], TableTerritoryMaster.[Name],TableCityMaster.[Name],TableTerritoryManagerMaster.[Name],TableAccountManagerMaster.[Name], TablePartyMaster.[Name] ASC
Go