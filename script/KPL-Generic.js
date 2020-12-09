"use strict";
(function() {

    let xl = domos.excel();
    let objects = domos.objects();

    for (let sheet of xl.sheets()) {
        // Fund Entities data
        if (sheet.name() == "Entities") {
            let fundCompteur = 0;
            let companyCompteur = 0;
            let corporateCompteur = 0;
            let advisorCompteur =0;

            // Find header row
            let mainIndex = sheet.defaultIndex();

            // Entities data headers
            let entityClientIdIndex = mainIndex.get("Client ID");
            let entityInternalIDIndex = mainIndex.get("Internal ID");
            let entityEntryNumberIndex = mainIndex.get("Entry number");
            let entityShortNameIndex = mainIndex.get("Short name");
            let entityFullNameIndex = mainIndex.get("Full name");
            let entityLegalFormIndex = mainIndex.get("Legal form");
            let entityRoleIndex = mainIndex.get("Role in structure");
            let entityAddressIndex = mainIndex.get("Registered address");
            let entityCountryIndex = mainIndex.get("Country");
            let entityIncorporationDateIndex = mainIndex.get("Incorporation date");
            let entityCurrencyIndex = mainIndex.get("Reference currency");
            let entityWebsiteIndex = mainIndex.get("Website url");
            let entityRegistrationIDIndex = mainIndex.get("Registration ID");
            let entityRegistrationNumberIndex = mainIndex.get("Registration number");

            for (let row of mainIndex.dataRows()) {
                // Retrieve entity data
                let entityClientId = row.cell(entityClientIdIndex);
                let entityInternalID = row.cell(entityInternalIDIndex);
                let entityEntryNumber = row.cell(entityEntryNumberIndex);
                let entityFullName = row.cell(entityFullNameIndex);
                let entityShortName = row.cell(entityShortNameIndex);
                let entityLegalForm = row.cell(entityLegalFormIndex);
                let entityRole = row.cell(entityRoleIndex);
                let entityIncorporationDate = util.extractDate(row.cell(entityIncorporationDateIndex), "dd.MM.yyyy");
                let entityAddress = row.cell(entityAddressIndex);
                let entityCountry = row.cell(entityCountryIndex);
                let entityCurrency = row.cell(entityCurrencyIndex);
                let entityWebsite = row.cell(entityWebsiteIndex);
                let entityRegistrationID = row.cell(entityRegistrationIDIndex);
                let entityRegistrationNumber = row.cell(entityRegistrationNumberIndex);

                // Create entity
                var entity;
                //Create Fund
                let name;
                if (!entityFullName.empty()) {
                    name = entityFullName.stringValue();
                } else if (!entityShortName.empty()) {
                    name = entityShortName.stringValue();
                }
                if(name && !entityRole.empty()){
                if (!entityRole.empty() && entityRole.stringOrNull().toUpperCase().includes("FUND")) {
                    entity = objects.findOrCreateVehicle(name);
                    fundCompteur += 1;
                }
                //Create Company
                if (!entityRole.empty() && entityRole.stringOrNull().toUpperCase().includes("COMPANY")) {
                    entity = objects.findOrCreateCompany(name);
                    if (entity && (entityRole.stringOrNull().toUpperCase().includes("HOLDING")||entityRole.stringOrNull().toUpperCase().includes("SPV"))) {
                        entity.objectType("HOLDING");
                    }
                    companyCompteur += 1;
                }
                //Create Corporate
                if (!entityRole.empty() && (entityRole.stringOrNull().includes("Land Developer") || entityRole.stringOrNull().includes("Land Owner") ||
                        entityRole.stringOrNull().includes("Lessor") || entityRole.stringOrNull().includes("Financing Institution") || entityRole.stringOrNull().includes("Platform") ||
                        entityRole.stringOrNull().includes("Project Developer") || entityRole.stringOrNull().includes("Contractor"))) {
                    //entity = objects.newCorporate(name);
                    entity = objects.findOrCreateObjectType(name,11005);
                    corporateCompteur += 1;
                }
                //Create Advisor
                if (!entityRole.empty() && entityRole.stringOrNull().includes(" Investment Advisor")){
                    entity = objects.findOrCreateObjectType(name,"INVESTMENT_ADVISER");
                    advisorCompteur += 1;
                }
                if (entity && !entityCurrency.empty()) entity.currency(entityCurrency.stringValue());
                if (entity && !entityCountry.empty()) {
                    let country = domos.inferCountryCodeOrNull(entityCountry.stringOrNull());
                    if (country) {
                        entity.country(country);
                        if (entity.isEntity()) entity.addresses().legal(true).country(country);
                    } else {
                        domos.log().warn("Country " + entityCountry.stringOrNull() + " not supported");
                    }
                }
                if (entity && !entityShortName.empty()) entity.shortName(entityShortName);
                if (entity && entity.isEntity() && !entityAddress.empty())  entity.addresses().legal(true).street(entityAddress);
                //Add Reference
				let registrationID;
                if (!entityRegistrationID.empty()){
                    registrationID = entityRegistrationID.stringOrNull();
                }else if (!entityRegistrationNumber.empty()){
                    registrationID = entityRegistrationNumber.stringOrNull();
                }
                if (entity && registrationID)entity.customReference("RegistrationID", registrationID);
                if (entity && !entityInternalID.empty()) entity.customReference("InternalID", entityInternalID);
                if (entity && !entityClientId.empty()) entity.customReference("ReferenceNoclinodos", entityClientId);
                }
            }
            domos.log().info("Number of fund position is " + fundCompteur);
            domos.log().info("Number of company position is " + companyCompteur);
            domos.log().info("Number of corporate position is " + corporateCompteur);
            domos.log().info("Number of advisor position is " + advisorCompteur);
            domos.log().info("Entities data has been imported");
        }

        // Fund Equities data
        if (sheet.name() == "Equities" || sheet.name() == "Shares") {
            let companyShareCompteur = 0;
            // Find header row
            let mainIndex = sheet.defaultIndex();
            // shares data headers
            let clientIdIndex = mainIndex.get("Client ID");
            let shareInvestorIndex = mainIndex.get("Investor");
            let shareInvesteeIndex = mainIndex.get("Investee");
            let sharePurchaseDateIndex = mainIndex.get("Purchase date");
            let shareInstrumentTypeIndex = mainIndex.get("Instrument type");
            let shareInstrumentNameIndex = mainIndex.get("Instrument name");
            let shareOperationCurrencyIndex = mainIndex.get("Operation currency");
            let shareTotalCostInValIndex = mainIndex.get("Total cost in val ccy_Memo");
            let shareValuationCurrencyIndex = mainIndex.get("Valuation currency");
            let shareQuantityMemoIndex = mainIndex.get("Quantity_Memo");
            let shareUnitPriceOpeCcyIndex = mainIndex.get("Unit price ope ccy_Memo");
            let shareValuationOpeCcyIndex = mainIndex.get("Valuation in ope ccy");
            let shareShortNameIndex = mainIndex.get("Short name");
            let shareFullNameIndex = mainIndex.get("Full name");
            let shareCommittedAmountOpeIndex = mainIndex.get("Committed amount ope ccy_Memo");
            let shareFundedAmountOpeIndex = mainIndex.get("Funded amount ope ccy_Memo");
            let shareCommittedAmountValIndex = mainIndex.get("Committed amount val ccy_Memo");
            let shareFundedAmountValIndex = mainIndex.get("Funded amount val ccy_Memo");
            let shareNameOwningEntityIndex = mainIndex.get("Name of owning entity");
            let shareOwningEntityIDIndex = mainIndex.get("Owning entity ID");
            let shareNameOwnedEntityIndex = mainIndex.get("Name of owned entity");
            let shareOwnedEntityIDIndex = mainIndex.get("Owned entity ID");

            for (let row of mainIndex.dataRows()) {
                // Retrieve entity data
                let clientId = row.cell(clientIdIndex);
                let shareInvestor = row.cell(shareInvestorIndex);
                let shareInvestee = row.cell(shareInvesteeIndex);
                let sharePurchaseDate = util.extractDate(row.cell(sharePurchaseDateIndex), "dd.MM.yyyy");
                let shareInstrumentType = row.cell(shareInstrumentTypeIndex);
                let shareInstrumentName = row.cell(shareInstrumentNameIndex);
                let shareOperationCurrency = row.cell(shareOperationCurrencyIndex);
                let shareTotalCostInVal = row.cell(shareTotalCostInValIndex);
                let shareValuationCurrency = row.cell(shareValuationCurrencyIndex);
                let shareUnitPrice = row.cell(shareUnitPriceOpeCcyIndex);
                let shareQuantityMemo = row.cell(shareQuantityMemoIndex);
                let shareValuationOpeCcy = row.cell(shareValuationOpeCcyIndex);
                let shareShortName = row.cell(shareShortNameIndex);
                let shareFullName = row.cell(shareFullNameIndex);
                let shareCommittedAmountOpe = row.cell(shareCommittedAmountOpeIndex);
                let shareFundedAmountOpe = row.cell(shareFundedAmountOpeIndex);
                let shareCommittedAmountVal = row.cell(shareCommittedAmountValIndex);
                let shareFundedAmountVal = row.cell(shareFundedAmountValIndex);
                let shareNameOwningEntity = row.cell(shareNameOwningEntityIndex);
                let shareOwningEntityID = row.cell(shareOwningEntityIDIndex);
                let shareNameOwnedEntity = row.cell(shareNameOwnedEntityIndex);
                let shareOwnedEntityID = row.cell(shareOwnedEntityIDIndex);

                // Create entity
                let share;
                //Create Company share
                let companyShareName;
                if (!shareShortName.empty()) {
                    companyShareName = shareShortName.stringValue();
                } else if (!shareFullName.empty()) {
                    companyShareName = shareFullName.stringValue();
                } else if (!shareInvestee.empty()) {
                    companyShareName = shareInvestee.stringValue() + " " + shareInstrumentType.stringValue();
                }
                if (companyShareName) {
                    if (!shareInstrumentName.empty()) companyShareName += " " + shareInstrumentName.stringValue();
                    share = objects.findOrCreateCompanyShare(companyShareName);
                    let shareType;
                    if (!shareInstrumentType.empty() && shareInstrumentType.stringValue().toUpperCase().includes("SHARES")) {
                        shareType = 1;
                    } else if (!shareInstrumentType.empty() && shareInstrumentType.stringValue().toUpperCase().includes("EQUITY")) {
                        shareType = 5;
                    }
                    if (shareType) share.shareType(shareType);
                    if (!shareOperationCurrency.empty()) {
                        share.currency(shareOperationCurrency);
                    } else if (!shareTotalCostInVal.empty()) {
                        share.currency(shareTotalCostInVal);
                    } else if (!shareValuationCurrency.empty()) {
                        share.currency(shareValuationCurrency);
                    }
                    let company;
                    if (!shareInvestee.empty()) {
                        company = objects.findCompanyOrNull(shareInvestee);
                    } else if (!shareNameOwnedEntity.empty()) {
                        company = objects.findCompanyOrNull(shareNameOwnedEntity);
                    } else if (!shareNameOwnedEntity.empty()) {
                        company = objects.findCompanyOrNull(shareOwnedEntityID);
                    }
                    if (company) company.addCompanyShare(share);
                    let investorName;
                    if(!shareInvestor.empty()){
                        investorName = shareInvestor;
                    }else if (!shareOwningEntityID.empty()){
                        investorName = shareOwningEntityID;
                    }
                    let investor;
                    if(investorName){
                        objects.findVehicleOrNull(investorName);
                        if (!investor) {
                            investor = objects.findCompanyOrNull(investorName);
                        }
                    }
                    /* if(investor){
                        let defaultDate = util.extractDate("31.12.2019", "dd.MM.yyyy");
                        share.findOrCreateCompanySubscription(defaultDate, investor, null)
                        .quantity(shareQuantity)
                        .amount(sharePositionARS);
                    } */
                    //	share.setLast(shareUnitPrice, shareOperationCurrency, null);
                    companyShareCompteur += 1;
                }
            }
            domos.log().info("Number of company share position is " + companyShareCompteur);
            domos.log().info("Equities data has been imported");
        }
        // Fund Loans data
        if (sheet.name() == "Loans") {
            let loansCompteur = 0;
            // Find header row
            let mainIndex = sheet.defaultIndex();
            // Loans data headers
            let clientIdIndex = mainIndex.get("Client ID");
            let loanFullNameIndex = mainIndex.get("Full name");
            let loanShortNameIndex = mainIndex.get("Short name");
            let loanLenderIndex = mainIndex.get("Lender");
            let loanBorrowerIndex = mainIndex.get("Borrower");
			let loanLenderNameIndex = mainIndex.get("Name of lending entity");
            let loanLenderIDIndex = mainIndex.get("Lending entity ID");
            let loanBorrowerNameIndex = mainIndex.get("Name of borrowing entity");
            let loanBorrowerIDIndex = mainIndex.get("Borrowing entity ID");
			let operationInvestorIndex = mainIndex.get("Investor");
            let operationInvesteeIndex = mainIndex.get("Investee");
            let loanInstrumentTypeIndex = mainIndex.get("Instrument type");
            let loanIssueDateIndex = mainIndex.get("Issue date");
            let loanMaturityDateIndex = mainIndex.get("Maturity date");
            let loanOperationCurrencyIndex = mainIndex.get("Operation currency");
            let loanInterestRateIndex = mainIndex.get("Interest rate %");
            let loanFundedAmountOpeIndex = mainIndex.get("Funded amount ope ccy_Memo");
			let loanOverallFacilityIndex = mainIndex.get("Overall facility in operation currency");
			let loanFundedAmountValIndex = mainIndex.get("Funded amount val ccy_Memo");
            for (let row of mainIndex.dataRows()) {
                // Retrieve operation data
                let loanFullName = row.cell(loanFullNameIndex);
				let loanShortName = row.cell(loanShortNameIndex);
				let loanLender = row.cell(loanLenderIndex);
				let loanBorrower = row.cell(loanBorrowerIndex);
				let loanLenderName = row.cell(loanLenderNameIndex);
				let loanLenderID = row.cell(loanLenderIDIndex);
				let loanBorrowerName = row.cell(loanBorrowerNameIndex);
				let loanBorrowerID = row.cell(loanBorrowerIDIndex);
				let operationInvestor = row.cell(operationInvestorIndex);
				let operationInvestee = row.cell(operationInvesteeIndex);
				let loanInstrumentType = row.cell(loanInstrumentTypeIndex);
				let loanIssueDate = util.extractDate(row.cell(loanIssueDateIndex), "dd.MM.yyyy");
				let loanMaturityDate = util.extractDate(row.cell(loanMaturityDateIndex), "dd.MM.yyyy");
				let loanOperationCurrency = row.cell(loanOperationCurrencyIndex);
				let loanInterestRate = row.cell(loanInterestRateIndex);
				let loanFundedAmountOpe = row.cell(loanFundedAmountOpeIndex);
				let loanOverallFacility = row.cell(loanOverallFacilityIndex);
				let loanFundedAmountVal = row.cell(loanFundedAmountValIndex);
                
                let nominalAmount;
				if (!loanFundedAmountOpe.empty()){
					nominalAmount = loanFundedAmountOpe.doubleOrNull();
				}else if (!loanOverallFacility.empty()){
					nominalAmount = loanOverallFacility.doubleOrNull();
				}else if (!loanFundedAmountVal.empty()){
					nominalAmount = loanFundedAmountVal.doubleOrNull();
				}
                if (nominalAmount && nominalAmount !== 0){
                    
                let loanName;
				if (!loanFullName.empty()){
					loanName = loanFullName.stringValue();
				}else if (!loanShortName.empty()){
					loanName = loanShortName.stringValue();
				}else continue;
				loansCompteur += 1;
                let loan = objects.findOrCreateDebt(loanName + " - Loan " + loansCompteur);
                loan.shortName(loanName);
                if(!loanOperationCurrency.empty()) loan.currency(loanOperationCurrency);
                if(loanIssueDate) loan.startDate(loanIssueDate, "dd.MM.yyyy");
                if(loanMaturityDate) loan.maturityDate(loanMaturityDate, "dd.MM.yyyy");
                if(!loanInterestRate.empty()) loan.ratePercent(loanInterestRate.doubleOrNull());
                loan.nominal(nominalAmount);
				let borrowerName;
				if (!loanBorrower.empty()){
					borrowerName = loanBorrower;
				}else if (!loanBorrowerID.empty()){
					borrowerName = loanBorrowerID;
				}else if (!operationInvestee.empty()){
					borrowerName = operationInvestee;
				}
                let borrower = objects.findCompanyOrNull(borrowerName);
                if(!borrower) borrower =objects.findVehicleOrNull(borrowerName);
				if (borrower) loan.issuer(borrower);
				let lenderName;
				if (!loanLender.empty()){
					lenderName = loanLender;
				}else if (!loanLenderID.empty()){
					lenderName = loanLenderID;
				}else if (!operationInvestor.empty()){
					lenderName = operationInvestor;
				}
                let lender = objects.findCompanyOrNull(lenderName);
                if(!lender) lender =objects.findVehicleOrNull(lenderName);
                if (lender) loan.lender(lender);
                }
                //domos.log().debug("operationInvestee : "+operationInvestee+" lender: "+lender+" || borrower : "+borrower);
            }
            domos.log().info("Number of loan is " + loansCompteur);
        }
		// Fund Properties data
        if (sheet.name() == "Non-financial assets") {
            // Find header row
            let mainIndex = sheet.defaultIndex();
	
            // Property data headers
            let propertyEntryNumberIndex = mainIndex.get("Entry number");
            let propertyInternalIdIndex = mainIndex.get("Internal ID");
            let propertyContractReferenceBOIndex = mainIndex.get("Contract reference - Blatt no");
            let propertyShortNameIndex = mainIndex.get("Short name");
            let propertyFullNameIndex = mainIndex.get("Full name");
            let propertyLocationIndex = mainIndex.get("Location");
            let propertyCityIndex = mainIndex.get("City");
            let propertyCountryIndex = mainIndex.get("Country");
            let propertyIndustryOrSectorIndex = mainIndex.get("Industry or sector");
            let propertyNameOwningEntityIndex = mainIndex.get("Name of owning entity");
            let propertyOwningEntityIdIndex = mainIndex.get("Owning entity ID");
            let propertyQuantityAreaIndex = mainIndex.get("Quantity Area (sq m)");
            let propertyMeasuringUnitIndex = mainIndex.get("Measuring unit");
            let propertyOperationCurrencyIndex = mainIndex.get("Operation currency");
            let propertyTotalCostOIndex = mainIndex.get("Total cost in operation currency");
            let propertyOwnershipERAKDateIndex = mainIndex.get("Ownership evidence received at knowledge date");
            let propertyCommentsIndex = mainIndex.get("Comments");
            let propertyLastLRRIndex = mainIndex.get("Last Land Registry received");
            let propertyDocumentationFLIndex = mainIndex.get("Documentation folder link");
            let propertyDescriptionIndex = mainIndex.get("Description");
            let propertyQuantityMemoIndex = mainIndex.get("Quantity_Memo");
            let propertyOwnerIndex = mainIndex.get("Owner");
	
            let propertyCompteur = 0;
            for (let row of mainIndex.dataRows()) {
                // Retrieve property data
                let propertyEntryNumber = row.cell(propertyEntryNumberIndex);
                let propertyInternalId = row.cell(propertyInternalIdIndex);
                let propertyContractReferenceBO = row.cell(propertyContractReferenceBOIndex);
                let propertyShortName = row.cell(propertyShortNameIndex);
                let propertyFullName = row.cell(propertyFullNameIndex);
                let propertyLocation = row.cell(propertyLocationIndex);
                let propertyCity = row.cell(propertyCityIndex);
                let propertyCountry = row.cell(propertyCountryIndex);
                let propertyIndustryOrSector = row.cell(propertyIndustryOrSectorIndex);
                let propertyNameOwningEntity = row.cell(propertyNameOwningEntityIndex);
                let propertyOwningEntityId = row.cell(propertyOwningEntityIdIndex);
                let propertyQuantityArea = row.cell(propertyQuantityAreaIndex);
                let propertyMeasuringUnit = row.cell(propertyMeasuringUnitIndex);
                let propertyOperationCurrency = row.cell(propertyOperationCurrencyIndex);
                let propertyTotalCostOC = row.cell(propertyTotalCostOIndex);
                let propertyOwnershipERAKDate = row.cell(propertyOwnershipERAKDateIndex);
                let propertyComments = row.cell(propertyCommentsIndex);
                let propertyLastLRR = row.cell(propertyLastLRRIndex);
                let propertyDocumentationFL = row.cell(propertyDocumentationFLIndex);
                let propertyDescription = row.cell(propertyDescriptionIndex);
                let propertyQuantityMemo = row.cell(propertyQuantityMemoIndex);
                let propertyOwner = row.cell(propertyOwnerIndex);
	
                // Create property
				let propertyName;
                if (propertyFullName.empty() && propertyShortName.empty()) continue;
				if (!propertyFullName.empty()){
					propertyName = propertyFullName.stringOrNull();
				}else {
					propertyName = propertyShortName.stringOrNull();
				}
				if(!propertyLocation.empty()) propertyName += " - " + propertyLocation.stringOrNull();
				if(!propertyContractReferenceBO.empty()) propertyName += " - " + propertyContractReferenceBO.stringOrNull();
                let property = objects.findOrCreateObjectType(propertyName, "PROPERTY");
                if (!propertyCountry.empty()) {
                    let country = domos.inferCountryCodeOrNull(propertyCountry.stringOrNull());
                    if (country) {
                        property.country(country);
                    } else {
                        domos.log().error("Country " + country + " not supported");
                    }
                }
                if (!propertyOperationCurrency.empty()) property.currency(propertyOperationCurrency);
                //if (!propertyShortName.empty()) property = property.shortName(propertyShortName);
                if (!propertyLocation.empty()) property.address(propertyLocation);
                if (!propertyCity.empty()) property.city(propertyCity);
                if (!propertyDescription.empty()) property.description(propertyDescription);
                //Add Measurement
                if (!propertyQuantityArea.empty()) {
                    property.addMeasurement().value(propertyQuantityArea)
                        .unit("SQUARE_METER");
                } else if (!propertyQuantityMemo.empty()){
					property.addMeasurement().value(propertyQuantityMemo)
                        .unit("SQUARE_METER");
				}
                //Add Reference
                if(!propertyInternalId.empty()) property.customReference("InternalID", propertyInternalId);
                if(!propertyContractReferenceBO.empty()) property.customReference("ContractReference", propertyContractReferenceBO);
                
                let defaultDate = util.extractDate("31.12.2019", "dd.MM.yyyy");
            
                let ownerName;
				if (!propertyOwningEntityId.empty()){
					ownerName = propertyOwningEntityId;
				} else if (!propertyOwner.empty()){
					ownerName = propertyOwner;
				}
				let owner = objects.findCompanyOrNull(ownerName);
				if (!owner) owner = objects.findVehicleOrNull(ownerName);
                if (owner) owner.addPosition(property, 1, null, defaultDate);
                propertyCompteur += 1;
            }
            domos.log().info("Number of property position is " + propertyCompteur);
            domos.log().info("Properties data has been imported");
        }
		
		// Fund Collaterals data
        if (sheet.name() == "Collateral") {
            // Find header row
            let mainIndex = sheet.defaultIndex();
	
            // Collateral data headers
			let collateralShortNameIndex = mainIndex.get("Short name");
			let collateralFullNameIndex = mainIndex.get("Full name");
			let collateralContractRefIndex = mainIndex.get("Contract reference");
			let collateralDescriptionIndex = mainIndex.get("Description");
			let collateralCommentsIndex = mainIndex.get("Comments");
			let collateralGrantorIndex = mainIndex.get("Grantor");
			let collateralcurrencyIndex = mainIndex.get("Operation currency");
			
			let collateralCompteur = 0;
            for (let row of mainIndex.dataRows()) {
                // Retrieve collateral data
                let collateralShortName = row.cell(collateralShortNameIndex);
                let collateralFullName = row.cell(collateralFullNameIndex);
                let collateralContractRef = row.cell(collateralContractRefIndex);
                let collateralDescription = row.cell(collateralDescriptionIndex);
                let collateralComments = row.cell(collateralCommentsIndex);
                let collateralGrantor = row.cell(collateralGrantorIndex);
                let collateralcurrency = row.cell(collateralcurrencyIndex);
				// Create collateral
				
				let collateralName;
                if (collateralFullName.empty() && collateralShortName.empty()) continue;
				if (!collateralFullName.empty()){
					collateralName = collateralFullName.stringOrNull();
				}else {
					collateralName = collateralShortName.stringOrNull();
				}
				
    			let collateral = objects.findOrCreateOtherAsset(collateralName);
				if (!collateralShortName.empty())collateral.shortName(collateralShortName);
			//	if (!collateralDescription.empty())collateral.description(collateralDescription);
				if (!collateralcurrency.empty())collateral.currency(collateralcurrency);
				collateralCompteur += 1;
			}
			domos.log().info("Number of collateral position is " + collateralCompteur);
            domos.log().info("Collaterals data has been imported");
		}
    }
})();