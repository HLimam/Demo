"use strict";
(function() {

    let xl = domos.excel();
    let objects = domos.objects();
	
    for (let sheet of xl.sheets()) {
        /*if (sheet.name() == "Client information") {
            // Find header row
            let informationheetDataIndex = sheet.defaultIndex();
            let parameterIndex = informationheetDataIndex.get("Parameter");
            let valueIndex = informationheetDataIndex.get("Value");
            let fund;
            for (let row of informationheetDataIndex.dataRows()) {
                // Retrieve operation data
                let parameter = row.cell(parameterIndex);
                let value = row.cell(valueIndex);
                if (parameter.stringValue() == "Fund name") {
                    fund = objects.findOrCreateVehicle(value);
                    fund.masterFeeder("MASTER");
                }
                if (fund && parameter.stringValue() == "AIFM") {
                    let fundManager = objects.findOrCreateObjectType(value, "MANAGEMENT_COMPANY");
                    fund.company(fundManager);
                }
            }
        }*/
	
        // Fund Entities data
        if (sheet.name() == "Entities") {
            let fundCompteur = 0;
            let companyCompteur = 0;
            // Find header row
            let mainIndex = sheet.defaultIndex();
	
            // Entities data headers
            let entityClientIdIndex = mainIndex.get("Client ID");
            let entityInternalIDIndex = mainIndex.get("Internal ID");
            let entityEntryNumberIndex = mainIndex.get("Entry number");
            let entityFullNameIndex = mainIndex.get("Full name");
            let entityShortNameIndex = mainIndex.get("Short name");
            let entityRoleIndex = mainIndex.get("Role in structure");
            let entityActiveIndex = mainIndex.get("Active");
            let entityRequiredIndex = mainIndex.get("Look-through required");
            let entityRegistrationIdIndex = mainIndex.get("Registration ID");
            let entityIncorporationDateIndex = mainIndex.get("Incorporation date");
            let entityTotalNOSSIndex = mainIndex.get("Total number of shares subscribed");
            let entityCurrencyIndex = mainIndex.get("Reference currency");
            let entityTotalSCIndex = mainIndex.get("Total subscribed capital");
	
            // Entities Adresse data headers
            let entityAddressIndex = mainIndex.get("Registered address");
            let entityCountryIndex = mainIndex.get("Country");
	
            for (let row of mainIndex.dataRows()) {
                // Retrieve entity data
                let entityClientId = row.cell(entityClientIdIndex);
                let entityInternalID = row.cell(entityInternalIDIndex);
                let entityEntryNumber = row.cell(entityEntryNumberIndex);
                if (entityClientId.empty() && entityEntryNumber.empty()) break;
                let entityFullName = row.cell(entityFullNameIndex);
                let entityShortName = row.cell(entityShortNameIndex);
                let entityRole = row.cell(entityRoleIndex);
                let entityActive = row.cell(entityActiveIndex);
                let entityRequired = row.cell(entityRequiredIndex);
                let entityRegistrationId = row.cell(entityRegistrationIdIndex);
                let entityIncorporationDate = util.extractDate(row.cell(entityIncorporationDateIndex), "dd.MM.yyyy");
                let entityTotalNOSS = row.cell(entityTotalNOSSIndex);
                let entityCurrency = row.cell(entityCurrencyIndex);
                let entityTotalSC = row.cell(entityTotalSCIndex);
                let entityAddress = row.cell(entityAddressIndex);
                let entityCountry = row.cell(entityCountryIndex);
	
                // Create entity
                var entity;
                //Create Fund
                if (entityRole.stringOrNull() == "Fund (Client)" || entityRole.stringOrNull() == "Investment Fund") {
                    if (!entityFullName.empty()) {
                        entity = objects.findOrCreateVehicle(entityFullName.stringValue());
                    } else {
                        if (!entityShortName.empty()) entity = objects.findOrCreateVehicle(entityShortName.stringValue());
                    }
	

                    fundCompteur = fundCompteur + 1;
                }
                //Create Company
                if (!entityRole.empty() && entityRole.stringOrNull().includes("Holding Company") || entityRole.stringOrNull() == "Operating Company" || entityRole.stringOrNull() == "Lending company") {
                    if (!entityFullName.empty()) {
                        entity = objects.findOrCreateCompany(entityFullName.stringValue());
                    } else {
                        if (!entityShortName.empty()) entity = objects.findOrCreateCompany(entityShortName.stringValue());
                    }
                    if (entity && entityRole.stringOrNull().includes("Holding Company")) {
                        entity.objectType("HOLDING");
                    }
                    companyCompteur = companyCompteur + 1;
                }
                //let adresse = entity.addresses()._create("FREE_QNAME");
                if (entity && !entityCurrency.empty()) entity.currency(entityCurrency.stringValue());
                if (entity && !entityCountry.empty()) {
                    let country = domos.inferCountryCodeOrNull(entityCountry.stringOrNull());
                    if (country) {
                        entity.country(country);
                        if (entity.isEntity()) {
                        entity.addresses().legal(true).country(country);
                        }
                    } else {
                        domos.log().error("Country " + country + " not supported");
                    }
                }
                if (entity && entity.isEntity() && !entityAddress.empty())  entity.addresses().legal(true).street(entityAddress);
             
                
                //Add Reference
                if (!entityInternalID.empty()) entity.customReference("InternalID", entityInternalID);
            }
            domos.log().info("Number of fund position is " + fundCompteur);
            domos.log().info("Number of company position is " + companyCompteur);
            domos.log().info("Entities data has been imported");
        }
        
        // Fund Shares data
        if (sheet.name() == "Shares") {
            // Find header row
            let mainIndex = sheet.defaultIndex();
	
            // Property data headers
            let shareEntryNumberIndex = mainIndex.get("Entry number");
            let shareInternalIDIndex = mainIndex.get("Internal ID");
            let shareShortNameIndex = mainIndex.get("Short name");
            let shareFullNameIndex = mainIndex.get("Full name");
            let shareNameOwningEntityIndex = mainIndex.get("Name of owning entity");
            let shareOwningEntityIDIndex = mainIndex.get("Owning entity ID");
            let shareNameOwnedEntityIndex = mainIndex.get("Name of owned entity");
            let shareOwnedEntityIDIndex = mainIndex.get("Owned entity ID");
            let shareInstrumentTypeIndex = mainIndex.get("Instrument type");
            let shareQuantityIndex = mainIndex.get("Quantity");
            let shareStakeHeldIndex = mainIndex.get("Stake held %");
            let sharePositionARSIndex = mainIndex.get("Position according to reconciliation source");
            let shareDifferencePositionIndex = mainIndex.get("Difference on position");
            let shareReconciliationStatusIndex = mainIndex.get("Reconciliation status");
	
            let shareCompteur = 0;
            for (let row of mainIndex.dataRows()) {
                // Retrieve property data
                let shareEntryNumber = row.cell(shareEntryNumberIndex);
                if (shareEntryNumber.empty()) break;
                let shareInternalID = row.cell(shareInternalIDIndex);
                let shareShortName = row.cell(shareShortNameIndex);
                let shareFullName = row.cell(shareFullNameIndex);
                let shareNameOwningEntity = row.cell(shareNameOwningEntityIndex);
                let shareOwningEntityID = row.cell(shareOwningEntityIDIndex);
                let shareNameOwnedEntity = row.cell(shareNameOwnedEntityIndex);
                let shareOwnedEntityID = row.cell(shareOwnedEntityIDIndex);
                let shareInstrumentType = row.cell(shareInstrumentTypeIndex);
                let shareQuantity = row.cell(shareQuantityIndex);
                let shareStakeHeld = row.cell(shareStakeHeldIndex);
                let sharePositionARS = row.cell(sharePositionARSIndex);
                let shareDifferencePosition = row.cell(shareDifferencePositionIndex);
                let shareReconciliationStatus = row.cell(shareReconciliationStatusIndex);
	
                // Create share
                let share;
                if (!shareShortName.empty()) {
                    share = objects.findOrCreateCompanyShare(shareShortName);
                    //asset.shareType(shareInstrumentType);
                    let company = objects.findCompanyOrNull(shareOwnedEntityID);
                    company.addCompanyShare(share);
                    let investor = objects.findVehicleOrNull(shareOwningEntityID);
                    if (!investor) {
                        investor = objects.findCompanyOrNull(shareOwningEntityID);
                    }
                    let defaultDate = util.extractDate("31.12.2019", "dd.MM.yyyy");
                    if(investor) share.findOrCreateCompanySubscription(defaultDate, investor, null)
                        .quantity(shareQuantity)
                        .amount(sharePositionARS);
                    shareCompteur = +1;
                }
            }
            domos.log().info("Number of share position is " + shareCompteur);
            domos.log().info("Shares data has been imported");
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
	
            let propertyCompteur = 0;
            for (let row of mainIndex.dataRows()) {
                // Retrieve property data
                let propertyEntryNumber = row.cell(propertyEntryNumberIndex);
                if (propertyEntryNumber.empty()) break;
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
	
                // Create property
                if (propertyFullName.empty()) break;
                let property = objects.findOrCreateObjectType(propertyFullName.stringOrNull() + " - " + propertyLocation.stringOrNull() + " - " + propertyContractReferenceBO.stringOrNull(), "PROPERTY");
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
                if (!propertyComments.empty()) property.description(propertyComments);
                //Add Measurement
                if (!propertyQuantityArea.empty()) {
                    property.addMeasurement().value(propertyQuantityArea)
                        .unit("SQUARE_METER");
                }
                //Add Reference
                property.customReference("InternalID", propertyInternalId);
                property.customReference("ContractReference", propertyContractReferenceBO);
                //
              
                let defaultDate = util.extractDate("31.12.2019", "dd.MM.yyyy");
            
                let portfolio = objects.findCompanyOrNull(propertyOwningEntityId);
                portfolio.addPosition(property, 1, null, defaultDate);
                propertyCompteur += 1;
            }
            domos.log().info("Number of property position is " + propertyCompteur);
            domos.log().info("Properties data has been imported");
        }
        
        // Fund Loans data
        if (sheet.name().includes("Loans")) {
            let loansCompteur = 0;
            // Find header row
            let mainIndex = sheet.defaultIndex();
            // Loans data headers
            let entryNumberIndex = mainIndex.get("Entry number");
            let clientIdIndex = mainIndex.get("Client ID");
            let loanFullNameIndex = mainIndex.get("Full name");
            let loanShortNameIndex = mainIndex.get("Short name");
            let loanLenderNameIndex = mainIndex.get("Name of lending entity");
            let loanLenderIDIndex = mainIndex.get("Lending entity ID");
            let loanBorrowerNameIndex = mainIndex.get("Name of borrowing entity");
            let loanBorrowerIDIndex = mainIndex.get("Borrowing entity ID");
            let loanInstrumentTypeIndex = mainIndex.get("Instrument type");
            let loanIssueDateIndex = mainIndex.get("Issue date");
            let loanMaturityDateIndex = mainIndex.get("Maturity date");
            let loanOperationCurrencyIndex = mainIndex.get("Operation currency");
            let loanInterestRateIndex = mainIndex.get("Interest rate %");
            let loanCommittedAmountIndex = mainIndex.get("Committed amount ope ccy_Memo");
            let loanFundedAmountIndex = mainIndex.get("Funded amount ope ccy_Memo");
            let loanOverallFacilityIndex = mainIndex.get("Overall facility in operation currency");
            let loanOverallDrawdownIndex = mainIndex.get("Overall drawdown in operation currency");
            let loanAvailableFacilityIndex = mainIndex.get("Available facility in operation currency");
            let loanOutstandingIndex = mainIndex.get("Outstanding");
            for (let row of mainIndex.dataRows()) {
                // Retrieve operation data
                let entryNumber = row.cell(entryNumberIndex);
                let clientId = row.cell(clientIdIndex);
                if (entryNumber.empty() && clientId.empty()) break;
                let loanShortName = row.cell(loanShortNameIndex);
                let loanFullName = row.cell(loanFullNameIndex);
                let loanLenderName = row.cell(loanLenderNameIndex);
                let loanLenderID = row.cell(loanLenderIDIndex);
                let loanBorrowerName = row.cell(loanBorrowerNameIndex);
                let loanBorrowerID = row.cell(loanBorrowerIDIndex);
                let loanInstrumentType = row.cell(loanInstrumentTypeIndex);
                let loanIssueDate = row.cell(loanIssueDateIndex);
                let loanMaturityDate = row.cell(loanMaturityDateIndex);
                let loanOperationCurrency = row.cell(loanOperationCurrencyIndex);
                let loanInterestRate = row.cell(loanInterestRateIndex);
                let loanCommittedAmount = row.cell(loanCommittedAmountIndex);
                let loanFundedAmount = row.cell(loanFundedAmountIndex);
                let loanOverallDrawdown = row.cell(loanOverallDrawdownIndex);
                let loanOverallFacility = row.cell(loanOverallFacilityIndex);
                let loanAvailableFacility = row.cell(loanAvailableFacilityIndex);
                let loanOutstanding = row.cell(loanOutstandingIndex);
                loansCompteur += 1;
                let loanName;
                if (!loanFullName.empty()) {
                    loanName = loanFullName.stringValue() + " - Loan " + loansCompteur;
                } else {
                    loanName = loanShortName.stringValue() + " - Loan " + loansCompteur;
                }
                let loan = objects.findOrCreateDebt(loanName);
                if (!loanFullName.empty()) {
                    loan.shortName(loanFullName);
                } else {
                    loan.shortName(loanShortName);
                }
                //loan.objectType("CREDIT_LINE");
                loan.currency(loanOperationCurrency);
                loan.startDate(loanIssueDate);
                loan.maturityDate(loanMaturityDate);
                let rate = loanInterestRate.doubleOrNull();
                if (rate) loan.ratePercent(rate);
                //if (!loanOverallDrawdown.empty()) loan.payment(loanOverallDrawdown);
                if (!loanOverallFacility.empty()) loan.nominal(loanOverallFacility);
                let borrower = objects.findCompanyOrNull(loanBorrowerID);
                if (!borrower) borrower.findVehicleOrNull(loanBorrowerID);
                let lender = objects.findCompanyOrNull(loanLenderID);
                if (!lender) lender = objects.findVehicleOrNull(loanLenderID);
                if (borrower) loan = loan.lender(borrower);
                if (lender) loan = loan.issuer(lender);
            }
            domos.log().info("Number of loan is " + loansCompteur);
        }
	
        // Fund Collateral data
       /* if (sheet.name().includes("Collateral")) {
            let collateralCompteur = 0;
            // Find header row
            let mainIndex = sheet.defaultIndex();
            // Loans data headers
            let entryNumberNumberIndex = mainIndex.get("Entry number");
            let clientIdIndex = mainIndex.get("Client ID");
            let collateralContractRefIndex = mainIndex.get("Contract reference");
            let collateralShortNameIndex = mainIndex.get("Short name");
            let collateralFullNameIndex = mainIndex.get("Full name");
            let collateralGrantorNameIndex = mainIndex.get("Name of grantor");
            let collateralGrantorIDIndex = mainIndex.get("Grantor ID");
            let collateralSecuredPartyNameIndex = mainIndex.get("Name of secured party");
            let collateralSecuredPartyIDIndex = mainIndex.get("Secured party ID");
            let collateralIssueDateIndex = mainIndex.get("Issue date");
            let collateralMaturityDateIndex = mainIndex.get("Maturity date");
            let collateralKnowledgeDateIndex = mainIndex.get("Knowledge date");
            let collateralCollateralManagerIndex = mainIndex.get("Collateral manager");
            let collateralValuationCurrencyIndex = mainIndex.get("Valuation currency");
            let collateralInstrumentTypeIndex = mainIndex.get("Instrument type");
            for (let row of mainIndex.dataRows()) {
                // Retrieve operation data
                let entryNumber = row.cell(entryNumberIndex);
                let clientId = row.cell(clientIdIndex);
                if (entryNumber.empty() && clientId.empty()) break;
                let collateralShortName = row.cell(collateralShortNameIndex);
                let collateralFullName = row.cell(collateralFullNameIndex);
                let collateralGrantorName = row.cell(collateralGrantorNameIndex);
                let collateralGrantorID = row.cell(collateralGrantorIDIndex);
                let collateralSecuredPartyName = row.cell(collateralSecuredPartyNameIndex);
                let collateralSecuredPartyID = row.cell(collateralSecuredPartyIDIndex);
                let collateralIssueDate = util.extractDate(row.cell(collateralIssueDateIndex), "dd.MM.yyyy");
                let collateralMaturityDate = util.extractDate(row.cell(collateralMaturityDateIndex), "dd.MM.yyyy");
                let collateralKnowledgeDate = util.extractDate(row.cell(collateralKnowledgeDateIndex), "dd.MM.yyyy");
                let collateralCollateralManager = row.cell(collateralCollateralManagerIndex);
                let collateralValuationCurrency = row.cell(collateralValuationCurrencyIndex);
                let collateralInstrumentType = row.cell(collateralInstrumentTypeIndex);
				
	
	
				collateralCompteur += 1;
            }
			domos.log().info("Number of collateral is " + collateralCompteur);
        }*/
    }
})()