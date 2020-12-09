"use strict";
(function() {

    let xl = domos.excel();
    let objects = domos.objects();

    for (let sheet of xl.sheets()) {
        // Fund Entities data
        if (sheet.name() == "Inventory") {
            // Find header row
            let mainIndex = sheet.defaultIndex();
            // Entities data headers
            let entityPositionDateIndex = mainIndex.get("Position Date");
            let entityStructureLevelIndex = mainIndex.get("Structure Level");
            let entityCategoryIndex = mainIndex.get("Category");
            let entityDirectionIndex = mainIndex.get("Direction");
            let entityAssetNameIndex = mainIndex.get("Asset Name");
            let entityInvestorIndex = mainIndex.get("Investor");
            let entityIssuerIndex = mainIndex.get("Issuer");
            let entityInvestorTypeIndex = mainIndex.get("Investor Type");
            let entityIssuerTypeIndex = mainIndex.get("Issuer Type");
            let entityInstrumentTypeIndex = mainIndex.get("Instrument Type");
            let entityInstrumentNameIndex = mainIndex.get("Instrument Name");
            let entityQuantityIndex = mainIndex.get("Quantity");
            let entityNominalPerUnitIndex = mainIndex.get("Nominal Per Unit");
            let entityNominalAmountIndex = mainIndex.get("Nominal Amount");
            let entityCommitmentIndex = mainIndex.get("Commitment");
            let entityCurrencyIndex = mainIndex.get("Currency");
            let entityInterestRateIndex = mainIndex.get("Interest Rate");
            let entityIssueDateIndex = mainIndex.get("Issue Date");
            let entityMaturityDateIndex = mainIndex.get("Maturity date");
            let entityOwnershipPercentageIndex = mainIndex.get("Ownership Percentage");
            let entityInstrumentISINIndex = mainIndex.get("Instrument ID_ISIN");
            let entityInvestorIDIndex = mainIndex.get("Investor ID");
            let entityIssuerIDIndex = mainIndex.get("Issuer ID");
            let entityIssuerCountryIndex = mainIndex.get("Issuer Country");
            let entityAddPortfolioPositionsIndex = mainIndex.get("Add To Portfolio Positions List Y/N");
            let entityAddressIndex = mainIndex.get("Address");
            let entityCountryIndex = mainIndex.get("Country");
            let entityDescriptionIndex = mainIndex.get("Description");
            for (let row of mainIndex.dataRows()) {
                // Retrieve entity data
                let entityPositionDate = row.cell(entityPositionDateIndex);
                let entityStructureLevel = row.cell(entityStructureLevelIndex);
                let entityCategory = row.cell(entityCategoryIndex);
                let entityDirection = row.cell(entityDirectionIndex);
                let entityAssetName = row.cell(entityAssetNameIndex);
                let entityInvestor = row.cell(entityInvestorIndex);
                let entityIssuer = row.cell(entityIssuerIndex);
                let entityInvestorType = row.cell(entityInvestorTypeIndex);
                let entityIssuerType = row.cell(entityIssuerTypeIndex);
                let entityInstrumentType = row.cell(entityInstrumentTypeIndex);
                let entityInstrumentName = row.cell(entityInstrumentNameIndex);
                let entityQuantity = row.cell(entityQuantityIndex);
                let entityNominalPerUnit = row.cell(entityNominalPerUnitIndex);
                let entityNominalAmount = row.cell(entityNominalAmountIndex);
                let entityCommitment = row.cell(entityCommitmentIndex);
                let entityCurrency = row.cell(entityCurrencyIndex);
                let entityInterestRate = row.cell(entityInterestRateIndex);
                let entityIssueDate = row.cell(entityIssueDateIndex);
                let entityMaturityDate = row.cell(entityMaturityDateIndex);
                let entityOwnershipPercentage = row.cell(entityOwnershipPercentageIndex); 
                let entityInstrumentISIN = row.cell(entityInstrumentISINIndex);
                let entityInvestorID = row.cell(entityInvestorIDIndex);
                let entityIssuerID = row.cell(entityIssuerIDIndex);
                let entityIssuerCountry = row.cell(entityIssuerCountryIndex);
                let entityAddPortfolioPositions = row.cell(entityAddPortfolioPositionsIndex);
                let entityAddress = row.cell(entityAddressIndex);
                let entityCountry = row.cell(entityCountryIndex);
                let entityDescription = row.cell(entityDescriptionIndex);
                
                //Find or create investor
                let investor;
                if(!entityInvestorID.empty() && entityInvestorType.stringOrNull().toUpperCase().includes("FUND")) investor = objects.findVehicleOrNull(entityInvestorID);
                if (!investor && !entityInvestor.empty() && !entityInvestorType.empty() && entityInvestorType.stringOrNull().toUpperCase().includes("FUND")) {
                    investor = objects.findOrCreateVehicle(entityInvestor);
                     if(!entityInvestorID.empty()) investor.customReference("ReferenceNoclinodos",entityInvestorID);
                }
                if(investor && !entityCurrency.empty()) investor.currency(entityCurrency);
                //Find or create issuer
                let issuer;
                if(!entityIssuerID.empty() && !entityIssuerType.empty() && entityIssuerType.stringOrNull().toUpperCase().includes("FUND")) issuer = objects.findVehicleOrNull(entityIssuerID);
                if (!issuer && !entityIssuer.empty() && !entityIssuerType.empty() && entityIssuerType.stringOrNull().toUpperCase().includes("FUND")) {
                    issuer = objects.findOrCreateVehicle(entityIssuer);
                    if(!entityIssuerID.empty()) issuer.customReference("ReferenceNoclinodos",entityIssuerID);
                }
                if(!entityIssuerID.empty() && !entityIssuerType.empty() && entityIssuerType.stringOrNull().toUpperCase().includes("FUND")) issuer = objects.findVehicleOrNull(entityIssuerID);
                if (!entityIssuer.empty() && !entityIssuerType.empty() && (entityIssuerType.stringOrNull().toUpperCase().includes("COMPANY") || entityIssuerType.stringOrNull().toUpperCase().includes("HOLDING"))) {
                    issuer = objects.findOrCreateCompany(entityIssuer);
                    if (issuer && (entityIssuerType.stringOrNull().toUpperCase().includes("HOLDING"))) {
                        issuer.objectType("HOLDING");
                    }
                    if(!entityIssuerID.empty()) issuer.customReference("ReferenceNoclinodos",entityIssuerID);
                }
                if(issuer && !entityCurrency.empty()) issuer.currency(entityCurrency);
                if(issuer && !entityIssuerCountry.empty()) {
					let country = domos.inferCountryCodeOrNull(entityIssuerCountry.stringOrNull());
                    if (country) {
				        issuer.country(country);
				        if (issuer.isEntity()) issuer.addresses().legal(true).country(country);
			        }else {
				        domos.log().warn("Country " + entityIssuerCountry.stringOrNull() + " not supported for entity" + issuer);
			        }
                }
                
                //Create Fund share
                if(!entityCategory.empty() && entityCategory.stringOrNull().includes("Fund share")){
                    let commitment;
                    let ownership;
                    if(!entityOwnershipPercentage.empty()) ownership = entityOwnershipPercentage;
                    let quantity = 1;
                    if(!entityQuantity.empty()) quantity = entityQuantity;
                    let amount = 1;
                    if(!entityCommitment.empty()) {
                        amount = entityCommitment;
                    }else if(!entityQuantity.empty() && !entityNominalPerUnit.empty()){
                        amount = entityQuantity.doubleValue() * entityNominalPerUnit.doubleValue();
                    }
                    if(!entityAssetName.empty()){
                        commitment = objects.findOrCreateFundShare(entityAssetName);
                      if(ownership){
                        issuer.addFundShare(commitment,quantity,ownership);  
                      } else {
                          issuer.addFundShare(commitment,quantity,null);
                      }
                    }
                    if(commitment && !entityCurrency.empty()) commitment.currency(entityCurrency);
                    //Add Subscription
                    if(investor && issuer){
                        let defaultDate = util.extractDate(entityPositionDate, "dd.MM.yyyy");
                        commitment.findOrCreateSubscription(amount,entityPositionDate.dateOrNull(), investor)
                                .currency(commitment.currency())
                                .quantity(quantity);
                    } 
                    //Add tags
                    if (!entityInstrumentType.empty()) commitment.addTag("Equity_Type",entityInstrumentType);
                    if (!entityNominalPerUnit.empty()) commitment.addTag("Purchase_Price",entityNominalPerUnit);
                    if (!entityNominalAmount.empty()) commitment.addTag("Equity_Nominal",entityNominalAmount);
                    //Add ISIN
                    if (!entityInstrumentISIN.empty()) commitment.isin(entityInstrumentISIN);
                    //ADD to position list
                    if(!entityAddPortfolioPositions.empty() && entityAddPortfolioPositions.stringOrNull()== "Y"){
                        let portfolio = investor.portfolioOrNull();
                        if(portfolio) portfolio.addPosition(commitment, quantity , amount, entityPositionDate.dateOrNull());
                    }
                }
                //Create shares
                if(!entityCategory.empty() && entityCategory.stringOrNull().includes("Company share")){
                    let share;
                    let ownership = 100;
                    if(!entityOwnershipPercentage.empty()) ownership = entityOwnershipPercentage;
                    let quantity = 1;
                    if(!entityQuantity.empty()) quantity = entityQuantity;
                    let amount = 1;
                    if(!entityNominalAmount.empty()) {
                        amount = entityNominalAmount;
                    }else if(!entityQuantity.empty() && !entityNominalPerUnit.empty()){
                        amount = entityQuantity.doubleValue() * entityNominalPerUnit.doubleValue();
                    }
                    if(!entityAssetName.empty()){
                        share = objects.findOrCreateCompanyShare(entityAssetName);
                        issuer.addCompanyShare(share,quantity,ownership);
                    }
                    if(share && !entityCurrency.empty()) share.currency(entityCurrency);
                    //Add Subscription
                    if(investor && issuer){
                        let defaultDate = util.extractDate(entityPositionDate, "dd.MM.yyyy"); 
						share.findOrCreateCompanySubscription(entityPositionDate.dateOrNull(), investor, null)
                             .quantity(quantity)
                             .amount(amount);
                    } 
                    //Add tags
                    if (!entityInstrumentType.empty()) share.addTag("Equity_Type",entityInstrumentType);
                    if (!entityNominalPerUnit.empty()) share.addTag("Purchase_Price",entityNominalPerUnit);
                    if (!entityNominalAmount.empty()) share.addTag("Equity_Nominal",entityNominalAmount);
                    //Add ISIN
                    if (!entityInstrumentISIN.empty()) share.isin(entityInstrumentISIN);
                    //ADD to position list
                    if(!entityAddPortfolioPositions.empty() && entityAddPortfolioPositions.stringOrNull()== "Y"){
                        let portfolio = investor.portfolioOrNull();
                        if(portfolio) portfolio.addPosition(share, quantity , amount, entityPositionDate.dateOrNull());
                    }
                }
                
                //Create depts
                if(!entityCategory.empty() && (entityCategory.stringOrNull().includes("Debt") || entityCategory.stringOrNull().includes("Credit line"))){
                    let loan;
                    if(!entityAssetName.empty()){
                        loan = objects.findOrCreateDebt(entityAssetName);
                        loan.repaymentFrequency("FINAL");
                    }
                    if(!entityIssueDate.empty()){
                        loan.startDate(entityIssueDate.dateValue());
                    } 
                    if(!entityMaturityDate.empty()){
                        loan.maturityDate(entityMaturityDate.dateValue());
                    } 
                    if(!entityCurrency.empty()) loan.currency(entityCurrency);
                    if(!entityInterestRate.empty()) loan.ratePercent(entityInterestRate.doubleOrNull());
                    let quantity = 1;
                    if(!entityQuantity.empty()) quantity = entityQuantity;
                    let amount = 1;
                    if(!entityNominalAmount.empty()) {
                        amount = entityNominalAmount;
                    }else if(!entityQuantity.empty() && !entityNominalPerUnit.empty()){
                        amount = entityQuantity.doubleValue() * entityNominalPerUnit.doubleValue();
                    }
                    loan.nominal(amount);
                   
                    if(entityCategory.stringOrNull().includes("Credit line")) loan.objectType("CREDIT_LINE");
                    if(entityOwnershipPercentage.empty() || entityOwnershipPercentage.doubleValue()== 1){
                        if(!entityDirection.empty() && entityDirection.stringOrNull().includes("Lending")){
                            loan.direction(2);
                            if (issuer) loan.issuer(issuer);
			                if (investor) loan.lender(investor);
                        }else{
                            loan.direction(1);
                            if (investor) loan.issuer(investor);
			                if (issuer) loan.lender(issuer);
                        }
                    }else{
                        let debtSubscription = loan.createDebtSubscription();
                        debtSubscription = debtSubscription.nominal(amount)
                                               .value(amount)
                                               .investor(investor)
                                               .date(loan.startDate());
                    }
                    //Add ISIN
                    if (!entityInstrumentISIN.empty()) loan.isin(entityInstrumentISIN);
                    if(!entityAddPortfolioPositions.empty() && entityAddPortfolioPositions.stringOrNull()== "Y"){
                        let portfolio = investor.portfolioOrNull();
                        if(portfolio) portfolio.addPosition(loan, quantity , amount, entityPositionDate.dateOrNull());
                    }
                }
            //Create properties
            if(!entityCategory.empty() && entityCategory.stringOrNull().includes("Property")){
                let property;
                if(!entityAssetName.empty()){
                        property = objects.findOrCreateObjectType(entityAssetName, "PROPERTY");
                    }
                if(property && !entityCountry.empty()) {
                    
                    let country = domos.inferCountryCodeOrNull(entityIssuerCountry.stringOrNull());
                    if (country) {
				        property.country(country);
				        if (property.isEntity()) property.addresses().legal(true).country(country);
			        }else {
				        domos.log().warn("Country " + entityIssuerCountry.stringOrNull() + " not supported for entity" + property);
			        }
                }entityAddress
                if (!entityAddress.empty()) property.address(entityAddress);
                let currency ="EUR";
                if (!entityCurrency.empty()) currency = entityCurrency.stringOrNull();
                property.currency(currency);
                if (!entityDescription.empty()) property.description(entityDescription);
                let amount = 1;
                    if(!entityNominalAmount.empty()) {
                        amount = entityNominalAmount;
                    }else if(!entityQuantity.empty() && !entityNominalPerUnit.empty()){
                        amount = entityQuantity.doubleValue() * entityNominalPerUnit.doubleValue();
                    }
                if (amount) property.addValuation(amount,currency,entityPositionDate.dateOrNull());
            }
            
            }
        }
    }
})();