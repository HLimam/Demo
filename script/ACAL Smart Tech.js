"use strict";

(function() {

    let xl = domos.excel();
    let objects = domos.objects();

    // Retrieve File name
    let fileName = domos.name();
    let date = util.extractDate(fileName, "dd.MM.yyyy");

    if (!fileName) {
        domos.log().error("No fund found in file name " + fileName);
        return;
    }

    for (let sheet of xl.sheets()) {

        // Fund Entities data
        if (sheet.name() == "Entities") {
            let fundCompteur = 0;
            let companyCompteur = 0;

            // Find header row
            let mainIndex = sheet.defaultIndex();

            // Entities data headers
            let entityClientIdIndex = mainIndex.get("Client ID");
            let entityEntryNumberIndex = mainIndex.get("Entry number");
            let entityNameIndex = mainIndex.get("Full name");
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
                if (entityClientId.empty()) break;
                let entityEntryNumber = row.cell(entityEntryNumberIndex);
                let entityName = row.cell(entityNameIndex);
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
                if (entityRole.stringOrNull() == "Fund (Client)") {
                    entity = objects.findOrCreateVehicle(entityName.stringValue());
                    fundCompteur = fundCompteur + 1;
                }
                //Create Company
                if (entityRole.stringOrNull() == "Holding Company" || entityRole.stringOrNull() == "Operating Company") {
                    entity = objects.findOrCreateCompany(entityName.stringValue());
                    if (entityRole.stringOrNull() == "Holding Company") {
                        entity.objectType("HOLDING");
                                            }
                    companyCompteur = companyCompteur + 1;
                }
                if(entity){
                if (!entityCurrency.empty()) entity.currency(entityCurrency.stringValue());
                if (!entityCountry.empty()) {
                    let country = domos.inferCountryCodeOrNull(entityCountry.stringOrNull());
                    if (country) {
                        entity.country(country);
                        if (entity.isEntity()) entity.addresses().legal(true).country(country);
                    } else {
                        domos.log().error("Country " + country + " not supported for entity" + entityName.stringValue());
                    }
                }
                if (entity && entity.isEntity() && !entityAddress.empty())  entity.addresses().legal(true).street(entityAddress);
                
                //Add Reference
                if (!entityClientId.empty()) entity.customReference("ReferenceNoclinodos", entityClientId);
                }
            }
            domos.log().info("Number of fund position is " + fundCompteur);
            domos.log().info("Number of company position is " + companyCompteur);
            domos.log().info("Entities data has been imported");
        }

        // Fund Equities data
        if (sheet.name() == "Equities") {
            let companyShareCompteur = 0;
            // Find header row
            let mainIndex = sheet.defaultIndex();

            // Assets data headers
            let clientIdIndex = mainIndex.get("Client ID");
            let assetInvestorIndex = mainIndex.get("Investor");
            let assetInvesteeIndex = mainIndex.get("Investee");
            let assetInstrumentTypeIndex = mainIndex.get("Instrument type");
            let assetInstrumentNameIndex = mainIndex.get("Instrument name");
            let assetOperationCurrencyIndex = mainIndex.get("Operation currency");
            let assetQuantityMemoIndex = mainIndex.get("Quantity_Memo");
            let assetUnitPriceOpeCcyIndex = mainIndex.get("Unit price ope ccy_Memo");
            let assetValuationOpeCcyIndex = mainIndex.get("Valuation in ope ccy");
            for (let row of mainIndex.dataRows()) {
                // Retrieve entity data
                let clientId = row.cell(clientIdIndex);
                if (clientId.empty()) break;
                let assetInvestor = row.cell(assetInvestorIndex);
                let assetInvestee = row.cell(assetInvesteeIndex);
                let assetInstrumentType = row.cell(assetInstrumentTypeIndex);
                let assetInstrumentName = row.cell(assetInstrumentNameIndex);
                let assetOperationCurrency = row.cell(assetOperationCurrencyIndex);
                let assetUnitPrice = row.cell(assetUnitPriceOpeCcyIndex);
                let assetQuantityMemo = row.cell(assetQuantityMemoIndex);
                let assetValuationOpeCcy = row.cell(assetValuationOpeCcyIndex);
                if (!assetInvestor.empty() && !assetInvestee.empty()) {
                    // Create entity
                    let asset;
                    //domos.log().info("asset " + assetInvestee.stringValue() + " " + assetInstrumentType.stringValue());
                    //Create Company share
                    let companyShareName = assetInvestee.stringValue() + " " + assetInstrumentType.stringValue();
                    asset = objects.findOrCreateCompanyShare(companyShareName);
                    let shareType;
                    if(assetInstrumentType.stringValue() == "Shares"){
                        shareType = 1;
                    }else if (assetInstrumentType.stringValue() == "Equity"){
                        shareType = 5;
                    }
                    if(shareType) asset.shareType(shareType);
                    asset.currency(assetOperationCurrency);
                    let company = objects.findCompanyOrNull(assetInvestee.stringValue());
                    company.addCompanyShare(asset);
                    //Add a transaction for the company d'Twin because it haven't yet a equity movements sheet 
                    if ((assetInvestee.stringValue() == "d'Twin Alpha Sàrl" && assetInvestor.stringValue() == "Acal Private Equtiy Sicar Smart Tech")
                    || (assetInvestee.stringValue() == "AB Motion Tech Sàrl" && assetInvestor.stringValue() == "Acal Private Equtiy Sicar Smart Tech"
                    && assetInstrumentType.stringValue() =="Shares")) {
                        let investor = objects.findVehicleOrNull(assetInvestor.stringValue());
                        if (!investor) {
                            investor = objects.findCompanyOrNull(assetInvestor.stringValue());
                        }
                        let defaultDate = util.extractDate("31/12/2016", "dd.MM.yyyy");
                        if(investor && !assetValuationOpeCcy.empty() && !assetQuantityMemo.empty()){
                            asset.findOrCreateCompanySubscription(defaultDate, investor, null)
                            .quantity(assetQuantityMemo)
                            .amount(assetValuationOpeCcy);
                            
                        }
                        asset.setLast(assetUnitPrice, assetOperationCurrency, null);
                    }
                    companyShareCompteur += 1;
                }
            }
            domos.log().info("Number of company share position is " + companyShareCompteur);
            domos.log().info("Equities data has been imported");
        }

        // Fund Transactions data
        if (sheet.name().includes("Equity movements")) {
            let transactionsCompteur = 0;

            // Find header row
            let mainIndex = sheet.defaultIndex();
            // Operations data headers
            let clientIdIndex = mainIndex.get("Client ID");
            let operationDateIndex = mainIndex.get("Operation date");
            let operationInvestorIndex = mainIndex.get("Investor");
            let operationInvesteeIndex = mainIndex.get("Investee");
            let operationCurrencyIndex = mainIndex.get("Operation ccy");
            let operationFundedAmountIndex = mainIndex.get("Funded amount ope ccy");
            let operationConsiderationAmountIndex = mainIndex.get("Consideration amount ope ccy");
            let operationQuantityIndex = mainIndex.get("Quantity");
            let operationUnitPriceIndex = mainIndex.get("Unit price ope ccy");
            for (let row of mainIndex.dataRows()) {
                // Retrieve operation data
                let clientId = row.cell(clientIdIndex);
                if (clientId.empty()) break;
                let operationDate = row.cell(operationDateIndex);
                let operationInvestor = row.cell(operationInvestorIndex);
                let operationInvestee = row.cell(operationInvesteeIndex);
                let operationCurrency = row.cell(operationCurrencyIndex);
                let operationQuantity = row.cell(operationQuantityIndex);
                let operationAmount = row.cell(operationFundedAmountIndex);
                let operationUnitPrice = row.cell(operationUnitPriceIndex);
                if (operationAmount.empty()) operationAmount = row.cell(operationConsiderationAmountIndex);
                if (operationUnitPrice.empty() && (operationQuantity !== 0)) operationUnitPrice = operationAmount / operationQuantity;
                let companyShare;
                if(operationInvestee.stringValue() != "AB Motion Tech Sàrl")companyShare=objects.findCompanyShareOrNull(operationInvestee.stringValue() + " Shares");
                //domos.log().debug("companyShare : "+companyShare);
                if (sheet.name() == "Equity movements ac 115") {
                    if(!companyShare) companyShare = objects.findCompanyShareOrNull("AB Motion Tech Sàrl Equity");
                    companyShare.shareType(5);
                }
                let company = objects.findCompanyOrNull(operationInvestee.stringValue());
                let investor = objects.findVehicleOrNull(operationInvestor.stringValue());
                if (!investor) {
                    investor = objects.findCompanyOrNull(operationInvestor.stringValue());
                }
                if (companyShare && !operationQuantity.empty() && !operationAmount.empty()) {
                    companyShare.findOrCreateCompanySubscription(operationDate.dateValue(), investor, null)
                        .quantity(operationQuantity)
                        .amount(operationAmount);
                    //  companyShare.setLast(operationUnitPrice, operationCurrency, operationDate.dateValue());
                } else {
                    domos.log().error("sheet: " + sheet.name() + " company: " + company + " investor:" + investor);
                }
            }
        }


        // Fund Loans data
        if (sheet.name().includes("Loans")) {
            let loansCompteur = 0;
            // Find header row
            let mainIndex = sheet.defaultIndex();
            // Loans data headers
            let clientIdIndex = mainIndex.get("Client ID");
            let loanFullNameIndex = mainIndex.get("Full name");
            let loanLenderIndex = mainIndex.get("Lender");
            let loanBorrowerIndex = mainIndex.get("Borrower");
            let loanInstrumentTypeIndex = mainIndex.get("Instrument type");
            let loanIssueDateIndex = mainIndex.get("Issue date");
            let loanMaturityDateIndex = mainIndex.get("Maturity date");
            let loanOperationCurrencyIndex = mainIndex.get("Operation currency");
            let loanInterestRateIndex = mainIndex.get("Interest rate %");
            let loanCommittedAmountIndex = mainIndex.get("Committed amount ope ccy_Memo");
            let loanFundedAmountIndex = mainIndex.get("Funded amount ope ccy_Memo");
            for (let row of mainIndex.dataRows()) {
                // Retrieve operation data
                let clientId = row.cell(clientIdIndex);
                if (clientId.empty()) break;
                let loanFullName = row.cell(loanFullNameIndex);
                let loanLender = row.cell(loanLenderIndex);
                let loanBorrower = row.cell(loanBorrowerIndex);
                let loanInstrumentType = row.cell(loanInstrumentTypeIndex);
                let loanIssueDate = row.cell(loanIssueDateIndex);
                let loanMaturityDate = row.cell(loanMaturityDateIndex);
                let loanOperationCurrency = row.cell(loanOperationCurrencyIndex);
                let loanInterestRate = row.cell(loanInterestRateIndex);
                let loanCommittedAmount = row.cell(loanCommittedAmountIndex);
                let loanFundedAmount = row.cell(loanFundedAmountIndex);
                loansCompteur += 1;
                let loanName = loanFullName.stringValue() + " - Loan " + loansCompteur;
                let loan = objects.findOrCreateDebt(loanName);
                loan.shortName(loanFullName);
                loan.currency(loanOperationCurrency);
                loan = loan.startDate(loanIssueDate);
                loan = loan.maturityDate(loanMaturityDate);
                loan = loan.ratePercent(loanInterestRate);
                loan = loan.nominal(loanFundedAmount);
                let borrower = objects.findCompanyOrNull(loanBorrower);
                let lender = objects.findCompanyOrNull(loanLender);
                if (borrower) loan = loan.issuer(borrower);
                if (lender) loan = loan.lender(lender);
            }

            domos.log().info("Number of loan is " + loansCompteur);
        }

    }
})();