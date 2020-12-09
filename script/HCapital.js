"use strict";
(function() {

    let xl = domos.excel();
    let objects = domos.objects();

    for (let sheet of xl.sheets()) {
        // Fund Entities data
        if (sheet.name() == "ESID") {
            let fundCompteur = 1;
            let companyCompteur = 1;

            // Find header row
            let mainIndex = sheet.defaultIndex();

            // Entities data headers
            let entityLevelIndex = mainIndex.get("Level");
			let entityFIndex = mainIndex.get("F");
			let entityL1Index = mainIndex.get("L1");
			let entityL2Index = mainIndex.get("L2");
			let entityL3Index = mainIndex.get("L3");
			let entityL4Index = mainIndex.get("L4");
			let entityTypeIndex = mainIndex.get("Type");
			let entityLocationCountryIndex = mainIndex.get("Location (country)");
			let entityCurrencyIndex = mainIndex.get("Currency");
			let entitypPercentOwnershipIndex = mainIndex.get("% of ownership");
			let entityQuantityHeldSharesIndex = mainIndex.get("Quantity of held shares");
			let entityPicePerSharesIndex = mainIndex.get("Pice per shares");
			let entityEquityAmountIndex = mainIndex.get("equity amount");
			let entityDebtAmountIndex = mainIndex.get("debt amount");
			let entityTotalIndex = mainIndex.get("Total");
			let entitySalePriceIndex = mainIndex.get("Sale price");

            for (let row of mainIndex.dataRows()) {
                // Retrieve entity data
                let entityLevel = row.cell(entityLevelIndex);
				let entityF  = row.cell(entityFIndex);
				let entityL1 = row.cell(entityL1Index);
				let entityL2 = row.cell(entityL2Index);
				let entityL3 = row.cell(entityL3Index);
				let entityL4 = row.cell(entityL4Index);
				let entityType = row.cell(entityTypeIndex);
				let entityLocationCountry = row.cell(entityLocationCountryIndex);
				let entityCurrency = row.cell(entityCurrencyIndex);
				let entitypPercentOwnership = row.cell(entitypPercentOwnershipIndex);
				let entityQuantityHeldShares = row.cell(entityQuantityHeldSharesIndex);
				let entityPicePerShares = row.cell(entityPicePerSharesIndex);
				let entityEquityAmount = row.cell(entityEquityAmountIndex);
				let entityDebtAmount = row.cell(entityDebtAmountIndex);
				let entityTotal = row.cell(entityTotalIndex);
				let entitySalePrice = row.cell(entitySalePriceIndex);
				// Create entity
		        let entity;
		        let entityName;
		        if (!entityF.empty()) {
		            entityName = entityF.stringOrNull();
		        } else if (!entityL1.empty()) {
		            entityName = entityL1.stringOrNull();
		        } else if (!entityL2.empty()) {
		            entityName = entityL2.stringOrNull();
		        } else if (!entityL3.empty()) {
		            entityName = entityL3.stringOrNull();
		        } else if (!entityL4.empty()) {
		            entityName = entityL4.stringOrNull();
		        }
				if(!entityType.empty() && entityType.stringOrNull() == "Sub-Fund"){
			        entity = objects.findOrCreateVehicle(entityName);
			        fundCompteur += 1;
				}else if(!entityType.empty() && entityType.stringOrNull() == "Company"){
				    entity = objects.findOrCreateCompany(entityName);
				     companyCompteur += 1;
				}else if(!entityType.empty() && entityType.stringOrNull() == "SPV"){
				    entity = objects.findOrCreateCompany(entityName);
				    entity.objectType("HOLDING");
				    companyCompteur += 1;
				}else{
			        continue;
		        }
		        if (!entityLocationCountry.empty()) {
			        let country = domos.inferCountryCodeOrNull(entityLocationCountry.stringOrNull());
			        if (country) {
				        entity.country(country);
				        if (entity.isEntity()) entity.addresses().legal(true).country(country);
			        }else {
				        domos.log().warn("Country " + entityLocationCountry.stringOrNull() + " not supported for entity" + entityName);
			        }
		        }
		        if (!entityCurrency.empty()) entity.currency(entityCurrency);
			}
			domos.log().info("Number of fund position is " + fundCompteur);
            domos.log().info("Number of company position is " + companyCompteur);
		}
	}
})();
