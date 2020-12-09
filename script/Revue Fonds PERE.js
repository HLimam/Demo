"use strict";
(function() {

    let xl = domos.excel();
    let objects = domos.objects();

    for (let sheet of xl.sheets()) {
        // Fund Entities data
        if (sheet.name() == "Funds") {
            let elementCompteur = 0;

            // Find header row
            let mainIndex = sheet.defaultIndex();

            // Entities data headers
            let entityNocliNodosIndex = mainIndex.get("NocliNodos");
            let entityNomOPCIndex = mainIndex.get("Nom de l"OPC");
            let entityNomCompartimentIndex = mainIndex.get("Nom du compartiment");
            let entitySectorIndex = mainIndex.get("Sector");
            let entityDetailSectorIndex = mainIndex.get("Detail Sector");
            let entityOnboardingIndex = mainIndex.get("Onboarding");

            for (let row of mainIndex.dataRows()) {
                // Retrieve entity data
                let entityNocliNodos = row.cell(entityNocliNodosIndex);
                let entityNomOPC = row.cell(entityNomOPCIndex);
                let entityNomCompartiment = row.cell(entityNomCompartimentIndex);
                let entitySector = row.cell(entitySectorIndex);
                let entityDetailSector = row.cell(entityDetailSectorIndex);
                let entityOnboarding = row.cell(entityOnboardingIndex);
                
                if(!entityNomCompartiment.empty() && entityNomCompartiment.stringOrNull() != "?"){
                    // Find entity
                    let entity = objects.findVehicleOrNull(entityNomCompartiment);
                    if(entity){
                        elementCompteur +=1;
                        // Add Reference
		                if (!entityNocliNodos.empty() && entityNocliNodos.stringOrNull() != "?") entity.customReference("ReferenceNoclinodos", entityNocliNodos);
		                // Add Tags       
		                if (!entityNomOPC.empty()) entity.addTag("nomOPC",entityNomOPC);
		                if (!entitySector.empty()){
		                    let sector;
		                    switch (entitySector.stringValue()) {
                                case "AEROSPACE"                   : sector = "aerospace"              ; break;
								case "AGRICULTURE & FORESTRY"      : sector = "agricultureForestry"    ; break;
								case "ART & LUXURY GOODS"          : sector = "artLuxuryGoods"         ; break;
								case "AUTOMOBILE & TRANSPORTS"     : sector = "automobileTransports"   ; break;
								case "BANK & INSURANCE"            : sector = "bankInsurance"          ; break;
								case "COMMODITIES / ROW MATERIALS" : sector = "commoditiesRowMaterials"; break;
								case "CONSTRUCTION"                : sector = "construction"           ; break;
								case "ENVIRONMENT"                 : sector = "environment"            ; break;
								case "FACTORING"                   : sector = "factoring"              ; break;
								case "HEALTH CARE & EDUCATION"     : sector = "healthCareEducation"    ; break;
								case "INFORMATION TECHNOLOGY"      : sector = "informationTechnology"  ; break;
								case "LEISURE HOTEL & RESORTS"     : sector = "leisureHotelResorts"    ; break;
								case "MANUFACTURING"               : sector = "manufacturing"          ; break;
								case "MEDIA & ENTERTAINMENT"       : sector = "mediaEntertainment"     ; break;
								case "MICRO FINANCE"               : sector = "microFinance"           ; break;
								case "PRECIOUS METALS"             : sector = "preciousMetals"         ; break;
								case "REAL ESTATE"                 : sector = "realEstate"             ; break;
								case "SECURITIZATION"              : sector = "securitization"         ; break;
								case "TANGIBLE ASSETS"             : sector = "tangibleAssets"         ; break;
								case "TELECOMMUNICATIONS"          : sector = "telecommunications"     ; break;
								case "UTILITIES & ENERGY"          : sector = "utilitiesEnergy"        ; break;
							}
							entity.addTag("sector",sector);	
		                } 
		                if (!entityDetailSector.empty()) entity.addTag("detailSector",entityDetailSector);
		                
		                if (!entityOnboarding.empty()){
		                    let onboarding;
		                    switch (entityOnboarding.stringValue()) {
                                case "EXISTING FUND": onboarding = "existingFund"; break;
                                case "MIGRATION": onboarding = "migration"; break;
                                case "LAUNCHING": onboarding = "launching"; break;
                                case "PASSATION DOSSIER": onboarding = "passationDossier"; break;
		                    }
		                    domos.log().debug(entityOnboarding.stringValue()+" || "+onboarding);
		                    entity.addTag("onboardingType",onboarding);
		                } 
                    }
                }
            }
            domos.log().info("number of item to process is " + elementCompteur);
        }
    }
})();