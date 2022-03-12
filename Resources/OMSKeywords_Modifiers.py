from robot.libraries.BuiltIn import BuiltIn
from robot.api.deco import keyword
import AccessRepository as AR
import CommonFunctions as CF

def getData():
    loc = 'Data\\FalconTestData.xlsx'
    SheetName = BuiltIn().get_variable_value('${Company}')
    testcase = 'Modifiers'
    return CF.fncGetValues(loc, SheetName, testcase)

@keyword
def Go_To_Modifiers():
    sl = BuiltIn().get_library_instance('SeleniumLibrary')
    try:
        if sl.get_text(AR.xPathOMSSubmissionBanner3) != 'Coverage Modifiers for Rerate Policy':
            sl.click_element(AR.ModifiersTab)

    except Exception as e1:
        BuiltIn().log('Modifiers Tab already open')

@keyword
def Add_Coverage_Modifiers():
    myData = getData()
    uCreditForSideA = myData['Option1']
    uCreditForIDL = myData['Option2']
    uNoOfOperations = myData['Option3']
    uNYOAppliedFactor = myData['Option4']
    uSNFMAActivity = myData['Option5']
    uSNFMAAppliedFactor = myData['Option6']
    uSECOffering = myData['Option7']
    uSECOfferingAppliedFactor = myData['Option8']
    uDOLitigation = myData['Option9']
    uDOLitigationAppliedFactor = myData['Option10']
    uOtherLitigation = myData['Option11']
    uOtherLitigationAppliedFactor = myData['Option12']
    uCoinsurance = myData['Option13']
    uCOSECClaims = myData['Option14']
    uCOSECClaimsAppliedFactor = myData['Option15']
    uIndustryRisk = myData['Option16']
    uIndustryRiskAppliedFactor = myData['Option17']
    uDiscovery = myData['Option18']
    uDiscoveryAppliedFactor = myData['Option19']

    sl = BuiltIn().get_library_instance('SeleniumLibrary')

    # Credit for Side A Only with DIC Coverage
    if uCreditForSideA != "" and uCreditForSideA != 1:
        sl.input_text(AR.idtxtMSideAOnly, uCreditForSideA)

    # Credit for IDL Only with DIC Coverage
    if uCreditForIDL != "" and uCreditForIDL != 1:
        sl.input_text(AR.idtxtMIDEOnly, uCreditForIDL)

    # Number of Years in Operations and Applied Factor
    if uNoOfOperations != "":
        CF.fncSelectDDElement(sl, AR.idddMNumberOfYearsInOperation, uNoOfOperations)
        sl.input_text(AR.idtxtMNumberOfYearsInOperationsFactors, uNYOAppliedFactor)

    # Significant M&A Activity / Level of M&A Concern and Applied Factor
    if uSNFMAActivity != "":
        CF.fncSelectDDElement(sl, AR.idddMLevelOfMAConcern, uSNFMAActivity)
        sl.input_text(AR.idtxtMLevelOfMAConcernFactor, uSNFMAAppliedFactor)

    # SEC Offering and Applied Factor
    if uSECOffering != "":
        CF.fncSelectDDElement(sl, AR.idddMSECOffering, uSECOffering)
        sl.input_text(AR.idtxtMSECOfferingFactor, uSECOfferingAppliedFactor)

    # D&O Litigation and Applied Factor
    if uDOLitigation != "":
        CF.fncSelectDDElement(sl, AR.idddMDOLitigation, uDOLitigation)
        sl.input_text(AR.idtxtMDOLitigationFactor, uDOLitigationAppliedFactor)

    # Other Litigation and Applied Factor
    if uOtherLitigation != "":
        CF.fncSelectDDElement(sl, AR.idddMOtherLitigation, uOtherLitigation)
        sl.input_text(AR.idtxtMOtherLitigationFactor, uOtherLitigationAppliedFactor)

    # Coinsurance
    if uCoinsurance != "" and uCoinsurance != 1:
        sl.input_text(AR.idtxtMCoinsurance, uCoinsurance)

    # Coinsurance re SEC Claims
    if uCOSECClaims != "":
        CF.fncSelectDDElement(sl, AR.idddMCOSECClaims, uCOSECClaims)
        sl.input_text(AR.idtxtMCOSECClaimsFactor, uCOSECClaimsAppliedFactor)

    # Industry Risk/Level of Confidence in Industry and Applied Factor
    if uIndustryRisk != "":
        CF.fncSelectDDElement(sl, AR.idddMIndustryRisk, uIndustryRisk)
        sl.input_text(AR.idtxtMIndustryRiskFactor, uIndustryRiskAppliedFactor)

    # Discovery (Extended Reporting) and Applied Factor
    if uDiscovery != "":
        CF.fncSelectDDElement(sl, AR.idddMDiscovery, uDiscovery)
        sl.input_text(AR.idtxtMDiscoveryFactor, uDiscoveryAppliedFactor)

@keyword
def Add_Other_RiskFactors():
    myData = getData()
    uClaimLitigationHistory = myData['Option20']
    uCorpGovernanceProcedure = myData['Option21']
    uEarningConsistency = myData['Option22']
    uEffectedByRecession = myData['Option23']
    uEnvironmentalIssues = myData['Option24']
    uFinancialSolvency = myData['Option25']
    uInsiderTradingActivity = myData['Option26']
    uJointVentures = myData['Option27']
    uLaborRelations = myData['Option28']
    uLiquidity = myData['Option29']
    uMajorCustomers = myData['Option30']
    uManagementStability = myData['Option31']
    uManagementOwnership = myData['Option32']
    uOtherFinFactors = myData['Option33']
    uQualityExtBrdMembers = myData['Option34']
    uRegulatoryExposure = myData['Option35']
    uStabilityofWorkforce = myData['Option36']
    uStockMrktSensitivity = myData['Option37']
    uStockVolatility = myData['Option38']
    uTakeoverPotential = myData['Option39']
    uTerrorismRiskDiscount = myData['Option40']
    uTransactionEvent = myData['Option41']

    sl = BuiltIn().get_library_instance('SeleniumLibrary')
    # Claims litigation history (severity) (0.80-1.50)
    if uClaimLitigationHistory != "":
        sl.input_text(AR.idtxtMClaimLitigationHistory, uClaimLitigationHistory)

    # Earnings consistency (0.75-4.00)
    if uEarningConsistency != "":
        sl.input_text(AR.idtxtMEarningConsistency, uEarningConsistency)

    # Environmental issues (0.80-1.50)
    if uEnvironmentalIssues != "":
        sl.input_text(AR.idtxtMEnvironmentalIssues, uEnvironmentalIssues)

    # Insider trading activity (1.00-1.50)
    if uInsiderTradingActivity != "":
        sl.input_text(AR.idtxtMInsiderTradingActivity, uInsiderTradingActivity)

    # Labor relations (0.80-1.50)
    if uLaborRelations != "":
        sl.input_text(AR.idtxtMLaborRelations, uLaborRelations)

    # Major customers (0.80-1.50)
    if uMajorCustomers != "":
        sl.input_text(AR.idtxtMMajorCustomers, uMajorCustomers)

    # Management ownership (0.80-1.25)
    if uManagementOwnership != "":
        sl.input_text(AR.idtxtMManagementOwnership, uManagementOwnership)

    # Quality of external board members (0.80-1.50)
    if uQualityExtBrdMembers != "":
        sl.input_text(AR.idtxtMQualityExtBrdMembers, uQualityExtBrdMembers)

    # Stability of workforce (0.80-1.50)
    if uStabilityofWorkforce != "":
        sl.input_text(AR.idtxtMStabilityofWorkforce, uStabilityofWorkforce)

    # Stock volatility (0.75-2.50)
    if uStockVolatility != "":
        sl.input_text(AR.idtxtMStockVolatility, uStockVolatility)

    # Terrorism risk discount (0.90-1.00)
    if uTerrorismRiskDiscount != "":
        sl.input_text(AR.idtxtMTerrorismRiskDiscount, uTerrorismRiskDiscount)

    # Corporate governance procedures(0.75 - 1.50)
    if uCorpGovernanceProcedure != "":
        sl.input_text(AR.idtxtMCorpGovernanceProcedure, uCorpGovernanceProcedure)

    # Effected by recession (0.80-1.50)
    if uEffectedByRecession != "":
        sl.input_text(AR.idtxtMEffectedByRecession, uEffectedByRecession)

    # Financial solvency (0.75-2.50)
    if uFinancialSolvency != "":
        sl.input_text(AR.idtxtMFinancialSolvency, uFinancialSolvency)

    # Joint ventures/Limited partnerships/Significant subsidiary operations (incl. SVPs) (1.00-2.50)
    if uLiquidity != "":
        sl.input_text(AR.idtxtMLiquidity, uLiquidity)

    # Management experience or stability (0.80-1.50)
    if uManagementStability != "":
        sl.input_text(AR.idtxtMManagementStability, uManagementStability)

    # Other financial factors (0.80-1.50)
    if uOtherFinFactors != "":
        sl.input_text(AR.idtxtMOtherFinFactors, uOtherFinFactors)

    # Regulatory exposure/experience (0.90-2.00)
    if uRegulatoryExposure != "":
        sl.input_text(AR.idtxtMRegulatoryExposure, uRegulatoryExposure)

    # Stock market sensitiviy (0.80-1.50)
    if uStockMrktSensitivity != "":
        sl.input_text(AR.idtxtMStockMrktSensitivity, uStockMrktSensitivity)

    # Takeover potential (0.80-1.50)
    if uTakeoverPotential != "":
        sl.input_text(AR.idtxtMTakeoverPotential, uTakeoverPotential)

    # Transaction Event (e.g., bankruptcy, credit downgrade) (1.00-2.00)
    if uTransactionEvent != "":
        sl.input_text(AR.idtxtMTransactionEvent, uTransactionEvent)

@keyword
def Save_Modifiers():
    sl = BuiltIn().get_library_instance('SeleniumLibrary')
    try:
        # Click Next to save UW Questions
        sl.click_button(AR.idbtnOneShieldNext)
        BuiltIn().sleep(2)

        # Verify if the record has been successfully created
        sl.element_should_contain(AR.idSubmissionBanner, 'Record has been saved')

    except Exception as e44:
        BuiltIn().fail('Save UW Questions Failed')


