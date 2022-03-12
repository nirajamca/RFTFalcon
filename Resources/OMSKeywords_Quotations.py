from robot.libraries.BuiltIn import BuiltIn
from robot.api.deco import keyword
import AccessRepository as AR
import CommonFunctions as CF


@keyword
def Go_To_Quotations():
    sl = BuiltIn().get_library_instance('SeleniumLibrary')
    sl.click_element(AR.QuotationsTab)
    # try:
    #     if sl.get_text(AR.xPathQuotationsTableHeader) != 'Organization':
    #         sl.click_element(AR.QuotationsTab)
    #
    # except Exception as e1:
    #     BuiltIn().log('Quotations Tab already open')

# @keyword
def Add_Organization():
    uCompany = BuiltIn().get_variable_value('${Company}')
    myData = CF.getData(uCompany, 'Quotations')
    uQCarrier = myData['Option1']
    uQDescription = myData['Option2']
    uQBrokerCommission = myData['Option3']
    uQInstallmentPlan = myData['Option4']

    sl = BuiltIn().get_library_instance('SeleniumLibrary')

    try:
        sl.element_should_contain(AR.idQuotationsTableHeader, 'Excess Follow form')
        BuiltIn().log('Select the existing record, do not create new one')

    except Exception as e721:
        print('No Carrier found')
        BuiltIn().sleep(3)

        # Click Add button to add an Organization
        sl.click_button(AR.idbtnQuotationsAddOrg)
        BuiltIn().sleep(3)

        # Select Carrier
        CF.fncSelectDDElement(sl, AR.idddQCarrier, uQCarrier)

        # Add Description
        sl.input_text(AR.idtxtQCarrierDesc, uQDescription)

        # Broker Commission
        sl.input_text(AR.idtxtQBrokerCommission, uQBrokerCommission)

        # Select Installment plan type
        CF.fncSelectDDElement(sl, AR.idddQInstallmentPlan, uQInstallmentPlan)

        # Click Save to add the ord record
        sl.click_button(AR.idbtnQuotationsSave)
        BuiltIn().sleep(3)

        sl.element_should_contain(AR.idQuotationsTableHeader, 'Excess Follow form')
        BuiltIn().sleep(3)

@keyword
def Select_Quote():
    sl = BuiltIn().get_library_instance('SeleniumLibrary')
    sl.click_element(AR.idtbrow1QOrganization)

@keyword
def Enter_Range_State_Elements():
    uCompany = BuiltIn().get_variable_value('${Company}')
    myData = CF.getData(uCompany, 'Quotations')
    uRangeStateFlag = myData['Option5']
    uRSRFactor = myData['Option6']
    uRSRFactorApplied = myData['Option7']
    uNRSRFactor = myData['Option8']
    uNRSRFactorApplied = myData['Option9']

    sl = BuiltIn().get_library_instance('SeleniumLibrary')

    # uRangeState can be either Yes or No based on data from State List of
    # Range states vs Non Range states.xlsx file
    if uRangeStateFlag == 'Yes':
        # If its a range state, enter values in Range State Risk Factor and Range State Risk Factor Applied
        CF.fncSelectDDElement(sl, AR.idddlQRangeStateRiskFactor, uRSRFactor)
        sl.input_text(AR.idtxtQRangeStateRiskFactorApplied, uRSRFactorApplied)

    elif uRangeStateFlag == 'No':
        # If its not a range state, enter values in Non-Range State Risk Factor and Non-Range State Risk Factor Applied
        CF.fncSelectDDElement(sl, AR.idddlQNonRangeStateRiskFactor, uNRSRFactor)
        sl.input_text(AR.idtxtQNonRangeStateRiskFactorApplied, uNRSRFactorApplied)

@keyword
def Add_Schedule_Rating():
    uCompany = BuiltIn().get_variable_value('${Company}')
    myData = CF.getData(uCompany, 'Quotations')
    uComplexityofRisk = myData['Option10']
    uRevenueSource = myData['Option11']
    uCvrgEnhancementsRestrictions = myData['Option12']
    uPrimaryCoverageTerms = myData['Option13']

    uAppScheduleRating = myData['Option14']
    uMktPricingAdjFactor = myData['Option15']

    uExtendedReportingPeriod = myData['Option16']
    uExtendedReportingPeriodFactor = myData['Option17']

    uRunoffPolicies = myData['Option18']
    uRunoffPoliciesFactor = myData['Option19']

    uIDLCoverage = myData['Option20']
    uSideACoverage = myData['Option21']

    uFurtherAdjPercentageOfULInsurance = myData['Option22']
    uFurtherAdjAvgScoreOfRiskFactors = myData['Option23']

    uRejectTerrorismCoverage = myData['Option24']
    uTerrorismDiscount = myData['Option25']

    sl = BuiltIn().get_library_instance('SeleniumLibrary')

    # Category		                            CW (x CA & NY)	CA	        NY	        GA
    # i.    Complexity of Risk		            +/- 25%	        +/- 15%     +/- 10%	    +/- 25%
    # ii.   Revenue Source		                +/- 25%	        +/- 15%	    +/- 10%	    +/- 25%
    # iii.  Coverage Enhancements/Restrictions	+/- 25%	        +/- 15%	    +/- 10%	    N/A
    # iv.   Primary Coverage Terms		        +/- 25%	        +/- 15%	    +/- 10%	    +/- 25%

    # Complexity of Risk
    sl.input_text(AR.idtxtQComplexityofRisk, uComplexityofRisk)

    # Revenue Source
    sl.input_text(AR.idtxtQRevenueSource, uRevenueSource)

    # Coverage Enhancements/Restrictions
    sl.input_text(AR.idtxtQCvrgEnhancementsRestrictions, uCvrgEnhancementsRestrictions)

    # Primary Coverage Terms
    sl.input_text(AR.idtxtQPrimaryCoverageTerms, uPrimaryCoverageTerms)

    # Applied Schedule Rating
    sl.input_text(AR.idtxtQAppScheduleRating, uAppScheduleRating)

    # Market Pricing Adjustment Factor (60% - 85%)
    sl.input_text(AR.idtxtQMktPricingAdjFactor, uMktPricingAdjFactor)

    # Extended Reporting Period and factor applied (Non-admitted)
    CF.fncSelectDDElement(sl, AR.idddQExtendedReportingPeriod, uExtendedReportingPeriod)
    sl.input_text(AR.idtxtQExtendedReportingPeriodFactor, uExtendedReportingPeriodFactor)

    # Run off Policies and factor applied (Non-admitted)
    CF.fncSelectDDElement(sl, AR.idddQRunoffPolicies, uRunoffPolicies)
    sl.input_text(AR.idtxtQRunoffPoliciesFactor, uRunoffPoliciesFactor)

    # Independent Directors Liability
    if uIDLCoverage is not None:
        sl.input_text(AR.idtxtQIDLCoverage, uIDLCoverage)

    # Side A Coverage
    if uSideACoverage is not None:
        sl.input_text(AR.idtxtQSideACoverage, uSideACoverage)

    # Further adjustment to percentage of the Underlying Insurance
    if uFurtherAdjPercentageOfULInsurance is not None:
        sl.input_text(AR.idtxtQFurtherAdjPercentageOfULInsurance, uFurtherAdjPercentageOfULInsurance)

    # Further percentage adjustment to Average Score of Risk
    if uFurtherAdjAvgScoreOfRiskFactors is not None:
        sl.input_text(AR.idtxtQFurtherAdjAvgScoreOfRiskFactors, uFurtherAdjAvgScoreOfRiskFactors)

    # Reject Terrorism Coverage
    CF.fncSelectDDElement(sl, AR.idddQRejectTerrorismCoverage, uRejectTerrorismCoverage)

    if uRejectTerrorismCoverage == 'Yes':
        if uTerrorismDiscount is None:
            uTerrorismDiscount = '1'
        sl.input_text(AR.idtxtQTerrorismDiscount, uTerrorismDiscount)

@keyword
def Enter_Coverage():
    uCompany = BuiltIn().get_variable_value('${Company}')
    myData = CF.getData(uCompany, 'Quotations')
    uLOLLimit = myData['Option26']
    uLOLDedRet = myData['Option27']

    sl = BuiltIn().get_library_instance('SeleniumLibrary')

    # Coverage table
    # Limit of Liability
    sl.click_element(AR.idchkQCoverageLimitofLiabilty)

    # Limit
    sl.input_text(AR.idtxtQCoverageLOLLimit, uLOLLimit)

    # Deductible/Retention
    sl.input_text(AR.idtxtQCoverageLOLDedRet, uLOLDedRet)
    BuiltIn().sleep(2)

@keyword
def Save_and_Price():
    sl = BuiltIn().get_library_instance('SeleniumLibrary')

    # Click Next to save UW Questions
    sl.click_button(AR.idbtnQSaveandPrice)
    BuiltIn().sleep(2)

    # Verify if the quotation has been successfully saved
    sl.element_should_contain(AR.idSubmissionBanner, 'Quotation has been saved !!')



