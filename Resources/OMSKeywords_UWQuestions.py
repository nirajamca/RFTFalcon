from robot.libraries.BuiltIn import BuiltIn
from robot.api.deco import keyword
import AccessRepository as AR
import CommonFunctions as CF

def getData():
    loc = 'Data\\FalconTestData.xlsx'
    SheetName = BuiltIn().get_variable_value('${Company}')
    testcase = 'UWQuestions'
    return CF.fncGetValues(loc, SheetName, testcase)

@keyword
def Go_to_UW_Questions():
    sl = BuiltIn().get_library_instance('SeleniumLibrary')
    try:
        sl.click_element(AR.UWQuestionsTab)

    except Exception as e1:
        BuiltIn().log('Not able to find the button')

@keyword
def Add_Policy_Information():
    myData = getData()
    uMarketCap = myData['Option1']
    uTotalAssets = myData['Option2']
    uAnnualRevenues = myData['Option3']
    uTicker = myData['Option4']
    uOrgCategory = myData['Option5']
    uIncludeFeesandTaxes = myData['Option6']

    sl = BuiltIn().get_library_instance('SeleniumLibrary')

    # Market Cap
    sl.input_text(AR.idtxtUWQMarketCap, uMarketCap)

    # Total Assets
    sl.input_text(AR.idtxtUWQTotalAssets, uTotalAssets)

    # Annual Revenues
    sl.input_text(AR.idtxtUWQAnnualRevenues, uAnnualRevenues)

    # Organization Category
    CF.fncSelectDDElement(sl, AR.idddUWQOrgCategory, uOrgCategory)

    # Ticker Symbol
    sl.input_text(AR.idtxtUWQTicker, uTicker)

    # Select Yes or No to Include fees and taxes to total Premium
    CF.fncSelectDDElement(sl, AR.idddUWIncludeFeesandTaxes, uIncludeFeesandTaxes)

@keyword
def Add_Underlying_Insurance_Risk_Factors():
    myData = getData()
    uFinCondition = myData['Option7']
    uNoOperations = myData['Option8']
    uAcqHistory = myData['Option9']
    uMgtQuality = myData['Option10']
    uLtgHistory = myData['Option11']
    uTimeInBusiness = myData['Option12']
    uClsBusiness = myData['Option13']

    sl = BuiltIn().get_library_instance('SeleniumLibrary')
    # Financial Condition
    CF.fncSelectDDElement(sl, AR.idddUWQFinCondition, uFinCondition)

    # Nature of Operations
    CF.fncSelectDDElement(sl, AR.idddUWQNoOperations, uNoOperations)

    # Acquisition History
    CF.fncSelectDDElement(sl, AR.idddUWQAcqHistory, uAcqHistory)

    # Management Quality
    CF.fncSelectDDElement(sl, AR.idddUWQMgtQuality, uMgtQuality)

    # Litigation History
    CF.fncSelectDDElement(sl, AR.idddUWQLtgHistory, uLtgHistory)

    # Time in Business
    CF.fncSelectDDElement(sl, AR.idddUWQTimeInBusiness, uTimeInBusiness)

    # Class of Business
    CF.fncSelectDDElement(sl, AR.idddUWQClsBusiness, uClsBusiness)

@keyword
def Save_UW_Questions():
    sl = BuiltIn().get_library_instance('SeleniumLibrary')
    # Click Next to save UW Questions
    sl.click_button(AR.idbtnOneShieldNext)
    BuiltIn().sleep(2)

    # Verify if the record has been successfully created
    sl.element_should_contain(AR.idSubmissionBanner, 'Record has been saved')



