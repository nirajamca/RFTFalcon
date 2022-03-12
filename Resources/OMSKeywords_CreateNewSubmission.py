from robot.libraries.BuiltIn import BuiltIn
from robot.api.deco import keyword
from datetime import date, time
import AccessRepository as AR
import CommonFunctions as CF


def getData():
    loc = 'Data\\FalconTestData.xlsx'
    SheetName = BuiltIn().get_variable_value('${Company}')
    testcase = 'NewSubmission'
    return CF.fncGetValues(loc, SheetName, testcase)

@keyword
def Go_to_New_Submissions_Tab():
    sl = BuiltIn().get_library_instance('SeleniumLibrary')
    try:
        # Go to New Submission menu item
        sl.click_element(AR.xPathMenuNewSubmission)

    except Exception as e1:
        BuiltIn().log('Not able to find the button')

@keyword
def Enter_Details_in_Overview():
    myData = getData()
    uInsuredOrg = myData['Option1']
    uUnderwriter = myData['Option2']
    uProgram = myData['Option3']
    uBroker = myData['Option4']
    uBrokerContact = myData['Option5']
    uEffDate = date.today().strftime("%m/%d/%Y")

    sl = BuiltIn().get_library_instance('SeleniumLibrary')

    # Verify if the Overview tab is opened
    sl.element_should_contain(AR.xPathtabNSOverview, 'Overview')

    # Select Insured Org from the dropdwon
    sl.input_text(AR.xPathtxtNSOInsured, str(uInsuredOrg))
    BuiltIn().sleep(1)
    sl.press_keys(AR.xPathtxtNSOInsured, 'ARROW_DOWN')
    BuiltIn().sleep(1)
    sl.press_keys(AR.xPathtxtNSOInsured, 'RETURN')

    if uBroker is not None:
        # Select Broker company from dropdown
        CF.fncSelectDDElement(sl, AR.idddNSBroker, uBroker)

    if uBrokerContact is not None:
        # Select Broker company from dropdown
        CF.fncSelectDDElement(sl, AR.idddNSBrokerContact, uBrokerContact)

    # Enter Program Code
    CF.fncSelectDDElement(sl, AR.idddNSOProgram, uProgram)

    # Enter Proposed Effective Date
    sl.input_text(AR.idtxtNSOEffDate, uEffDate)

    # Underwriter
    CF.fncSelectDDElement(sl, AR.idddNSUnderwriter, uUnderwriter)

@keyword
def Save_New_Submission():
    sl = BuiltIn().get_library_instance('SeleniumLibrary')
    sl.click_button(AR.idbtnOneShieldNext)
