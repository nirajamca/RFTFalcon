from robot.libraries.BuiltIn import BuiltIn
from robot.api.deco import keyword
import AccessRepository as AR
import CommonFunctions as CF


def getData():
    loc = 'Data\\FalconTestData.xlsx'
    SheetName = BuiltIn().get_variable_value('${Company}')
    testcase = 'SearchSubmission'
    return CF.fncGetValues(loc, SheetName, testcase)

@keyword
def Go_to_Submissions_Tab():
    sl = BuiltIn().get_library_instance('SeleniumLibrary')
    try:
        sl.click_element(AR.xPathMenuSubmissions)

    except Exception as e1:
        BuiltIn().log('Not able to find the button')

@keyword
def Enter_Submission_Number():
    myData = getData()
    uSubmissionNumber = myData['Option1']
    sl = BuiltIn().get_library_instance('SeleniumLibrary')
    sl.input_text(AR.idtxtSubmissionSearchSubNumber, uSubmissionNumber)

@keyword
def Enter_Prefix_if_Necessary():
    myData = getData()
    uSubmissionPrefix = myData['Option2']
    if uSubmissionPrefix is None:
        y = 1
    else:
        sl = BuiltIn().get_library_instance('SeleniumLibrary')
        sl.input_text(AR.idtxtSubmissionSearchSubSuffix, uSubmissionPrefix)

@keyword
def Search_for_Submission():
    sl = BuiltIn().get_library_instance('SeleniumLibrary')
    sl.click_button(AR.xPathbtnSubmissionsRunSearch)
