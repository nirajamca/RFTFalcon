from robot.libraries.BuiltIn import BuiltIn
from robot.api.deco import keyword
import AccessRepository as AR

@keyword
def Go_To_Actions():
    sl = BuiltIn().get_library_instance('SeleniumLibrary')

    # Click botton next to Actions
    sl.click_button(AR.xPathbtnSubmissionActions)
    BuiltIn().sleep(2)

    # Verify if Copy Submission option is available
    sl.element_should_be_visible(AR.idOptSubmissionCopy)

@keyword
def Copy_Submission():
    sl = BuiltIn().get_library_instance('SeleniumLibrary')

    # Select Copy Submission option from the list
    sl.click_element(AR.idOptSubmissionCopy)
    BuiltIn().sleep(2)

@keyword
def Verify_that_Submission_is_copied():
    sl = BuiltIn().get_library_instance('SeleniumLibrary')

    # Verify is a message is displayed that submission is copied successfully
    sl.element_should_contain(AR.idlblSubmissionBanner, 'submission has been Copied')
    BuiltIn().log(sl.get_text(AR.idlblSummaryText))



