from robot.libraries.BuiltIn import BuiltIn
from robot.api.deco import keyword
import AccessRepository as AR
import CommonFunctions as CF

@keyword
def Change_Range_State_if_required():
    uCompany = BuiltIn().get_variable_value('${Company}')
    myData = CF.getData(uCompany, 'ChangeRangeState')
    uChangeRangeState = myData['Option1']
    uRangeState = myData['Option2']

    sl = BuiltIn().get_library_instance('SeleniumLibrary')
    if uChangeRangeState == 'Yes':
        sl.click_element(AR.xPathSubMenuOverviewSubmission)
        BuiltIn().sleep(2)
        CF.fncSelectDDElement(sl, AR.idddRangeState, uRangeState)
        BuiltIn().sleep(2)
        sl.click_button(AR.idbtnOneShieldNext)
        BuiltIn().sleep(2)

        sl.element_should_contain(AR.idSubmissionBanner, 'Record has been saved')


