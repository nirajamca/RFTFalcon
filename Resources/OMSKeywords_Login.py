from robot.libraries.BuiltIn import BuiltIn
from robot.api.deco import keyword
import AccessRepository as AR
import CommonFunctions as CF

def getLoginCreds():
    uEnvironment = BuiltIn().get_variable_value('${Env}')
    loc = 'Data\\FalconTestData.xlsx'
    SheetName = 'EnvLoginDetails'
    testcase = uEnvironment
    return CF.fncGetValues(loc, SheetName, testcase)

@keyword
def Open_Browser_To_Login_Page():
    uLoginCreds = getLoginCreds()
    sl = BuiltIn().get_library_instance('SeleniumLibrary')
    sl.create_webdriver('Chrome')
    sl.maximize_browser_window()
    sl.go_to(uLoginCreds['URL'])
    BuiltIn().sleep(3)
    try:
        sl.title_should_be('Login')
    except Exception as e1:
        sl.title_should_be('Oneshield | Log in')

@keyword
def Enter_Login_Credentials():
    uLoginCreds = getLoginCreds()
    sl = BuiltIn().get_library_instance('SeleniumLibrary')
    try:
        sl.input_text(AR.txtOMS_ClientID,uLoginCreds['Client'])
        sl.input_text(AR.txtOMS_LoginID, uLoginCreds['LoginID'])
        sl.input_password(AR.txtOMS_Password, uLoginCreds['Password'])
    except Exception as e1:
        sl.input_text(AR.txtOMS_ClientID1, uLoginCreds['Client'])
        sl.input_text(AR.txtOMS_LoginID1, uLoginCreds['LoginID'])
        sl.input_password(AR.txtOMS_Password1, uLoginCreds['Password'])

    BuiltIn().sleep(3)

@keyword
def Submit_Credentials():
    sl = BuiltIn().get_library_instance('SeleniumLibrary')
    try:
        sl.click_button(AR.btnOMS_Login)
    except Exception as e1:
        sl.click_button(AR.btnOMS_Login1)
    BuiltIn().sleep(3)

@keyword
def Welcome_Page_should_be_Open():
    sl = BuiltIn().get_library_instance('SeleniumLibrary')
    sl.title_should_be('OMS')
