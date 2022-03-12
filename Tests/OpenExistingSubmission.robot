*** Settings ***
Documentation       Simple example using SeleniumLibrary
Library             SeleniumLibrary
#Variables           ../Data/OMSData.py
Library            ../Resources/OMSKeywords_Login.py
Library            ../Resources/OMSKeywords_SearchSubmission.py
Library            ../Resources/OMSKeywords_ChangeRangeState.py
Library            ../Resources/OMSKeywords_UWQuestions.py
Library            ../Resources/OMSKeywords_Modifiers.py
Library            ../Resources/OMSKeywords_UnderlyingPolicies.py
Library            ../Resources/OMSKeywords_Quotations.py
Library            ../Resources/OMSKeywords_VerifyPremiumInRater.py


*** Variables ***
${Env}       PreProd
${Company}   Bally

*** Keywords ***
    myNote   [Documentation]  All Keywords are available in their respective resource files

*** Test Cases ***
OMS Login
    Open Browser To Login Page
    Enter Login Credentials
    Submit Credentials
    Welcome Page should be Open

Search for Submission
    Go to Submissions Tab
    Enter Submission Number
    Enter Prefix if Necessary
    Search for Submission

Change Range State
    Change Range State if required

UW Questions
    Go to UW Questions
    Add Policy Information
    Add Underlying Insurance Risk Factors
    Save UW Questions

Modifiers
    Go To Modifiers
    Add Coverage Modifiers
    Add Other RiskFactors
    Save Modifiers

Underlying Polocies
    Go To UL Policies
    Delete Existing Policies if Required
    Add Layers
    Save Layers

Quotations
    Go To Quotations
    Add Organization
    Select Quote
    Enter Range State Elements
    Add Schedule Rating
    Enter Coverage
    Save and Price

#Verify Premium in Rater
#    Verify Premium with Rater
#    [Teardown]      Close Browser


