*** Settings ***
Documentation       Insert the sales data for the week and export it as a PDF

Library             RPA.Browser.Selenium    auto_close=${False}
Library             RPA.HTTP
Library             RPA.Excel.Files
Library             RPA.Tables
Library             RPA.PDF
Library             RPA.Robocorp.Vault


*** Variables ***
${excel_example}    https://robotsparebinindustries.com/SalesData.xlsx


*** Tasks ***
Insert the sales data for the week and export it as a PDF
    Open the intranet website
    Log in
    Download the excel file
    Fill the form using the data from the excel file
    Collect the result
    Export the table as PDF
    [Teardown]    Log out and close the browser


*** Keywords ***
Log out and close the browser
    Click Button    //*[@id="logout"]
    Close Browser

Export the table as PDF
    Wait Until Element Is Visible    id:sales-results
    ${sales_result_html}=    Get Element Attribute    id:sales-results    outerHTML
    Html To Pdf    ${sales_result_html}    ${OUTPUT_DIR}${/}Sales_Result.PDF

Collect the result
    Screenshot    css:#root > div > div > div > div:nth-child(2)    ${OUTPUT_DIR}${/}sales_summary.png

Fill the form using the data from the excel file
    Open Workbook    SalesData.xlsx
    @{sales_reps}=    Read Worksheet As Table    header=${True}
    Close Workbook
    FOR    ${sales_rep}    IN    @{sales_reps}
        Log    ${sales_rep}
        Fill and submit the form for one person    ${sales_rep}
    END

Download the excel file
    Download    ${excel_example}    overwrite=${True}

Fill and submit the form for one person
    [Arguments]    ${sales_rep}
    Input Text    firstname    ${sales_rep}[First Name]
    Input Text    lastname    ${sales_rep}[Last Name]
    Select From List By Value    salestarget    ${sales_rep}[Sales Target]
    Input Text    salesresult    ${sales_rep}[Sales]
    Click Button    Submit

Log in
    ${account}=    Get Secret    robotsparebin
    Input Text    username    ${account}[username]
    Input Password    password    ${account}[password]
    Submit Form
    Wait Until Page Contains Element    id:sales-form

Open the intranet website
    Open Available Browser    https://robotsparebinindustries.com/
