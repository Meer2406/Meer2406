*** Settings ***
Library     RPA.Browser.Selenium    auto_close=${FALSE}
Library     RPA.HTTP
Library     RPA.Excel.Files
Library     RPA.PDF


*** Variables ***
@{sales_reps}
${sales_rep}
${sales_results_html}


*** Tasks ***
Insert The Sales Data For The Week And Export It As A Pdf
    [Documentation]    Inserts the sales data for the week into the intranet and exports it as a PDF file.
    Open The Intranet Website
    Log In
    Download The Excel File
    Fill The Form Using The Data From The Excel File
    Collect The Results
    Export The Table As A Pdf
    Log Out And Close The Browser


*** Keywords ***
Open The Intranet Website
    Open Available Browser    https://robotsparebinindustries.com/    maximized=${TRUE}

Log In
    Input Text    alias:Username    maria
    Input Text    alias:Password    thoushallnotpass
    Click Element    alias:Btn
    Wait Until Page Contains Element    alias:Salesform

Download The Excel File
    Download    https://robotsparebinindustries.com/SalesData.xlsx    overwrite=${TRUE}

Fill And Submit The Form For One Person
    [Arguments]    ${sales_rep}
    Input Text    alias:Firstname    ${sales_rep}[First Name]
    Input Text    alias:Lastname    ${sales_rep}[Last Name]
    Select From List By Value    alias:Salestarget    ${sales_rep}[Sales Target]
    Input Text    alias:Salesresult    ${sales_rep}[Sales]
    Click Element    alias:Btnprimary

Fill The Form Using The Data From The Excel File
    Open Workbook    %{ROBOT_ROOT}${/}SalesData.xlsx
    ${sales_reps}=    Read Worksheet As Table    header=${TRUE}
    FOR    ${sales_rep}    IN    @{sales_reps}
        Fill And Submit The Form For One Person    ${sales_rep}
    END
    Close Workbook

Collect The Results
    Screenshot    alias:SalesSummary    ${OUTPUT_DIR}${/}sales_summary.png

Export The Table As A Pdf
    Wait Until Element Is Visible    alias:SalesResultTable
    ${sales_results_html}=    Get Element Attribute    alias:SalesResultTable    outerHTML
    Html To Pdf    ${sales_results_html}    ${OUTPUT_DIR}${/}sales_results.pdf

Log Out And Close The Browser
    Click Element    alias:Logout
    Close Browser
