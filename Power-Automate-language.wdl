@@copilotGeneratedAction: 'True'
**REGION Get current date and time, convert to string, launch Excel, and write date and month values
@@copilotGeneratedAction: 'True'
DateTime.GetCurrentDateTime.Local DateTimeFormat: DateTime.DateTimeFormat.DateAndTime CurrentDateTime=> CurrentDateTime
@@copilotGeneratedAction: 'True'
Text.ConvertDateTimeToText.FromCustomDateTime DateTime: CurrentDateTime CustomFormat: $'''dd''' Result=> Date
@@copilotGeneratedAction: 'True'
Text.ConvertDateTimeToText.FromCustomDateTime DateTime: CurrentDateTime CustomFormat: $'''MMM''' Result=> Month
@@copilotGeneratedAction: 'True'
Text.ConvertDateTimeToText.FromCustomDateTime DateTime: CurrentDateTime CustomFormat: $'''yyyy''' Result=> Year
@@copilotGeneratedAction: 'True'
Excel.LaunchExcel.LaunchAndOpen Path: $'''C:\\Users\\aditya.sagave\\Desktop\\Booking form-offline\\ORDER BOOKING FORM INTLOPS 20.06.2025.xlsx''' Visible: True ReadOnly: False LoadAddInsAndMacros: False Instance=> ExcelInstance
@@copilotGeneratedAction: 'True'
Excel.WriteToExcel.WriteCell Instance: ExcelInstance Column: $'''M''' Row: $'''10''' Value: Month
@@copilotGeneratedAction: 'True'
Excel.WriteToExcel.WriteCell Instance: ExcelInstance Column: $'''N''' Row: $'''10''' Value: Date
Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: Year Column: $'''O''' Row: 10
@@copilotGeneratedAction: 'True'
DateTime.Add DateTime: CurrentDateTime TimeToAdd: 2 TimeUnit: DateTime.TimeUnit.Days ResultedDate=> FutureDate
Text.ConvertDateTimeToText.FromCustomDateTime DateTime: FutureDate CustomFormat: $'''dd''' Result=> Date
Text.ConvertDateTimeToText.FromCustomDateTime DateTime: FutureDate CustomFormat: $'''MMM''' Result=> Month
Text.ConvertDateTimeToText.FromCustomDateTime DateTime: FutureDate CustomFormat: $'''yyyy''' Result=> Year
Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: Month Column: $'''M''' Row: 12
Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: Date Column: $'''N''' Row: 12
Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: Year Column: $'''O''' Row: 12
@@copilotGeneratedAction: 'True'
**REGION Ask the user to enter the weight of the shipment, write the user input to cell I27 in the Excel file, and close and save the Excel instance
@@copilotGeneratedAction: 'True'
Display.InputDialog Title: $'''Shipment Weight''' Message: $'''What will be the weight of the shipment?''' InputType: Display.InputType.SingleLine IsTopMost: False UserInput=> ShipmentWeight ButtonPressed=> ButtonPressed
@@copilotGeneratedAction: 'True'
Excel.WriteToExcel.WriteCell Instance: ExcelInstance Column: $'''I''' Row: 27 Value: ShipmentWeight
@@copilotGeneratedAction: 'True'
**ENDREGION
@@copilotGeneratedAction: 'True'
**REGION Ask the user to enter the dimensions of the shipment, split the dimensions into length, breath, and height, store the dimensions in separate variables, and write the dimensions to the specified cells in the Excel file
@@copilotGeneratedAction: 'True'
Display.InputDialog Title: $'''Shipment Dimensions''' Message: $'''Please enter the dimensions of the shipment in the format: length,breath,height (e.g., 0,0,0)''' InputType: Display.InputType.SingleLine IsTopMost: False UserInput=> Dimensions ButtonPressed=> ButtonPressed
@@copilotGeneratedAction: 'True'
Text.SplitText.SplitWithDelimiter Text: Dimensions CustomDelimiter: $''',''' IsRegEx: False Result=> DimensionsList
@@copilotGeneratedAction: 'True'
SET Length TO DimensionsList[0]
@@copilotGeneratedAction: 'True'
SET Breath TO DimensionsList[1]
@@copilotGeneratedAction: 'True'
SET Height TO DimensionsList[2]
@@copilotGeneratedAction: 'True'
Excel.WriteToExcel.WriteCell Instance: ExcelInstance Column: $'''M''' Row: 27 Value: Length
@@copilotGeneratedAction: 'True'
Excel.WriteToExcel.WriteCell Instance: ExcelInstance Column: $'''N''' Row: 27 Value: Breath
@@copilotGeneratedAction: 'True'
Excel.WriteToExcel.WriteCell Instance: ExcelInstance Column: $'''O''' Row: 27 Value: Height
@@copilotGeneratedAction: 'True'
**ENDREGION
@@copilotGeneratedAction: 'True'
**REGION Ask the user if there is another shipment, if yes, ask for the weight and update the Excel file
@@copilotGeneratedAction: 'True'
Display.ShowMessageDialog.ShowMessage Title: $'''Shipment''' Message: $'''Is there another shipment?''' Buttons: Display.Buttons.YesNo DefaultButton: Display.DefaultButton.Button1 IsTopMost: False ButtonPressed=> UserResponse
@@copilotGeneratedAction: 'True'
IF UserResponse = $'''Yes''' THEN
    @@copilotGeneratedAction: 'True'
Display.InputDialog Title: $'''Shipment Weight''' Message: $'''What is the weight of the shipment?''' InputType: Display.InputType.SingleLine IsTopMost: False UserInput=> ShipmentWeight ButtonPressed=> ButtonPressed
    @@copilotGeneratedAction: 'True'
Excel.WriteToExcel.WriteCell Instance: ExcelInstance Column: $'''H''' Row: 28 Value: $'''1'''
    @@copilotGeneratedAction: 'True'
Excel.WriteToExcel.WriteCell Instance: ExcelInstance Column: $'''I''' Row: 28 Value: ShipmentWeight
    @@copilotGeneratedAction: 'True'
**REGION Ask the user to enter the dimensions of the shipment, split the dimensions into length, breath, and height, store the dimensions in separate variables, and write the dimensions to the specified cells in the Excel file
    @@copilotGeneratedAction: 'True'
Display.InputDialog Title: $'''Shipment Dimensions''' Message: $'''Please enter the dimensions of the shipment in the format: length,breath,height (e.g., 0,0,0)''' InputType: Display.InputType.SingleLine IsTopMost: False UserInput=> Dimensions ButtonPressed=> ButtonPressed
    @@copilotGeneratedAction: 'True'
Text.SplitText.SplitWithDelimiter Text: Dimensions CustomDelimiter: $''',''' IsRegEx: False Result=> DimensionsList
    @@copilotGeneratedAction: 'True'
SET Length TO DimensionsList[0]
    @@copilotGeneratedAction: 'True'
SET Breath TO DimensionsList[1]
    @@copilotGeneratedAction: 'True'
SET Height TO DimensionsList[2]
    Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: Length Column: $'''M''' Row: 28
    Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: Breath Column: $'''N''' Row: 28
    Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: Height Column: $'''O''' Row: 28
    @@copilotGeneratedAction: 'True'
**ENDREGION
@@copilotGeneratedAction: 'True'
END
@@copilotGeneratedAction: 'True'
IF UserResponse = $'''No''' THEN
@@copilotGeneratedAction: 'True'
END
@@copilotGeneratedAction: 'True'
**ENDREGION
Text.ConvertDateTimeToText.FromCustomDateTime DateTime: CurrentDateTime CustomFormat: $'''dd.MM.yyyy''' Result=> bookingFormDate
Excel.CloseExcel.CloseAndSaveAs Instance: ExcelInstance DocumentFormat: Excel.ExcelFormat.FromExtension DocumentPath: $'''C:\\Users\\aditya.sagave\\Desktop\\Booking form-online\\ORDER BOOKING FORM INTLOPS %bookingFormDate%.xlsx'''
@@copilotGeneratedAction: 'True'
**ENDREGION
@@copilotGeneratedAction: 'True'
**REGION Launch and open the Word document, find and replace all occurrences of the specified text, and close and save the Word instance
@@copilotGeneratedAction: 'True'
Word.LaunchWord.LaunchAndOpen Path: $'''C:\\Users\\aditya.sagave\\Desktop\\CSD-offline\\Consigment Security Declaration -20.06.2025.docx''' Visible: True ReadOnly: False Instance=> WordInstance
Display.SelectFromListDialog.SelectFromList Title: $'''Booking Person''' Message: $'''Please select the person who is booking the shipment''' List: $'''Aditya SAGAVE
Ravi ANTON
Lisa JOHNS
Heather BURTON''' IsTopMost: False AllowEmpty: False SelectedItem=> SelectedItem ButtonPressed=> ButtonPressed
Word.FindAndReplaceWord.FindAndReplaceAllWithoutWildcards Instance: WordInstance TextToFind: $'''Aditya SAGAVE''' TextToReplaceWith: SelectedItem MatchCase: False MatchEntireWord: False
@@copilotGeneratedAction: 'True'
**REGION Display a dialog box to select an item from a list, check if the selected item is "Aditya Sagave", and find and replace "VEA0405014" in the Word document
IF SelectedItem = $'''Aditya SAGAVE''' THEN
    Word.FindAndReplaceWord.FindAndReplaceAllWithoutWildcards Instance: WordInstance TextToFind: $'''VEA0405014''' TextToReplaceWith: $'''VEA0405014''' MatchCase: False MatchEntireWord: False
    Word.InsertImageToWord.InsertImageEndOfTextFromFile Instance: WordInstance TextToFind: $'''Signed:''' Path: $'''C:\\Users\\aditya.sagave\\Desktop\\CSD-offline\\aditya.sagave.png'''
END
IF SelectedItem = $'''Lisa JOHNS''' THEN
    Word.FindAndReplaceWord.FindAndReplaceAllWithoutWildcards Instance: WordInstance TextToFind: $'''VEA0405014''' TextToReplaceWith: $'''VEA0405015''' MatchCase: False MatchEntireWord: False
    Word.InsertImageToWord.InsertImageEndOfTextFromFile Instance: WordInstance TextToFind: $'''Signed:''' Path: $'''C:\\Users\\aditya.sagave\\Desktop\\CSD-offline\\lisa.johns.png'''
END
IF SelectedItem = $'''Ravi ANTON''' THEN
    Word.FindAndReplaceWord.FindAndReplaceAllWithoutWildcards Instance: WordInstance TextToFind: $'''VEA0405014''' TextToReplaceWith: $'''VEA0405016''' MatchCase: False MatchEntireWord: False
    Word.InsertImageToWord.InsertImageEndOfTextFromFile Instance: WordInstance TextToFind: $'''Signed:''' Path: $'''C:\\Users\\aditya.sagave\\Desktop\\CSD-offline\\ravi.anton.png'''
END
IF SelectedItem = $'''Heather BURTON''' THEN
    Word.FindAndReplaceWord.FindAndReplaceAllWithoutWildcards Instance: WordInstance TextToFind: $'''VEA0405014''' TextToReplaceWith: $'''VEA0405017''' MatchCase: False MatchEntireWord: False
    Word.InsertImageToWord.InsertImageEndOfTextFromFile Instance: WordInstance TextToFind: $'''Signed:''' Path: $'''C:\\Users\\aditya.sagave\\Desktop\\CSD-offline\\heather.burton.png'''
END
@@copilotGeneratedAction: 'True'
DateTime.GetCurrentDateTime.Local DateTimeFormat: DateTime.DateTimeFormat.DateAndTime CurrentDateTime=> CurrentDateTime
@@copilotGeneratedAction: 'True'
Text.ConvertDateTimeToText.FromCustomDateTime DateTime: CurrentDateTime CustomFormat: $'''dd/MM/yyyy''' Result=> FormattedDateTime
Word.FindAndReplaceWord.FindAndReplaceAllWildcards Instance: WordInstance TextToFind: $'''20/06/2025''' TextToReplaceWith: FormattedDateTime
@@copilotGeneratedAction: 'True'
Display.InputDialog Title: $'''Tamper Tape Initial Value''' Message: $'''Please enter the tamper tape initial value:''' InputType: Display.InputType.SingleLine IsTopMost: False UserInput=> TamperTapeInitialValue ButtonPressed=> ButtonPressed
Text.ToNumber Text: TamperTapeInitialValue Number=> TamperNumber
Word.FindAndReplaceWord.FindAndReplaceAllWildcards Instance: WordInstance TextToFind: $'''E19-0065152''' TextToReplaceWith: $'''E19-00%TamperNumber%'''
Variables.IncreaseVariable Value: TamperNumber IncrementValue: 1
Word.FindAndReplaceWord.FindAndReplaceAllWildcards Instance: WordInstance TextToFind: $'''E19-0065153''' TextToReplaceWith: $'''E19-00%TamperNumber%'''
Variables.IncreaseVariable Value: TamperNumber IncrementValue: 1
Word.FindAndReplaceWord.FindAndReplaceAllWildcards Instance: WordInstance TextToFind: $'''E19-0065154''' TextToReplaceWith: $'''E19-00%TamperNumber%'''
Variables.IncreaseVariable Value: TamperNumber IncrementValue: 1
Word.FindAndReplaceWord.FindAndReplaceAllWildcards Instance: WordInstance TextToFind: $'''E19-0065155''' TextToReplaceWith: $'''E19-00%TamperNumber%'''
@@copilotGeneratedAction: 'True'
**ENDREGION
@@copilotGeneratedAction: 'True'
**ENDREGION
@@copilotGeneratedAction: 'True'
**REGION Launch and open the Word document, save the Word document with a different name in a different folder
@@copilotGeneratedAction: 'True'
**REGION Get current date and time, convert to string
@@copilotGeneratedAction: 'True'
DateTime.GetCurrentDateTime.Local DateTimeFormat: DateTime.DateTimeFormat.DateAndTime CurrentDateTime=> CurrentDateTime
Text.ConvertDateTimeToText.FromCustomDateTime DateTime: CurrentDateTime CustomFormat: $'''dd.MM.yyyy''' Result=> CsdFileDate
@@copilotGeneratedAction: 'True'
**ENDREGION
Word.CloseWord.CloseAndSaveAs Instance: WordInstance DocumentFormat: Word.WordFormat.DOCX DocumentPath: $'''C:\\Users\\aditya.sagave\\Desktop\\CSD-online\\Consigment Security Declaration -%CsdFileDate%.docx'''
@@copilotGeneratedAction: 'True'
**ENDREGION
@@copilotGeneratedAction: 'True'
**REGION Launch Word, find and replace specified words, and save the document
@@copilotGeneratedAction: 'True'
Word.LaunchWord.LaunchAndOpen Path: $'''C:\\Users\\aditya.sagave\\Desktop\\CI-offline\\LDR Inc Customs Invoice 20.06.2025.docx''' ReadOnly: False Visible: True Instance=> WordInstance
DateTime.GetCurrentDateTime.Local DateTimeFormat: DateTime.DateTimeFormat.DateAndTime CurrentDateTime=> CurrentDateTime
Text.ConvertDateTimeToText.FromCustomDateTime DateTime: CurrentDateTime CustomFormat: $'''dd/MM/yyyy''' Result=> CiInFiledate
Word.FindAndReplaceWord.FindAndReplaceAllWithoutWildcards Instance: WordInstance TextToFind: $'''20/06/2025''' TextToReplaceWith: CiInFiledate MatchCase: False MatchEntireWord: True
Word.FindAndReplaceWord.FindAndReplaceAllWithoutWildcards Instance: WordInstance TextToFind: $'''Total Weight:''' TextToReplaceWith: $'''Total Weight: %ShipmentWeight% KG''' MatchCase: False MatchEntireWord: True
Text.ConvertDateTimeToText.FromCustomDateTime DateTime: CurrentDateTime CustomFormat: $'''dd.MM.yyyy''' Result=> CiFileDate
Word.CloseWord.CloseAndSaveAs Instance: WordInstance DocumentFormat: Word.WordFormat.FromExtension DocumentPath: $'''C:\\Users\\aditya.sagave\\Desktop\\CI-online\\LDR Inc Customs Invoice %CiFileDate%.docx'''
@@copilotGeneratedAction: 'True'
**ENDREGION
@@copilotGeneratedAction: 'True'
**REGION Launch Outlook, send an email with the attached Word document, and close Outlook
@@copilotGeneratedAction: 'True'
Outlook.Launch Instance=> OutlookInstance
DateTime.GetCurrentDateTime.Local DateTimeFormat: DateTime.DateTimeFormat.DateOnly CurrentDateTime=> CurrentDateTime2
Text.ConvertDateTimeToText.FromCustomDateTime DateTime: CurrentDateTime2 CustomFormat: $'''dd-MMM-yyyy''' Result=> FormattedDateTime2
Text.ConvertDateTimeToText.FromCustomDateTime DateTime: CurrentDateTime2 CustomFormat: $'''dddd''' Result=> day
Outlook.SendEmailThroughOutlook.SendEmail Instance: OutlookInstance Account: $'''aditya.sagave@landauer.com''' SendTo: $'''aditya.sagave@landauer.com''' CC: $'''orders@ldrp.com.au''' Subject: $'''NEW BOOKING %FormattedDateTime2% test14''' Body: $'''Dear Team,
 
Please find attached the booking form for a shipment we want to send out today (%FormattedDateTime2%).
 
Please confirm if a pickup could be arranged for this shipment today after 1 PM (%day%)
 
Kind regards,
Aditya Sagave
Customer Services Representative – Landauer Australasia | Fluke Health Solutions
P: +612 8651 4000
M: aditya.sagavelandauer.com''' IsBodyHtml: False IsDraft: False Attachments: $'''\"C:\\Users\\aditya.sagave\\Desktop\\Booking form-online\\ORDER BOOKING FORM INTLOPS %bookingFormDate%.xlsx\"'''
@@copilotGeneratedAction: 'True'
LABEL FindEmail
Outlook.RetrieveEmailMessages.RetrieveEmails Instance: OutlookInstance Account: $'''aditya.sagave@landauer.com''' MailFolder: $'''Inbox''' EmailsToRetrieve: Outlook.RetrieveMessagesMode.All MarkAsRead: True ReadBodyAsHtml: False SubjectContains: $'''NEW BOOKING %FormattedDateTime2% test14''' Messages=> RetrievedEmails
@@copilotGeneratedAction: 'True'
IF RetrievedEmails.Count = 0 THEN
    @@copilotGeneratedAction: 'True'
GOTO FindEmail
@@copilotGeneratedAction: 'True'
END
LOOP FOREACH CurrentItem IN RetrievedEmails
    Outlook.RespondToMailMessage.ReplyAllToEmail Instance: OutlookInstance Account: $'''aditya.sagave@landauer.com''' MailMessage: CurrentItem Body: $'''Dear Team,<br><br>

Please find attached template for <strong>Custom Invoice</strong> and <strong>Consignment Security Declaration</strong>.<br><br>

Please provide <strong>Job ID</strong> and <strong>Air Way Bill #</strong> at your earliest convenience.<br><br>

Kind regards,<br>
<strong>Aditya Sagave</strong><br>
Customer Services Representative – <strong>Landauer Australasia</strong> | <strong>Fluke Health Solutions</strong><br>
P: +612 8651 4000<br>
M: <a href=\"mailto:aditya.sagave@landauer.com\">aditya.sagave@landauer.com</a>''' Attachments: $'''\"C:\\Users\\aditya.sagave\\Desktop\\CSD-online\\Consigment Security Declaration -%CsdFileDate%.docx\" \"C:\\Users\\aditya.sagave\\Desktop\\CI-online\\LDR Inc Customs Invoice %CiFileDate%.docx\"'''
END
@@copilotGeneratedAction: 'True'
**ENDREGION
@@copilotGeneratedAction: 'True'
**REGION Launch Outlook, retrieve emails, loop through each email to parse for job ID, remove duplicates, and display unique matches
@@copilotGeneratedAction: 'True'
Outlook.Launch Instance=> OutlookInstance
LABEL 'find-jobID'
Outlook.RetrieveEmailMessages.RetrieveEmails Instance: OutlookInstance Account: $'''aditya.sagave@landauer.com''' MailFolder: $'''Inbox''' EmailsToRetrieve: Outlook.RetrieveMessagesMode.All MarkAsRead: False ReadBodyAsHtml: False SubjectContains: $'''NEW BOOKING %FormattedDateTime2% test14''' Messages=> RetrievedEmails
Text.ParseText.RegexParseForFirstOccurrence Text: RetrievedEmails[0].Body TextToFind: $'''(?:DHL\\s*SD\\s*#?\\s*)(\\d{7})''' StartingPosition: 0 IgnoreCase: True OccurrencePosition=> Position Match=> JobIdMatch
IF IsEmpty(JobIdMatch) THEN
    GOTO 'find-jobID'
END
Text.ParseText.RegexParseForFirstOccurrence Text: JobIdMatch TextToFind: $'''(\\d{7})''' StartingPosition: 0 IgnoreCase: True OccurrencePosition=> Position Match=> jobID
Display.ShowMessageDialog.ShowMessage Title: $'''emails''' Message: $'''The Job ID is: %jobID%''' Icon: Display.Icon.None Buttons: Display.Buttons.OK DefaultButton: Display.DefaultButton.Button1 IsTopMost: False ButtonPressed=> ButtonPressed
@@copilotGeneratedAction: 'True'
**ENDREGION
**REGION Launch Outlook, retrieve emails, loop through each email to parse for AWB number, remove duplicates, and display unique matches
LABEL 'find-awbID'
Outlook.RetrieveEmailMessages.RetrieveEmails Instance: OutlookInstance Account: $'''aditya.sagave@landauer.com''' MailFolder: $'''Inbox''' EmailsToRetrieve: Outlook.RetrieveMessagesMode.All MarkAsRead: False ReadBodyAsHtml: False SubjectContains: $'''NEW BOOKING %FormattedDateTime2% test14''' Messages=> RetrievedEmails
Text.ParseText.RegexParseForFirstOccurrence Text: RetrievedEmails[0].Body TextToFind: $'''(?:AWB\\s*016\\s*-?\\s*)(\\d+)''' StartingPosition: 0 IgnoreCase: True OccurrencePosition=> Position Match=> AwbMatch
IF IsEmpty(AwbMatch) THEN
    GOTO 'find-awbID'
END
Text.ParseText.RegexParseForFirstOccurrence Text: AwbMatch TextToFind: $'''(\\d{8})''' StartingPosition: 0 IgnoreCase: True OccurrencePosition=> Position Match=> awbID
Display.ShowMessageDialog.ShowMessage Title: $'''ID Information''' Message: $'''The AWB number is: %awbID%''' Icon: Display.Icon.Information Buttons: Display.Buttons.OK DefaultButton: Display.DefaultButton.Button1 IsTopMost: False ButtonPressed=> ButtonPressed
**ENDREGION
**REGION update CI with Job ID and awb ID
DateTime.GetCurrentDateTime.Local DateTimeFormat: DateTime.DateTimeFormat.DateAndTime CurrentDateTime=> CurrentDateTime
Text.ConvertDateTimeToText.FromCustomDateTime DateTime: CurrentDateTime CustomFormat: $'''dd.MM.yyyy''' Result=> CiFileDate
Word.LaunchWord.LaunchAndOpen Path: $'''C:\\Users\\aditya.sagave\\Desktop\\CI-online\\LDR Inc Customs Invoice %CiFileDate%.docx''' Visible: True ReadOnly: False Instance=> CiInstance
Word.FindAndReplaceWord.FindAndReplaceSingle Instance: CiInstance TextToFind: $'''Shipment ID: ''' TextToReplaceWith: $'''Shipment ID: %jobID%''' MatchCase: False MatchEntireWord: False
Word.FindAndReplaceWord.FindAndReplaceSingle Instance: CiInstance TextToFind: $'''Waybill Number: 016-''' TextToReplaceWith: $'''Waybill Number: 016- %awbID%''' MatchCase: False MatchEntireWord: False
Word.CloseWord.CloseAndSave Instance: CiInstance
**ENDREGION
**REGION update CSD with awb ID
DateTime.GetCurrentDateTime.Local DateTimeFormat: DateTime.DateTimeFormat.DateAndTime CurrentDateTime=> CurrentDateTime
Text.ConvertDateTimeToText.FromCustomDateTime DateTime: CurrentDateTime CustomFormat: $'''dd.MM.yyyy''' Result=> CsdFileDate
Word.LaunchWord.LaunchAndOpen Path: $'''C:\\Users\\aditya.sagave\\Desktop\\CSD-online\\Consigment Security Declaration -%CsdFileDate%.docx''' Visible: True ReadOnly: False Instance=> csdInstance
Word.FindAndReplaceWord.FindAndReplaceSingle Instance: csdInstance TextToFind: $'''Master Bill Number:  016-''' TextToReplaceWith: $'''Master Bill Number:  016- %awbID%''' MatchCase: False MatchEntireWord: False
Word.CloseWord.CloseAndSave Instance: csdInstance
**ENDREGION
LABEL FindEmail_update_CI_CSD
Outlook.RetrieveEmailMessages.RetrieveEmails Instance: OutlookInstance Account: $'''aditya.sagave@landauer.com''' MailFolder: $'''Inbox''' EmailsToRetrieve: Outlook.RetrieveMessagesMode.All MarkAsRead: True ReadBodyAsHtml: False SubjectContains: $'''NEW BOOKING %FormattedDateTime2% test14''' Messages=> RetrievedEmails
@@copilotGeneratedAction: 'True'
IF RetrievedEmails.Count = 0 THEN
    GOTO FindEmail_update_CI_CSD
@@copilotGeneratedAction: 'True'
END
LOOP FOREACH CurrentItem IN RetrievedEmails
    Outlook.RespondToMailMessage.ReplyAllToEmail Instance: OutlookInstance Account: $'''aditya.sagave@landauer.com''' MailMessage: CurrentItem Body: $'''<strong>Dear Team,</strong><br><br>

Thank you for your prompt responses.<br><br>

Please find attached updated copies of <strong>Consignment Security Declaration</strong> and <strong>Customs Invoice</strong>.<br><br>

<strong>Waybill Number:</strong> 016-%awbID%<br>
<strong>Shipment ID:</strong>%jobID%<br>
<strong>Date:</strong>%FormattedDateTime2%<br><br>

The shipment will be ready for pick up by <strong>1 PM</strong>.<br>
Thank you!''' Attachments: $'''\"C:\\Users\\aditya.sagave\\Desktop\\CSD-online\\Consigment Security Declaration -%CsdFileDate%.docx\" \"C:\\Users\\aditya.sagave\\Desktop\\CI-online\\LDR Inc Customs Invoice %CiFileDate%.docx\"'''
END
WAIT (File.WaitForFile.Created File: $'''C:\\Users\\aditya.sagave\\Desktop\\CI-online\\LDR Inc Customs Invoice %CiFileDate%.docx''')
WAIT (File.WaitForFile.Created File: $'''C:\\Users\\aditya.sagave\\Desktop\\CSD-online\\Consigment Security Declaration -%CsdFileDate%.docx''')
Workstation.PrintDocument DocumentPath: $'''C:\\Users\\aditya.sagave\\Desktop\\CI-online\\LDR Inc Customs Invoice %CiFileDate%.docx'''
Workstation.PrintDocument DocumentPath: $'''C:\\Users\\aditya.sagave\\Desktop\\CSD-online\\Consigment Security Declaration -%CsdFileDate%.docx'''
Workstation.PrintDocument DocumentPath: $'''C:\\Users\\aditya.sagave\\Desktop\\CSD-online\\Consigment Security Declaration -%CsdFileDate%.docx'''
