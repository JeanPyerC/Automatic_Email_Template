Attribute VB_Name = "Module1"
Sub Email_Template()

    Dim emailApplication As Object
    Dim emailItem As Object
    Dim emailRng As Range, CL As Range
    Dim sTO As String
    Dim signature As String

    ' C - Primary Email / D - Secondary Email
    Set emailRng = Worksheets("My Vendor List").Range("C6:D6")

    Set emailApplication = CreateObject("Outlook.Application")


    While Not emailRng.Cells(1, 1).Value = vbNullString

        For Each CL In emailRng
            sTO = sTO & ";" & CL.Value
        Next
        sTO = Mid(sTO, 2)

        Set emailItem = emailApplication.CreateItem(0)

        ' Building the email'
        emailItem.to = sTO
        
        emailItem.Subject = "MONTHLY ORDER UPDATES"
        
        emailItem.Body = "Hi," _
        & vbNewLine & vbNewLine & _
        "Can you please send us a file of our current active order(s)? Please note which materials are available, production date if any, and the pallet quantity." _
        & vbNewLine & vbNewLine & _
        "Please do not include materials we are planning to load, already loaded with a forwarder, or not approved by us." _
        & vbNewLine & vbNewLine & _
        "We will prioritize each item and load them as soon as possible. Thank you!" _
        & vbNewLine & vbNewLine & _
        "Also, please take the time to view this month's newsletter for our company."

        
        emailItem.Attachments.Add ("C:\Users\Owner\OneDrive\PROJECTS\EXCEL PROJECTS\AUTOMAIC EMAIL TEMPLATE\NEWSLETTER.pdf")

        emailItem.Display
        'emailItem.Send

        sTO = vbNullString
        Set emailRng = emailRng.Offset(1, 0)

    Wend

    Set emailItem = Nothing
    Set emailApplication = Nothing

End Sub
