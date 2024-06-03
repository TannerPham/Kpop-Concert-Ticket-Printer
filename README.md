# Code Breakdown 
## 1. "DataPrep" Module:
### a. "Auto Printer" Sub:
    Sub Auto_Printer()
    Dim i As Long
    Dim k As Long
    Sold_out = False
    Entry_Point ' turn off screen updates
    'import info from print_form to sheet prep
    For i = 1 To 5 Step 2 ' loop with an increment of 2 in order to skip even row numbers
        Sel_Area = ShPrep.Range("i" & i).Value
        Sel_Qty = ShPrep.Range("i" & i + 1).Value


        If Sel_Area = "" Then GoTo skip:
        Generate_Tickets 'a sub that generates random ticket_id
        If Sold_out = True Then Exit Sub ' check if tickets of the area is not sufficient
        
        For k = 0 To Sel_Qty - 1
            If ShPrep.Range("StartTickRow").Value = 7 Then ShPrep.Range("StartTickRow").Value = ShPrep.Range("StartTickRow").Value - 1
            Print_Tickets (ShPrep.Range("StartTickRow").Value + k) ' the parameter of the function is the starting row num
        
        Next k
    
        'update the current numbers of seats and rows in the master database
        Update_MasterDB
        'load the updated numbers into the ticket inventory section of the print_form
        Load_Inventory
    skip:
    Next i
    Use_PreviousFP = False
    MsgBox "The printing process has been completed!", , "Printing Completed"
    Exit_Point 'turn on screen updates
### b. "Generate Tickets" Sub:
    Sub Generate_Tickets()
    Dim TicketTbl As ListObject, MasterTbl As ListObject, AreaList As Range
    Dim CurSeatNum As Long, CurRowNum As Long
    Dim TicketID As String, TicketRow As Long, CurDateTime As Date
    Dim i As Long
    
    Set TicketTbl = ShPrep.ListObjects("Ticket_Records")
    Set MasterTbl = ShDB.ListObjects("Master_Table")
    ' get a list of all the existing areas
    Set AreaList = MasterTbl.ListColumns("Area").Range
    'identify the starting row in the ticket table
    ShPrep.Range("StartTickRow").Value = ShPrep.ListObjects(1).Range(1, 1).End(xlDown).row
    If ShPrep.Range("StartTickRow").Value = 7 Then
        TicketRow = 1
    Else
        TicketRow = ShPrep.Range("StartTickRow").Value - 5
    End If
    ' assign relevant information of the ticket to the named ranges in the Prep sheet
    ShPrep.Range("AreaRow").Value = AreaList.Find(Sel_Area, , xlValues, xlWhole).row - 2
    ShPrep.Range("SeatCol").Value = MasterTbl.ListColumns("Current Seat Number").Index
    ShPrep.Range("CurSeatNum").Value = MasterTbl.Range(ShPrep.Range("AreaRow").Value, ShPrep.Range("SeatCol").Value).Value
    ShPrep.Range("CurRowNum").Value = MasterTbl.Range(ShPrep.Range("AreaRow").Value, MasterTbl.ListColumns("Current Row Number").Index).Value
    ShPrep.Range("Ticket_Class").Value = MasterTbl.Range(ShPrep.Range("AreaRow").Value, 2).Value
    ShPrep.Range("Ticket_Class").Interior.Color = MasterTbl.Range(ShPrep.Range("AreaRow").Value, 2).Interior.Color
    ' assign named ranges to variables
    CurDateTime = VBA.DateAdd("m", -16, VBA.Now)
    CurSeatNum = ShPrep.Range("CurSeatNum").Value
    CurRowNum = ShPrep.Range("CurRowNum").Value
    ' check if there is enough tickets in the specified area
    If ShPrep.Range("CurSeatNum").Value + Sel_Qty > MasterTbl.Range(ShPrep.Range("AreaRow").Value, ShPrep.Range("SeatCol").Value - 2).Value Then
        MsgBox "There are not enough tickets left in area " & Sel_Area & "." & vbNewLine _
        & "Please look at the ticket inventory table below to check the remaining tickets."
        Sold_out = True
        Exit Sub
    End If
    ' Run a loop to create as many tickets as required
    For i = 1 To Sel_Qty
        'move to the next row and reset seat number if the current row runs out of seats
        If ShPrep.Range("CurSeatNum").Value + i > MasterTbl.Range(ShPrep.Range("AreaRow").Value, ShPrep.Range("SeatCol").Value - 2).Value Then
            CurSeatNum = ShPrep.Range("CurSeatNum").Value + i - MasterTbl.Range(ShPrep.Range("AreaRow").Value, ShPrep.Range("SeatCol").Value - 2).Value
            If CurSeatNum = 1 Then CurRowNum = CurRowNum + 1
        Else
            CurSeatNum = ShPrep.Range("CurSeatNum").Value + i
        End If
        'assign respective values to a record of the ticket table
        TicketTbl.Range(TicketRow + i, 1) = CurDateTime 'the datetime when tickets were issued
        TicketTbl.Range(TicketRow + i, 2) = Sel_Area & GetTicketID(6) 'generate random ticket_id
        TicketTbl.Range(TicketRow + i, 3) = Sel_Area ' selected_area
        TicketTbl.Range(TicketRow + i, 4) = CurRowNum ' the current row number
        TicketTbl.Range(TicketRow + i, 5) = CurSeatNum ' the current seat number
    Next i
    
    End Sub
### c. "ShowPrintForm" Sub:
    Sub ShowPrintForm()
    Load_Inventory
    Print_Form.Show
    End Sub
## 2. "Database" Module:
### a. "Update MasterDB" Sub:
    Sub Update_MasterDB()
    Dim AreaRow As Byte, SeatNumCol As Long, MaxSeatCol As Byte
    Dim MasterTbl As ListObject, CurSeatnumber As Long
    Dim AreaList As Range
    
    Set MasterTbl = ShDB.ListObjects("Master_Table")
    On Error GoTo handling
    AreaRow = ShPrep.Range("AreaRow").Value 'get the row num of the area in the table
    SeatNumCol = ShPrep.Range("SeatCol").Value ' get the column num of the area in the table
    MaxSeatCol = MasterTbl.ListColumns("Max Seats per Row").Index ' get the index of the 'maximum seats per row' column
    CurSeatnumber = MasterTbl.Range(AreaRow, SeatNumCol).Value 'find the current seat number
    
    ' move to the next row after the current row runs out of seats
    If CurSeatnumber + Sel_Qty > MasterTbl.Range(AreaRow, MaxSeatCol).Value Then
        MasterTbl.Range(AreaRow, SeatNumCol - 1).Value = MasterTbl.Range(AreaRow, SeatNumCol - 1).Value + 1
        MasterTbl.Range(AreaRow, SeatNumCol).Value = (CurSeatnumber + Sel_Qty) - MasterTbl.Range(AreaRow, MaxSeatCol).Value
    Else
        MasterTbl.Range(AreaRow, SeatNumCol).Value = CurSeatnumber + Sel_Qty
    End If
    Exit Sub
    'this error occurs because the subs might not be executed in the correct order
    handling:
    MsgBox "Some crucial information is missing, please try running the 'Auto Printer' sub"
    End Sub

### b. "Reset_MasterDB" Sub:
    Sub Reset_MasterDB()
    Dim MasterTbl As ListObject, TicketTbl As ListObject
    Dim TblRow As Long
    Dim confirm As VbMsgBoxResult
    Set MasterTbl = ShDB.ListObjects("Master_Table")
    Set TicketTbl = ShPrep.ListObjects("Ticket_Records")
    'add a confirmation step to prevent the risk of mis-clicking the buttons
    confirm = MsgBox("Are you sure?", vbYesNo, "Delete all data???")
    If confirm = vbNo Then Exit Sub
    
    'change values of all the rows back to its default value
    TblRow = MasterTbl.ListRows.count
    MasterTbl.ListColumns("Current Row Number").DataBodyRange.Range(ShDB.Cells(1, 1), ShDB.Cells(TblRow, 1)).Value = 1
    MasterTbl.ListColumns("Current Seat Number").DataBodyRange.Range(ShDB.Cells(1, 1), ShDB.Cells(TblRow, 1)).Value = 0
    On Error Resume Next ' ignore the error if the ticket table has already been deleted
    TicketTbl.DataBodyRange.Delete
    End Sub
### c. "Load Inventory" Sub:
    Sub Load_Inventory()
    Dim ColArray(4) As Integer, LowerB As Integer, UpperB As Integer
    Dim MasterTbl As ListObject
    Dim MasterRow As Long, MasterCol As Long, i As Integer, k As Integer
    Set MasterTbl = ShDB.ListObjects("Master_Table")
    'assign the indexes of relevant columns to the array
    ColArray(1) = MasterTbl.ListColumns("Area").Index
    ColArray(2) = MasterTbl.ListColumns("Quantity").Index
    ColArray(3) = MasterTbl.ListColumns("Issued Tickets").Index
    ColArray(4) = MasterTbl.ListColumns("Remaining Tickets").Index
    'define upperbound and lowerbound
    LowerB = LBound(ColArray)
    UpperB = UBound(ColArray)
    'modify some attributes of the listbox
    Print_Form.InvLB.ColumnCount = 5
    Print_Form.InvLB.ColumnWidths = "1,25,60,40,50"
    Print_Form.InvLB.TextAlign = fmTextAlignCenter

    
    For MasterRow = 1 To MasterTbl.ListRows.count 'the number of rows of the Mastertbl
        For i = LowerB + 1 To UpperB
           Print_Form.InvLB.AddItem
           Print_Form.InvLB.List(MasterRow - 1, i) = MasterTbl.DataBodyRange(MasterRow, ColArray(i)).Value
            
        Next i
    Next MasterRow
    
    
    For k = Print_Form.InvLB.ListCount - 1 To 0 Step -1
        ' check each if it contains meaningful text
        If Trim(Print_Form.InvLB.List(k, 2) & vbNullString) = vbNullString Then
            ' if not, delete that item
            Print_Form.InvLB.RemoveItem (k)
        End If
    Next k
    
    End Sub
## 3.  "Functions" Module:
### a. "Get Ticket ID" Function:
    Function GetTicketID(n As Long) As String
        
    Dim i As Long, j As Long, text_len As Byte, number_len As Byte
    Dim Barcode As String, number_pool As String, text_pool As String
    
    number_pool = "0123456789"
    text_pool = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
    text_len = Len(text_pool) ' the number of chars in the "text_pool" list
    number_len = Len(number_pool) ' the number of chars in the "number_pool" list
    
    For i = 1 To n
        
        If WorksheetFunction.RandBetween(1, 2) = 1 Then 'randomize between 1 and 2 to choose number of text
        j = 1 + Int(text_len * Rnd()) ' 1 = text
        Barcode = Barcode & Mid(text_pool, j, 1)
        Else
        j = 1 + Int(number_len * Rnd()) '2 = number
        Barcode = Barcode & Mid(number_pool, j, 1)
        End If
    Next i
    GetTicketID = Barcode ' the function return the randomized characters
    
    End Function

### b. "Print Ticket" Function:
    Function Print_Tickets(StartRow As Long)
    Dim Barcode As String
    Dim Filepath As String
    Dim Folderpath As String
    Dim Folder As String
    Dim TickRow As String
    Dim Answer As VbMsgBoxResult
    TickRow = StartRow - 4 ' find the row in the ticket table
    Barcode = ShPrep.ListObjects(1).Range(TickRow, 2).Value ' get the unique barcode (ticket_id) of the ticket
    
    'modify the information displaying on the sheet "ticket form" to get ready for printing
    ShTF.Shapes("Barcode").TextFrame2.TextRange.Text = "*" & Barcode & "*"
    ShTF.Shapes("Hour_Text").TextFrame2.TextRange.Text = ShPrep.Range("b2").Text
    ShTF.Shapes("Date_Text").TextFrame2.TextRange.Text = ShPrep.Range("b1").Text
    ShTF.Shapes("Location_Text").TextFrame2.TextRange.Text = ShPrep.Range("b3").Text
    ShTF.Shapes("Area_Text").TextFrame2.TextRange.Text = ShPrep.ListObjects(1).Range(TickRow, 3).Value
    ShTF.Shapes("Row_Text").TextFrame2.TextRange.Text = ShPrep.ListObjects(1).Range(TickRow, 4).Value
    ShTF.Shapes("Seat_Text").TextFrame2.TextRange.Text = ShPrep.ListObjects(1).Range(TickRow, 5).Value
    ShTF.Shapes("TicketClass").TextFrame2.TextRange.Text = ShPrep.Range("Ticket_Class").Text
    ShTF.Shapes("TicketClass").TextFrame2.TextRange.Font.Fill.ForeColor.RGB = ShPrep.Range("Ticket_Class").Interior.Color
    
    'check if users want to save new batches of tickets at the previous folder path
    If Use_PreviousFP = False Then ' this variable's value is False by default
        'generate name for the new folder path
        Folderpath = ShPrep.Range("b4").Text & VBA.Format(ShPrep.ListObjects(1).Range(TickRow, 1).Text, "yyyy-mm-dd hhmmss")
    Else
        'use the old directory
        Folderpath = ShPrep.Range("PreviousFP").Value
    End If
    
    
    Folder = Dir(Folderpath, vbDirectory) 'check if the newly created folder path exists
    If Folder = vbNullString Then ' this means the folder path does not exists
        'ask if users want use the previous folder path or not
        Answer = MsgBox("Do you want to save this batch of tickets (Area: " & ShPrep.ListObjects(1).Range(TickRow, 3).Value & ") in the previous folder (" & Dir(ShPrep.Range("PreviousFP").Value, vbDirectory) & ")?" & vbNewLine _
            & "If you select NO, new folder will be created for this batch.", vbYesNo, "Save Folder?")
        Select Case Answer
            Case vbNo ' users select NO => create new folder path
                VBA.FileSystem.MkDir (Folderpath)
            Case Else ' users select YES => use the previous folder path
                'check if there is any previous folder path from the named range
                Folder = Dir(ShPrep.Range("PreviousFP").Value, vbDirectory)
                
                If Folder = vbNullString Then
                    'create a default folder in case there is no previous folder path
                    VBA.FileSystem.MkDir (ShPrep.Range("b4").Text & "Default_Ticket_Folder")
                    Folderpath = ShPrep.Range("b4").Text & "Default_Ticket_Folder"
                Else
                    Folderpath = ShPrep.Range("PreviousFP").Value
                End If
                Use_PreviousFP = True
        End Select
    
    End If
    
    ShPrep.Range("PreviousFP").Value = Folderpath 'store the folderpath to the named range
    Filepath = Folderpath & "\" & VBA.Left(Barcode, 6) & "___" 'create a filepath
    ShTF.ExportAsFixedFormat xlTypePDF, Filepath ' export the sheet Ticket Form as PDF
    End Function


## 4. "Admin" Module:
### a. "Entry Point" Function:
    Sub Entry_Point()
    With Application
    .ScreenUpdating = False
    .DisplayAlerts = False
    .StatusBar = "Printing..."
    .Calculation = xlCalculationManual
    
    End With
    
    End Sub
### b. "Exit Point" Function:
    Sub Exit_Point()
    With Application
    .ScreenUpdating = True
    .DisplayAlerts = True
    .StatusBar = ""
    .Calculation = xlCalculationAutomatic
    .CutCopyMode = False
    
    End With
    End Sub
## 5. "Print_Form" UserForm:
### a. "CancelButton_Click" Private Sub:
    Private Sub CancelButton_Click()
    Unload Me
    End Sub
### b. "PrintButton_Click" Private Sub:
    Private Sub PrintButton_Click()
    Dim i As Byte
    'import values from comboboxes to named ranges of the prep sheet
    For i = 1 To 6
        ShPrep.Range("i" & i).Value = Print_Form.Controls("Field" & i).Value
    
    Next i
    Auto_Printer ' run the main sub of the file
    'remove values from the comboboxes after printing is done
    For i = 1 To 6
        Print_Form.Controls("Field" & i).Value = ""
    
    Next i
    
    End Sub
