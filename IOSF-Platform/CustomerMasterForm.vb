Imports Microsoft.Data.SqlClient
Imports System.Windows.Forms
Imports System.Drawing

''' <summary>
''' Ports the "Customer_Master" PowerApp - a master-detail editor for Customer_Ops_Header
''' (one row per customer, business/service-term fields) and Customer_Ops_Item (many rows
''' per customer, keyed by Account_Num + Cont_Num - individual contacts/users under that
''' customer, including RemoteLock_Manual and other fields that likely connect to
''' RemoteLockUsersJob/PaperCutSyncJob elsewhere in this port, though that connection
''' wasn't independently traced).
'''
''' REVISED after Al's review of the first version - see individual method remarks for
''' what changed and why. Confirmed with Al before the original build:
'''  - Creating a new customer still auto-prompts for its first Item entry, matching the
'''    original's own Header New -> Item New Entry navigation on successful save.
'''
''' Customer list now queries Customer_QB directly (not Customer_Ops_Header.Company_Evo,
''' which is itself just a synced/computed copy) - matches Al's own PowerFX formula for
''' the original gallery: IsActive=true AND Value(AccountNumber) < 9000, sorted by
''' CompanyName. See LoadCustomerList for the exact SQL.
'''
''' Items section (Customer_Ops_Item) is UNCHANGED from the first version - Al asked to
''' defer that discussion ("We'll work through the contact level after").
''' </summary>
Public Class CustomerMasterForm
    Inherits Form

    Private ReadOnly headerColumns As New List(Of String)
    Private ReadOnly headerControls As New Dictionary(Of String, Control)

    ' Fields needing non-default treatment, per Al's review. Discovered by field name
    ' convention/direct feedback, not by checking real SQL types - same "don't guess a
    ' type wrong" caution as the rest of this port, applied to CONTROL CHOICE rather than
    ' parameter type this time.
    Private ReadOnly dateFields As New HashSet(Of String) From {"Date_Sold", "Date_Renewed", "Term_Exp", "Terminated"}
    Private ReadOnly readOnlyFields As New HashSet(Of String) From {"Term_Exp", "Company_Evo"} ' Term_Exp computed elsewhere; Company_Evo is a synced/computed field per Al ("comes from the view")
    Private ReadOnly currencyFields As New HashSet(Of String) From {"Service_Amt"}
    Private ReadOnly blankOrXFields As New HashSet(Of String) From {"Member", "Scan_Package", "Autorenew"}
    Private ReadOnly excludedFields As New HashSet(Of String) From {"Version_Stamp"} ' SQL Server rowversion/timestamp column - system-managed, cannot be written to directly, and is almost certainly what caused the earlier binary-data grid errors

    Private searchBox As TextBox
    Private customerList As ListBox
    Private customerListTable As DataTable

    Private headerPanel As Panel
    Private saveCustomerButton As Button
    Private newCustomerButton As Button

    Private itemsGrid As DataGridView
    Private itemsAdapter As SqlDataAdapter
    Private itemsTable As DataTable
    Private saveItemsButton As Button
    Private deleteItemButton As Button

    Private statusLabel As Label

    Private currentAccountNum As Integer?
    Private isNewCustomer As Boolean

    Public Sub New()
        Text = "Customer Master"
        Width = 1150
        Height = 850
        StartPosition = FormStartPosition.CenterScreen

        Dim topSection = BuildTopSection()
        BuildItemsGrid()
        Dim itemsButtonPanel = BuildItemsButtonPanel()
        Dim statusBar = BuildStatusBar()

        ' Fill-first-then-edges rule (see TableEditorForm/LandingPageForm remarks for why
        ' this order matters). Visual order top-to-bottom: topSection (search+header),
        ' [ItemsGrid fill], ItemButtons, StatusBar.
        Controls.Add(itemsGrid) ' Fill - added first
        Controls.Add(itemsButtonPanel) ' Bottom - innermost
        Controls.Add(statusBar) ' Bottom - outermost
        Controls.Add(topSection) ' Top

        LoadHeaderSchema()
        LoadCustomerList()
    End Sub

    ' ===================== Top section: left customer list + right header form =====================

    ''' <summary>
    ''' REDESIGNED per Al's review: was a stacked search-grid-on-top, header-below layout;
    ''' now a left/right split matching the original PowerApp's own "gallery on the left"
    ''' layout - a plain company-name list on the left, header form on the right. Also
    ''' fixes a real bug: the old search filtered across every loaded column including
    ''' Term_Exp (date) and Terminated, and converting those to strings inside a
    ''' DataView.RowFilter expression was throwing, silently leaving a stale/overly-narrow
    ''' filter in place (caught by the Try/Catch, but the practical effect was "typing
    ''' anything shows nothing"). Now filters on company name only, which is also all Al
    ''' said he actually needs from this list.
    ''' </summary>
    Private Function BuildTopSection() As Panel
        Dim outer As New Panel With {.Dock = DockStyle.Top, .Height = 560}

        ' --- Right: header form (added first per the Fill-first rule within this panel) ---
        Dim rightPanel As New Panel With {.Dock = DockStyle.Fill}

        Dim headerButtonPanel As New FlowLayoutPanel With {.Dock = DockStyle.Top, .Height = 40, .Padding = New Padding(8)}
        newCustomerButton = New Button With {.Text = "New Customer", .Width = 120, .Margin = New Padding(0, 0, 8, 0)}
        AddHandler newCustomerButton.Click, AddressOf NewCustomerClicked
        saveCustomerButton = New Button With {.Text = "Save Customer", .Width = 120, .Margin = New Padding(0, 0, 8, 0)}
        AddHandler saveCustomerButton.Click, AddressOf SaveCustomerClicked
        headerButtonPanel.Controls.Add(newCustomerButton)
        headerButtonPanel.Controls.Add(saveCustomerButton)

        ' REAL BUG FIXED: was a TableLayoutPanel, which produced a real, persistent
        ' alignment bug across two attempted fixes (explicit row/col computation, then
        ' explicit RowStyles) - neither resolved it. Switched to a plain Panel with fully
        ' manual (X,Y) positioning per field in LoadHeaderSchema, the same approach that
        ' resolved LandingPageForm's layout bugs earlier this session - removes all
        ' ambiguity about where each field ends up, rather than relying on
        ' TableLayoutPanel's auto-sizing behavior.
        headerPanel = New Panel With {
            .Dock = DockStyle.Fill,
            .AutoScroll = True
        }

        rightPanel.Controls.Add(headerPanel) ' Fill - added first within rightPanel
        rightPanel.Controls.Add(headerButtonPanel) ' Top

        ' --- Left: company list (added second, so it correctly claims the left strip) ---
        Dim leftPanel As New Panel With {.Dock = DockStyle.Left, .Width = 320}

        customerList = New ListBox With {
            .Dock = DockStyle.Fill
        }
        AddHandler customerList.SelectedIndexChanged, AddressOf CustomerListSelectionChanged

        Dim searchPanel As New FlowLayoutPanel With {.Dock = DockStyle.Top, .Height = 36, .Padding = New Padding(8)}
        Dim searchLabel As New Label With {.Text = "Search:", .AutoSize = True, .Padding = New Padding(0, 6, 4, 0)}
        searchBox = New TextBox With {.Width = 230}
        AddHandler searchBox.TextChanged, AddressOf SearchTextChanged
        searchPanel.Controls.Add(searchLabel)
        searchPanel.Controls.Add(searchBox)

        leftPanel.Controls.Add(customerList) ' Fill - added first within leftPanel
        leftPanel.Controls.Add(searchPanel) ' Top

        outer.Controls.Add(rightPanel) ' Fill - added first within outer
        outer.Controls.Add(leftPanel) ' Left
        Return outer
    End Function

    ''' <summary>
    ''' REBUILT per Al: the original gallery is built from Customer_QB directly, not
    ''' Customer_Ops_Header - confirmed from Al's own PowerFX formula:
    ''' Filter(Search('Customer_QB', TextInput1.Text, CompanyName), IsActive=true &&
    ''' Value(AccountNumber) < 9000). Customer_Ops_Header.Company_Evo is itself just a
    ''' synced/computed copy ("comes from the view", per Al), which explains why the
    ''' earlier version's list was blank - it was reading a column that isn't the
    ''' authoritative source and can be blank/stale.
    '''
    ''' Customer_QB.AccountNumber is stored as a string (confirmed elsewhere in this
    ''' port), so the "< 9000" filter needs an explicit numeric conversion - TRY_CAST
    ''' rather than CAST, so a non-numeric AccountNumber value (if one ever exists) is
    ''' excluded rather than throwing and breaking the whole list.
    ''' </summary>
    Private Sub LoadCustomerList()
        Try
            Const sql As String =
                "SELECT AccountNumber, CompanyName FROM Customer_QB " &
                "WHERE IsActive = 1 AND TRY_CAST(AccountNumber AS INT) < 9000 " &
                "ORDER BY CompanyName"
            Using conn As New SqlConnection(ConfigHelper.ConnectionString)
                Using adapter As New SqlDataAdapter(sql, conn)
                    customerListTable = New DataTable()
                    adapter.Fill(customerListTable)
                    customerList.DisplayMember = "CompanyName"
                    customerList.ValueMember = "AccountNumber"
                    customerList.DataSource = customerListTable
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show(Me, $"Error loading customer list: {ex.Message}", "Customer Master", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub SearchTextChanged(sender As Object, e As EventArgs)
        If customerListTable Is Nothing Then Return
        Try
            Dim term = searchBox.Text
            If String.IsNullOrEmpty(term) Then
                customerListTable.DefaultView.RowFilter = String.Empty
            Else
                customerListTable.DefaultView.RowFilter = $"CompanyName LIKE '%{term.Replace("'", "''")}%'"
            End If
            customerList.DataSource = customerListTable.DefaultView
        Catch
            ' mid-keystroke, not yet a valid filter expression - ignore
        End Try
    End Sub

    ''' <summary>Customer_QB.AccountNumber is a string - converted here before use as the integer Customer_Ops_Header.Account_Num.</summary>
    Private Sub CustomerListSelectionChanged(sender As Object, e As EventArgs)
        If customerList.SelectedValue Is Nothing OrElse customerList.SelectedValue Is DBNull.Value Then Return
        LoadCustomer(Convert.ToInt32(customerList.SelectedValue))
    End Sub

    ' ===================== Header (Customer_Ops_Header) =====================

    ''' <summary>
    ''' Builds one Label+input pair per real Customer_Ops_Header column (excluding
    ''' Version_Stamp - see excludedFields remarks), using the table's actual schema
    ''' rather than a hardcoded field list. The input control type varies per Al's review:
    ''' DateTimePicker (date-only, with a checkbox for allowing blank/NULL) for date
    ''' fields, a locked two-item ComboBox for blank-or-"X" fields, a read-only TextBox for
    ''' Term_Exp (computed elsewhere), and a plain TextBox for everything else.
    ''' </summary>
    Private Sub LoadHeaderSchema()
        Try
            headerColumns.Clear()
            Using conn As New SqlConnection(ConfigHelper.ConnectionString)
                Using cmd As New SqlCommand("SELECT TOP 0 * FROM Customer_Ops_Header", conn)
                    conn.Open()
                    Using reader = cmd.ExecuteReader()
                        Dim schemaTable = reader.GetSchemaTable()
                        For Each row As DataRow In schemaTable.Rows
                            Dim colName = row("ColumnName").ToString()
                            If Not excludedFields.Contains(colName) Then headerColumns.Add(colName)
                        Next
                    End Using
                End Using
            End Using

            headerPanel.Controls.Clear()

            ' Fully manual (X,Y) positioning - see BuildTopSection's remarks for why this
            ' replaced TableLayoutPanel. Every field's position is computed explicitly
            ' from its index, removing any dependency on an auto-sizing algorithm.
            Const marginX As Integer = 8
            Const marginY As Integer = 8
            Const rowHeight As Integer = 32
            Const labelWidth As Integer = 120
            Const inputWidth As Integer = 260
            Const columnGap As Integer = 24
            Dim col2X = marginX + labelWidth + inputWidth + columnGap

            For i = 0 To headerColumns.Count - 1
                Dim colName = headerColumns(i)
                Dim rowIdx = i \ 2
                Dim isSecondColumn = (i Mod 2) = 1
                Dim baseX = If(isSecondColumn, col2X, marginX)
                Dim y = marginY + rowIdx * rowHeight

                Dim label As New Label With {
                    .Text = colName,
                    .Location = New Point(baseX, y + 4),
                    .Size = New Size(labelWidth, 20),
                    .TextAlign = ContentAlignment.MiddleLeft
                }

                Dim input As Control = CreateHeaderInput(colName)
                input.Location = New Point(baseX + labelWidth, y)
                input.Width = inputWidth
                headerControls(colName) = input

                headerPanel.Controls.Add(label)
                headerPanel.Controls.Add(input)
            Next
        Catch ex As Exception
            MessageBox.Show(Me, $"Error loading Customer_Ops_Header schema: {ex.Message}", "Customer Master", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Function CreateHeaderInput(colName As String) As Control
        ' Read-only date fields (Term_Exp) use a plain read-only TextBox instead - per
        ' Al, a date picker with a checkbox is more UI than a never-editable field needs,
        ' and DateTimePicker.Value can't represent true blankness (it's always a valid
        ' date internally, even when unchecked) the way a TextBox's empty string can.
        ' Editable date fields (Date_Sold, Date_Renewed, Terminated) keep the picker.
        If dateFields.Contains(colName) AndAlso Not readOnlyFields.Contains(colName) Then
            Dim dtp As New DateTimePicker With {
                .Format = DateTimePickerFormat.Short, ' date only, no time - per Al
                .ShowCheckBox = True ' lets the field represent NULL/blank (unchecked) rather than forcing a date
            }
            Return dtp
        End If

        If blankOrXFields.Contains(colName) Then
            Dim combo As New ComboBox With {.DropDownStyle = ComboBoxStyle.DropDownList}
            combo.Items.Add("") ' blank
            combo.Items.Add("X")
            combo.SelectedIndex = 0
            Return combo
        End If

        Dim textBox As New TextBox With {.ReadOnly = readOnlyFields.Contains(colName)}
        Return textBox
    End Function

    ''' <summary>
    ''' Company_Evo is loaded separately from Evo_Customer_XRef, not from
    ''' Customer_Ops_Header itself - REAL BUG FIXED, per Al's correction: this originally
    ''' assumed Company_Evo was a column on Customer_Ops_Header (blank in practice), but
    ''' the original's own formula is LookUp('Evo_Customer_XRef',
    ''' ThirdPartyAccountId=varCompanyNum, Name). Fetched here instead of via the main
    ''' SELECT * below, so headerColumns still includes "Company_Evo" for layout/save
    ''' purposes (it's read-only, so it's never actually written back), but its VALUE
    ''' comes from this separate lookup.
    ''' </summary>
    ''' <summary>
    ''' Company_Evo (from Evo_Customer_XRef) and Term_Exp (from Customer_Term_Exp) are both
    ''' loaded separately here, not from Customer_Ops_Header's own row - confirmed by Al's
    ''' own PowerFX formulas for both. Customer_Term_Exp's real column is "Term Expiration"
    ''' (with a literal space) - PowerFX's Term_x0020_Expiration is its own encoding for a
    ''' space character (_x0020_ is the Unicode escape for space), not the real SQL name.
    ''' </summary>
    Private Sub LoadCustomer(accountNum As Integer)
        Try
            Using conn As New SqlConnection(ConfigHelper.ConnectionString)
                Using cmd As New SqlCommand("SELECT * FROM Customer_Ops_Header WHERE Account_Num = @AccountNum", conn)
                    cmd.Parameters.AddWithValue("@AccountNum", accountNum)
                    conn.Open()
                    Using reader = cmd.ExecuteReader()
                        If reader.Read() Then
                            For Each colName In headerColumns
                                If colName = "Company_Evo" OrElse colName = "Term_Exp" Then Continue For ' loaded separately below
                                Dim value = reader(colName)
                                SetHeaderValue(colName, value)
                            Next
                        End If
                    End Using
                End Using

                Using cmd As New SqlCommand("SELECT Name FROM Evo_Customer_XRef WHERE ThirdPartyAccountId = @AccountNum", conn)
                    cmd.Parameters.AddWithValue("@AccountNum", accountNum.ToString())
                    Dim evoName = cmd.ExecuteScalar()
                    SetHeaderValue("Company_Evo", evoName)
                End Using

                Using cmd As New SqlCommand("SELECT [Term Expiration] FROM Customer_Term_Exp WHERE Account_Num = @AccountNum", conn)
                    cmd.Parameters.AddWithValue("@AccountNum", accountNum)
                    Dim termExp = cmd.ExecuteScalar()
                    SetHeaderValue("Term_Exp", termExp)
                End Using
            End Using

            currentAccountNum = accountNum
            isNewCustomer = False
            CType(headerControls("Account_Num"), TextBox).ReadOnly = True ' primary key of an existing record - not editable
            LoadItemsForAccount(accountNum)
            statusLabel.Text = $"Loaded customer Account_Num={accountNum}."
        Catch ex As Exception
            MessageBox.Show(Me, $"Error loading customer: {ex.Message}", "Customer Master", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>Writes a DB value into whichever control type this field uses.</summary>
    Private Sub SetHeaderValue(colName As String, value As Object)
        Dim ctrl = headerControls(colName)
        Dim isNull = (value Is Nothing OrElse value Is DBNull.Value)

        If TypeOf ctrl Is DateTimePicker Then
            Dim dtp = CType(ctrl, DateTimePicker)
            If isNull Then
                ' REAL BUG FIXED: DateTimePicker.ShowCheckBox unchecking does NOT clear
                ' the displayed date - it only greys out whatever .Value was last set, so
                ' a customer with a blank date was showing the PREVIOUS customer's stale
                ' date instead. Explicitly resetting .Value (not just .Checked) fixes
                ' this - now every customer with no date consistently shows today's date
                ' greyed out, rather than carrying over an unrelated customer's value.
                dtp.Value = DateTime.Today
                dtp.Checked = False
            Else
                dtp.Checked = True
                dtp.Value = Convert.ToDateTime(value)
            End If
        ElseIf TypeOf ctrl Is ComboBox Then
            Dim combo = CType(ctrl, ComboBox)
            Dim text = If(isNull, "", value.ToString().Trim())
            combo.SelectedIndex = If(text = "X", 1, 0)
        Else
            Dim textBox = CType(ctrl, TextBox)
            If isNull Then
                textBox.Text = String.Empty
            ElseIf currencyFields.Contains(colName) Then
                textBox.Text = Convert.ToDecimal(value).ToString("F2") ' two decimals, per Al
            ElseIf dateFields.Contains(colName) Then
                textBox.Text = Convert.ToDateTime(value).ToString("d") ' short date, no time - this is a read-only date field rendered as a TextBox (see CreateHeaderInput), not the DateTimePicker branch above
            Else
                textBox.Text = value.ToString()
            End If
        End If
    End Sub

    ''' <summary>Reads whichever control type this field uses back into a save-ready value (String, Date, or DBNull).</summary>
    Private Function GetHeaderValue(colName As String) As Object
        Dim ctrl = headerControls(colName)

        If TypeOf ctrl Is DateTimePicker Then
            Dim dtp = CType(ctrl, DateTimePicker)
            Return If(dtp.Checked, CObj(dtp.Value.Date), DBNull.Value)
        ElseIf TypeOf ctrl Is ComboBox Then
            Dim text = CType(ctrl, ComboBox).Text
            Return If(String.IsNullOrEmpty(text), CObj(DBNull.Value), text)
        Else
            Dim text = CType(ctrl, TextBox).Text
            Return If(String.IsNullOrEmpty(text), CObj(DBNull.Value), text)
        End If
    End Function

    Private Sub NewCustomerClicked(sender As Object, e As EventArgs)
        For Each colName In headerColumns
            SetHeaderValue(colName, DBNull.Value)
        Next
        CType(headerControls("Account_Num"), TextBox).ReadOnly = False
        currentAccountNum = Nothing
        isNewCustomer = True
        itemsGrid.DataSource = Nothing
        itemsTable = Nothing
        statusLabel.Text = "Enter a new Account_Num and other details, then Save Customer."
    End Sub

    ''' <summary>
    ''' Matches the original's own Header New -> Item New Entry cascade (confirmed with
    ''' Al to keep this flow): after successfully creating a NEW customer, prompts and adds
    ''' a blank new row to the Items grid so the first item can be entered immediately.
    ''' Version_Stamp is excluded from headerColumns entirely (see excludedFields remarks),
    ''' so it's never part of this INSERT/UPDATE - including a rowversion column in an
    ''' explicit column list would fail, since SQL Server manages that value itself.
    ''' </summary>
    Private Sub SaveCustomerClicked(sender As Object, e As EventArgs)
        Try
            Dim accountNumText = CType(headerControls("Account_Num"), TextBox).Text
            If String.IsNullOrWhiteSpace(accountNumText) Then
                MessageBox.Show(Me, "Account_Num is required.", "Customer Master")
                Return
            End If
            Dim accountNum = Convert.ToInt32(accountNumText)

            Using conn As New SqlConnection(ConfigHelper.ConnectionString)
                conn.Open()
                ' Company_Evo and Term_Exp both excluded, same reasoning as
                ' Version_Stamp/Account_Num exclusions - their displayed values come from
                ' Evo_Customer_XRef and Customer_Term_Exp respectively (see LoadCustomer),
                ' not from this table, so writing them back here could conflict with
                ' whatever process actually maintains those columns.
                Dim writableColumns = headerColumns.Where(Function(c) c <> "Company_Evo" AndAlso c <> "Term_Exp").ToList()

                Dim sql As String
                If isNewCustomer Then
                    Dim colList = String.Join(", ", writableColumns)
                    Dim paramList = String.Join(", ", writableColumns.Select(Function(c) $"@{c}"))
                    sql = $"INSERT INTO Customer_Ops_Header ({colList}) VALUES ({paramList})"
                Else
                    Dim setList = String.Join(", ", writableColumns.Where(Function(c) c <> "Account_Num").Select(Function(c) $"{c} = @{c}"))
                    sql = $"UPDATE Customer_Ops_Header SET {setList} WHERE Account_Num = @Account_Num"
                End If

                Using cmd As New SqlCommand(sql, conn)
                    For Each colName In writableColumns
                        cmd.Parameters.AddWithValue($"@{colName}", GetHeaderValue(colName))
                    Next
                    cmd.ExecuteNonQuery()
                End Using
            End Using

            Dim wasNew = isNewCustomer
            currentAccountNum = accountNum
            isNewCustomer = False
            CType(headerControls("Account_Num"), TextBox).ReadOnly = True
            LoadCustomerList()
            ' Re-selects the same customer per Al, rather than losing the selection back
            ' to the full list. This naturally re-triggers CustomerListSelectionChanged ->
            ' LoadCustomer(accountNum), which re-fetches every header/item value fresh
            ' from the server (LoadCustomer already calls LoadItemsForAccount internally,
            ' so no separate call is needed here anymore).
            customerList.SelectedValue = accountNum

            If wasNew Then
                statusLabel.Text = "Customer created. Add their first item below."
                MessageBox.Show(Me, "Customer created. Add their first item in the Items grid below.", "Customer Master")
                itemsGrid.Focus()
            Else
                statusLabel.Text = $"Saved customer Account_Num={accountNum}."
            End If
        Catch ex As Exception
            MessageBox.Show(Me, $"Error saving customer: {ex.Message}", "Customer Master", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' ===================== Items (Customer_Ops_Item) - unchanged, Al deferred this =====================

    Private Sub BuildItemsGrid()
        itemsGrid = New DataGridView With {
            .Dock = DockStyle.Fill,
            .AllowUserToAddRows = True,
            .AllowUserToDeleteRows = True,
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells,
            .SelectionMode = DataGridViewSelectionMode.CellSelect
        }
        AddHandler itemsGrid.DefaultValuesNeeded, AddressOf ItemsGridDefaultValuesNeeded
        AddHandler itemsGrid.DataError, AddressOf GridDataError
        AddHandler itemsGrid.CellContextMenuStripNeeded, AddressOf ItemsGridCellContextMenuStripNeeded
    End Sub

    ''' <summary>
    ''' Real gap found and fixed: the standard "select cell, press Delete" gesture doesn't
    ''' clear a DataGridViewDateTimePickerColumn cell (confirmed by Al) - its ValueType is
    ''' DateTime, which can't represent "empty" the way a plain text cell can, so
    ''' DataGridView's built-in clear-on-Delete behavior appears to fail silently against
    ''' it. Adds an explicit right-click "Clear Date" option instead, generic to any
    ''' DataGridViewDateTimePickerColumn cell (not hardcoded to Terminated_Cont), since the
    ''' same limitation applies to any future use of this custom column type.
    ''' </summary>
    Private Sub ItemsGridCellContextMenuStripNeeded(sender As Object, e As DataGridViewCellContextMenuStripNeededEventArgs)
        If e.RowIndex < 0 OrElse e.ColumnIndex < 0 Then Return
        If Not TypeOf itemsGrid.Columns(e.ColumnIndex) Is DataGridViewDateTimePickerColumn Then Return

        Dim menu As New ContextMenuStrip()
        Dim clearItem As New ToolStripMenuItem("Clear Date")
        AddHandler clearItem.Click, Sub()
                                         itemsGrid.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = DBNull.Value
                                     End Sub
        menu.Items.Add(clearItem)
        e.ContextMenuStrip = menu
    End Sub

    Private Sub GridDataError(sender As Object, e As DataGridViewDataErrorEventArgs)
        e.ThrowException = False
    End Sub

    Private Function BuildItemsButtonPanel() As FlowLayoutPanel
        Dim buttonPanel As New FlowLayoutPanel With {.Dock = DockStyle.Bottom, .Height = 44, .Padding = New Padding(8)}
        saveItemsButton = New Button With {.Text = "Save Items", .Width = 110, .Margin = New Padding(0, 0, 8, 0)}
        AddHandler saveItemsButton.Click, AddressOf SaveItemsClicked
        deleteItemButton = New Button With {.Text = "Delete Selected Item(s)", .Width = 160, .Margin = New Padding(0, 0, 8, 0)}
        AddHandler deleteItemButton.Click, AddressOf DeleteItemClicked
        buttonPanel.Controls.Add(saveItemsButton)
        buttonPanel.Controls.Add(deleteItemButton)
        Return buttonPanel
    End Function

    Private Sub LoadItemsForAccount(accountNum As Integer)
        Try
            Dim conn As New SqlConnection(ConfigHelper.ConnectionString)
            itemsAdapter = New SqlDataAdapter("SELECT * FROM Customer_Ops_Item WHERE Account_Num = @AccountNum", conn)
            itemsAdapter.SelectCommand.Parameters.AddWithValue("@AccountNum", accountNum)

            Dim builder As New SqlCommandBuilder(itemsAdapter)
            itemsAdapter.InsertCommand = builder.GetInsertCommand()
            itemsAdapter.UpdateCommand = builder.GetUpdateCommand()
            itemsAdapter.DeleteCommand = builder.GetDeleteCommand()

            itemsTable = New DataTable()
            itemsAdapter.Fill(itemsTable)
            itemsGrid.DataSource = itemsTable
            HideBinaryColumns(itemsGrid)
            ApplyItemsColumnEditors()
        Catch ex As Exception
            MessageBox.Show(Me, $"Error loading items: {ex.Message}", "Customer Master", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' Per Al: blank/"X" dropdown for Frequent_Scans and RemoteLock_Manual (same pattern
    ''' as Member/Scan_Package/Autorenew on the header form), and a date picker for
    ''' Terminated_Cont (see DataGridViewDateTimePickerColumn.vb). Applied by replacing the
    ''' auto-generated column at the same index/DataPropertyName, rather than turning off
    ''' AutoGenerateColumns and hand-defining every column - keeps this working
    ''' automatically if Customer_Ops_Item's schema ever changes elsewhere.
    ''' </summary>
    Private Sub ApplyItemsColumnEditors()
        ReplaceWithComboBoxColumn("Frequent_Scans")
        ReplaceWithComboBoxColumn("RemoteLock_Manual")
        ReplaceWithDateTimePickerColumn("Terminated_Cont")
    End Sub

    Private Sub ReplaceWithComboBoxColumn(columnName As String)
        Dim index = itemsGrid.Columns(columnName)?.Index
        If Not index.HasValue Then Return

        Dim combo As New DataGridViewComboBoxColumn With {
            .Name = columnName,
            .DataPropertyName = columnName,
            .HeaderText = columnName
        }
        combo.Items.Add("")
        combo.Items.Add("X")

        itemsGrid.Columns.RemoveAt(index.Value)
        itemsGrid.Columns.Insert(index.Value, combo)
    End Sub

    Private Sub ReplaceWithDateTimePickerColumn(columnName As String)
        Dim index = itemsGrid.Columns(columnName)?.Index
        If Not index.HasValue Then Return

        Dim dateColumn As New DataGridViewDateTimePickerColumn With {
            .Name = columnName,
            .DataPropertyName = columnName,
            .HeaderText = columnName
        }

        itemsGrid.Columns.RemoveAt(index.Value)
        itemsGrid.Columns.Insert(index.Value, dateColumn)
    End Sub

    Private Shared Sub HideBinaryColumns(grid As DataGridView)
        For Each col As DataGridViewColumn In grid.Columns
            If col.ValueType Is GetType(Byte()) Then col.Visible = False
        Next
    End Sub

    Private Sub ItemsGridDefaultValuesNeeded(sender As Object, e As DataGridViewRowEventArgs)
        If Not currentAccountNum.HasValue Then Return
        e.Row.Cells("Account_Num").Value = currentAccountNum.Value
        e.Row.Cells("Cont_Num").Value = GetNextContNum(currentAccountNum.Value)
    End Sub

    ''' <summary>
    ''' REAL BUG FIXED: originally queried Max_Contact_Num directly (matching the
    ''' original PowerApp's own LookUp formula), but this was confirmed broken via a real
    ''' test - a brand-new Account_Num with zero existing items returned Cont_Num=2 from
    ''' that view, meaning it isn't actually scoped per-account the way the PowerFX
    ''' formula's Account_Num filter implied (or its underlying definition doesn't behave
    ''' as expected). Rather than continue depending on a view whose real behavior turned
    ''' out to be unreliable, this computes the next Cont_Num directly and unambiguously
    ''' from Customer_Ops_Item itself - the actual table the value needs to be unique
    ''' against.
    ''' </summary>
    Private Function GetNextContNum(accountNum As Integer) As Integer
        Try
            Using conn As New SqlConnection(ConfigHelper.ConnectionString)
                Using cmd As New SqlCommand("SELECT ISNULL(MAX(Cont_Num), 0) + 1 FROM Customer_Ops_Item WHERE Account_Num = @AccountNum", conn)
                    cmd.Parameters.AddWithValue("@AccountNum", accountNum)
                    conn.Open()
                    Return Convert.ToInt32(cmd.ExecuteScalar())
                End Using
            End Using
        Catch
            Return 1
        End Try
    End Function

    Private Sub SaveItemsClicked(sender As Object, e As EventArgs)
        If itemsTable Is Nothing Then Return
        Try
            itemsGrid.EndEdit()
            Dim changed = itemsAdapter.Update(itemsTable)
            statusLabel.Text = $"Saved {changed} item change(s)."
        Catch ex As Exception
            MessageBox.Show(Me, $"Error saving items: {ex.Message}", "Customer Master", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub DeleteItemClicked(sender As Object, e As EventArgs)
        Dim rowsToDelete = itemsGrid.SelectedCells.
            Cast(Of DataGridViewCell)().
            Select(Function(c) c.OwningRow).
            Where(Function(r) Not r.IsNewRow).
            Distinct().
            ToList()

        If rowsToDelete.Count = 0 Then
            MessageBox.Show(Me, "Select an item row first.", "Customer Master")
            Return
        End If

        Dim confirm = MessageBox.Show(Me, $"Delete {rowsToDelete.Count} item(s)? Click Save Items afterward to make it permanent.",
            "Customer Master", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
        If confirm <> DialogResult.Yes Then Return

        For Each row In rowsToDelete
            itemsGrid.Rows.Remove(row)
        Next
    End Sub

    ' ===================== Status bar =====================

    Private Function BuildStatusBar() As Label
        statusLabel = New Label With {.Dock = DockStyle.Bottom, .Height = 24, .Padding = New Padding(8, 4, 0, 0)}
        Return statusLabel
    End Function

End Class