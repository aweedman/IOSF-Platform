Imports Microsoft.Data.SqlClient

''' <summary>
''' Direct port of Landing Page.cls: Spheremail_Customers().
''' Pulls active SphereMail customers and re-populates the Spheremail_Customers table.
'''
''' Changes from the VBA original:
'''  - The "For i = 1 To 1" loop with an unreachable Else branch for a "Burlingame"
'''    location has been removed - that branch never executed (loop only ever ran i=1),
'''    so this is dead code from an apparently-abandoned second-location rollout. Only
'''    the live "San Francisco" path is ported. Let me know if Burlingame needs reviving.
'''  - "location" is now an explicit constant instead of an implicit global variable that
'''    Spheremail_Storage() silently depended on being set here first.
'''  - MsgBox("API Call Error...") replaced with a thrown exception - a blocking dialog
'''    is wrong for anything that might run headless via Task Scheduler; let the caller
'''    (or EmailHelper.EmailError, once wired) decide how to surface the failure.
'''  - INSERT is now parameterized (original string-concatenated full_name/email
'''    unescaped - a customer name containing an apostrophe would have broken the query).
''' </summary>
Public Module SphereMailCustomersJob

    Private Const AdminBaseUrl As String = "https://api.spheremail.co/v1/admin"
    Private Const Location As String = "San Francisco" ' see remarks above re: dead Burlingame branch

    Public Async Function RunAsync() As Task
        DeleteExistingRows()

        Dim token = Await SphereMailAuth.GetTokenAsync()
        Dim headers = New Dictionary(Of String, String) From {{"Authorization", token}}
        Dim queryParams = New Dictionary(Of String, String) From {
            {"limit", "1000"},
            {"is_active", "true"}
        }

        Dim response = Await ApiClient.GetAsync($"{AdminBaseUrl}/customers", queryParams, headers, timeoutSeconds:=60)
        response.EnsureSuccess()

        Dim data = response.DataAs(Of SphereMailCustomersResponse)()

        For Each customer In data.Customers
            Dim email = StripPlusAddressing(customer.Email)
            InsertCustomer(customer.Pmb, Location, customer.FullName, email)
        Next
    End Function

    ''' <summary>
    ''' Strips Gmail-style "+tag" addressing: "user+tag@domain.com" -> "user@domain.com".
    ''' </summary>
    Private Function StripPlusAddressing(email As String) As String
        Dim plusPos = email.IndexOf("+"c)
        Dim atPos = email.IndexOf("@"c)
        If plusPos >= 0 AndAlso atPos > plusPos Then
            Return email.Substring(0, plusPos) & email.Substring(atPos)
        End If
        Return email
    End Function

    Private Sub DeleteExistingRows()
        Using conn As New SqlConnection(ConfigHelper.ConnectionString)
            Using cmd As New SqlCommand("DELETE FROM Spheremail_Customers", conn)
                conn.Open()
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Private Sub InsertCustomer(mailBox As String, location As String, fullName As String, email As String)
        Const sql As String =
            "INSERT INTO Spheremail_Customers (Mail_Box, Location, Full_Name, Email) " &
            "VALUES (@MailBox, @Location, @FullName, @Email)"

        Using conn As New SqlConnection(ConfigHelper.ConnectionString)
            Using cmd As New SqlCommand(sql, conn)
                cmd.Parameters.AddWithValue("@MailBox", mailBox)
                cmd.Parameters.AddWithValue("@Location", location)
                cmd.Parameters.AddWithValue("@FullName", fullName)
                cmd.Parameters.AddWithValue("@Email", email)
                conn.Open()
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

End Module