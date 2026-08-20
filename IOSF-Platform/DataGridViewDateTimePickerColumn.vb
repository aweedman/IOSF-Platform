Imports System.Windows.Forms
Imports System.ComponentModel

''' <summary>
''' DataGridView has a built-in DataGridViewComboBoxColumn but no equivalent for a date
''' picker - this is the standard, Microsoft-documented pattern for hosting a custom
''' control (IDataGridViewEditingControl) as a cell's editor, applied here specifically
''' for a DateTimePicker. Used for Customer_Ops_Item.Terminated_Cont per Al's request.
'''
''' NOT COMPILE-TESTED (same caveat as other new files in this port that need a NuGet
''' package or specialized API this sandbox can't verify) - this is a well-established,
''' standard pattern, but written from knowledge of it rather than verified against a
''' real build.
''' </summary>
Public Class DataGridViewDateTimePickerColumn
    Inherits DataGridViewColumn

    Public Sub New()
        MyBase.New(New DataGridViewDateTimePickerCell())
    End Sub

    Public Overrides Property CellTemplate As DataGridViewCell
        Get
            Return MyBase.CellTemplate
        End Get
        Set(value As DataGridViewCell)
            If value IsNot Nothing AndAlso Not TypeOf value Is DataGridViewDateTimePickerCell Then
                Throw New InvalidCastException("Must be a DataGridViewDateTimePickerCell")
            End If
            MyBase.CellTemplate = value
        End Set
    End Property
End Class

Public Class DataGridViewDateTimePickerCell
    Inherits DataGridViewTextBoxCell

    Public Sub New()
        Style.Format = "d" ' short date, no time
    End Sub

    Public Overrides Sub InitializeEditingControl(rowIndex As Integer, initialFormattedValue As Object, dataGridViewCellStyle As DataGridViewCellStyle)
        MyBase.InitializeEditingControl(rowIndex, initialFormattedValue, dataGridViewCellStyle)
        Dim ctl = TryCast(DataGridView.EditingControl, DateTimePickerEditingControl)
        If ctl IsNot Nothing Then
            ctl.Value = If(Value Is Nothing OrElse Value Is DBNull.Value, DateTime.Today, Convert.ToDateTime(Value))
        End If
    End Sub

    Public Overrides ReadOnly Property EditType As Type
        Get
            Return GetType(DateTimePickerEditingControl)
        End Get
    End Property

    Public Overrides ReadOnly Property ValueType As Type
        Get
            Return GetType(DateTime)
        End Get
    End Property

    ''' <summary>
    ''' REAL RISK FIXED per Al: this previously defaulted new rows to DateTime.Today,
    ''' which pre-filled every new row's cell with today's date before the user had
    ''' touched it at all - a real danger for Terminated_Cont specifically, where an
    ''' unnoticed default could get a brand-new contact saved as already terminated. New
    ''' rows now start genuinely blank (DBNull), consistent with this column existing
    ''' specifically to support clearable/nullable dates (see the "Clear Date" context
    ''' menu in CustomerMasterForm).
    ''' </summary>
    Public Overrides ReadOnly Property DefaultNewRowValue As Object
        Get
            Return DBNull.Value
        End Get
    End Property
End Class

Public Class DateTimePickerEditingControl
    Inherits DateTimePicker
    Implements IDataGridViewEditingControl

    Private dataGridViewControl As DataGridView
    Private valueIsChanged As Boolean
    Private rowIndexNum As Integer

    Public Sub New()
        Format = DateTimePickerFormat.Short
    End Sub

    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Property EditingControlFormattedValue As Object Implements IDataGridViewEditingControl.EditingControlFormattedValue
        Get
            Return Value.ToShortDateString()
        End Get
        Set(value As Object)
            If TypeOf value Is String Then
                Dim parsed As DateTime
                If DateTime.TryParse(CStr(value), parsed) Then Value = parsed
            End If
        End Set
    End Property

    Public Function GetEditingControlFormattedValue(context As DataGridViewDataErrorContexts) As Object Implements IDataGridViewEditingControl.GetEditingControlFormattedValue
        Return EditingControlFormattedValue
    End Function

    Public Sub ApplyCellStyleToEditingControl(dataGridViewCellStyle As DataGridViewCellStyle) Implements IDataGridViewEditingControl.ApplyCellStyleToEditingControl
        Font = dataGridViewCellStyle.Font
    End Sub

    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Property EditingControlRowIndex As Integer Implements IDataGridViewEditingControl.EditingControlRowIndex
        Get
            Return rowIndexNum
        End Get
        Set(value As Integer)
            rowIndexNum = value
        End Set
    End Property

    Public Function EditingControlWantsInputKey(keyData As Keys, dataGridViewWantsInputKey As Boolean) As Boolean Implements IDataGridViewEditingControl.EditingControlWantsInputKey
        Return True
    End Function

    Public Sub PrepareEditingControlForEdit(selectAll As Boolean) Implements IDataGridViewEditingControl.PrepareEditingControlForEdit
    End Sub

    Private repositionOnValueChange As Boolean

    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Property RepositionEditingControlOnValueChange As Boolean Implements IDataGridViewEditingControl.RepositionEditingControlOnValueChange
        Get
            Return repositionOnValueChange
        End Get
        Set(value As Boolean)
            repositionOnValueChange = value
        End Set
    End Property

    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Property EditingControlDataGridView As DataGridView Implements IDataGridViewEditingControl.EditingControlDataGridView
        Get
            Return dataGridViewControl
        End Get
        Set(value As DataGridView)
            dataGridViewControl = value
        End Set
    End Property

    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Property EditingControlValueChanged As Boolean Implements IDataGridViewEditingControl.EditingControlValueChanged
        Get
            Return valueIsChanged
        End Get
        Set(value As Boolean)
            valueIsChanged = value
        End Set
    End Property

    Public ReadOnly Property EditingPanelCursor As Cursor Implements IDataGridViewEditingControl.EditingPanelCursor
        Get
            Return MyBase.Cursor
        End Get
    End Property

    Protected Overrides Sub OnValueChanged(eventargs As EventArgs)
        valueIsChanged = True
        dataGridViewControl?.NotifyCurrentCellDirty(True)
        MyBase.OnValueChanged(eventargs)
    End Sub

End Class