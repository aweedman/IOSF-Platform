''' <summary>
''' Shape of the ListReservationsRequest response from HostedSuite's json/reply API.
''' Only the fields EarlyMeeting actually reads are mapped - extend if other jobs need more.
'''
''' NOTE: this endpoint returns a raw JSON array at the root, not an object wrapping a
''' "Data" property - deserialize directly as List(Of ReservationItem). Same fix as
''' CustomerXrefJob/EvoCustomerXrefModels - VBA-Web's "Response.Data" in the original
''' refers to the whole parsed body, not a nested field, and this API family (io.hostedsuite.com
''' json/reply) returns arrays directly rather than object-wrapped ones. This was caught via
''' testing on the ListClientNamesRequest endpoint; fixed here preemptively since it's the
''' same API family and hadn't been tested yet.
''' </summary>
Public Class ReservationItem
    Public Property MeetingRoomName As String
    Public Property StartTime As String ' comes back as a string like "2026-08-06 08:45:00"
End Class