''' <summary>
''' Shared markup formula constants for mail forward pricing, per Al - extracted into one
''' place so MailForwardsToEvoJob (charge posting) and MailForwardsReportJob (display
''' report) can't drift apart, and so either can be updated in one spot. Formula:
''' per-shipment marked-up cost = ROUND((Postage + MIN(Postage * MarkupPercentage,
''' MarkupCap)) / RoundingIncrement) * RoundingIncrement.
''' </summary>
Public Module MailForwardMarkup

    Public Const MarkupPercentage As Decimal = 0.2D ' 20% surcharge
    Public Const MarkupCap As Decimal = 3D ' surcharge never exceeds $3 per shipment
    Public Const RoundingIncrement As Decimal = 0.05D ' rounds each shipment's marked-up cost to the nearest nickel

End Module