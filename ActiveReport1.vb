Imports System
Imports DataDynamics.ActiveReports
Imports DataDynamics.ActiveReports.Document

Public Class ActiveReport1
Inherits ActiveReport
	Public Sub New()
	MyBase.New()
		InitializeReport()
	End Sub
	#Region "ActiveReports Designer generated code"
    Private WithEvents PageHeader As DataDynamics.ActiveReports.PageHeader = Nothing
    Private WithEvents Detail As DataDynamics.ActiveReports.Detail = Nothing
    Private WithEvents PageFooter As DataDynamics.ActiveReports.PageFooter = Nothing
	Private Label As DataDynamics.ActiveReports.Label = Nothing
	Private Label1 As DataDynamics.ActiveReports.Label = Nothing
	Private Label2 As DataDynamics.ActiveReports.Label = Nothing
	Private Label3 As DataDynamics.ActiveReports.Label = Nothing
	Private Label4 As DataDynamics.ActiveReports.Label = Nothing
	Private Label5 As DataDynamics.ActiveReports.Label = Nothing
	Private Label6 As DataDynamics.ActiveReports.Label = Nothing
	Private Label7 As DataDynamics.ActiveReports.Label = Nothing
	Private Label8 As DataDynamics.ActiveReports.Label = Nothing
	Private Label9 As DataDynamics.ActiveReports.Label = Nothing
	Private Label10 As DataDynamics.ActiveReports.Label = Nothing
	Private Label11 As DataDynamics.ActiveReports.Label = Nothing
	Private Label12 As DataDynamics.ActiveReports.Label = Nothing
	Private Label13 As DataDynamics.ActiveReports.Label = Nothing
	Private Label14 As DataDynamics.ActiveReports.Label = Nothing
	Private Label15 As DataDynamics.ActiveReports.Label = Nothing
	Private Label16 As DataDynamics.ActiveReports.Label = Nothing
	Private Label17 As DataDynamics.ActiveReports.Label = Nothing
	Private Label18 As DataDynamics.ActiveReports.Label = Nothing
	Private Label19 As DataDynamics.ActiveReports.Label = Nothing
	Private Label20 As DataDynamics.ActiveReports.Label = Nothing
	Private Label21 As DataDynamics.ActiveReports.Label = Nothing
	Private Label22 As DataDynamics.ActiveReports.Label = Nothing
	Private Line As DataDynamics.ActiveReports.Line = Nothing
	Private Line1 As DataDynamics.ActiveReports.Line = Nothing
	Private Line2 As DataDynamics.ActiveReports.Line = Nothing
	Private Line3 As DataDynamics.ActiveReports.Line = Nothing
	Private Line4 As DataDynamics.ActiveReports.Line = Nothing
	Private Line5 As DataDynamics.ActiveReports.Line = Nothing
	Private Line6 As DataDynamics.ActiveReports.Line = Nothing
	Private Line7 As DataDynamics.ActiveReports.Line = Nothing
	Private Line8 As DataDynamics.ActiveReports.Line = Nothing
	Private Line9 As DataDynamics.ActiveReports.Line = Nothing
	Private Line10 As DataDynamics.ActiveReports.Line = Nothing
	Private Line11 As DataDynamics.ActiveReports.Line = Nothing
	Private Line27 As DataDynamics.ActiveReports.Line = Nothing
	Private Line28 As DataDynamics.ActiveReports.Line = Nothing
	Private Line29 As DataDynamics.ActiveReports.Line = Nothing
	Private Line35 As DataDynamics.ActiveReports.Line = Nothing
	Private TextBox14 As DataDynamics.ActiveReports.TextBox = Nothing
	Private Label23 As DataDynamics.ActiveReports.Label = Nothing
	Private TextBox As DataDynamics.ActiveReports.TextBox = Nothing
	Private TextBox1 As DataDynamics.ActiveReports.TextBox = Nothing
	Private TextBox2 As DataDynamics.ActiveReports.TextBox = Nothing
	Private TextBox3 As DataDynamics.ActiveReports.TextBox = Nothing
	Private TextBox4 As DataDynamics.ActiveReports.TextBox = Nothing
	Private TextBox5 As DataDynamics.ActiveReports.TextBox = Nothing
	Private TextBox6 As DataDynamics.ActiveReports.TextBox = Nothing
	Private TextBox7 As DataDynamics.ActiveReports.TextBox = Nothing
	Private TextBox8 As DataDynamics.ActiveReports.TextBox = Nothing
	Private TextBox9 As DataDynamics.ActiveReports.TextBox = Nothing
	Private TextBox10 As DataDynamics.ActiveReports.TextBox = Nothing
	Private TextBox11 As DataDynamics.ActiveReports.TextBox = Nothing
	Private TextBox12 As DataDynamics.ActiveReports.TextBox = Nothing
	Private TextBox13 As DataDynamics.ActiveReports.TextBox = Nothing
	Private TextBox15 As DataDynamics.ActiveReports.TextBox = Nothing
	Private TextBox17 As DataDynamics.ActiveReports.TextBox = Nothing
	Private TextBox18 As DataDynamics.ActiveReports.TextBox = Nothing
	Private TextBox19 As DataDynamics.ActiveReports.TextBox = Nothing
	Private TextBox20 As DataDynamics.ActiveReports.TextBox = Nothing
	Private TextBox21 As DataDynamics.ActiveReports.TextBox = Nothing
	Private TextBox23 As DataDynamics.ActiveReports.TextBox = Nothing
	Private TextBox24 As DataDynamics.ActiveReports.TextBox = Nothing
	Private TextBox26 As DataDynamics.ActiveReports.TextBox = Nothing
	Private TextBox27 As DataDynamics.ActiveReports.TextBox = Nothing
	Private TextBox28 As DataDynamics.ActiveReports.TextBox = Nothing
	Private Line12 As DataDynamics.ActiveReports.Line = Nothing
	Private Line13 As DataDynamics.ActiveReports.Line = Nothing
	Private Line14 As DataDynamics.ActiveReports.Line = Nothing
	Private Line16 As DataDynamics.ActiveReports.Line = Nothing
	Private Line17 As DataDynamics.ActiveReports.Line = Nothing
	Private Line18 As DataDynamics.ActiveReports.Line = Nothing
	Private Line19 As DataDynamics.ActiveReports.Line = Nothing
	Private Line20 As DataDynamics.ActiveReports.Line = Nothing
	Private Line22 As DataDynamics.ActiveReports.Line = Nothing
	Private Line23 As DataDynamics.ActiveReports.Line = Nothing
	Private Line25 As DataDynamics.ActiveReports.Line = Nothing
	Private Line30 As DataDynamics.ActiveReports.Line = Nothing
	Private Line31 As DataDynamics.ActiveReports.Line = Nothing
	Private Line32 As DataDynamics.ActiveReports.Line = Nothing
	Private Line33 As DataDynamics.ActiveReports.Line = Nothing
	Private Line34 As DataDynamics.ActiveReports.Line = Nothing
	Public Sub InitializeReport()
		Me.LoadLayout(Me.GetType, "BIC_Import_bat.ActiveReport1.rpx")
		Me.PageHeader = CType(Me.Sections("PageHeader"),DataDynamics.ActiveReports.PageHeader)
		Me.Detail = CType(Me.Sections("Detail"),DataDynamics.ActiveReports.Detail)
		Me.PageFooter = CType(Me.Sections("PageFooter"),DataDynamics.ActiveReports.PageFooter)
		Me.Label = CType(Me.PageHeader.Controls(0),DataDynamics.ActiveReports.Label)
		Me.Label1 = CType(Me.PageHeader.Controls(1),DataDynamics.ActiveReports.Label)
		Me.Label2 = CType(Me.PageHeader.Controls(2),DataDynamics.ActiveReports.Label)
		Me.Label3 = CType(Me.PageHeader.Controls(3),DataDynamics.ActiveReports.Label)
		Me.Label4 = CType(Me.PageHeader.Controls(4),DataDynamics.ActiveReports.Label)
		Me.Label5 = CType(Me.PageHeader.Controls(5),DataDynamics.ActiveReports.Label)
		Me.Label6 = CType(Me.PageHeader.Controls(6),DataDynamics.ActiveReports.Label)
		Me.Label7 = CType(Me.PageHeader.Controls(7),DataDynamics.ActiveReports.Label)
		Me.Label8 = CType(Me.PageHeader.Controls(8),DataDynamics.ActiveReports.Label)
		Me.Label9 = CType(Me.PageHeader.Controls(9),DataDynamics.ActiveReports.Label)
		Me.Label10 = CType(Me.PageHeader.Controls(10),DataDynamics.ActiveReports.Label)
		Me.Label11 = CType(Me.PageHeader.Controls(11),DataDynamics.ActiveReports.Label)
		Me.Label12 = CType(Me.PageHeader.Controls(12),DataDynamics.ActiveReports.Label)
		Me.Label13 = CType(Me.PageHeader.Controls(13),DataDynamics.ActiveReports.Label)
		Me.Label14 = CType(Me.PageHeader.Controls(14),DataDynamics.ActiveReports.Label)
		Me.Label15 = CType(Me.PageHeader.Controls(15),DataDynamics.ActiveReports.Label)
		Me.Label16 = CType(Me.PageHeader.Controls(16),DataDynamics.ActiveReports.Label)
		Me.Label17 = CType(Me.PageHeader.Controls(17),DataDynamics.ActiveReports.Label)
		Me.Label18 = CType(Me.PageHeader.Controls(18),DataDynamics.ActiveReports.Label)
		Me.Label19 = CType(Me.PageHeader.Controls(19),DataDynamics.ActiveReports.Label)
		Me.Label20 = CType(Me.PageHeader.Controls(20),DataDynamics.ActiveReports.Label)
		Me.Label21 = CType(Me.PageHeader.Controls(21),DataDynamics.ActiveReports.Label)
		Me.Label22 = CType(Me.PageHeader.Controls(22),DataDynamics.ActiveReports.Label)
		Me.Line = CType(Me.PageHeader.Controls(23),DataDynamics.ActiveReports.Line)
		Me.Line1 = CType(Me.PageHeader.Controls(24),DataDynamics.ActiveReports.Line)
		Me.Line2 = CType(Me.PageHeader.Controls(25),DataDynamics.ActiveReports.Line)
		Me.Line3 = CType(Me.PageHeader.Controls(26),DataDynamics.ActiveReports.Line)
		Me.Line4 = CType(Me.PageHeader.Controls(27),DataDynamics.ActiveReports.Line)
		Me.Line5 = CType(Me.PageHeader.Controls(28),DataDynamics.ActiveReports.Line)
		Me.Line6 = CType(Me.PageHeader.Controls(29),DataDynamics.ActiveReports.Line)
		Me.Line7 = CType(Me.PageHeader.Controls(30),DataDynamics.ActiveReports.Line)
		Me.Line8 = CType(Me.PageHeader.Controls(31),DataDynamics.ActiveReports.Line)
		Me.Line9 = CType(Me.PageHeader.Controls(32),DataDynamics.ActiveReports.Line)
		Me.Line10 = CType(Me.PageHeader.Controls(33),DataDynamics.ActiveReports.Line)
		Me.Line11 = CType(Me.PageHeader.Controls(34),DataDynamics.ActiveReports.Line)
		Me.Line27 = CType(Me.PageHeader.Controls(35),DataDynamics.ActiveReports.Line)
		Me.Line28 = CType(Me.PageHeader.Controls(36),DataDynamics.ActiveReports.Line)
		Me.Line29 = CType(Me.PageHeader.Controls(37),DataDynamics.ActiveReports.Line)
		Me.Line35 = CType(Me.PageHeader.Controls(38),DataDynamics.ActiveReports.Line)
		Me.TextBox14 = CType(Me.PageHeader.Controls(39),DataDynamics.ActiveReports.TextBox)
		Me.Label23 = CType(Me.PageHeader.Controls(40),DataDynamics.ActiveReports.Label)
		Me.TextBox = CType(Me.Detail.Controls(0),DataDynamics.ActiveReports.TextBox)
		Me.TextBox1 = CType(Me.Detail.Controls(1),DataDynamics.ActiveReports.TextBox)
		Me.TextBox2 = CType(Me.Detail.Controls(2),DataDynamics.ActiveReports.TextBox)
		Me.TextBox3 = CType(Me.Detail.Controls(3),DataDynamics.ActiveReports.TextBox)
		Me.TextBox4 = CType(Me.Detail.Controls(4),DataDynamics.ActiveReports.TextBox)
		Me.TextBox5 = CType(Me.Detail.Controls(5),DataDynamics.ActiveReports.TextBox)
		Me.TextBox6 = CType(Me.Detail.Controls(6),DataDynamics.ActiveReports.TextBox)
		Me.TextBox7 = CType(Me.Detail.Controls(7),DataDynamics.ActiveReports.TextBox)
		Me.TextBox8 = CType(Me.Detail.Controls(8),DataDynamics.ActiveReports.TextBox)
		Me.TextBox9 = CType(Me.Detail.Controls(9),DataDynamics.ActiveReports.TextBox)
		Me.TextBox10 = CType(Me.Detail.Controls(10),DataDynamics.ActiveReports.TextBox)
		Me.TextBox11 = CType(Me.Detail.Controls(11),DataDynamics.ActiveReports.TextBox)
		Me.TextBox12 = CType(Me.Detail.Controls(12),DataDynamics.ActiveReports.TextBox)
		Me.TextBox13 = CType(Me.Detail.Controls(13),DataDynamics.ActiveReports.TextBox)
		Me.TextBox15 = CType(Me.Detail.Controls(14),DataDynamics.ActiveReports.TextBox)
		Me.TextBox17 = CType(Me.Detail.Controls(15),DataDynamics.ActiveReports.TextBox)
		Me.TextBox18 = CType(Me.Detail.Controls(16),DataDynamics.ActiveReports.TextBox)
		Me.TextBox19 = CType(Me.Detail.Controls(17),DataDynamics.ActiveReports.TextBox)
		Me.TextBox20 = CType(Me.Detail.Controls(18),DataDynamics.ActiveReports.TextBox)
		Me.TextBox21 = CType(Me.Detail.Controls(19),DataDynamics.ActiveReports.TextBox)
		Me.TextBox23 = CType(Me.Detail.Controls(20),DataDynamics.ActiveReports.TextBox)
		Me.TextBox24 = CType(Me.Detail.Controls(21),DataDynamics.ActiveReports.TextBox)
		Me.TextBox26 = CType(Me.Detail.Controls(22),DataDynamics.ActiveReports.TextBox)
		Me.TextBox27 = CType(Me.Detail.Controls(23),DataDynamics.ActiveReports.TextBox)
		Me.TextBox28 = CType(Me.Detail.Controls(24),DataDynamics.ActiveReports.TextBox)
		Me.Line12 = CType(Me.Detail.Controls(25),DataDynamics.ActiveReports.Line)
		Me.Line13 = CType(Me.Detail.Controls(26),DataDynamics.ActiveReports.Line)
		Me.Line14 = CType(Me.Detail.Controls(27),DataDynamics.ActiveReports.Line)
		Me.Line16 = CType(Me.Detail.Controls(28),DataDynamics.ActiveReports.Line)
		Me.Line17 = CType(Me.Detail.Controls(29),DataDynamics.ActiveReports.Line)
		Me.Line18 = CType(Me.Detail.Controls(30),DataDynamics.ActiveReports.Line)
		Me.Line19 = CType(Me.Detail.Controls(31),DataDynamics.ActiveReports.Line)
		Me.Line20 = CType(Me.Detail.Controls(32),DataDynamics.ActiveReports.Line)
		Me.Line22 = CType(Me.Detail.Controls(33),DataDynamics.ActiveReports.Line)
		Me.Line23 = CType(Me.Detail.Controls(34),DataDynamics.ActiveReports.Line)
		Me.Line25 = CType(Me.Detail.Controls(35),DataDynamics.ActiveReports.Line)
		Me.Line30 = CType(Me.Detail.Controls(36),DataDynamics.ActiveReports.Line)
		Me.Line31 = CType(Me.Detail.Controls(37),DataDynamics.ActiveReports.Line)
		Me.Line32 = CType(Me.Detail.Controls(38),DataDynamics.ActiveReports.Line)
		Me.Line33 = CType(Me.Detail.Controls(39),DataDynamics.ActiveReports.Line)
		Me.Line34 = CType(Me.Detail.Controls(40),DataDynamics.ActiveReports.Line)
	End Sub

	#End Region
End Class
