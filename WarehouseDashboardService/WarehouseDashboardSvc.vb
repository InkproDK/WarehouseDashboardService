Imports System.Data.SqlClient
Imports System.Configuration
Public Class WarehouseDashboardSvc

    Dim objSQL01Conn As New SqlConnection(ConfigurationManager.ConnectionStrings("sql01conn").ConnectionString)
    Dim objWEB01Conn As New SqlConnection(ConfigurationManager.ConnectionStrings("web01conn").ConnectionString)

    Dim boolSnapshotTaken As Boolean = True
    Dim intAll_DK As Integer = 0
    Dim intAll_SE As Integer = 0
    Dim intAll_NO As Integer = 0
    Dim intCreatedToday_DK As Integer = 0
    Dim intCreatedToday_SE As Integer = 0
    Dim intCreatedToday_NO As Integer = 0
    Dim intSplitAll As Integer = 0
    Dim intSplitSentToday As Integer = 0
    Dim intReadyToPickAll As Integer = 0
    Dim intReadyToPickCreatedToday As Integer = 0
    Dim intPickedReadyToShip As Integer = 0
    Dim intShippedToday_DK As Integer = 0
    Dim intShippedToday_SE As Integer = 0
    Dim intShippedToday_NO As Integer = 0
    Dim intPickZone_Bestilling As Integer = 0
    Dim intPickZone_Kompatibel As Integer = 0
    Dim intPickZone_Office As Integer = 0
    Dim intPickZone_Original As Integer = 0
    Dim intPickZone_Toner As Integer = 0
    Dim dtSnapshotCreatedAt As Date = Now()
    Public Sub New()
        MyBase.New()
        InitializeComponent()
        EventLog1 = New EventLog
        If Not EventLog.SourceExists("WarehouseDashboardService") Then
            EventLog.CreateEventSource("WarehouseDashboardService", "InkproIT")
        End If
        EventLog1.Source = "WarehouseDashboardService"
        EventLog1.Log = "InkproIT"
    End Sub

    Protected Overrides Sub OnStart(ByVal args() As String)
        EventLog1.WriteEntry("Starting service...", EventLogEntryType.Information, 1)
        ' Add code here to start your service. This method should set things
        ' in motion so your service can do its work.
    End Sub

    Protected Overrides Sub OnStop()
        EventLog1.WriteEntry("Stopping service...", EventLogEntryType.Information, 2)
        Timer1.Stop()
        ' Add code here to perform any tear-down necessary to stop your service.
    End Sub

    Private Sub Timer1_Elapsed(sender As Object, e As Timers.ElapsedEventArgs) Handles Timer1.Elapsed

        If boolSnapshotTaken = True Then
            boolSnapshotTaken = False
            CreateSnapshot()
        End If

    End Sub

    Private Sub CreateSnapshot()

        GetOrdersAll()
        GetOrdersForToday()
        GetOrdersShippedToday()
        GetOrdersPickByZone()
        GetVariousOrderData()

        Try

            Dim SqlCmdInsertSnapshop As SqlCommand = objWEB01Conn.CreateCommand
            SqlCmdInsertSnapshop.CommandText = "INSERT INTO wh_snapshots
                                            (All_DK, ALL_SE, ALL_NO, CreatedToday_DK, CreatedToday_SE, CreatedToday_NO, SplitAll, SplitSentToday, ReadyToPickAll, ReadyToPickCreatedToday, 
                                            PickedReadyToShip, ShippedToday_DK, ShippedToday_SE, ShippedToday_NO, PickZone_Bestilling, PickZone_Kompatibel, PickZone_Office, 
                                            PickZone_Original, PickZone_Toner, SnapshotCreatedAt) VALUES
                                            (@All_DK, @ALL_SE, @ALL_NO, @CreatedToday_DK, @CreatedToday_SE, @CreatedToday_NO, @SplitAll, @SplitSentToday, @ReadyToPickAll,
                                            @ReadyToPickCreatedToday, @PickedReadyToShip, @ShippedToday_DK, @ShippedToday_SE, @ShippedToday_NO, @PickZone_Bestilling,
                                            @PickZone_Kompatibel, @PickZone_Office, @PickZone_Original, @PickZone_Toner, @SnapshotCreatedAt)"
            SqlCmdInsertSnapshop.Parameters.AddWithValue("@All_DK", intAll_DK)
            SqlCmdInsertSnapshop.Parameters.AddWithValue("@ALL_SE", intAll_SE)
            SqlCmdInsertSnapshop.Parameters.AddWithValue("@ALL_NO", intAll_NO)
            SqlCmdInsertSnapshop.Parameters.AddWithValue("@CreatedToday_DK", intCreatedToday_DK)
            SqlCmdInsertSnapshop.Parameters.AddWithValue("@CreatedToday_SE", intCreatedToday_SE)
            SqlCmdInsertSnapshop.Parameters.AddWithValue("@CreatedToday_NO", intCreatedToday_NO)
            SqlCmdInsertSnapshop.Parameters.AddWithValue("@SplitAll", intSplitAll)
            SqlCmdInsertSnapshop.Parameters.AddWithValue("@SplitSentToday", intSplitSentToday)
            SqlCmdInsertSnapshop.Parameters.AddWithValue("@ReadyToPickAll", intReadyToPickAll)
            SqlCmdInsertSnapshop.Parameters.AddWithValue("@ReadyToPickCreatedToday", intReadyToPickCreatedToday)
            SqlCmdInsertSnapshop.Parameters.AddWithValue("@PickedReadyToShip", intPickedReadyToShip)
            SqlCmdInsertSnapshop.Parameters.AddWithValue("@ShippedToday_DK", intShippedToday_DK)
            SqlCmdInsertSnapshop.Parameters.AddWithValue("@ShippedToday_SE", intShippedToday_SE)
            SqlCmdInsertSnapshop.Parameters.AddWithValue("@ShippedToday_NO", intShippedToday_NO)
            SqlCmdInsertSnapshop.Parameters.AddWithValue("@PickZone_Bestilling", intPickZone_Bestilling)
            SqlCmdInsertSnapshop.Parameters.AddWithValue("@PickZone_Kompatibel", intPickZone_Kompatibel)
            SqlCmdInsertSnapshop.Parameters.AddWithValue("@PickZone_Office", intPickZone_Office)
            SqlCmdInsertSnapshop.Parameters.AddWithValue("@PickZone_Original", intPickZone_Original)
            SqlCmdInsertSnapshop.Parameters.AddWithValue("@PickZone_Toner", intPickZone_Toner)
            SqlCmdInsertSnapshop.Parameters.AddWithValue("@SnapshotCreatedAt", dtSnapshotCreatedAt)

            If Not objSQL01Conn.State = ConnectionState.Open Then
                objWEB01Conn.Open()
            End If

            SqlCmdInsertSnapshop.ExecuteNonQuery()
            SqlCmdInsertSnapshop.Parameters.Clear()

            objWEB01Conn.Close()

        Catch ex As Exception
            EventLog1.WriteEntry("Exception in GetVariousOrderData():" & vbCrLf & ex.ToString(), EventLogEntryType.Error, 15)
        Finally
            boolSnapshotTaken = True
        End Try

    End Sub

    Protected Sub GetOrdersAll()

        Dim strSqlCmdGetOrdersAll As SqlCommand = objSQL01Conn.CreateCommand
        strSqlCmdGetOrdersAll.CommandText =
            "With group1 AS
                (SELECT COUNT(*) As AntalIalt, CASE WHEN [Currency Code] = '' THEN 'DKK' ELSE [Currency Code] END AS 'IaltPrLand'
                    FROM [BC18_InkPro].[dbo].[Inkpro$Sales Header$437dbf0e-84ff-417a-965d-ed2bb9650972] sh
                    WHERE [Document Type] = 1 AND [Sell-to Customer No_] <> '888'
                    GROUP BY [Currency Code]
                ),
            group2 AS
                (SELECT COUNT(DISTINCT sl.[Document No_]) AS AntalPrLand, CASE WHEN [Currency Code] = '' THEN 'DKK' ELSE [Currency Code] END AS 'MinusIaltPrLand'
                    FROM [BC18_InkPro].[dbo].[Inkpro$Sales Line$437dbf0e-84ff-417a-965d-ed2bb9650972] sl
	                    INNER JOIN(
		                    SELECT [Document No_], COUNT(DISTINCT [Document No_]) AS colCount
		                        FROM [BC18_InkPro].[dbo].[Inkpro$Sales Line$437dbf0e-84ff-417a-965d-ed2bb9650972]
		                        WHERE [Document Type] = 1 AND [Sell-to Customer No_] <> '888' AND NOT
			                    [Document No_] IN (SELECT [Document No_] FROM [BC18_InkPro].[dbo].[Inkpro$Sales Line$437dbf0e-84ff-417a-965d-ed2bb9650972] sl2
			                    WHERE Quantity = 0)
		                        GROUP BY [Document No_] HAVING SUM(Quantity) - SUM([Quantity Shipped]) = 0 AND NOT SUM([Quantity Shipped]) = 0) 
                                as cnt ON sl.[Document No_] = cnt.[Document No_]
                    GROUP BY [Currency Code])
		        SELECT (ISNULL([AntalIalt], 0) - ISNULL(AntalPrLand, 0)) AS 'Antal', COALESCE(group1.IaltPrLand, group2.MinusIaltPrLand) AS 'Land' FROM group1 FULL OUTER JOIN group2 ON group1.IaltPrLand = group2.MinusIaltPrLand
		        GROUP BY group1.IaltPrLand, group2.MinusIaltPrLand, AntalIalt, AntalPrLand"

        If Not objSQL01Conn.State = ConnectionState.Open Then
            objSQL01Conn.Open()
        End If

        Dim strSqlReader As SqlDataReader = strSqlCmdGetOrdersAll.ExecuteReader()
        Try

            Dim arrValues As List(Of String()) = New List(Of String())

            While strSqlReader.Read()



                arrValues.Add({strSqlReader("Antal"), strSqlReader("Land")})

            End While

            objSQL01Conn.Close()

            For Each orders_all In arrValues
                If orders_all(1) = "DKK" Then
                    intAll_DK = orders_all(0)
                End If
                If orders_all(1) = "SEK" Then
                    intAll_SE = orders_all(0)
                End If
                If orders_all(1) = "NOK" Then
                    intAll_NO = orders_all(0)
                End If
            Next

        Catch ex As Exception
            EventLog1.WriteEntry("Exception in GetOrdersAll():" & vbCrLf & ex.ToString(), EventLogEntryType.Error, 15)
        End Try

    End Sub

    Protected Sub GetOrdersForToday()

        Dim strSqlCmdGetOrdersForToday As SqlCommand = objSQL01Conn.CreateCommand
        strSqlCmdGetOrdersForToday.CommandText =
            "With group1 As
                (SELECT COUNT(DISTINCT No_) AS 'Antal', CASE WHEN [Currency Code] = '' THEN 'DKK' ELSE [Currency Code] END AS 'Land'
                FROM [BC18_InkPro].[dbo].[Inkpro$Sales Header$437dbf0e-84ff-417a-965d-ed2bb9650972] 
                    WHERE DATEADD(MI, DATEDIFF(MI, GETUTCDATE(), GETDATE()), [$systemCreatedAt]) >= CAST(GETDATE() AS DATE)
                    AND [Document Type] = 1 AND [Sell-to Customer No_] <> '888'
                GROUP BY [Currency Code]), 

            group2 As
                (SELECT COUNT(DISTINCT No_) AS 'Antal', CASE WHEN [Currency Code] = '' THEN 'DKK' ELSE [Currency Code] END AS 'Land'
                FROM [BC18_InkPro].[dbo].[Inkpro$Sales Invoice Header$437dbf0e-84ff-417a-965d-ed2bb9650972] 
                    WHERE DATEADD(MI, DATEDIFF(MI, GETUTCDATE(), GETDATE()), [Order Date]) >= CAST(GETDATE() AS DATE)
                    AND [Sell-to Customer No_] <> '888'
                    AND [Order No_] NOT IN (SELECT No_ FROM [BC18_InkPro].[dbo].[Inkpro$Sales Header$437dbf0e-84ff-417a-965d-ed2bb9650972])
                GROUP BY [Currency Code])

                SELECT SUM(ISNULL(group1.Antal, 0) + ISNULL(group2.Antal, 0)) As 'Antal', COALESCE(group1.Land, group2.Land) AS 'Land'
                FROM group1 FULL OUTER JOIN group2 ON group1.Land = group2.Land
                GROUP BY group1.Land, group2.Land"

        If Not objSQL01Conn.State = ConnectionState.Open Then
            objSQL01Conn.Open()
        End If

        Dim strSqlReader As SqlDataReader = strSqlCmdGetOrdersForToday.ExecuteReader()
        Try

            Dim arrValues As List(Of String()) = New List(Of String())

            While strSqlReader.Read()

                arrValues.Add({strSqlReader("Antal"), strSqlReader("Land")})

            End While

            For Each orders_all In arrValues
                If orders_all(1) = "DKK" Then
                    intCreatedToday_DK = orders_all(0)
                End If
                If orders_all(1) = "SEK" Then
                    intCreatedToday_SE = orders_all(0)
                End If
                If orders_all(1) = "NOK" Then
                    intCreatedToday_NO = orders_all(0)
                End If
            Next

        Catch ex As Exception
            EventLog1.WriteEntry("Exception in GetOrdersForToday():" & vbCrLf & ex.ToString(), EventLogEntryType.Error, 15)
        End Try

        objSQL01Conn.Close()

    End Sub

    Protected Sub GetOrdersShippedToday()

        Dim strSqlCmdGetOrdersShippedToday As SqlCommand = objSQL01Conn.CreateCommand
        strSqlCmdGetOrdersShippedToday.CommandText =
            "SELECT COUNT(DISTINCT slv.[Order No_]) AS 'Antal', CASE WHEN [Currency Code] = '' THEN 'DKK' ELSE [Currency Code] END AS 'Land'
                FROM [BC18_InkPro].[dbo].[Inkpro$Sales Shipment Line$437dbf0e-84ff-417a-965d-ed2bb9650972] slv
                LEFT JOIN [BC18_InkPro].[dbo].[Inkpro$Sales Shipment Header$437dbf0e-84ff-417a-965d-ed2bb9650972] so ON slv.[Order No_] = so.[Order No_]
                WHERE DATEADD(MI, DATEDIFF(MI, GETUTCDATE(), GETDATE()), slv.[$systemCreatedAt]) >= CAST(GETDATE() AS DATE)
                AND slv.[Sell-to Customer No_] <> '888'
                GROUP BY [Currency Code]"

        If Not objSQL01Conn.State = ConnectionState.Open Then
            objSQL01Conn.Open()
        End If

        Dim strSqlReader As SqlDataReader = strSqlCmdGetOrdersShippedToday.ExecuteReader()
        Try

            Dim arrValues As List(Of String()) = New List(Of String())

            While strSqlReader.Read()

                arrValues.Add({strSqlReader("Antal"), strSqlReader("Land")})

            End While

            For Each orders_all In arrValues
                If orders_all(1) = "DKK" Then
                    intShippedToday_DK = orders_all(0)
                End If
                If orders_all(1) = "SEK" Then
                    intShippedToday_SE = orders_all(0)
                End If
                If orders_all(1) = "NOK" Then
                    intShippedToday_NO = orders_all(0)
                End If
            Next

        Catch ex As Exception
            EventLog1.WriteEntry("Exception in GetOrdersShippedToday():" & vbCrLf & ex.ToString(), EventLogEntryType.Error, 15)
        End Try

        objSQL01Conn.Close()

    End Sub

    Protected Sub GetOrdersPickByZone()

        Dim strSqlCmdGetOrdersByZone As SqlCommand = objSQL01Conn.CreateCommand
        strSqlCmdGetOrdersByZone.CommandText =
            "SELECT COUNT(DISTINCT [Whse_ Document No_]) As 'Antal', ll2.[C2IT Assigned Zone] As 'Plukzone'
                FROM [BC18_InkPro].[dbo].[Inkpro$Warehouse Activity Line$437dbf0e-84ff-417a-965d-ed2bb9650972] ll 
                LEFT JOIN [BC18_InkPro].[dbo].[Inkpro$Warehouse Shipment Header$911dbd37-a000-4768-96cd-e49de6a45f4d] ll2 ON ll.[Whse_ Document No_] = ll2.No_
                WHERE [Activity Type] = 2 AND ll.[Zone Code] NOT IN ('LEVER', 'MONTAGE') AND ll.[Destination No_] <> '888'
                GROUP BY ll2.[C2IT Assigned Zone]"

        If Not objSQL01Conn.State = ConnectionState.Open Then
            objSQL01Conn.Open()
        End If

        Dim strSqlReader As SqlDataReader = strSqlCmdGetOrdersByZone.ExecuteReader()
        Try

            Dim arrValues As List(Of String()) = New List(Of String())

            While strSqlReader.Read()

                arrValues.Add({strSqlReader("Antal"), strSqlReader("Plukzone")})

            End While

            For Each orders_all In arrValues
                If orders_all(1) = "BESTILLING" Then
                    intPickZone_Bestilling = orders_all(0)
                End If
                If orders_all(1) = "KOMPATIBEL" Then
                    intPickZone_Kompatibel = orders_all(0)
                End If
                If orders_all(1) = "OFFICE" Then
                    intPickZone_Office = orders_all(0)
                End If
                If orders_all(1) = "ORIGINAL" Then
                    intPickZone_Original = orders_all(0)
                End If
                If orders_all(1) = "TONER" Then
                    intPickZone_Toner = orders_all(0)
                End If
            Next

        Catch ex As Exception
            EventLog1.WriteEntry("Exception in GetOrdersPickByZone():" & vbCrLf & ex.ToString(), EventLogEntryType.Error, 15)
        End Try

        objSQL01Conn.Close()

    End Sub

    Protected Sub GetVariousOrderData()

        Dim strSqlCmdVariousOrderData As SqlCommand = objSQL01Conn.CreateCommand
        strSqlCmdVariousOrderData.CommandText =
            "SELECT 

                (SELECT COUNT(DISTINCT sl.[Document No_])
                FROM [BC18_InkPro].[dbo].[Inkpro$Sales Line$437dbf0e-84ff-417a-965d-ed2bb9650972] sl
	                INNER JOIN(
		                SELECT [Document No_], COUNT(DISTINCT [Document No_]) AS colCount
		                FROM [BC18_InkPro].[dbo].[Inkpro$Sales Line$437dbf0e-84ff-417a-965d-ed2bb9650972]
		                WHERE [Document Type] = 1 AND [Sell-to Customer No_] <> '888'
		            GROUP BY [Document No_] HAVING SUM(Quantity) - SUM([Quantity Shipped]) > 0 AND NOT SUM([Quantity Shipped]) = 0) as cnt ON sl.[Document No_] = cnt.[Document No_]
                ) As 'DelordreriAlt',

                (SELECT COUNT(DISTINCT [Order No_]) FROM [BC18_InkPro].[dbo].[Inkpro$Sales Shipment Line$437dbf0e-84ff-417a-965d-ed2bb9650972]
                WHERE DATEADD(MI, DATEDIFF(MI, GETUTCDATE(), GETDATE()), [$systemCreatedAt]) >= CAST(GETDATE() AS DATE) AND Quantity = 0
                AND [Sell-to Customer No_] <> '888' AND [Location Code] = 'LAGER'
                AND [Order No_] NOT IN 
                    (SELECT [Document No_] AS colCount
		            FROM [BC18_InkPro].[dbo].[Inkpro$Sales Line$437dbf0e-84ff-417a-965d-ed2bb9650972]
		            WHERE [Document Type] = 1 AND [Sell-to Customer No_] <> '888' AND NOT
			        [Document No_] IN (SELECT [Document No_] FROM [BC18_InkPro].[dbo].[Inkpro$Sales Line$437dbf0e-84ff-417a-965d-ed2bb9650972] sl2
			    WHERE Quantity = 0)
    		    GROUP BY [Document No_] HAVING SUM(Quantity) - SUM([Quantity Shipped]) = 0 AND NOT SUM([Quantity Shipped]) = 0)
                AND [Order No_] IN
		            (SELECT [Document No_] AS colCount
		            FROM [BC18_InkPro].[dbo].[Inkpro$Sales Line$437dbf0e-84ff-417a-965d-ed2bb9650972])
                ) As 'DelordrerSendtiDag',

                (SELECT COUNT(DISTINCT [Whse_ Document No_]) FROM [BC18_InkPro].[dbo].[Inkpro$Warehouse Activity Line$437dbf0e-84ff-417a-965d-ed2bb9650972]
                WHERE [Activity Type] = 2 AND [Destination No_] <> '888'
                ) As 'OrdrerTilPlukialt',

                (SELECT COUNT(DISTINCT [Whse_ Document No_]) FROM [BC18_InkPro].[dbo].[InkproIT_LogWarehouseActivityLine]
                WHERE DATEADD(MI, DATEDIFF(MI, GETUTCDATE(), GETDATE()), [$systemCreatedAt]) >= CAST(GETDATE() AS DATE)
                AND ([Whse_ Document No_] IN 
                    (SELECT [Whse_ Document No_] FROM [BC18_InkPro].[dbo].[Inkpro$Warehouse Activity Line$437dbf0e-84ff-417a-965d-ed2bb9650972]
                    WHERE [Activity Type] = 2 AND [Destination No_] <> '888' AND 
                    DATEADD(MI, DATEDIFF(MI, GETUTCDATE(), GETDATE()), [$systemCreatedAt]) >= CAST(GETDATE() AS DATE))
                    OR [Whse_ Document No_] IN (SELECT [Whse_ Document No_] FROM [BC18_InkPro].[dbo].[Inkpro$Registered Whse_ Activity Line$437dbf0e-84ff-417a-965d-ed2bb9650972]
                    WHERE [Activity Type] = 2 AND [Destination No_] <> '888' AND 
                    DATEADD(MI, DATEDIFF(MI, GETUTCDATE(), GETDATE()), [$systemCreatedAt]) >= CAST(GETDATE() AS DATE)
                ))) As 'OrdrerSendtTilPlukidag',

                (SELECT COUNT(DISTINCT No_)
                FROM [BC18_InkPro].[dbo].[Inkpro$Warehouse Shipment Line$437dbf0e-84ff-417a-965d-ed2bb9650972] ll
                    LEFT JOIN [BC18_InkPro].[dbo].[Inkpro$MOB WMS Registration$a5727ce6-368c-49e2-84cb-1a6052f0551c] mob
                    ON ll.No_ = mob.[Whse_ Document No_] AND ll.[Line No_] = mob.[Whse_ Document Line No_]
                WHERE [Whse_ Document No_] <> ''
                ) As 'PlukketKlarTilPak',

                (SELECT GETDATE()
                ) As 'SnapshotCreatedAt'"

        If Not objSQL01Conn.State = ConnectionState.Open Then
            objSQL01Conn.Open()
        End If

        Dim strSqlReader As SqlDataReader = strSqlCmdVariousOrderData.ExecuteReader()
        Try

            Dim arrValues As List(Of String()) = New List(Of String())

            While strSqlReader.Read()

                arrValues.Add({strSqlReader("DelordreriAlt"), strSqlReader("DelordrerSendtiDag"), strSqlReader("OrdrerTilPlukialt"),
                              strSqlReader("OrdrerSendtTilPlukidag"), strSqlReader("PlukketKlarTilPak"), strSqlReader("SnapshotCreatedAt")})

            End While

            For Each orders_all In arrValues

                intSplitAll = orders_all(0)
                intSplitSentToday = orders_all(1)
                intReadyToPickAll = orders_all(2)
                intReadyToPickCreatedToday = orders_all(3)
                intPickedReadyToShip = orders_all(4)
                dtSnapshotCreatedAt = orders_all(5)

            Next

        Catch ex As Exception
            EventLog1.WriteEntry("Exception in GetVariousOrderData():" & vbCrLf & ex.ToString(), EventLogEntryType.Error, 15)
        End Try

        objSQL01Conn.Close()

    End Sub

End Class
