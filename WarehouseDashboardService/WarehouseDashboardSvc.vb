Imports System.Data.SqlClient
Imports System.Configuration
Public Class WarehouseDashboardSvc

    Dim objSQL01Conn As New SqlConnection(ConfigurationManager.ConnectionStrings("sql01conn").ConnectionString)
    Dim objWEB01Conn As New SqlConnection(ConfigurationManager.ConnectionStrings("web01conn").ConnectionString)

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
            EventLog.CreateEventSource("WarehouseDashboardService", "Inkpro IT")
        End If
        EventLog1.Source = "WarehouseDashboardService"
        EventLog1.Log = "Inkpro IT"
    End Sub

    Protected Overrides Sub OnStart(ByVal args() As String)
        EventLog1.WriteEntry("Starting service...", EventLogEntryType.Information, 1)
        Timer1.Start()
    End Sub

    Protected Overrides Sub OnStop()
        EventLog1.WriteEntry("Stopping service...", EventLogEntryType.Information, 2)
        Timer1.Stop()
    End Sub

    Private Sub Timer1_Elapsed(sender As Object, e As Timers.ElapsedEventArgs) Handles Timer1.Elapsed

        CreateSnapshot()

    End Sub

    Private Sub CreateSnapshot()

        GetOrdersAll()
        GetOrdersForToday()
        GetOrdersShippedToday()
        GetOrdersByPickZone()
        GetVariousOrderData()

        Try

            Dim SqlCmdInsertSnapshop As SqlCommand = objWEB01Conn.CreateCommand
            SqlCmdInsertSnapshop.CommandText = "
                                            IF NOT EXISTS (SELECT TOP 1 * FROM [insight].[dbo].[wh_snapshots]
                                                WHERE All_DK = @All_DK AND All_SE = @All_SE AND All_NO = @All_NO AND CreatedToday_DK = @CreatedToday_DK AND CreatedToday_SE = @CreatedToday_SE 
                                                AND CreatedToday_NO = @CreatedToday_NO AND SplitAll = @SplitAll AND SplitSentToday = @SplitSentToday AND ReadyToPickAll = @ReadyToPickAll
		                                        AND ReadyToPickCreatedToday = @ReadyToPickCreatedToday AND PickedReadyToShip = @PickedReadyToShip AND ShippedToday_DK = @ShippedToday_DK
                                                AND ShippedToday_SE = @ShippedToday_SE AND ShippedToday_NO = @ShippedToday_NO
                                                AND PickZone_Bestilling = @PickZone_Bestilling AND PickZone_Kompatibel = @PickZone_Kompatibel AND PickZone_Office = @PickZone_Office
                                                AND PickZone_Original = @PickZone_Original AND PickZone_Toner = @PickZone_Toner
                                                AND Id = (SELECT MAX(Id) FROM [insight].[dbo].[wh_snapshots]) ORDER BY Id DESC)
                                            BEGIN
                                            INSERT INTO wh_snapshots
                                                (All_DK, All_SE, All_NO, CreatedToday_DK, CreatedToday_SE, CreatedToday_NO, SplitAll, SplitSentToday, ReadyToPickAll,
                                                ReadyToPickCreatedToday, PickedReadyToShip, ShippedToday_DK, ShippedToday_SE, ShippedToday_NO, PickZone_Bestilling, PickZone_Kompatibel,
                                                PickZone_Office, PickZone_Original, PickZone_Toner, SnapshotCreatedAt) VALUES
                                                (@All_DK, @All_SE, @All_NO, @CreatedToday_DK, @CreatedToday_SE, @CreatedToday_NO, @SplitAll, @SplitSentToday, @ReadyToPickAll,
                                                @ReadyToPickCreatedToday, @PickedReadyToShip, @ShippedToday_DK, @ShippedToday_SE, @ShippedToday_NO, @PickZone_Bestilling,
                                                @PickZone_Kompatibel, @PickZone_Office, @PickZone_Original, @PickZone_Toner, @SnapshotCreatedAt)
                                            END"
            SqlCmdInsertSnapshop.Parameters.AddWithValue("@All_DK", intAll_DK)
            SqlCmdInsertSnapshop.Parameters.AddWithValue("@All_SE", intAll_SE)
            SqlCmdInsertSnapshop.Parameters.AddWithValue("@All_NO", intAll_NO)
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

            If Not objWEB01Conn.State = ConnectionState.Open Then
                objWEB01Conn.Open()
            End If

            SqlCmdInsertSnapshop.ExecuteNonQuery()

            If Not objWEB01Conn.State = ConnectionState.Closed Then
                objWEB01Conn.Close()
            End If

            SqlCmdInsertSnapshop.Parameters.Clear()

        Catch ex As Exception
            EventLog1.WriteEntry("Exception in GetVariousOrderData():" & vbCrLf & ex.ToString(), EventLogEntryType.Error, 15)
        Finally
            If Not objWEB01Conn.State = ConnectionState.Closed Then
                objWEB01Conn.Close()
            End If
        End Try

    End Sub

    Protected Sub GetOrdersAll()

        ' Reset variables
        intAll_DK = 0
        intAll_SE = 0
        intAll_NO = 0

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

            While strSqlReader.Read()

                If strSqlReader("Land") = "DKK" Then
                    intAll_DK = strSqlReader("Antal")
                End If

                If strSqlReader("Land") = "SEK" Then
                    intAll_SE = strSqlReader("Antal")
                End If

                If strSqlReader("Land") = "NOK" Then
                    intAll_NO = strSqlReader("Antal")
                End If

            End While

        Catch ex As Exception
            EventLog1.WriteEntry("Exception in GetOrdersAll():" & vbCrLf & ex.ToString(), EventLogEntryType.Error, 15)
        Finally
            If Not objSQL01Conn.State = ConnectionState.Closed Then
                objSQL01Conn.Close()
            End If
        End Try

    End Sub

    Protected Sub GetOrdersForToday()

        ' Reset variables
        intCreatedToday_DK = 0
        intCreatedToday_SE = 0
        intCreatedToday_NO = 0

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

            While strSqlReader.Read()

                If strSqlReader("Land") = "DKK" Then
                    intCreatedToday_DK = strSqlReader("Antal")
                End If

                If strSqlReader("Land") = "SEK" Then
                    intCreatedToday_SE = strSqlReader("Antal")
                End If

                If strSqlReader("Land") = "NOK" Then
                    intCreatedToday_NO = strSqlReader("Antal")
                End If

            End While

        Catch ex As Exception
            EventLog1.WriteEntry("Exception in GetOrdersForToday():" & vbCrLf & ex.ToString(), EventLogEntryType.Error, 15)
        Finally
            If Not objSQL01Conn.State = ConnectionState.Closed Then
                objSQL01Conn.Close()
            End If
        End Try

    End Sub

    Protected Sub GetOrdersShippedToday()

        ' Reset variables
        intShippedToday_DK = 0
        intShippedToday_SE = 0
        intShippedToday_NO = 0

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

            While strSqlReader.Read()

                If strSqlReader("Land") = "DKK" Then
                    intShippedToday_DK = strSqlReader("Antal")
                End If

                If strSqlReader("Land") = "SEK" Then
                    intShippedToday_SE = strSqlReader("Antal")
                End If

                If strSqlReader("Land") = "NOK" Then
                    intShippedToday_NO = strSqlReader("Antal")
                End If

            End While

        Catch ex As Exception
            EventLog1.WriteEntry("Exception in GetOrdersShippedToday():" & vbCrLf & ex.ToString(), EventLogEntryType.Error, 15)
        Finally
            If Not objSQL01Conn.State = ConnectionState.Closed Then
                objSQL01Conn.Close()
            End If
        End Try

    End Sub

    Protected Sub GetOrdersByPickZone()

        ' Reset variables
        intPickZone_Bestilling = 0
        intPickZone_Kompatibel = 0
        intPickZone_Office = 0
        intPickZone_Original = 0
        intPickZone_Toner = 0

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

            While strSqlReader.Read()

                If strSqlReader("Plukzone") = "BESTILLING" Then
                    intPickZone_Bestilling = strSqlReader("Antal")
                End If

                If strSqlReader("Plukzone") = "KOMPATIBEL" Then
                    intPickZone_Kompatibel = strSqlReader("Antal")
                End If

                If strSqlReader("Plukzone") = "OFFICE" Then
                    intPickZone_Office = strSqlReader("Antal")
                End If

                If strSqlReader("Plukzone") = "ORIGINAL" Then
                    intPickZone_Original = strSqlReader("Antal")
                End If

                If strSqlReader("Plukzone") = "TONER" Then
                    intPickZone_Toner = strSqlReader("Antal")
                End If

            End While

        Catch ex As Exception
            EventLog1.WriteEntry("Exception in GetOrdersPickByZone():" & vbCrLf & ex.ToString(), EventLogEntryType.Error, 15)
        Finally
            If Not objSQL01Conn.State = ConnectionState.Closed Then
                objSQL01Conn.Close()
            End If
        End Try

    End Sub

    Protected Sub GetVariousOrderData()

        ' Reset variables
        intSplitAll = 0
        intSplitSentToday = 0
        intReadyToPickAll = 0
        intReadyToPickCreatedToday = 0
        intPickedReadyToShip = 0

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

            While strSqlReader.Read()

                intSplitAll = strSqlReader("DelordreriAlt")
                intSplitSentToday = strSqlReader("DelordrerSendtiDag")
                intReadyToPickAll = strSqlReader("OrdrerTilPlukialt")
                intReadyToPickCreatedToday = strSqlReader("OrdrerSendtTilPlukidag")
                intPickedReadyToShip = strSqlReader("PlukketKlarTilPak")
                dtSnapshotCreatedAt = strSqlReader("SnapshotCreatedAt")

            End While

        Catch ex As Exception
            EventLog1.WriteEntry("Exception in GetVariousOrderData():" & vbCrLf & ex.ToString(), EventLogEntryType.Error, 15)
        Finally
            If Not objSQL01Conn.State = ConnectionState.Closed Then
                objSQL01Conn.Close()
            End If
        End Try

    End Sub

End Class
