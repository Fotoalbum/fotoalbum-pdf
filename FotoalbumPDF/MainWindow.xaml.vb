Imports System.IO
Imports System.Data
Imports MySql.Data.MySqlClient
Imports System.Xml
Imports System.Net
Imports System.Windows.Threading
Imports System.Threading
Imports System.Windows.Markup
Imports PDFlib_dotnet
Imports System.Windows.Media.Effects
Imports FAEffects.Effects
Imports System.Net.Mail
Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Drawing.Text
Imports System.Runtime

Class MainWindow

    'Public mySqlConnection As String = "Server=vm2526.vellance.net;Database=xhibit_2_0;Uid=fotoalbum_root;Pwd=3wnvfEX4psGrp7BU;"
    Public mySqlConnection As String = "Server=vm2526.vellance.net;Database=xhibit_2_0;Uid=fotoalbum_write;Pwd=JMYdhjAKyhHQ4DYJ;"

    Public searchpath_imagefolder As String = "http://api.xhibit.com/v2/"
    Public searchpath_fontfolder As String = "C:\DO_NOT_REMOVE_THIS_FOLDER\FONTS_ORIGINAL\" 'PRODUCTION
    Public exportfolder As String = ""

    '******** DEBUG *********
    'Public original_exportfolder As String = "C:\PDF_Debug\"
    'Public pdf_createfolder As String = "C:\PDF_Debug\"

    Public targetDPI As Integer = 300

    '******** PRODUCTION *********
    Public original_exportfolder As String = "C:\Users\Administrator\Documents\DONOTREMOVE\FA_ASSETS_PRERENDER\"
    Public pdf_createfolder As String = "N:\XHIBIT\www.xhibit.com\pdfexports_xhibit\"

    '******** PDF *********
    Private p As PDFlib_dotnet.PDFlib
    Public cover_filename As String
    Public bblock_filename As String
    Public textlines As XmlDocument
    Public textlinecontainers As XmlNodeList
    Public numPages As Integer
    Public fonts As DataTable
    Public usecover As Boolean = False
    Public singlepageproduct As Boolean = False
    Public trimbox As String
    Public currentpageorientation As String = "topdown"
    Public pdfcompatible As String = "1.3"

    Public orders As ArrayList = New ArrayList
    Public currentOrder As order
    Public spreads As XmlDocument
    Public spreadLst As XmlNodeList
    Public colors As XmlDocument
    Public colorlist As XmlNodeList
    Public res As Integer = 96
    Public renderarray As ArrayList
    Public doneloading As Boolean = False
    Public renderindex As Integer = 0
    Private savecovertimer As DispatcherTimer
    Private savebblocktimer As DispatcherTimer
    Private starttimer As DispatcherTimer
    Public spineX As Double = 0
    Public isLayFlat As Boolean = False

    Public ordertimer As DispatcherTimer

    Public Event SaveGrid(sender As Object, e As SaveGridEventArgs)

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)

        ordertimer = New DispatcherTimer
        AddHandler ordertimer.Tick, AddressOf GetOrders
        ordertimer.Interval = New TimeSpan(0, 0, 5)
        ordertimer.Start()

    End Sub

    Private Sub GetOrders(sender As Object, e As EventArgs)

        Try

            ordertimer.Stop()

            Me.Content = Nothing
            Me.UpdateLayout()

            Dim dt As DataTable = New DataTable

            Dim query As String = "SELECT o.id, o.status, o.order_id,  u.product_id, u.user_product_id, u.pages_xml, u.textflow_xml, u.textlines_xml, u.photo_xml, u.color_xml FROM pdfengine_order_pdfs o " &
                                  "LEFT OUTER JOIN pdfengine_order_pdf_user_products u ON u.id = o.order_pdf_user_product_id " &
                                  "WHERE o.status='start'"

            Dim connStringSQL As New MySqlConnection(mySqlConnection)
            Dim myAdapter As New MySqlDataAdapter(query, connStringSQL)
            myAdapter.Fill(dt)

            For Each row As DataRow In dt.Rows
                Dim neworder As New order()
                neworder.id = row("id")
                neworder.status = row("status")
                neworder.order_id = row("order_id")
                neworder.product_id = row("product_id")
                neworder.user_product_id = row("user_product_id")
                If Not IsDBNull(row("pages_xml")) Then
                    neworder.pages_xml = row("pages_xml")
                Else
                    ExitOrder("Geen pages_xml gevonden!!")
                End If
                If Not IsDBNull(row("textflow_xml")) Then
                    neworder.textflow_xml = row("textflow_xml")
                Else
                    neworder.textflow_xml = Nothing
                End If
                If Not IsDBNull(row("textlines_xml")) Then
                    neworder.textlines_xml = row("textlines_xml")
                Else
                    neworder.textlines_xml = Nothing
                End If
                If Not IsDBNull(row("photo_xml")) Then
                    neworder.photo_xml = row("photo_xml")
                Else
                    neworder.photo_xml = Nothing
                End If
                If Not IsDBNull(row("color_xml")) Then
                    neworder.color_xml = row("color_xml")
                Else
                    neworder.color_xml = Nothing
                End If
                orders.Add(neworder)
            Next

            If orders.Count > 0 Then
                'Process the order
                currentOrder = orders.Item(0)
                exportfolder = original_exportfolder & currentOrder.id & "\"
                CreatePDFRender()
            Else
                ordertimer.Start()
            End If

        Catch ex As Exception

            ExitOrder(ex.Message)

        End Try

    End Sub

    Private Sub CreatePDFRender()

        Try

            isLayFlat = False

            'Check if this is a lay flat book
            Dim dt As DataTable = New DataTable

            Dim query As String = "SELECT product_papertype_id FROM xhibit_products " & _
                                  "WHERE id=" & currentOrder.product_id

            Dim connStringSQL As New MySqlConnection(mySqlConnection)
            Dim myAdapter As New MySqlDataAdapter(query, connStringSQL)
            myAdapter.Fill(dt)
            'Papertype 6 and 7 are LAYFLAT books -> Create PDF as Spreads
            For Each row As DataRow In dt.Rows
                If row("product_papertype_id").ToString = "6" Or row("product_papertype_id").ToString = "7" Then
                    isLayFlat = True
                End If
            Next

            renderarray = New ArrayList

            fonts = New DataTable

            query = "SELECT * FROM cms_app_fonts"
            connStringSQL = New MySqlConnection(mySqlConnection)
            myAdapter = New MySqlDataAdapter(query, connStringSQL)
            myAdapter.Fill(fonts)

            'Check if we need to create a cover or not
            spreads = New XmlDocument()
            spreads.LoadXml(currentOrder.pages_xml)

            spreadLst = spreads.GetElementsByTagName("spread")
            Dim firstspread As XmlNode = spreadLst(0)
            Dim pages As XmlNodeList = firstspread.SelectNodes("descendant::page")

            usecover = False
            If (pages(0).Attributes.GetNamedItem("pageType").Value = "coverback") Then
                usecover = True
                cover_filename = Guid.NewGuid.ToString & ".pdf"
            Else
                cover_filename = ""
            End If

            bblock_filename = Guid.NewGuid.ToString & ".pdf"

            'Check single page product
            singlepageproduct = False
            If Not IsNothing(firstspread.Attributes.GetNamedItem("spe")) Then
                If firstspread.Attributes.GetNamedItem("spe").Value = "true" Then
                    singlepageproduct = True
                End If
            End If

            'Initiate the textlines collection
            textlines = New XmlDocument
            textlines.PreserveWhitespace = True

            If Not IsNothing(currentOrder.textlines_xml) Then
                If currentOrder.textlines_xml <> "" Then
                    textlines.LoadXml(currentOrder.textlines_xml)
                    textlinecontainers = textlines.SelectNodes("descendant::container")
                End If
            End If

            'Set the color reference collection
            colors = New XmlDocument
            If Not IsNothing(currentOrder.color_xml) Then
                If currentOrder.color_xml <> "" Then
                    colors.LoadXml(currentOrder.color_xml)
                    colorlist = colors.SelectNodes("descendant::color")
                End If
            End If

            '==============================================================================
            ' Remove content or create the image directory
            '==============================================================================
            If Directory.Exists(exportfolder) Then
                Try
                    Directory.Delete(exportfolder, True)
                Catch ex As Exception
                    Debug.Print("Error deleting directory! " & ex.Message)
                End Try
            End If

            'Create the new directory after a short pause
            starttimer = New DispatcherTimer()
            AddHandler starttimer.Tick, AddressOf CreateDir
            starttimer.Interval = New TimeSpan(0, 0, 1)
            starttimer.Start()

        Catch ex As System.Exception

            MsgBox(ex.Message)

            ExitOrder(ex.Message)

        End Try

    End Sub

    Public Sub CreateDir()

        starttimer.Stop()
        starttimer = Nothing

        Dim result As Boolean = True

        Directory.CreateDirectory(exportfolder)

        '==============================================================================
        ' Start the PDF Creation process
        '==============================================================================

        renderindex = 0
        spineX = 0

        CreatePDF()

    End Sub

    Public Sub CreatePDF()

        Try

            If renderindex = 0 And usecover = True Then

                CreateCover()

            Else

                CreatePageBlock()

            End If

        Catch ex As Exception

            ExitOrder(ex.Message)

        End Try

    End Sub

    Public Sub CreateCover()

        Dim scale As New ScaleTransform()
        scale.ScaleX = 0.1
        scale.ScaleY = 0.1
        Me.LayoutTransform = scale

        Try

            Dim grid As New Canvas
            AddChild(grid)
            Me.UpdateLayout()

            Dim spread As XmlNode = spreadLst(0)
            Dim spreadID As String = spread.Attributes.GetNamedItem("spreadID").Value

            Dim pages As XmlNodeList = spread.SelectNodes("descendant::page")

            Dim coverbleed As Double = XmlConvert.ToDouble(pages(0).Attributes.GetNamedItem("horizontalBleed").Value.ToString()) * (targetDPI / 96)
            Dim coverwrap As Double = XmlConvert.ToDouble(pages(0).Attributes.GetNamedItem("horizontalWrap").Value.ToString()) * (targetDPI / 96)
            Dim trimbox As String = ""

            Dim width As Double = XmlConvert.ToDouble(spread.Attributes.GetNamedItem("totalWidth").Value.ToString()) * (targetDPI / 96)
            Dim height As Double = XmlConvert.ToDouble(spread.Attributes.GetNamedItem("totalHeight").Value.ToString()) * (targetDPI / 96)

            Dim mainspread As New Canvas
            mainspread.Width = width
            mainspread.Height = height
            mainspread.HorizontalAlignment = Windows.HorizontalAlignment.Left
            mainspread.VerticalAlignment = Windows.VerticalAlignment.Top
            mainspread.Background = New SolidColorBrush(System.Windows.Media.Color.FromRgb(255, 255, 255))
            mainspread.ClipToBounds = True
            grid.Children.Add(mainspread)

            mainspread.UpdateLayout()

            If spread.Attributes.GetNamedItem("backgroundColor").Value.ToString <> "-1" Then

                Dim fillArr As ArrayList = GetColorRgb(spread.Attributes.GetNamedItem("backgroundColor").Value.ToString)

                Dim bg As New Canvas
                bg.Width = width
                bg.Height = height
                bg.Background = New SolidColorBrush(System.Windows.Media.Color.FromRgb(fillArr(0), fillArr(1), fillArr(2)))
                bg.HorizontalAlignment = Windows.HorizontalAlignment.Left
                bg.VerticalAlignment = Windows.VerticalAlignment.Top
                bg.Margin = New Thickness(0, 0, 0, 0)
                mainspread.Children.Add(bg)

            End If

            If Not spread.SelectSingleNode("background") Is Nothing Then

                Dim background As XmlNode = spread.SelectSingleNode("background")

                'Load display and save the image, then put it in the PDF
                Dim filename As String
                Dim continueImage As Boolean = True

                If background.Attributes.GetNamedItem("status").Value <> "done" Then
                    If background.Attributes.GetNamedItem("status").Value = "new" Then
                        'Get the filedata from the upload, if its available
                        Dim arr As ArrayList = GetFileUrlFromUpload(background.Attributes.GetNamedItem("id").Value)
                        If arr.Count > 0 Then
                            filename = searchpath_imagefolder & arr(0)
                            filename = filename.Replace("\", "/")
                        Else
                            continueImage = False
                        End If
                    ElseIf background.Attributes.GetNamedItem("status").Value = "empty" Then
                        continueImage = False
                    End If
                Else
                    filename = searchpath_imagefolder & background.Attributes.GetNamedItem("hires_url").Value
                    filename = filename.Replace("\", "/")
                End If

                If (continueImage) Then

                    Dim alpha As Double = XmlConvert.ToDouble(spread.Attributes.GetNamedItem("backgroundAlpha").Value.ToString)
                    Dim fliphorizontal As Integer = 0
                    If background.Attributes.GetNamedItem("fliphorizontal").Value <> "" Then
                        fliphorizontal = XmlConvert.ToInt32(background.Attributes.GetNamedItem("fliphorizontal").Value)
                    End If
                    Dim imgRotation As Integer = XmlConvert.ToInt32(background.Attributes.GetNamedItem("imageRotation").Value)
                    Dim imgWidth As Double = XmlConvert.ToDouble(background.Attributes.GetNamedItem("width").Value) * (targetDPI / 96)
                    Dim imgHeight As Double = XmlConvert.ToDouble(background.Attributes.GetNamedItem("height").Value) * (targetDPI / 96)
                    Dim offsetX As Double = XmlConvert.ToDouble(background.Attributes.GetNamedItem("x").Value) * (targetDPI / 96)
                    Dim offsetY As Double = XmlConvert.ToDouble(background.Attributes.GetNamedItem("y").Value) * (targetDPI / 96)

                    Dim c As New WebClient()
                    Dim bytes = c.DownloadData(New Uri(filename))
                    Dim ms As New MemoryStream(bytes)

                    Dim src As BitmapImage = New BitmapImage
                    src.BeginInit()
                    src.CreateOptions = BitmapCreateOptions.PreservePixelFormat Or BitmapCreateOptions.IgnoreColorProfile
                    src.StreamSource = ms
                    src.CacheOption = BitmapCacheOption.None
                    Select Case imgRotation
                        Case 90

                            src.Rotation = Rotation.Rotate90

                            Dim iw As Double = imgWidth
                            Dim xo As Double = offsetX

                            imgWidth = imgHeight
                            imgHeight = iw

                            'Check if the image was flipped
                            If fliphorizontal = 0 Then
                                offsetX = width - (imgWidth + offsetY)
                                offsetY = xo
                            Else
                                offsetX = width - (imgWidth + offsetY)
                                offsetY = height - (imgHeight + xo)
                            End If
                        Case 180

                            src.Rotation = Rotation.Rotate180

                            If fliphorizontal = 0 Then
                                offsetX = width - (imgWidth + offsetX)
                                offsetY = height - (imgHeight + offsetY)
                            Else
                                offsetY = height - (imgHeight + offsetY)
                            End If

                        Case 270

                            src.Rotation = Rotation.Rotate270

                            Dim iw As Double = imgWidth
                            Dim xo As Double = offsetX

                            imgWidth = imgHeight
                            imgHeight = iw

                            'Check if the image was flipped
                            If fliphorizontal = 0 Then
                                offsetX = offsetY
                                offsetY = height - (imgHeight + xo)
                            Else
                                offsetX = offsetY
                                offsetY = xo
                            End If

                        Case -90

                            src.Rotation = Rotation.Rotate270

                            Dim iw As Double = imgWidth
                            Dim xo As Double = offsetX

                            imgWidth = imgHeight
                            imgHeight = iw

                            'Check if the image was flipped
                            If fliphorizontal = 0 Then
                                offsetX = offsetY
                                offsetY = height - (imgHeight + xo)
                            Else
                                offsetX = offsetY
                                offsetY = xo
                            End If

                        Case 0

                            src.Rotation = Rotation.Rotate0

                    End Select
                    src.EndInit()

                    Dim placeholder As New Canvas
                    placeholder.Width = width
                    placeholder.Height = height
                    placeholder.ClipToBounds = True
                    mainspread.Children.Add(placeholder)

                    Dim srcimg As New System.Windows.Controls.Image
                    RenderOptions.SetBitmapScalingMode(srcimg, BitmapScalingMode.Fant)
                    Select Case background.Attributes.GetNamedItem("imageFilter").Value
                        Case "" 'Default
                            srcimg.Source = src
                        Case "bw"
                            srcimg.Source = MakeGrayscale(src)
                        Case "sepia"
                            srcimg.Source = src
                    End Select

                    If fliphorizontal = 1 Then
                        srcimg.RenderTransformOrigin = New System.Windows.Point(0.5, 0.5)
                        Dim flip As New ScaleTransform()
                        Select Case imgRotation
                            Case 0
                                offsetX = width - (imgWidth + offsetX)
                                flip.ScaleX = -1
                            Case 90
                                flip.ScaleY = -1
                            Case 180
                                flip.ScaleX = -1
                            Case 270
                                flip.ScaleY = -1
                        End Select
                        srcimg.RenderTransform = flip
                    End If

                    srcimg.HorizontalAlignment = Windows.HorizontalAlignment.Left
                    srcimg.VerticalAlignment = Windows.VerticalAlignment.Top
                    srcimg.Width = imgWidth
                    srcimg.Height = imgHeight
                    srcimg.Margin = New Thickness(offsetX, offsetY, 0, 0)

                    If alpha < 1 Then
                        srcimg.Opacity = alpha
                    End If

                    If background.Attributes.GetNamedItem("imageFilter").Value = "sepia" Then
                        srcimg = MakeSepia(srcimg)
                    End If

                    placeholder.Children.Add(srcimg)
                    placeholder.UpdateLayout()
                Else
                    If background.Attributes.GetNamedItem("status").Value <> "empty" Then
                        Throw New Exception("No source found!! " + background.OuterXml)
                    End If
                End If
            End If

            'Add separate backgrounds if available
            Dim xpos As Double = 0

            For Each page As XmlNode In pages

                Dim position As String = "center"

                Dim pageID As String = page.Attributes.GetNamedItem("pageID").Value
                If page.Attributes.GetNamedItem("pageType").Value.ToString = "coverback" Then
                    spineX = (XmlConvert.ToDouble(page.Attributes.GetNamedItem("pageWidth").Value) * (targetDPI / 96)) + coverbleed + coverwrap
                End If

                If page.Attributes.GetNamedItem("pageType").Value.ToString = "coverspine" Then
                    spineX += (XmlConvert.ToDouble(page.Attributes.GetNamedItem("pageWidth").Value) * (targetDPI / 96)) / 2
                End If

                Dim pagewidth As Double = XmlConvert.ToDouble(page.Attributes.GetNamedItem("pageWidth").Value) * (targetDPI / 96)
                Dim pageheight As Double = (XmlConvert.ToDouble(page.Attributes.GetNamedItem("pageHeight").Value) * (targetDPI / 96)) + (coverbleed * 2) + (coverwrap * 2)
                If page.Attributes.GetNamedItem("pageType").Value.ToString <> "coverspine" Then
                    pagewidth += coverbleed + coverwrap
                End If

                If page.Attributes.GetNamedItem("backgroundColor").Value.ToString <> "-1" Then

                    Dim fillArr As ArrayList = GetColorRgb(page.Attributes.GetNamedItem("backgroundColor").Value.ToString)

                    Dim bg As New System.Windows.Shapes.Rectangle
                    bg.Width = pagewidth
                    If page.Attributes.GetNamedItem("pageType").Value.ToString = "coverspine" Then
                        bg.Width += 2 * (targetDPI / 96)
                    End If
                    bg.Height = pageheight
                    bg.Fill = New SolidColorBrush(System.Windows.Media.Color.FromRgb(fillArr(0), fillArr(1), fillArr(2)))
                    bg.HorizontalAlignment = Windows.HorizontalAlignment.Left
                    bg.VerticalAlignment = Windows.VerticalAlignment.Top
                    'bg.StrokeThickness = 0
                    'bg.Stroke = New SolidColorBrush(Color.FromRgb(0, 0, 0))
                    bg.RadiusX = 0
                    bg.RadiusY = 0
                    If page.Attributes.GetNamedItem("pageType").Value.ToString = "coverspine" Then
                        bg.Margin = New Thickness(xpos - (targetDPI / 96), 0, 0, 0)
                    Else
                        bg.Margin = New Thickness(xpos, 0, 0, 0)
                    End If
                    mainspread.Children.Add(bg)

                End If

                If Not page.SelectSingleNode("background") Is Nothing Then

                    Dim background As XmlNode = page.SelectSingleNode("background")

                    Dim filename As String
                    Dim continueImage As Boolean = True

                    If background.Attributes.GetNamedItem("status").Value <> "done" Then
                        If background.Attributes.GetNamedItem("status").Value = "new" Then
                            'Get the filedata from the upload, if its available
                            Dim arr As ArrayList = GetFileUrlFromUpload(background.Attributes.GetNamedItem("id").Value)
                            If arr.Count > 0 Then
                                filename = searchpath_imagefolder & arr(0)
                                filename = filename.Replace("\", "/")
                            Else
                                continueImage = False
                            End If
                        ElseIf background.Attributes.GetNamedItem("status").Value = "empty" Or background.Attributes.GetNamedItem("status").Value = "" Then
                            continueImage = False
                        End If
                    Else
                        filename = searchpath_imagefolder & background.Attributes.GetNamedItem("hires_url").Value
                        filename = filename.Replace("\", "/")
                    End If

                    If (continueImage) Then

                        Dim alpha As Double = XmlConvert.ToDouble(page.Attributes.GetNamedItem("backgroundAlpha").Value.ToString)
                        Dim fliphorizontal As Integer = 0
                        If background.Attributes.GetNamedItem("fliphorizontal").Value <> "" Then
                            fliphorizontal = XmlConvert.ToInt32(background.Attributes.GetNamedItem("fliphorizontal").Value)
                        End If
                        Dim imgRotation As Integer = XmlConvert.ToInt32(background.Attributes.GetNamedItem("imageRotation").Value)
                        Dim imgWidth As Double = XmlConvert.ToDouble(background.Attributes.GetNamedItem("width").Value) * (targetDPI / 96)
                        Dim imgHeight As Double = XmlConvert.ToDouble(background.Attributes.GetNamedItem("height").Value) * (targetDPI / 96)
                        Dim offsetX As Double = XmlConvert.ToDouble(background.Attributes.GetNamedItem("x").Value) * (targetDPI / 96)
                        Dim offsetY As Double = XmlConvert.ToDouble(background.Attributes.GetNamedItem("y").Value) * (targetDPI / 96)

                        Dim c As New WebClient()
                        Dim bytes = c.DownloadData(New Uri(filename))
                        Dim ms As New MemoryStream(bytes)

                        Dim src As BitmapImage = New BitmapImage
                        src.BeginInit()
                        src.CreateOptions = BitmapCreateOptions.PreservePixelFormat Or BitmapCreateOptions.IgnoreColorProfile
                        src.StreamSource = ms
                        src.CacheOption = BitmapCacheOption.None
                        Select Case imgRotation
                            Case 90

                                src.Rotation = Rotation.Rotate90

                                Dim iw As Double = imgWidth
                                Dim xo As Double = offsetX

                                imgWidth = imgHeight
                                imgHeight = iw

                                'Check if the image was flipped
                                If fliphorizontal = 0 Then
                                    offsetX = width - (imgWidth + offsetY)
                                    offsetY = xo
                                Else
                                    offsetX = width - (imgWidth + offsetY)
                                    offsetY = height - (imgHeight + xo)
                                End If

                            Case 180

                                src.Rotation = Rotation.Rotate180

                                If fliphorizontal = 0 Then
                                    offsetX = width - (imgWidth + offsetX)
                                    offsetY = height - (imgHeight + offsetY)
                                Else
                                    offsetY = height - (imgHeight + offsetY)
                                End If

                            Case 270

                                src.Rotation = Rotation.Rotate270

                                Dim iw As Double = imgWidth
                                Dim xo As Double = offsetX

                                imgWidth = imgHeight
                                imgHeight = iw

                                'Check if the image was flipped
                                If fliphorizontal = 0 Then
                                    offsetX = offsetY
                                    offsetY = height - (imgHeight + xo)
                                Else
                                    offsetX = offsetY
                                    offsetY = xo
                                End If

                            Case -90

                                src.Rotation = Rotation.Rotate270

                                Dim iw As Double = imgWidth
                                Dim xo As Double = offsetX

                                imgWidth = imgHeight
                                imgHeight = iw

                                'Check if the image was flipped
                                If fliphorizontal = 0 Then
                                    offsetX = offsetY
                                    offsetY = height - (imgHeight + xo)
                                Else
                                    offsetX = offsetY
                                    offsetY = xo
                                End If

                            Case 0

                                src.Rotation = Rotation.Rotate0

                        End Select
                        src.EndInit()

                        Dim placeholder As New Canvas
                        placeholder.Width = pagewidth
                        placeholder.Height = pageheight
                        placeholder.HorizontalAlignment = Windows.HorizontalAlignment.Left
                        placeholder.VerticalAlignment = Windows.VerticalAlignment.Top
                        placeholder.Margin = New Thickness(xpos, 0, 0, 0)
                        placeholder.ClipToBounds = True
                        mainspread.Children.Add(placeholder)

                        Dim srcimg As New System.Windows.Controls.Image
                        RenderOptions.SetBitmapScalingMode(srcimg, BitmapScalingMode.Fant)
                        Select Case background.Attributes.GetNamedItem("imageFilter").Value
                            Case "" 'Default
                                srcimg.Source = src
                            Case "bw"
                                srcimg.Source = MakeGrayscale(src)
                            Case "sepia"
                                srcimg.Source = src
                        End Select

                        If fliphorizontal = 1 Then
                            srcimg.RenderTransformOrigin = New System.Windows.Point(0.5, 0.5)
                            Dim flip As New ScaleTransform()
                            Select Case imgRotation
                                Case 0
                                    offsetX = width - (imgWidth + offsetX)
                                    flip.ScaleX = -1
                                Case 90
                                    flip.ScaleY = -1
                                Case 180
                                    flip.ScaleX = -1
                                Case 270
                                    flip.ScaleY = -1
                            End Select
                            srcimg.RenderTransform = flip
                        End If

                        srcimg.HorizontalAlignment = Windows.HorizontalAlignment.Left
                        srcimg.VerticalAlignment = Windows.VerticalAlignment.Top
                        srcimg.Width = XmlConvert.ToDouble(background.Attributes.GetNamedItem("width").Value) * (targetDPI / 96)
                        srcimg.Height = XmlConvert.ToDouble(background.Attributes.GetNamedItem("height").Value) * (targetDPI / 96)
                        srcimg.Margin = New Thickness(XmlConvert.ToDouble(background.Attributes.GetNamedItem("x").Value) * (targetDPI / 96),
                        XmlConvert.ToDouble(background.Attributes.GetNamedItem("y").Value) * (targetDPI / 96),
                                0, 0)

                        If alpha < 1 Then
                            srcimg.Opacity = alpha
                        End If

                        If background.Attributes.GetNamedItem("imageFilter").Value = "sepia" Then
                            srcimg = MakeSepia(srcimg)
                        End If

                        placeholder.Children.Add(srcimg)
                        placeholder.UpdateLayout()
                    Else
                        If background.Attributes.GetNamedItem("status").Value <> "empty" Then
                            Throw New Exception("No source found!! " + background.OuterXml)
                        End If
                    End If
                End If

                xpos += pagewidth

            Next

            'Add the elements
            Dim elements As XmlNodeList = spread.SelectSingleNode("elements").SelectNodes("element")

            For Each element As XmlNode In elements

                Dim elementID As String = element.Attributes.GetNamedItem("id").Value

                Debug.Print(element.Attributes.GetNamedItem("type").Value)

                If element.Attributes.GetNamedItem("type").Value = "photo" Then

                    Dim filename As String = ""

                    Dim continueImage As Boolean = True

                    If element.Attributes.GetNamedItem("status").Value <> "done" Then
                        If element.Attributes.GetNamedItem("status").Value = "new" Then
                            'Get the filedata from the upload, if its available
                            Dim arr As ArrayList = GetFileUrlFromUpload(element.Attributes.GetNamedItem("original_image_id").Value)
                            If arr.Count > 0 Then
                                filename = searchpath_imagefolder & arr(0)
                                filename = filename.Replace("\", "/")
                            Else
                                continueImage = False
                            End If
                        ElseIf element.Attributes.GetNamedItem("status").Value = "empty" Then
                            continueImage = False
                        End If
                    Else
                        If (element.Attributes.GetNamedItem("hires_url").Value <> "") Then
                            filename = searchpath_imagefolder & element.Attributes.GetNamedItem("hires_url").Value
                            filename = filename.Replace("\", "/")
                            Debug.Print(filename)
                        Else
                            Dim arr As ArrayList = GetFileUrlFromUpload(element.Attributes.GetNamedItem("original_image_id").Value)
                            If arr.Count > 0 Then
                                filename = searchpath_imagefolder & arr(0)
                                filename = filename.Replace("\", "/")
                                Debug.Print(filename)
                            Else
                                continueImage = False
                            End If
                        End If
                    End If

                    If continueImage Then

                        Dim posx As Double = XmlConvert.ToDouble(element.Attributes.GetNamedItem("objectX").Value.ToString) * (targetDPI / 96)
                        Dim posy As Double = XmlConvert.ToDouble(element.Attributes.GetNamedItem("objectY").Value.ToString) * (targetDPI / 96)
                        Dim w As Double = Math.Round(XmlConvert.ToDouble(element.Attributes.GetNamedItem("objectWidth").Value.ToString) * (targetDPI / 96))
                        Dim h As Double = Math.Round(XmlConvert.ToDouble(element.Attributes.GetNamedItem("objectHeight").Value.ToString) * (targetDPI / 96))
                        Dim r As Double = XmlConvert.ToDouble(element.Attributes.GetNamedItem("rotation").Value.ToString)

                        Dim offsetX As Double = Math.Round(XmlConvert.ToDouble(element.Attributes.GetNamedItem("offsetX").Value.ToString) * (targetDPI / 96))
                        Dim offsetY As Double = Math.Round(XmlConvert.ToDouble(element.Attributes.GetNamedItem("offsetY").Value.ToString) * (targetDPI / 96))
                        Dim imgWidth As Double = Math.Round(XmlConvert.ToDouble(element.Attributes.GetNamedItem("imageWidth").Value) * (targetDPI / 96))
                        Dim imgHeight As Double = Math.Round(XmlConvert.ToDouble(element.Attributes.GetNamedItem("imageHeight").Value) * (targetDPI / 96))
                        Dim imgRotation As Double = XmlConvert.ToDouble(element.Attributes.GetNamedItem("imageRotation").Value)

                        Dim imageMask As String = element.Attributes.GetNamedItem("mask_hires_url").Value.ToString
                        Dim imageOverlay As String = element.Attributes.GetNamedItem("overlay_hires_url").Value.ToString
                        Dim imageAlpha As Double = XmlConvert.ToDouble(element.Attributes.GetNamedItem("imageAlpha").Value.ToString)

                        Dim fliphorizontal As Integer = 0
                        If element.Attributes.GetNamedItem("fliphorizontal").Value <> "" Then
                            fliphorizontal = XmlConvert.ToInt32(element.Attributes.GetNamedItem("fliphorizontal").Value)
                        End If

                        Dim shadow As String = element.Attributes.GetNamedItem("shadow").Value
                        Dim useshadow As Boolean = False

                        If shadow <> "" Then
                            useshadow = True
                        End If

                        Dim borderweight As Integer = 0
                        If Not IsNothing(element.Attributes.GetNamedItem("borderweight")) Then
                            If element.Attributes.GetNamedItem("borderweight").Value > 0 Then
                                borderweight = Math.Round(XmlConvert.ToDouble(element.Attributes.GetNamedItem("borderweight").Value) * (targetDPI / 96))
                            End If
                        End If

                        Dim border As New Border()
                        border.Width = w
                        border.Height = h
                        border.BorderThickness = New Thickness(0)
                        border.HorizontalAlignment = Windows.HorizontalAlignment.Left
                        border.VerticalAlignment = Windows.VerticalAlignment.Top
                        border.Margin = New Thickness(posx - borderweight, posy - borderweight, 0, 0)
                        border.ClipToBounds = False
                        mainspread.Children.Add(border)

                        If borderweight > 0 Then
                            border.BorderThickness = New Thickness(borderweight)
                            border.Width += (borderweight * 2)
                            border.Height += (borderweight * 2)
                            Dim colorArr As ArrayList = GetColorRgb(element.Attributes.GetNamedItem("bordercolor").Value)
                            border.BorderBrush = New SolidColorBrush(System.Windows.Media.Color.FromRgb(colorArr(0), colorArr(1), colorArr(2)))

                            imgWidth += borderweight
                            imgHeight += borderweight
                        End If

                        'Add a placeholder for the image
                        Dim placeholder As New Canvas
                        placeholder.Width = w
                        placeholder.Height = h
                        placeholder.HorizontalAlignment = Windows.HorizontalAlignment.Left
                        placeholder.VerticalAlignment = Windows.VerticalAlignment.Top
                        placeholder.ClipToBounds = True
                        border.Child = placeholder

                        Dim c As New WebClient()
                        Dim bytes = c.DownloadData(New Uri(filename))
                        Dim ms As New MemoryStream(bytes)

                        Dim src As BitmapImage = New BitmapImage()
                        src.BeginInit()
                        src.CreateOptions = BitmapCreateOptions.PreservePixelFormat Or BitmapCreateOptions.IgnoreColorProfile
                        src.StreamSource = ms
                        src.CacheOption = BitmapCacheOption.None

                        Select Case imgRotation
                            Case 90

                                src.Rotation = Rotation.Rotate90

                                Dim iw As Double = imgWidth
                                Dim xo As Double = offsetX

                                imgWidth = imgHeight
                                imgHeight = iw

                                'Check if the image was flipped
                                If fliphorizontal = 0 Then
                                    offsetX = w - (imgWidth + offsetY)
                                    offsetY = xo
                                Else
                                    offsetX = w - (imgWidth + offsetY)
                                    offsetY = h - (imgHeight + xo)
                                End If

                            Case 180

                                src.Rotation = Rotation.Rotate180

                                If fliphorizontal = 0 Then
                                    offsetX = w - (imgWidth + offsetX)
                                    offsetY = h - (imgHeight + offsetY)
                                Else
                                    offsetY = h - (imgHeight + offsetY)
                                End If

                            Case 270

                                src.Rotation = Rotation.Rotate270

                                Dim iw As Double = imgWidth
                                Dim xo As Double = offsetX

                                imgWidth = imgHeight
                                imgHeight = iw

                                'Check if the image was flipped
                                If fliphorizontal = 0 Then
                                    offsetX = offsetY
                                    offsetY = h - (imgHeight + xo)
                                Else
                                    offsetX = offsetY
                                    offsetY = xo
                                End If

                            Case -90

                                src.Rotation = Rotation.Rotate270

                                Dim iw As Double = imgWidth
                                Dim xo As Double = offsetX

                                imgWidth = imgHeight
                                imgHeight = iw

                                'Check if the image was flipped
                                If fliphorizontal = 0 Then
                                    offsetX = offsetY
                                    offsetY = h - (imgHeight + xo)
                                Else
                                    offsetX = offsetY
                                    offsetY = xo
                                End If

                            Case 0

                                src.Rotation = Rotation.Rotate0

                        End Select
                        src.EndInit()

                        Dim srcimg As New System.Windows.Controls.Image
                        RenderOptions.SetBitmapScalingMode(srcimg, BitmapScalingMode.HighQuality)

                        Select Case element.Attributes.GetNamedItem("imageFilter").Value
                            Case "" 'Default
                                srcimg.Source = src
                            Case "bw"
                                srcimg.Source = MakeGrayscale(src)
                            Case "sepia"
                                srcimg.Source = src
                        End Select

                        If fliphorizontal = 1 Then
                            srcimg.RenderTransformOrigin = New System.Windows.Point(0.5, 0.5)
                            Dim flip As New ScaleTransform()
                            Select Case imgRotation
                                Case 0
                                    offsetX = w - (imgWidth + offsetX)
                                    flip.ScaleX = -1
                                Case 90
                                    flip.ScaleY = -1
                                Case 180
                                    flip.ScaleX = -1
                                Case 270
                                    flip.ScaleY = -1
                            End Select
                            srcimg.RenderTransform = flip
                        End If

                        srcimg.HorizontalAlignment = Windows.HorizontalAlignment.Left
                        srcimg.VerticalAlignment = Windows.VerticalAlignment.Top
                        srcimg.Width = imgWidth
                        srcimg.Height = imgHeight
                        srcimg.Margin = New Thickness(offsetX, offsetY, 0, 0)

                        If element.Attributes.GetNamedItem("imageFilter").Value = "sepia" Then
                            srcimg = MakeSepia(srcimg)
                        End If

                        If imageAlpha < 1 Then
                            placeholder.Opacity = imageAlpha
                        End If

                        placeholder.Children.Add(srcimg)
                        placeholder.UpdateLayout()

                        If imageMask <> "" Then

                            Dim mask As New ImageBrush()
                            Dim path As String = searchpath_imagefolder & imageMask
                            path = path.Replace("\", "/")
                            mask.ImageSource = New BitmapImage(New Uri(path))
                            mask.Viewport = New Rect(0, 0, w, h)
                            mask.TileMode = TileMode.None
                            mask.ViewportUnits = BrushMappingMode.Absolute
                            placeholder.OpacityMask = mask

                        End If

                        If imageOverlay <> "" Then

                            Dim overlaypath As String = searchpath_imagefolder & imageOverlay
                            overlaypath = overlaypath.Replace("\", "/")

                            Dim overlaywc As New WebClient()
                            Dim overlaybytes = overlaywc.DownloadData(New Uri(overlaypath))
                            Dim overlayms As New MemoryStream(overlaybytes)

                            Dim overlaysrc As BitmapImage = New BitmapImage
                            overlaysrc.BeginInit()
                            overlaysrc.CreateOptions = BitmapCreateOptions.PreservePixelFormat Or BitmapCreateOptions.IgnoreColorProfile
                            overlaysrc.CacheOption = BitmapCacheOption.None
                            overlaysrc.StreamSource = overlayms
                            overlaysrc.EndInit()

                            Dim overlay As New System.Windows.Controls.Image
                            RenderOptions.SetBitmapScalingMode(overlay, BitmapScalingMode.Fant)
                            overlay.Source = overlaysrc
                            overlay.Stretch = Stretch.Fill
                            overlay.HorizontalAlignment = Windows.HorizontalAlignment.Left
                            overlay.VerticalAlignment = Windows.VerticalAlignment.Top
                            overlay.Width = w
                            overlay.Height = h
                            overlay.Margin = New Thickness(posx, posy, 0, 0)

                            If imageAlpha < 1 Then
                                overlay.Opacity = imageAlpha
                            End If

                            mainspread.Children.Add(overlay)

                        End If

                        If shadow <> "" Then

                            Dim dse As System.Windows.Media.Effects.DropShadowEffect = New System.Windows.Media.Effects.DropShadowEffect
                            dse.Color = System.Windows.Media.Color.FromRgb(88, 89, 91)

                            dse.ShadowDepth = 5 * (targetDPI / 96)
                            dse.BlurRadius = 15 * (targetDPI / 96)

                            Select Case shadow
                                Case "left"
                                    dse.Direction = -135
                                Case "right"
                                    dse.Direction = -45
                                Case "bottom"
                                    dse.Direction = -90
                            End Select

                            border.Effect = dse

                        End If

                        If r <> 0 Then
                            Dim rt As New RotateTransform()
                            rt.Angle = r
                            border.RenderTransform = rt
                        End If

                        placeholder.UpdateLayout()
                        mainspread.UpdateLayout()
                        grid.UpdateLayout()

                    Else
                        If element.Attributes.GetNamedItem("status").Value <> "empty" Then
                            Throw New Exception("No source found!! " + element.OuterXml)
                        End If
                    End If

                End If

                If element.Attributes.GetNamedItem("type").Value = "clipart" Then

                    Dim posx As Double = XmlConvert.ToDouble(element.Attributes.GetNamedItem("objectX").Value.ToString) * (targetDPI / 96)
                    Dim posy As Double = XmlConvert.ToDouble(element.Attributes.GetNamedItem("objectY").Value.ToString) * (targetDPI / 96)
                    Dim w As Double = Math.Round(XmlConvert.ToDouble(element.Attributes.GetNamedItem("objectWidth").Value.ToString) * (targetDPI / 96))
                    Dim h As Double = Math.Round(XmlConvert.ToDouble(element.Attributes.GetNamedItem("objectHeight").Value.ToString) * (targetDPI / 96))
                    Dim r As Double = XmlConvert.ToDouble(element.Attributes.GetNamedItem("rotation").Value.ToString)

                    Dim imageAlpha As Double = XmlConvert.ToDouble(element.Attributes.GetNamedItem("imageAlpha").Value.ToString)

                    Dim fliphorizontal As Integer = 0
                    If element.Attributes.GetNamedItem("fliphorizontal").Value <> "" Then
                        fliphorizontal = XmlConvert.ToInt32(element.Attributes.GetNamedItem("fliphorizontal").Value)
                    End If

                    Dim shadow As String = element.Attributes.GetNamedItem("shadow").Value
                    Dim useshadow As Boolean = False

                    If shadow <> "" Then
                        useshadow = True
                    End If

                    Dim borderweight As Integer = 0
                    If Not IsNothing(element.Attributes.GetNamedItem("borderweight")) Then
                        If element.Attributes.GetNamedItem("borderweight").Value > 0 Then
                            borderweight = Math.Round(XmlConvert.ToDouble(element.Attributes.GetNamedItem("borderweight").Value) * (targetDPI / 96))
                        End If
                    End If

                    Dim border As New Border()
                    border.Width = w
                    border.Height = h
                    border.BorderThickness = New Thickness(0)
                    border.HorizontalAlignment = Windows.HorizontalAlignment.Left
                    border.VerticalAlignment = Windows.VerticalAlignment.Top
                    border.Margin = New Thickness(posx, posy, 0, 0)
                    border.ClipToBounds = False
                    mainspread.Children.Add(border)

                    If borderweight > 0 Then
                        border.BorderThickness = New Thickness(borderweight)
                        border.Width += (borderweight * 2)
                        border.Height += (borderweight * 2)
                        Dim colorArr As ArrayList = GetColorRgb(element.Attributes.GetNamedItem("bordercolor").Value)
                        border.BorderBrush = New SolidColorBrush(System.Windows.Media.Color.FromRgb(colorArr(0), colorArr(1), colorArr(2)))
                    End If

                    'Add a placeholder for the image
                    Dim placeholder As New Canvas
                    placeholder.Width = w
                    placeholder.Height = h
                    placeholder.HorizontalAlignment = Windows.HorizontalAlignment.Left
                    placeholder.VerticalAlignment = Windows.VerticalAlignment.Top
                    placeholder.ClipToBounds = True
                    border.Child = placeholder

                    Dim filename As String = searchpath_imagefolder & element.Attributes.GetNamedItem("hires_url").Value
                    filename = filename.Replace("\", "/")

                    Dim c As New WebClient()
                    Dim bytes = c.DownloadData(New Uri(filename))
                    Dim ms As New MemoryStream(bytes)

                    Dim src As BitmapImage = New BitmapImage
                    src.BeginInit()
                    src.CreateOptions = BitmapCreateOptions.PreservePixelFormat Or BitmapCreateOptions.IgnoreColorProfile
                    src.StreamSource = ms
                    src.CacheOption = BitmapCacheOption.None
                    src.EndInit()

                    Dim srcimg As New System.Windows.Controls.Image
                    RenderOptions.SetBitmapScalingMode(srcimg, BitmapScalingMode.Fant)
                    srcimg.Source = src
                    srcimg.HorizontalAlignment = Windows.HorizontalAlignment.Left
                    srcimg.VerticalAlignment = Windows.VerticalAlignment.Top
                    srcimg.Width = w
                    srcimg.Height = h
                    srcimg.Stretch = Stretch.Fill
                    srcimg.Margin = New Thickness(0, 0, 0, 0)

                    If imageAlpha < 1 Then
                        srcimg.Opacity = imageAlpha
                    End If

                    If fliphorizontal = 1 Then
                        srcimg.RenderTransformOrigin = New System.Windows.Point(0.5, 0.5)
                        Dim flip As New ScaleTransform()
                        flip.ScaleX = -1
                        srcimg.RenderTransform = flip
                    End If

                    placeholder.Children.Add(srcimg)
                    placeholder.UpdateLayout()

                    If shadow <> "" Then

                        Dim dse As System.Windows.Media.Effects.DropShadowEffect = New System.Windows.Media.Effects.DropShadowEffect
                        dse.Color = System.Windows.Media.Color.FromRgb(88, 89, 91)
                        dse.ShadowDepth = 5
                        dse.BlurRadius = 15

                        Select Case shadow
                            Case "left"
                                dse.Direction = -135
                            Case "right"
                                dse.Direction = -45
                            Case "bottom"
                                dse.Direction = -90
                        End Select

                        border.Effect = dse

                    End If

                    If r <> 0 Then
                        Dim rt As New RotateTransform()
                        rt.Angle = r
                        border.RenderTransform = rt
                    End If

                    placeholder.UpdateLayout()
                    mainspread.UpdateLayout()
                    grid.UpdateLayout()

                End If

                If element.Attributes.GetNamedItem("type").Value = "rectangle" Then

                    Dim fillArr As ArrayList = GetColorRgb(element.Attributes.GetNamedItem("fillcolor").Value.ToString)
                    Dim borderArr As ArrayList = GetColorRgb(element.Attributes.GetNamedItem("bordercolor").Value.ToString)
                    Dim borderweight As Double = Math.Round(XmlConvert.ToDouble(element.Attributes.GetNamedItem("borderweight").Value.ToString) * (targetDPI / 96))
                    Dim r As Double = XmlConvert.ToDouble(element.Attributes.GetNamedItem("rotation").Value.ToString)
                    Dim fillalpha As Double = XmlConvert.ToDouble(element.Attributes.GetNamedItem("fillalpha").Value.ToString)
                    Dim shadow As String = element.Attributes.GetNamedItem("shadow").Value.ToString

                    Dim rect As New System.Windows.Shapes.Rectangle
                    rect.Width = Math.Round(XmlConvert.ToDouble(element.Attributes.GetNamedItem("objectWidth").Value) * (targetDPI / 96))
                    rect.Height = Math.Round(XmlConvert.ToDouble(element.Attributes.GetNamedItem("objectHeight").Value) * (targetDPI / 96))
                    rect.HorizontalAlignment = Windows.HorizontalAlignment.Left
                    rect.VerticalAlignment = Windows.VerticalAlignment.Top
                    rect.Fill = New SolidColorBrush(System.Windows.Media.Color.FromRgb(fillArr(0), fillArr(1), fillArr(2)))
                    rect.Stroke = New SolidColorBrush(System.Windows.Media.Color.FromRgb(borderArr(0), borderArr(1), borderArr(2)))
                    rect.StrokeThickness = borderweight
                    If borderweight > 0 Then
                        rect.Width += borderweight
                        rect.Height += borderweight
                    End If
                    rect.Margin = New Thickness((XmlConvert.ToDouble(element.Attributes.GetNamedItem("objectX").Value) * (targetDPI / 96)) - (borderweight / 2),
                                                 (XmlConvert.ToDouble(element.Attributes.GetNamedItem("objectY").Value) * (targetDPI / 96)) - (borderweight / 2),
                                                 0, 0)

                    If fillalpha < 1 Then
                        rect.Opacity = fillalpha
                    End If

                    If r <> 0 Then
                        Dim rt As New RotateTransform()
                        rt.Angle = r
                        rect.RenderTransform = rt
                    End If

                    If shadow <> "" Then

                        Dim dse As System.Windows.Media.Effects.DropShadowEffect = New System.Windows.Media.Effects.DropShadowEffect
                        dse.Color = System.Windows.Media.Color.FromRgb(88, 89, 91)
                        dse.ShadowDepth = 5
                        dse.BlurRadius = 15

                        Select Case shadow
                            Case "left"
                                dse.Direction = -135
                            Case "right"
                                dse.Direction = -45
                            Case "bottom"
                                dse.Direction = -90
                        End Select

                        rect.Effect = dse

                    End If

                    mainspread.Children.Add(rect)

                    mainspread.UpdateLayout()
                    grid.UpdateLayout()

                End If

                If element.Attributes.GetNamedItem("type").Value = "circle" Then

                    Dim fillArr As ArrayList = GetColorRgb(element.Attributes.GetNamedItem("fillcolor").Value.ToString)
                    Dim borderArr As ArrayList = GetColorRgb(element.Attributes.GetNamedItem("bordercolor").Value.ToString)
                    Dim borderweight As Double = Math.Round(XmlConvert.ToDouble(element.Attributes.GetNamedItem("borderweight").Value.ToString) * (targetDPI / 96))
                    Dim r As Double = XmlConvert.ToDouble(element.Attributes.GetNamedItem("rotation").Value.ToString)
                    Dim fillalpha As Double = XmlConvert.ToDouble(element.Attributes.GetNamedItem("fillalpha").Value.ToString)
                    Dim shadow As String = element.Attributes.GetNamedItem("shadow").Value.ToString

                    Dim ellipse As New Ellipse
                    ellipse.Width = Math.Round(XmlConvert.ToDouble(element.Attributes.GetNamedItem("objectWidth").Value) * (targetDPI / 96))
                    ellipse.Height = Math.Round(XmlConvert.ToDouble(element.Attributes.GetNamedItem("objectHeight").Value) * (targetDPI / 96))
                    ellipse.HorizontalAlignment = Windows.HorizontalAlignment.Left
                    ellipse.VerticalAlignment = Windows.VerticalAlignment.Top
                    ellipse.Fill = New SolidColorBrush(System.Windows.Media.Color.FromRgb(fillArr(0), fillArr(1), fillArr(2)))
                    ellipse.Stroke = New SolidColorBrush(System.Windows.Media.Color.FromRgb(borderArr(0), borderArr(1), borderArr(2)))
                    ellipse.StrokeThickness = borderweight
                    If borderweight > 0 Then
                        ellipse.Width += borderweight
                        ellipse.Height += borderweight
                    End If
                    ellipse.Margin = New Thickness((XmlConvert.ToDouble(element.Attributes.GetNamedItem("objectX").Value) * (targetDPI / 96)) - (borderweight / 2),
                                                 (XmlConvert.ToDouble(element.Attributes.GetNamedItem("objectY").Value) * (targetDPI / 96)) - (borderweight / 2),
                                                 0, 0)

                    If fillalpha < 1 Then
                        ellipse.Opacity = fillalpha
                    End If

                    If r <> 0 Then
                        Dim rt As New RotateTransform()
                        rt.Angle = r
                        ellipse.RenderTransform = rt
                    End If

                    If shadow <> "" Then

                        Dim dse As System.Windows.Media.Effects.DropShadowEffect = New System.Windows.Media.Effects.DropShadowEffect
                        dse.Color = System.Windows.Media.Color.FromRgb(88, 89, 91)
                        dse.ShadowDepth = 5
                        dse.BlurRadius = 15

                        Select Case shadow
                            Case "left"
                                dse.Direction = -135
                            Case "right"
                                dse.Direction = -45
                            Case "bottom"
                                dse.Direction = -90
                        End Select

                        ellipse.Effect = dse

                    End If

                    mainspread.Children.Add(ellipse)

                    mainspread.UpdateLayout()
                    grid.UpdateLayout()

                End If

                If element.Attributes.GetNamedItem("type").Value = "line" Then

                    Dim fillArr As ArrayList = GetColorRgb(element.Attributes.GetNamedItem("fillcolor").Value.ToString)
                    Dim lineweight As Double = Math.Round(XmlConvert.ToDouble(element.Attributes.GetNamedItem("lineweight").Value.ToString) * (targetDPI / 96))
                    Dim r As Double = XmlConvert.ToDouble(element.Attributes.GetNamedItem("rotation").Value.ToString)
                    Dim fillalpha As Double = XmlConvert.ToDouble(element.Attributes.GetNamedItem("fillalpha").Value.ToString)
                    Dim shadow As String = element.Attributes.GetNamedItem("shadow").Value.ToString

                    Dim line As New System.Windows.Shapes.Rectangle
                    line.Width = Math.Round(XmlConvert.ToDouble(element.Attributes.GetNamedItem("objectWidth").Value) * (targetDPI / 96))
                    line.Height = lineweight
                    line.HorizontalAlignment = Windows.HorizontalAlignment.Left
                    line.VerticalAlignment = Windows.VerticalAlignment.Top
                    line.Fill = New SolidColorBrush(System.Windows.Media.Color.FromRgb(fillArr(0), fillArr(1), fillArr(2)))
                    line.Margin = New Thickness((XmlConvert.ToDouble(element.Attributes.GetNamedItem("objectX").Value) * (targetDPI / 96)),
                                                 (XmlConvert.ToDouble(element.Attributes.GetNamedItem("objectY").Value) * (targetDPI / 96)),
                                                 0, 0)
                    If fillalpha < 1 Then
                        line.Opacity = fillalpha
                    End If

                    If r <> 0 Then
                        Dim rt As New RotateTransform()
                        rt.Angle = r
                        line.RenderTransform = rt
                    End If

                    If shadow <> "" Then

                        Dim dse As System.Windows.Media.Effects.DropShadowEffect = New System.Windows.Media.Effects.DropShadowEffect
                        dse.Color = System.Windows.Media.Color.FromRgb(88, 89, 91)
                        dse.ShadowDepth = 5
                        dse.BlurRadius = 15

                        Select Case shadow
                            Case "left"
                                dse.Direction = -135
                            Case "right"
                                dse.Direction = -45
                            Case "bottom"
                                dse.Direction = -90
                        End Select

                        line.Effect = dse

                    End If

                    mainspread.Children.Add(line)

                    mainspread.UpdateLayout()
                    grid.UpdateLayout()

                End If

            Next

            Application.Current.Dispatcher.Invoke(New Action(AddressOf SaveCover), DispatcherPriority.ContextIdle)

            'Save the image
            'savecovertimer = New DispatcherTimer()
            'AddHandler savecovertimer.Tick, AddressOf SaveCover
            'savecovertimer.Interval = New TimeSpan(0, 0, 5)
            'savecovertimer.Start()

        Catch ex As Exception

            ExitOrder(ex.Message)

        End Try

    End Sub

    Public Sub CreatePageBlock()

        Dim scale As New ScaleTransform()
        scale.ScaleX = 0.1
        scale.ScaleY = 0.1
        Me.LayoutTransform = scale

        Try

            GC.Collect()

            'Debug.Print("Memory used after full collection:   {0:N0}", GC.GetTotalMemory(True))

            Dim grid As New Canvas
            AddChild(grid)
            Me.UpdateLayout()

            Dim spread As XmlNode = spreadLst(renderindex)
            Dim spreadID As String = spread.Attributes.GetNamedItem("spreadID").Value

            Dim pages As XmlNodeList = spread.SelectNodes("descendant::page")

            Dim bleed As Double = XmlConvert.ToDouble(pages(0).Attributes.GetNamedItem("horizontalBleed").Value.ToString()) * (targetDPI / 96)

            Dim wrap As Double = 0
            If Not IsNothing(pages(0).Attributes.GetNamedItem("horizontalWrap")) Then
                wrap = XmlConvert.ToDouble(pages(0).Attributes.GetNamedItem("horizontalWrap").Value.ToString()) * (targetDPI / 96)
            End If

            Dim trimbox As String = ""
            Dim layflatmargin As Double = 0

            Dim width As Double = XmlConvert.ToDouble(spread.Attributes.GetNamedItem("totalWidth").Value.ToString()) * (targetDPI / 96)
            Dim height As Double = XmlConvert.ToDouble(spread.Attributes.GetNamedItem("totalHeight").Value.ToString()) * (targetDPI / 96)

            If singlepageproduct = True Then
                width = XmlConvert.ToDouble(pages(0).Attributes.GetNamedItem("pageWidth").Value.ToString()) * (targetDPI / 96)
                height = XmlConvert.ToDouble(pages(0).Attributes.GetNamedItem("pageHeight").Value.ToString()) * (targetDPI / 96)
                width += 2 * (bleed + wrap)
                height += 2 * (bleed + wrap)
            End If

            If isLayFlat = True And (renderindex = 1 Or renderindex = spreadLst.Count - 1) Then
                width = XmlConvert.ToDouble(spreadLst(2).Attributes.GetNamedItem("totalWidth").Value.ToString()) * (targetDPI / 96)
                If renderindex = 1 Then
                    layflatmargin = XmlConvert.ToDouble(spread.Attributes.GetNamedItem("totalWidth").Value.ToString()) * (targetDPI / 96)
                    'Extra bleed correction for layflat
                    layflatmargin -= (2 * bleed)
                Else
                    layflatmargin = 0
                End If
            End If

            Dim mainspread As New Canvas
            mainspread.Width = width
            mainspread.Height = height
            mainspread.HorizontalAlignment = Windows.HorizontalAlignment.Left
            mainspread.VerticalAlignment = Windows.VerticalAlignment.Top
            mainspread.Background = New SolidColorBrush(System.Windows.Media.Color.FromRgb(255, 255, 255))
            mainspread.ClipToBounds = True
            grid.Children.Add(mainspread)

            mainspread.UpdateLayout()

            If spread.Attributes.GetNamedItem("backgroundColor").Value.ToString <> "-1" Then

                Dim fillArr As ArrayList = GetColorRgb(spread.Attributes.GetNamedItem("backgroundColor").Value.ToString)

                Dim bg As New Canvas
                bg.Width = width - layflatmargin
                bg.Height = height
                bg.Background = New SolidColorBrush(System.Windows.Media.Color.FromRgb(fillArr(0), fillArr(1), fillArr(2)))
                bg.HorizontalAlignment = Windows.HorizontalAlignment.Left
                bg.VerticalAlignment = Windows.VerticalAlignment.Top
                bg.Margin = New Thickness(layflatmargin, 0, 0, 0)
                mainspread.Children.Add(bg)

            End If

            If Not spread.SelectSingleNode("background") Is Nothing Then

                Dim background As XmlNode = spread.SelectSingleNode("background")
                Dim filename As String
                Dim continueImage As Boolean = True

                If background.Attributes.GetNamedItem("status").Value <> "done" Then
                    If background.Attributes.GetNamedItem("status").Value = "new" Or
                        background.Attributes.GetNamedItem("status").Value = "" Then
                        'Get the filedata from the upload, if its available
                        Dim arr As ArrayList = GetFileUrlFromUpload(background.Attributes.GetNamedItem("id").Value)
                        If arr.Count > 0 Then
                            filename = searchpath_imagefolder & arr(0)
                            filename = filename.Replace("\", "/")
                        Else
                            continueImage = False
                        End If
                    ElseIf background.Attributes.GetNamedItem("status").Value = "empty" Then
                        continueImage = False
                    End If
                Else
                    filename = searchpath_imagefolder & background.Attributes.GetNamedItem("hires_url").Value
                    filename = filename.Replace("\", "/")
                End If

                'Debug.Print(filename)

                If (continueImage) Then

                    Dim alpha As Double = XmlConvert.ToDouble(spread.Attributes.GetNamedItem("backgroundAlpha").Value.ToString)
                    Dim fliphorizontal As Integer = 0
                    If background.Attributes.GetNamedItem("fliphorizontal").Value <> "" Then
                        fliphorizontal = XmlConvert.ToInt32(background.Attributes.GetNamedItem("fliphorizontal").Value)
                    End If
                    Dim imgRotation As Integer = XmlConvert.ToInt32(background.Attributes.GetNamedItem("imageRotation").Value)
                    Dim imgWidth As Double = XmlConvert.ToDouble(background.Attributes.GetNamedItem("width").Value) * (targetDPI / 96)
                    Dim imgHeight As Double = XmlConvert.ToDouble(background.Attributes.GetNamedItem("height").Value) * (targetDPI / 96)
                    Dim offsetX As Double = XmlConvert.ToDouble(background.Attributes.GetNamedItem("x").Value) * (targetDPI / 96)
                    Dim offsetY As Double = XmlConvert.ToDouble(background.Attributes.GetNamedItem("y").Value) * (targetDPI / 96)

                    Dim c As New WebClient()
                    Dim bytes = c.DownloadData(New Uri(filename))
                    Dim ms As New MemoryStream(bytes)

                    Dim src As BitmapImage = New BitmapImage
                    src.BeginInit()
                    src.CreateOptions = BitmapCreateOptions.PreservePixelFormat Or BitmapCreateOptions.IgnoreColorProfile
                    src.StreamSource = ms
                    src.CacheOption = BitmapCacheOption.None

                    Select Case imgRotation
                        Case 90

                            src.Rotation = Rotation.Rotate90

                            Dim iw As Double = imgWidth
                            Dim xo As Double = offsetX

                            imgWidth = imgHeight
                            imgHeight = iw

                            'Check if the image was flipped
                            If fliphorizontal = 0 Then
                                offsetX = width - (imgWidth + offsetY)
                                offsetY = xo
                            Else
                                offsetX = width - (imgWidth + offsetY)
                                offsetY = height - (imgHeight + xo)
                            End If

                        Case 180

                            src.Rotation = Rotation.Rotate180

                            If fliphorizontal = 0 Then
                                offsetX = width - (imgWidth + offsetX)
                                offsetY = height - (imgHeight + offsetY)
                            Else
                                offsetY = height - (imgHeight + offsetY)
                            End If

                        Case 270

                            src.Rotation = Rotation.Rotate270

                            Dim iw As Double = imgWidth
                            Dim xo As Double = offsetX

                            imgWidth = imgHeight
                            imgHeight = iw

                            'Check if the image was flipped
                            If fliphorizontal = 0 Then
                                offsetX = offsetY
                                offsetY = height - (imgHeight + xo)
                            Else
                                offsetX = offsetY
                                offsetY = xo
                            End If

                        Case -90

                            src.Rotation = Rotation.Rotate270


                            Dim iw As Double = imgWidth
                            Dim xo As Double = offsetX

                            imgWidth = imgHeight
                            imgHeight = iw

                            'Check if the image was flipped
                            If fliphorizontal = 0 Then
                                offsetX = offsetY
                                offsetY = height - (imgHeight + xo)
                            Else
                                offsetX = offsetY
                                offsetY = xo
                            End If

                        Case 0

                            src.Rotation = Rotation.Rotate0

                    End Select
                    src.EndInit()

                    Dim placeholder As New Canvas
                    placeholder.Width = width - layflatmargin
                    placeholder.Height = height
                    placeholder.HorizontalAlignment = Windows.HorizontalAlignment.Left
                    placeholder.VerticalAlignment = Windows.VerticalAlignment.Top
                    placeholder.ClipToBounds = True
                    placeholder.Margin = New Thickness(layflatmargin, 0, 0, 0)
                    mainspread.Children.Add(placeholder)

                    Dim srcimg As New System.Windows.Controls.Image
                    RenderOptions.SetBitmapScalingMode(srcimg, BitmapScalingMode.Fant)
                    Select Case background.Attributes.GetNamedItem("imageFilter").Value
                        Case "" 'Default
                            srcimg.Source = src
                        Case "bw"
                            srcimg.Source = MakeGrayscale(src)
                        Case "sepia"
                            srcimg.Source = src
                    End Select

                    If fliphorizontal = 1 Then
                        srcimg.RenderTransformOrigin = New System.Windows.Point(0.5, 0.5)
                        Dim flip As New ScaleTransform()
                        Select Case imgRotation
                            Case 0
                                offsetX = width - (imgWidth + offsetX)
                                flip.ScaleX = -1
                            Case 90
                                flip.ScaleY = -1
                            Case 180
                                flip.ScaleX = -1
                            Case 270
                                flip.ScaleY = -1
                        End Select
                        srcimg.RenderTransform = flip
                    End If

                    srcimg.HorizontalAlignment = Windows.HorizontalAlignment.Left
                    srcimg.VerticalAlignment = Windows.VerticalAlignment.Top
                    srcimg.Width = imgWidth
                    srcimg.Height = imgHeight
                    srcimg.Margin = New Thickness(offsetX, offsetY, 0, 0)

                    If background.Attributes.GetNamedItem("imageFilter").Value = "sepia" Then
                        srcimg = MakeSepia(srcimg)
                    End If

                    If alpha < 1 Then
                        srcimg.Opacity = alpha
                    End If

                    placeholder.Children.Add(srcimg)
                    placeholder.UpdateLayout()

                    ms = Nothing
                    src = Nothing
                    srcimg = Nothing

                Else
                    If background.Attributes.GetNamedItem("status").Value <> "empty" Then
                        Throw New Exception("No source found!! " + background.OuterXml)
                    End If
                End If
            End If

            'Add separate backgrounds if available
            Dim xpos As Double = 0
            Dim spineX As Double = 0

            For Each page As XmlNode In pages

                Dim position As String = "center"

                Dim pageID As String = page.Attributes.GetNamedItem("pageID").Value
                Dim pagewidth As Double = XmlConvert.ToDouble(page.Attributes.GetNamedItem("width").Value) * (targetDPI / 96)
                Dim pageheight As Double = XmlConvert.ToDouble(page.Attributes.GetNamedItem("height").Value) * (targetDPI / 96)

                If page.Attributes.GetNamedItem("backgroundColor").Value.ToString <> "-1" Then

                    Dim fillArr As ArrayList = GetColorRgb(page.Attributes.GetNamedItem("backgroundColor").Value.ToString)

                    Dim bg As New Canvas
                    bg.Width = pagewidth
                    bg.Height = pageheight
                    bg.Background = New SolidColorBrush(System.Windows.Media.Color.FromRgb(fillArr(0), fillArr(1), fillArr(2)))
                    bg.HorizontalAlignment = Windows.HorizontalAlignment.Left
                    bg.VerticalAlignment = Windows.VerticalAlignment.Top
                    If isLayFlat And layflatmargin > 0 Then
                        bg.Margin = New Thickness(layflatmargin, 0, 0, 0)
                    Else
                        bg.Margin = New Thickness(xpos, 0, 0, 0)
                    End If

                    mainspread.Children.Add(bg)

                End If

                If isLayFlat = True And renderindex = spreadLst.Count - 1 Then
                    'Force a white background for the right page
                    Dim bg As New Canvas
                    bg.Width = pagewidth
                    bg.Height = pageheight
                    bg.Background = New SolidColorBrush(System.Windows.Media.Color.FromRgb(255, 255, 255))
                    bg.HorizontalAlignment = Windows.HorizontalAlignment.Right
                    bg.VerticalAlignment = Windows.VerticalAlignment.Top
                    bg.Margin = New Thickness(pagewidth, 0, 0, 0)
                    mainspread.Children.Add(bg)
                End If

                If Not page.SelectSingleNode("background") Is Nothing Then

                    Dim background As XmlNode = page.SelectSingleNode("background")

                    Dim filename As String
                    Dim continueImage As Boolean = True

                    If background.Attributes.GetNamedItem("status").Value <> "done" Then
                        If background.Attributes.GetNamedItem("status").Value = "new" Or
                            background.Attributes.GetNamedItem("status").Value = "" Then
                            'Get the filedata from the upload, if its available
                            Dim arr As ArrayList = GetFileUrlFromUpload(background.Attributes.GetNamedItem("id").Value)
                            If arr.Count > 0 Then
                                filename = searchpath_imagefolder & arr(0)
                                filename = filename.Replace("\", "/")
                            Else
                                continueImage = False
                            End If
                        ElseIf background.Attributes.GetNamedItem("status").Value = "empty" Then
                            continueImage = False
                        End If
                    Else
                        filename = searchpath_imagefolder & background.Attributes.GetNamedItem("hires_url").Value
                        filename = filename.Replace("\", "/")
                    End If

                    If (continueImage) Then

                        Dim alpha As Double = XmlConvert.ToDouble(page.Attributes.GetNamedItem("backgroundAlpha").Value.ToString)
                        Dim fliphorizontal As Integer = 0
                        If background.Attributes.GetNamedItem("fliphorizontal").Value <> "" Then
                            fliphorizontal = XmlConvert.ToInt32(background.Attributes.GetNamedItem("fliphorizontal").Value)
                        End If
                        Dim imgRotation As Integer = XmlConvert.ToInt32(background.Attributes.GetNamedItem("imageRotation").Value)
                        Dim imgWidth As Double = XmlConvert.ToDouble(background.Attributes.GetNamedItem("width").Value) * (targetDPI / 96)
                        Dim imgHeight As Double = XmlConvert.ToDouble(background.Attributes.GetNamedItem("height").Value) * (targetDPI / 96)
                        Dim offsetX As Double = XmlConvert.ToDouble(background.Attributes.GetNamedItem("x").Value) * (targetDPI / 96)
                        Dim offsetY As Double = XmlConvert.ToDouble(background.Attributes.GetNamedItem("y").Value) * (targetDPI / 96)

                        Dim c As New WebClient()
                        Dim bytes = c.DownloadData(New Uri(filename))
                        Dim ms As New MemoryStream(bytes)

                        Dim src As BitmapImage = New BitmapImage
                        src.BeginInit()
                        src.CreateOptions = BitmapCreateOptions.PreservePixelFormat Or BitmapCreateOptions.IgnoreColorProfile
                        src.StreamSource = ms
                        src.CacheOption = BitmapCacheOption.None

                        Select Case imgRotation

                            Case 90

                                src.Rotation = Rotation.Rotate90

                                Dim iw As Double = imgWidth
                                Dim xo As Double = offsetX

                                imgWidth = imgHeight
                                imgHeight = iw

                                'Check if the image was flipped
                                If fliphorizontal = 0 Then
                                    offsetX = pagewidth - (imgWidth + offsetY)
                                    offsetY = xo
                                Else
                                    offsetX = pagewidth - (imgWidth + offsetY)
                                    offsetY = pageheight - (imgHeight + xo)
                                End If

                            Case 180

                                src.Rotation = Rotation.Rotate180

                                If fliphorizontal = 0 Then
                                    offsetX = pagewidth - (imgWidth + offsetX)
                                    offsetY = pageheight - (imgHeight + offsetY)
                                Else
                                    offsetY = pageheight - (imgHeight + offsetY)
                                End If

                            Case 270

                                src.Rotation = Rotation.Rotate270

                                Dim iw As Integer = imgWidth
                                Dim xo As Integer = offsetX

                                imgWidth = imgHeight
                                imgHeight = iw

                                'Check if the image was flipped
                                If fliphorizontal = 0 Then
                                    offsetX = offsetY
                                    offsetY = pageheight - (imgHeight + xo)
                                Else
                                    offsetX = offsetY
                                    offsetY = xo
                                End If

                            Case -90

                                src.Rotation = Rotation.Rotate270

                                Dim iw As Integer = imgWidth
                                Dim xo As Integer = offsetX

                                imgWidth = imgHeight
                                imgHeight = iw

                                'Check if the image was flipped
                                If fliphorizontal = 0 Then
                                    offsetX = offsetY
                                    offsetY = pageheight - (imgHeight + xo)
                                Else
                                    offsetX = offsetY
                                    offsetY = xo
                                End If

                            Case 0

                                src.Rotation = Rotation.Rotate0

                        End Select
                        src.EndInit()

                        Dim placeholder As New Canvas
                        placeholder.Width = pagewidth
                        placeholder.Height = pageheight
                        placeholder.HorizontalAlignment = Windows.HorizontalAlignment.Left
                        placeholder.VerticalAlignment = Windows.VerticalAlignment.Top

                        If isLayFlat And layflatmargin > 0 Then
                            placeholder.Margin = New Thickness(layflatmargin, 0, 0, 0)
                        Else
                            placeholder.Margin = New Thickness(xpos, 0, 0, 0)
                        End If

                        placeholder.ClipToBounds = True
                        mainspread.Children.Add(placeholder)

                        Dim srcimg As New System.Windows.Controls.Image
                        RenderOptions.SetBitmapScalingMode(srcimg, BitmapScalingMode.Fant)
                        Select Case background.Attributes.GetNamedItem("imageFilter").Value
                            Case "" 'Default
                                srcimg.Source = src
                            Case "bw"
                                srcimg.Source = MakeGrayscale(src)
                            Case "sepia"
                                srcimg.Source = src
                        End Select

                        srcimg.HorizontalAlignment = Windows.HorizontalAlignment.Left
                        srcimg.VerticalAlignment = Windows.VerticalAlignment.Top
                        srcimg.Width = imgWidth
                        srcimg.Height = imgHeight

                        If fliphorizontal = 1 Then
                            srcimg.RenderTransformOrigin = New System.Windows.Point(0.5, 0.5)
                            Dim flip As New ScaleTransform()
                            Select Case imgRotation
                                Case 0
                                    offsetX = pagewidth - (imgWidth + offsetX)
                                    flip.ScaleX = -1
                                Case 90
                                    flip.ScaleY = -1
                                Case 180
                                    flip.ScaleX = -1
                                Case 270
                                    flip.ScaleY = -1
                            End Select
                            srcimg.RenderTransform = flip
                        End If

                        srcimg.Margin = New Thickness(offsetX, offsetY, 0, 0)

                        If background.Attributes.GetNamedItem("imageFilter").Value = "sepia" Then
                            srcimg = MakeSepia(srcimg)
                        End If

                        If alpha < 1 Then
                            srcimg.Opacity = alpha
                        End If

                        placeholder.Children.Add(srcimg)
                        placeholder.UpdateLayout()

                        ms = Nothing
                        src = Nothing
                        srcimg = Nothing

                    Else
                        If background.Attributes.GetNamedItem("status").Value <> "empty" Then
                            Throw New Exception("No source found!! " + background.OuterXml)
                        End If
                    End If
                End If

                xpos += pagewidth

            Next

            'Add the elements
            Dim elements As XmlNodeList = spread.SelectSingleNode("elements").SelectNodes("element")

            For Each element As XmlNode In elements

                Dim elementID As String = element.Attributes.GetNamedItem("id").Value

                If element.Attributes.GetNamedItem("type").Value = "photo" Then

                    Dim filename As String = ""

                    Dim continueImage As Boolean = True

                    If element.Attributes.GetNamedItem("status").Value <> "done" Then
                        If element.Attributes.GetNamedItem("status").Value = "new" Or
                            element.Attributes.GetNamedItem("status").Value = "" Then
                            'Get the filedata from the upload, if its available
                            Dim arr As ArrayList = GetFileUrlFromUpload(element.Attributes.GetNamedItem("original_image_id").Value)
                            If arr.Count > 0 Then
                                filename = searchpath_imagefolder & arr(0)
                                filename = filename.Replace("\", "/")
                                Debug.Print(filename)
                            Else
                                continueImage = False
                            End If
                        ElseIf element.Attributes.GetNamedItem("status").Value = "empty" Then
                            continueImage = False
                        End If
                    Else
                        If (element.Attributes.GetNamedItem("hires_url").Value <> "") Then
                            filename = searchpath_imagefolder & element.Attributes.GetNamedItem("hires_url").Value
                            filename = filename.Replace("\", "/")
                            Debug.Print(filename)
                        Else
                            Dim arr As ArrayList = GetFileUrlFromUpload(element.Attributes.GetNamedItem("original_image_id").Value)
                            If arr.Count > 0 Then
                                filename = searchpath_imagefolder & arr(0)
                                filename = filename.Replace("\", "/")
                                Debug.Print(filename)
                            Else
                                continueImage = False
                            End If
                        End If
                    End If

                    If continueImage Then

                        Dim posx As Double = XmlConvert.ToDouble(element.Attributes.GetNamedItem("objectX").Value.ToString) * (targetDPI / 96)
                        Dim posy As Double = XmlConvert.ToDouble(element.Attributes.GetNamedItem("objectY").Value.ToString) * (targetDPI / 96)
                        Dim w As Double = XmlConvert.ToDouble(element.Attributes.GetNamedItem("objectWidth").Value.ToString) * (targetDPI / 96)
                        Dim h As Double = XmlConvert.ToDouble(element.Attributes.GetNamedItem("objectHeight").Value.ToString) * (targetDPI / 96)
                        Dim r As Double = XmlConvert.ToDouble(element.Attributes.GetNamedItem("rotation").Value.ToString)

                        If singlepageproduct = True Then
                            w = width - (2 * wrap)
                            h = height - (2 * wrap)
                            posx += wrap
                            posy += wrap
                        End If

                        Dim offsetX As Double = XmlConvert.ToDouble(element.Attributes.GetNamedItem("offsetX").Value.ToString) * (targetDPI / 96)
                        Dim offsetY As Double = XmlConvert.ToDouble(element.Attributes.GetNamedItem("offsetY").Value.ToString) * (targetDPI / 96)
                        Dim imgWidth As Double = XmlConvert.ToDouble(element.Attributes.GetNamedItem("imageWidth").Value) * (targetDPI / 96)
                        Dim imgHeight As Double = XmlConvert.ToDouble(element.Attributes.GetNamedItem("imageHeight").Value) * (targetDPI / 96)
                        Dim imgRotation As Double = XmlConvert.ToDouble(element.Attributes.GetNamedItem("imageRotation").Value)

                        Dim imageMask As String = element.Attributes.GetNamedItem("mask_hires_url").Value.ToString
                        Dim imageOverlay As String = element.Attributes.GetNamedItem("overlay_hires_url").Value.ToString
                        Dim imageAlpha As Double = XmlConvert.ToDouble(element.Attributes.GetNamedItem("imageAlpha").Value.ToString)

                        Dim fliphorizontal As Integer = 0
                        If element.Attributes.GetNamedItem("fliphorizontal").Value <> "" Then
                            fliphorizontal = XmlConvert.ToInt32(element.Attributes.GetNamedItem("fliphorizontal").Value)
                        End If

                        Dim shadow As String = ""
                        If Not IsNothing(element.Attributes.GetNamedItem("shadow")) Then
                            shadow = element.Attributes.GetNamedItem("shadow").Value
                        End If

                        Dim useshadow As Boolean = False

                        If shadow <> "" Then
                            useshadow = True
                        End If

                        Dim borderweight As Integer = 0
                        If Not IsNothing(element.Attributes.GetNamedItem("borderweight")) Then
                            If element.Attributes.GetNamedItem("borderweight").Value > 0 Then
                                borderweight = Math.Round(XmlConvert.ToDouble(element.Attributes.GetNamedItem("borderweight").Value) * (targetDPI / 96))
                            End If
                        End If

                        Dim border As New Border()
                        border.Width = w
                        border.Height = h
                        border.BorderThickness = New Thickness(0)
                        border.HorizontalAlignment = Windows.HorizontalAlignment.Left
                        border.VerticalAlignment = Windows.VerticalAlignment.Top
                        border.Margin = New Thickness(layflatmargin + posx - borderweight, posy - borderweight, 0, 0)
                        border.ClipToBounds = False
                        mainspread.Children.Add(border)

                        If borderweight > 0 Then
                            border.BorderThickness = New Thickness(borderweight)
                            border.Width += (borderweight * 2)
                            border.Height += (borderweight * 2)
                            Dim colorArr As ArrayList = GetColorRgb(element.Attributes.GetNamedItem("bordercolor").Value)
                            border.BorderBrush = New SolidColorBrush(System.Windows.Media.Color.FromRgb(colorArr(0), colorArr(1), colorArr(2)))

                            imgWidth += borderweight
                            imgHeight += borderweight
                        End If

                        'Add a placeholder for the image
                        Dim placeholder As New Canvas
                        placeholder.Width = w
                        placeholder.Height = h
                        placeholder.HorizontalAlignment = Windows.HorizontalAlignment.Left
                        placeholder.VerticalAlignment = Windows.VerticalAlignment.Top
                        placeholder.ClipToBounds = True
                        border.Child = placeholder

                        Dim c As New WebClient()
                        Dim bytes = c.DownloadData(New Uri(filename))
                        Dim ms As New MemoryStream(bytes)

                        Dim src As BitmapImage = New BitmapImage
                        src.BeginInit()
                        src.CreateOptions = BitmapCreateOptions.PreservePixelFormat Or BitmapCreateOptions.IgnoreColorProfile
                        src.StreamSource = ms
                        src.CacheOption = BitmapCacheOption.None

                        Select Case imgRotation

                            Case 90

                                src.Rotation = Rotation.Rotate90

                                Dim iw As Integer = imgWidth
                                Dim xo As Integer = offsetX

                                imgWidth = imgHeight
                                imgHeight = iw

                                'Check if the image was flipped
                                If fliphorizontal = 0 Then
                                    If singlepageproduct = False Then
                                        offsetX = w - (imgWidth + offsetY)
                                        offsetY = xo
                                    End If

                                Else
                                    offsetX = w - (imgWidth + offsetY)
                                    offsetY = h - (imgHeight + xo)
                                End If

                            Case 180

                                src.Rotation = Rotation.Rotate180

                                If fliphorizontal = 0 Then
                                    offsetX = w - (imgWidth + offsetX)
                                    offsetY = h - (imgHeight + offsetY)
                                Else
                                    offsetY = h - (imgHeight + offsetY)
                                End If

                            Case 270

                                src.Rotation = Rotation.Rotate270

                                Dim iw As Integer = imgWidth
                                Dim xo As Integer = offsetX

                                imgWidth = imgHeight
                                imgHeight = iw

                                'Check if the image was flipped
                                If fliphorizontal = 0 Then
                                    If singlepageproduct = False Then
                                        offsetX = offsetY
                                        offsetY = h - (imgHeight + xo)
                                    End If
                                Else
                                    offsetX = offsetY
                                    offsetY = xo
                                End If

                            Case -90

                                src.Rotation = Rotation.Rotate270

                                Dim iw As Integer = imgWidth
                                Dim xo As Integer = offsetX

                                imgWidth = imgHeight
                                imgHeight = iw

                                'Check if the image was flipped
                                If fliphorizontal = 0 Then
                                    offsetX = offsetY
                                    offsetY = h - (imgHeight + xo)
                                Else
                                    offsetX = offsetY
                                    offsetY = xo
                                End If

                            Case 0

                                src.Rotation = Rotation.Rotate0

                        End Select
                        src.EndInit()

                        Dim srcimg As New System.Windows.Controls.Image
                        RenderOptions.SetBitmapScalingMode(srcimg, BitmapScalingMode.Fant)
                        Select Case element.Attributes.GetNamedItem("imageFilter").Value
                            Case "" 'Default
                                srcimg.Source = src
                            Case "bw"
                                srcimg.Source = MakeGrayscale(src)
                            Case "sepia"
                                srcimg.Source = src
                        End Select
                        srcimg.HorizontalAlignment = Windows.HorizontalAlignment.Left
                        srcimg.VerticalAlignment = Windows.VerticalAlignment.Top
                        srcimg.Width = imgWidth
                        srcimg.Height = imgHeight

                        If fliphorizontal = 1 Then
                            srcimg.RenderTransformOrigin = New System.Windows.Point(0.5, 0.5)
                            Dim flip As New ScaleTransform()
                            Select Case imgRotation
                                Case 0
                                    offsetX = w - (imgWidth + offsetX)
                                    flip.ScaleX = -1
                                Case 90
                                    flip.ScaleY = -1
                                Case 180
                                    flip.ScaleX = -1
                                Case 270
                                    flip.ScaleY = -1
                            End Select
                            srcimg.RenderTransform = flip
                        End If

                        srcimg.Margin = New Thickness(offsetX, offsetY, 0, 0)

                        If element.Attributes.GetNamedItem("imageFilter").Value = "sepia" Then
                            srcimg = MakeSepia(srcimg)
                        End If

                        If imageAlpha < 1 Then
                            placeholder.Opacity = imageAlpha
                        End If

                        placeholder.Children.Add(srcimg)
                        placeholder.UpdateLayout()

                        If imageMask <> "" Then

                            Dim mask As New ImageBrush()
                            Dim path As String = searchpath_imagefolder & imageMask
                            path = path.Replace("\", "/")
                            mask.ImageSource = New BitmapImage(New Uri(path))
                            mask.Viewport = New Rect(0, 0, w, h)
                            mask.TileMode = TileMode.None
                            mask.ViewportUnits = BrushMappingMode.Absolute
                            placeholder.OpacityMask = mask

                        End If

                        Dim overlay As New System.Windows.Controls.Image

                        If imageOverlay <> "" Then

                            Dim overlaypath As String = searchpath_imagefolder & imageOverlay
                            overlaypath = overlaypath.Replace("\", "/")

                            Dim overlaywc As New WebClient()
                            Dim overlaybytes = overlaywc.DownloadData(New Uri(overlaypath))
                            Dim overlayms As New MemoryStream(overlaybytes)

                            Dim overlaysrc As BitmapImage = New BitmapImage
                            overlaysrc.BeginInit()
                            overlaysrc.CreateOptions = BitmapCreateOptions.PreservePixelFormat Or BitmapCreateOptions.IgnoreColorProfile
                            overlaysrc.CacheOption = BitmapCacheOption.None
                            overlaysrc.StreamSource = overlayms
                            overlaysrc.EndInit()

                            overlay = New System.Windows.Controls.Image
                            RenderOptions.SetBitmapScalingMode(overlay, BitmapScalingMode.Fant)
                            overlay.Source = overlaysrc
                            overlay.Stretch = Stretch.Fill
                            overlay.HorizontalAlignment = Windows.HorizontalAlignment.Left
                            overlay.VerticalAlignment = Windows.VerticalAlignment.Top
                            overlay.Width = w
                            overlay.Height = h
                            overlay.Margin = New Thickness(layflatmargin + posx, posy, 0, 0)

                            If imageAlpha < 1 Then
                                overlay.Opacity = imageAlpha
                            End If

                            mainspread.Children.Add(overlay)

                            overlayms = Nothing
                            overlaysrc = Nothing

                        End If

                        If shadow <> "" Then

                            Dim dse As System.Windows.Media.Effects.DropShadowEffect = New System.Windows.Media.Effects.DropShadowEffect
                            dse.Color = System.Windows.Media.Color.FromRgb(88, 89, 91)
                            dse.ShadowDepth = 5 * (targetDPI / 96)
                            dse.BlurRadius = 15 * (targetDPI / 96)

                            Select Case shadow
                                Case "left"
                                    dse.Direction = -135
                                Case "right"
                                    dse.Direction = -45
                                Case "bottom"
                                    dse.Direction = -90
                            End Select

                            border.Effect = dse

                        End If

                        If r <> 0 Then
                            Dim rt As New RotateTransform()
                            rt.Angle = r
                            border.RenderTransform = rt

                            If imageOverlay <> "" Then
                                overlay.RenderTransform = rt
                            End If
                        End If

                        placeholder.UpdateLayout()
                        mainspread.UpdateLayout()
                        grid.UpdateLayout()

                        ms = Nothing
                        src = Nothing
                        srcimg = Nothing

                    Else

                        If element.Attributes.GetNamedItem("status").Value <> "empty" Then
                            Throw New Exception("No source found!! " + element.OuterXml)
                        End If

                    End If

                End If

                If element.Attributes.GetNamedItem("type").Value = "clipart" Then

                    Dim posx As Double = XmlConvert.ToDouble(element.Attributes.GetNamedItem("objectX").Value.ToString) * (targetDPI / 96)
                    Dim posy As Double = XmlConvert.ToDouble(element.Attributes.GetNamedItem("objectY").Value.ToString) * (targetDPI / 96)
                    Dim w As Double = Math.Round(XmlConvert.ToDouble(element.Attributes.GetNamedItem("objectWidth").Value.ToString) * (targetDPI / 96))
                    Dim h As Double = Math.Round(XmlConvert.ToDouble(element.Attributes.GetNamedItem("objectHeight").Value.ToString) * (targetDPI / 96))
                    Dim r As Double = XmlConvert.ToDouble(element.Attributes.GetNamedItem("rotation").Value.ToString)

                    Dim imageAlpha As Double = XmlConvert.ToDouble(element.Attributes.GetNamedItem("imageAlpha").Value.ToString)

                    Dim fliphorizontal As Integer = 0
                    If element.Attributes.GetNamedItem("fliphorizontal").Value <> "" Then
                        fliphorizontal = XmlConvert.ToInt32(element.Attributes.GetNamedItem("fliphorizontal").Value)
                    End If

                    Dim shadow As String = element.Attributes.GetNamedItem("shadow").Value
                    Dim useshadow As Boolean = False

                    If shadow <> "" Then
                        useshadow = True
                    End If

                    Dim borderweight As Integer = 0
                    If Not IsNothing(element.Attributes.GetNamedItem("borderweight")) Then
                        If element.Attributes.GetNamedItem("borderweight").Value > 0 Then
                            borderweight = Math.Round(XmlConvert.ToDouble(element.Attributes.GetNamedItem("borderweight").Value) * (targetDPI / 96))
                        End If
                    End If

                    Dim border As New Border()
                    border.Width = w
                    border.Height = h
                    border.BorderThickness = New Thickness(0)
                    border.HorizontalAlignment = Windows.HorizontalAlignment.Left
                    border.VerticalAlignment = Windows.VerticalAlignment.Top
                    border.Margin = New Thickness(layflatmargin + posx, posy, 0, 0)
                    border.ClipToBounds = False
                    mainspread.Children.Add(border)

                    If useshadow = True Then 'Draw the border if we also have a shadow
                        If borderweight > 0 Then
                            border.BorderThickness = New Thickness(borderweight)
                            border.Width += (borderweight * 2)
                            border.Height += (borderweight * 2)
                            Dim colorArr As ArrayList = GetColorRgb(element.Attributes.GetNamedItem("bordercolor").Value)
                            border.BorderBrush = New SolidColorBrush(System.Windows.Media.Color.FromRgb(colorArr(0), colorArr(1), colorArr(2)))
                        End If
                    End If

                    'Add a placeholder for the image
                    Dim placeholder As New Canvas
                    placeholder.Width = w
                    placeholder.Height = h
                    placeholder.HorizontalAlignment = Windows.HorizontalAlignment.Left
                    placeholder.VerticalAlignment = Windows.VerticalAlignment.Top
                    placeholder.ClipToBounds = True
                    border.Child = placeholder

                    Dim filename As String = searchpath_imagefolder & element.Attributes.GetNamedItem("hires_url").Value
                    filename = filename.Replace("\", "/")

                    Dim c As New WebClient()
                    Dim bytes = c.DownloadData(New Uri(filename))
                    Dim ms As New MemoryStream(bytes)

                    Dim src As BitmapImage = New BitmapImage
                    src.BeginInit()
                    src.CreateOptions = BitmapCreateOptions.PreservePixelFormat Or BitmapCreateOptions.IgnoreColorProfile
                    src.StreamSource = ms
                    src.CacheOption = BitmapCacheOption.None
                    src.EndInit()

                    Dim srcimg As New System.Windows.Controls.Image
                    RenderOptions.SetBitmapScalingMode(srcimg, BitmapScalingMode.Fant)
                    srcimg.Source = src
                    srcimg.HorizontalAlignment = Windows.HorizontalAlignment.Left
                    srcimg.VerticalAlignment = Windows.VerticalAlignment.Top
                    srcimg.Width = w
                    srcimg.Height = h
                    srcimg.Stretch = Stretch.Fill
                    srcimg.Margin = New Thickness(0, 0, 0, 0)

                    If imageAlpha < 1 Then
                        srcimg.Opacity = imageAlpha
                    End If

                    If fliphorizontal = 1 Then
                        srcimg.RenderTransformOrigin = New System.Windows.Point(0.5, 0.5)
                        Dim flip As New ScaleTransform()
                        flip.ScaleX = -1
                        srcimg.RenderTransform = flip
                    End If

                    placeholder.Children.Add(srcimg)
                    placeholder.UpdateLayout()

                    If shadow <> "" Then

                        Dim dse As System.Windows.Media.Effects.DropShadowEffect = New System.Windows.Media.Effects.DropShadowEffect
                        dse.Color = System.Windows.Media.Color.FromRgb(88, 89, 91)
                        dse.ShadowDepth = 5 * (targetDPI / 96)
                        dse.BlurRadius = 15 * (targetDPI / 96)

                        Select Case shadow
                            Case "left"
                                dse.Direction = -135
                            Case "right"
                                dse.Direction = -45
                            Case "bottom"
                                dse.Direction = -90
                        End Select

                        border.Effect = dse

                    End If

                    If r <> 0 Then
                        Dim rt As New RotateTransform()
                        rt.Angle = r
                        border.RenderTransform = rt
                    End If

                    placeholder.UpdateLayout()
                    mainspread.UpdateLayout()
                    grid.UpdateLayout()

                    ms = Nothing
                    src = Nothing
                    srcimg = Nothing

                End If

                If element.Attributes.GetNamedItem("type").Value = "rectangle" Then

                    Dim fillArr As ArrayList = GetColorRgb(element.Attributes.GetNamedItem("fillcolor").Value.ToString)
                    Dim borderArr As ArrayList = GetColorRgb(element.Attributes.GetNamedItem("bordercolor").Value.ToString)
                    Dim borderweight As Double = Math.Round(XmlConvert.ToDouble(element.Attributes.GetNamedItem("borderweight").Value.ToString) * (targetDPI / 96))
                    Dim r As Double = XmlConvert.ToDouble(element.Attributes.GetNamedItem("rotation").Value.ToString)
                    Dim fillalpha As Double = XmlConvert.ToDouble(element.Attributes.GetNamedItem("fillalpha").Value.ToString)
                    Dim shadow As String = element.Attributes.GetNamedItem("shadow").Value.ToString

                    Dim rect As New System.Windows.Shapes.Rectangle
                    rect.Width = Math.Round(XmlConvert.ToDouble(element.Attributes.GetNamedItem("objectWidth").Value) * (targetDPI / 96))
                    rect.Height = Math.Round(XmlConvert.ToDouble(element.Attributes.GetNamedItem("objectHeight").Value) * (targetDPI / 96))
                    rect.HorizontalAlignment = Windows.HorizontalAlignment.Left
                    rect.VerticalAlignment = Windows.VerticalAlignment.Top
                    rect.Fill = New SolidColorBrush(System.Windows.Media.Color.FromRgb(fillArr(0), fillArr(1), fillArr(2)))
                    rect.Stroke = New SolidColorBrush(System.Windows.Media.Color.FromRgb(borderArr(0), borderArr(1), borderArr(2)))
                    rect.StrokeThickness = borderweight
                    If borderweight > 0 Then
                        rect.Width += borderweight
                        rect.Height += borderweight
                    End If
                    rect.Margin = New Thickness(layflatmargin + (XmlConvert.ToDouble(element.Attributes.GetNamedItem("objectX").Value) * (targetDPI / 96)) - (borderweight / 2),
                                                 (XmlConvert.ToDouble(element.Attributes.GetNamedItem("objectY").Value) * (targetDPI / 96)) - (borderweight / 2),
                                                 0, 0)

                    If fillalpha < 1 Then
                        rect.Opacity = fillalpha
                    End If

                    If r <> 0 Then
                        Dim rt As New RotateTransform()
                        rt.Angle = r
                        rect.RenderTransform = rt
                    End If

                    If shadow <> "" Then

                        Dim dse As System.Windows.Media.Effects.DropShadowEffect = New System.Windows.Media.Effects.DropShadowEffect
                        dse.Color = System.Windows.Media.Color.FromRgb(88, 89, 91)
                        dse.ShadowDepth = 5 * (targetDPI / 96)
                        dse.BlurRadius = 15 * (targetDPI / 96)

                        Select Case shadow
                            Case "left"
                                dse.Direction = -135
                            Case "right"
                                dse.Direction = -45
                            Case "bottom"
                                dse.Direction = -90
                        End Select

                        rect.Effect = dse

                    End If

                    mainspread.Children.Add(rect)

                    mainspread.UpdateLayout()
                    grid.UpdateLayout()

                End If

                If element.Attributes.GetNamedItem("type").Value = "circle" Then

                    Dim fillArr As ArrayList = GetColorRgb(element.Attributes.GetNamedItem("fillcolor").Value.ToString)
                    Dim borderArr As ArrayList = GetColorRgb(element.Attributes.GetNamedItem("bordercolor").Value.ToString)
                    Dim borderweight As Double = Math.Round(XmlConvert.ToDouble(element.Attributes.GetNamedItem("borderweight").Value.ToString) * (targetDPI / 96))
                    Dim r As Double = XmlConvert.ToDouble(element.Attributes.GetNamedItem("rotation").Value.ToString)
                    Dim fillalpha As Double = XmlConvert.ToDouble(element.Attributes.GetNamedItem("fillalpha").Value.ToString)
                    Dim shadow As String = element.Attributes.GetNamedItem("shadow").Value.ToString

                    Dim ellipse As New Ellipse
                    ellipse.Width = Math.Round(XmlConvert.ToDouble(element.Attributes.GetNamedItem("objectWidth").Value) * (targetDPI / 96))
                    ellipse.Height = Math.Round(XmlConvert.ToDouble(element.Attributes.GetNamedItem("objectHeight").Value) * (targetDPI / 96))
                    ellipse.HorizontalAlignment = Windows.HorizontalAlignment.Left
                    ellipse.VerticalAlignment = Windows.VerticalAlignment.Top
                    ellipse.Fill = New SolidColorBrush(System.Windows.Media.Color.FromRgb(fillArr(0), fillArr(1), fillArr(2)))
                    ellipse.Stroke = New SolidColorBrush(System.Windows.Media.Color.FromRgb(borderArr(0), borderArr(1), borderArr(2)))
                    ellipse.StrokeThickness = borderweight
                    If borderweight > 0 Then
                        ellipse.Width += borderweight
                        ellipse.Height += borderweight
                    End If
                    ellipse.Margin = New Thickness(layflatmargin + (XmlConvert.ToDouble(element.Attributes.GetNamedItem("objectX").Value) * (targetDPI / 96)) - (borderweight / 2),
                                                 (XmlConvert.ToDouble(element.Attributes.GetNamedItem("objectY").Value) * (targetDPI / 96)) - (borderweight / 2),
                                                 0, 0)

                    If fillalpha < 1 Then
                        ellipse.Opacity = fillalpha
                    End If

                    If r <> 0 Then
                        Dim rt As New RotateTransform()
                        rt.Angle = r
                        ellipse.RenderTransform = rt
                    End If

                    If shadow <> "" Then

                        Dim dse As System.Windows.Media.Effects.DropShadowEffect = New System.Windows.Media.Effects.DropShadowEffect
                        dse.Color = System.Windows.Media.Color.FromRgb(88, 89, 91)
                        dse.ShadowDepth = 5 * (targetDPI / 96)
                        dse.BlurRadius = 15 * (targetDPI / 96)

                        Select Case shadow
                            Case "left"
                                dse.Direction = -135
                            Case "right"
                                dse.Direction = -45
                            Case "bottom"
                                dse.Direction = -90
                        End Select

                        ellipse.Effect = dse

                    End If

                    mainspread.Children.Add(ellipse)

                    mainspread.UpdateLayout()
                    grid.UpdateLayout()

                End If

                If element.Attributes.GetNamedItem("type").Value = "line" Then

                    Dim fillArr As ArrayList = GetColorRgb(element.Attributes.GetNamedItem("fillcolor").Value.ToString)
                    Dim lineweight As Double = Math.Round(XmlConvert.ToDouble(element.Attributes.GetNamedItem("lineweight").Value.ToString) * (targetDPI / 96))
                    Dim r As Double = XmlConvert.ToDouble(element.Attributes.GetNamedItem("rotation").Value.ToString)
                    Dim fillalpha As Double = XmlConvert.ToDouble(element.Attributes.GetNamedItem("fillalpha").Value.ToString)
                    Dim shadow As String = element.Attributes.GetNamedItem("shadow").Value.ToString

                    Dim line As New System.Windows.Shapes.Rectangle
                    line.Width = Math.Round(XmlConvert.ToDouble(element.Attributes.GetNamedItem("objectWidth").Value) * (targetDPI / 96))
                    line.Height = lineweight
                    line.HorizontalAlignment = Windows.HorizontalAlignment.Left
                    line.VerticalAlignment = Windows.VerticalAlignment.Top
                    line.Fill = New SolidColorBrush(System.Windows.Media.Color.FromRgb(fillArr(0), fillArr(1), fillArr(2)))
                    line.Margin = New Thickness(layflatmargin + (XmlConvert.ToDouble(element.Attributes.GetNamedItem("objectX").Value) * (targetDPI / 96)),
                                                 (XmlConvert.ToDouble(element.Attributes.GetNamedItem("objectY").Value) * (targetDPI / 96)),
                                                 0, 0)
                    If fillalpha < 1 Then
                        line.Opacity = fillalpha
                    End If

                    If r <> 0 Then
                        Dim rt As New RotateTransform()
                        rt.Angle = r
                        line.RenderTransform = rt
                    End If

                    If shadow <> "" Then

                        Dim dse As System.Windows.Media.Effects.DropShadowEffect = New System.Windows.Media.Effects.DropShadowEffect
                        dse.Color = System.Windows.Media.Color.FromRgb(88, 89, 91)
                        dse.ShadowDepth = 5 * (targetDPI / 96)
                        dse.BlurRadius = 15 * (targetDPI / 96)

                        Select Case shadow
                            Case "left"
                                dse.Direction = -135
                            Case "right"
                                dse.Direction = -45
                            Case "bottom"
                                dse.Direction = -90
                        End Select

                        line.Effect = dse

                    End If

                    mainspread.Children.Add(line)

                    mainspread.UpdateLayout()
                    grid.UpdateLayout()

                End If

            Next


            'Save the image
            savebblocktimer = New DispatcherTimer()
            AddHandler savebblocktimer.Tick, AddressOf SaveBBlock
            savebblocktimer.Interval = New TimeSpan(0, 0, 1)
            savebblocktimer.Start()

        Catch ex As Exception

            ExitOrder(ex.Message)

        End Try

    End Sub

    Public Sub SavePDFCover()

        'Create the PDF now
        If p Is Nothing Then
            p = New PDFlib()
            p.set_parameter("license", "W800102-010000-126165-VAUBF2-MAQC32")
            p.set_parameter("errorpolicy", "return")
            p.set_info("Creator", "Fotoalbum PDF Generator")
            p.set_info("Title", "Fotoalbum PDF Generator")
            p.set_parameter("charref", "true")
            p.set_parameter("SearchPath", exportfolder)
            p.set_parameter("SearchPath", searchpath_fontfolder)
            p.set_parameter("escapesequence", "false")
            p.set_parameter("usercoordinates", "true")
            p.begin_document(pdf_createfolder & cover_filename, "compatibility=" & pdfcompatible)

        End If

        Dim spread As XmlNode = spreadLst(0)
        Dim spreadID As String = spread.Attributes.GetNamedItem("spreadID").Value

        Dim pages As XmlNodeList = spread.SelectNodes("descendant::page")
        Dim coverbleed As Double = XmlConvert.ToDouble(pages(0).Attributes.GetNamedItem("horizontalBleed").Value.ToString())

        Dim width As Double = XmlConvert.ToDouble(spread.Attributes.GetNamedItem("totalWidth").Value.ToString())
        Dim height As Double = XmlConvert.ToDouble(spread.Attributes.GetNamedItem("totalHeight").Value.ToString())
        Dim trimbox As String = "trimbox={" & coverbleed & " " & coverbleed & " " & width + coverbleed & " " & height + coverbleed & "}"

        Dim filename As String = exportfolder & spreadID & ".jpg"

        p.begin_page_ext(width + (2 * coverbleed), height + (2 * coverbleed), currentpageorientation & " " & trimbox)

        'Now load the image in the PDF
        Dim img As Integer = -1
        img = p.load_image("auto", filename, "")
        If img = -1 Then
            ExitOrder(p.get_errmsg)
        End If

        Dim imageopt As String = "boxsize={" & width & " " & height & "} position={center} fitmethod=slice"
        p.fit_image(img, coverbleed, height + coverbleed, imageopt)
        p.close_image(img)
        img = Nothing

        Dim elements As XmlNodeList = spread.SelectSingleNode("elements").SelectNodes("element")

        For Each element As XmlNode In elements

            If element.Attributes.GetNamedItem("type").Value = "text" Then

                'Find the textelement and place it
                Dim posx As Double = XmlConvert.ToDouble(element.Attributes.GetNamedItem("objectX").Value.ToString)
                Dim posy As Double = XmlConvert.ToDouble(element.Attributes.GetNamedItem("objectY").Value.ToString)
                Dim rotation As Double = XmlConvert.ToDouble(element.Attributes.GetNamedItem("rotation").Value.ToString)

                Dim tw As Double = XmlConvert.ToDouble(element.Attributes.GetNamedItem("objectWidth").Value.ToString)
                Dim th As Double = XmlConvert.ToDouble(element.Attributes.GetNamedItem("objectHeight").Value.ToString)

                If CreateTextLines(element.Attributes.GetNamedItem("tfID").Value.ToString(), posx + coverbleed, posy + coverbleed, rotation) = False Then
                    ExitOrder("Problem with text")
                End If

            End If
        Next

        Dim cropwidth As Double = width + (2 * coverbleed)
        Dim cropheight As Double = height + (2 * coverbleed)

        'Draw the cropmarks etc
        p.setlinewidth(0.3)
        p.setcolor("stroke", "rgb", 1, 1, 1, 1)
        'left top vertical
        p.moveto(coverbleed, 0)
        p.lineto(coverbleed, coverbleed - 2)
        p.stroke()

        p.setlinewidth(0.1)
        p.setcolor("stroke", "rgb", 0, 0, 0, 0)
        'left top vertical
        p.moveto(coverbleed, 0)
        p.lineto(coverbleed, coverbleed - 2)
        p.stroke()

        p.setlinewidth(0.3)
        p.setcolor("stroke", "rgb", 1, 1, 1, 1)
        'left top horizontal
        p.moveto(0, coverbleed)
        p.lineto(coverbleed - 2, coverbleed)
        p.stroke()

        p.setlinewidth(0.1)
        p.setcolor("stroke", "rgb", 0, 0, 0, 0)
        'left top horizontal
        p.moveto(0, coverbleed)
        p.lineto(coverbleed - 2, coverbleed)
        p.stroke()

        p.setlinewidth(0.3)
        p.setcolor("stroke", "rgb", 1, 1, 1, 1)
        'spine top vertical
        p.moveto(spineX + coverbleed, 0)
        p.lineto(spineX + coverbleed, coverbleed - 2)
        p.stroke()

        p.setlinewidth(0.1)
        p.setcolor("stroke", "rgb", 0, 0, 0, 0)
        'spine top vertical
        p.moveto(spineX + coverbleed, 0)
        p.lineto(spineX + coverbleed, coverbleed - 2)
        p.stroke()

        p.setlinewidth(0.3)
        p.setcolor("stroke", "rgb", 1, 1, 1, 1)
        'right top vertical
        p.moveto(cropwidth - coverbleed, 0)
        p.lineto(cropwidth - coverbleed, coverbleed - 2)
        p.stroke()

        p.setlinewidth(0.1)
        p.setcolor("stroke", "rgb", 0, 0, 0, 0)
        'right top vertical
        p.moveto(cropwidth - coverbleed, 0)
        p.lineto(cropwidth - coverbleed, coverbleed - 2)
        p.stroke()

        p.setlinewidth(0.3)
        p.setcolor("stroke", "rgb", 1, 1, 1, 1)
        'right top horizontal
        p.moveto(cropwidth, coverbleed)
        p.lineto(cropwidth - coverbleed + 2, coverbleed)
        p.stroke()

        p.setlinewidth(0.1)
        p.setcolor("stroke", "rgb", 0, 0, 0, 0)
        'right top horizontal
        p.moveto(cropwidth, coverbleed)
        p.lineto(cropwidth - coverbleed + 2, coverbleed)
        p.stroke()

        p.setlinewidth(0.3)
        p.setcolor("stroke", "rgb", 1, 1, 1, 1)
        'right bottom horizontal
        p.moveto(cropwidth, cropheight - coverbleed)
        p.lineto(cropwidth - coverbleed + 2, cropheight - coverbleed)
        p.stroke()

        p.setlinewidth(0.1)
        p.setcolor("stroke", "rgb", 0, 0, 0, 0)
        'right bottom horizontal
        p.moveto(cropwidth, cropheight - coverbleed)
        p.lineto(cropwidth - coverbleed + 2, cropheight - coverbleed)
        p.stroke()

        p.setlinewidth(0.3)
        p.setcolor("stroke", "rgb", 1, 1, 1, 1)
        'right bottom vertical
        p.moveto(cropwidth - coverbleed, cropheight)
        p.lineto(cropwidth - coverbleed, cropheight - coverbleed + 2)
        p.stroke()

        p.setlinewidth(0.1)
        p.setcolor("stroke", "rgb", 0, 0, 0, 0)
        'right bottom vertical
        p.moveto(cropwidth - coverbleed, cropheight)
        p.lineto(cropwidth - coverbleed, cropheight - coverbleed + 2)
        p.stroke()

        p.setlinewidth(0.3)
        p.setcolor("stroke", "rgb", 1, 1, 1, 1)
        'spine bottom vertical
        p.moveto(spineX + coverbleed, cropheight)
        p.lineto(spineX + coverbleed, cropheight - coverbleed + 2)
        p.stroke()

        p.setlinewidth(0.1)
        p.setcolor("stroke", "rgb", 0, 0, 0, 0)
        'spine bottom vertical
        p.moveto(spineX + coverbleed, cropheight)
        p.lineto(spineX + coverbleed, cropheight - coverbleed + 2)
        p.stroke()

        p.setlinewidth(0.3)
        p.setcolor("stroke", "rgb", 1, 1, 1, 1)
        'left bottom vertical
        p.moveto(coverbleed, cropheight)
        p.lineto(coverbleed, cropheight - coverbleed + 2)
        p.stroke()

        p.setlinewidth(0.1)
        p.setcolor("stroke", "rgb", 0, 0, 0, 0)
        'left bottom vertical
        p.moveto(coverbleed, cropheight)
        p.lineto(coverbleed, cropheight - coverbleed + 2)
        p.stroke()

        p.setlinewidth(0.3)
        p.setcolor("stroke", "rgb", 1, 1, 1, 1)
        'left bottom horizontal
        p.moveto(0, cropheight - coverbleed)
        p.lineto(coverbleed - 2, cropheight - coverbleed)
        p.stroke()

        p.setlinewidth(0.1)
        p.setcolor("stroke", "rgb", 0, 0, 0, 0)
        'left bottom horizontal
        p.moveto(0, cropheight - coverbleed)
        p.lineto(coverbleed - 2, cropheight - coverbleed)
        p.stroke()

        p.end_page_ext("")

        p.end_document("destination={type=fixed zoom=.5}")

        p.Dispose()
        p = Nothing

        renderindex += 1

        Me.Content = Nothing

        CreatePageBlock()

    End Sub

    Private Function CreateTextLines(id As String, posx As Double, posy As Double, r As Double) As Boolean

        Dim loadfont As String

        Try

            Dim topmargin As Double = 0
            Dim firstline As Boolean = True

            'Show the bounding box
            If r <> 0 Then
                p.save()
                p.translate(posx, posy)
                p.rotate(r * -1)

                posx = 0
                posy = 0
            End If

            'Read and place the textlines
            If Not IsNothing(textlinecontainers) Then

                For Each e As XmlNode In textlinecontainers

                    If e.Attributes.GetNamedItem("id").Value.ToString() = id Then

                        Dim textlines As XmlNodeList = e.SelectNodes("textline")
                        For Each textline As XmlNode In textlines

                            Dim lastx As Double = posx

                            Dim nodes As XmlNodeList = textline.SelectNodes("descendant::span")
                            Dim finalStr As String = ""
                            Dim tf As Integer = -1
                            Dim leading As Double = 0

                            For Each node As XmlNode In nodes

                                If firstline = True Then
                                    firstline = False
                                End If

                                Dim fontname As String = GetFont(node)
                                loadfont = GetFontPath(node)

                                Dim check As Double = node.Attributes.GetNamedItem("corps").Value
                                If check > leading Then
                                    leading = check
                                End If

                                Dim tempoptlist As String = " encoding=unicode embedding=true fontsize=" & node.Attributes.GetNamedItem("corps").Value.ToString & " " & GetColorStringRgb(node.Attributes.GetNamedItem("color").Value.ToString)

                                Debug.Print(tempoptlist)

                                If Not IsNothing(node.Attributes.GetNamedItem("underline").Value) Then
                                    If node.Attributes.GetNamedItem("underline").Value.ToString = "true" Then
                                        tempoptlist += " underline"
                                    End If
                                End If

                                Dim str As String = node.InnerText

                                For Each c As Char In str

                                    If c <> " " Then
                                        If (p.info_textline(c, "unmappedchars", "fontname=" & fontname & " charref" & tempoptlist) = 1) Then
                                            'Replace it with Arial
                                            tf = p.add_textflow(tf, c, "fontname=Arial" & tempoptlist)
                                        Else
                                            tf = p.add_textflow(tf, c, "fontname=" & fontname & tempoptlist)
                                        End If
                                        finalStr += c
                                    Else
                                        tf = p.add_textflow(tf, "&nbsp;", "fontname=" & fontname & tempoptlist)
                                        finalStr += "<space>"
                                    End If

                                Next

                            Next

                            If (tf > -1) Then

                                p.fit_textflow(tf, lastx + XmlConvert.ToDouble(textline.Attributes.GetNamedItem("x").Value), posy + XmlConvert.ToDouble(textline.Attributes.GetNamedItem("y").Value) - leading, 10000, 10000, "verticalalign=top")

                                lastx = 0

                                p.delete_textflow(tf)

                            End If

                        Next

                        Exit For

                    End If

                Next
            End If

            If r <> 0 Then
                p.restore()
            End If

            Return True

        Catch ex As System.Exception

            'MsgBox(ex.Message & " -> " & loadfont)
            Return False

        End Try

    End Function

    Private Function GetFont(node As XmlNode) As String

        Dim SystemFontName As String = ""

        For Each row As DataRow In fonts.Rows

            Dim font As String = row("swfName").ToString
            font = "_" & font.Replace(".swf", "")

            Dim checkfont As String = node.Attributes.GetNamedItem("font").Value.ToString
            If checkfont.Substring(0, 1) <> "_" Then
                checkfont = "_" & checkfont.ToLower
            End If

            If checkfont = font Then
                SystemFontName = row("name").ToString
                SystemFontName = SystemFontName.Replace("." & row("extension"), "")
                Exit For
            End If
        Next

        Return SystemFontName

    End Function

    Private Function GetFontPath(node As XmlNode) As String

        Dim SystemFontName As String = ""

        For Each row As DataRow In fonts.Rows

            Dim font As String = row("swfName").ToString
            font = "_" & font.Replace(".swf", "")

            Dim checkfont As String = node.Attributes.GetNamedItem("font").Value.ToString
            If checkfont.Substring(0, 1) <> "_" Then
                checkfont = "_" & checkfont.ToLower
            End If

            If checkfont = font Then
                SystemFontName = row("name").ToString
                Exit For
            End If
        Next

        Return SystemFontName

    End Function

    Public Sub SavePDFBlock()

        numPages = 0

        'Create the PDF now
        If p Is Nothing Then
            p = New PDFlib()
            p.set_parameter("license", "W800102-010000-126165-VAUBF2-MAQC32")
            p.set_parameter("errorpolicy", "return")
            p.set_info("Creator", "Fotoalbum PDF Generator")
            p.set_info("Title", "Fotoalbum PDF Generator")
            p.set_parameter("charref", "true")
            p.set_parameter("SearchPath", exportfolder)
            p.set_parameter("SearchPath", searchpath_fontfolder)
            p.set_parameter("escapesequence", "false")
            p.set_parameter("usercoordinates", "true")

            p.begin_document(pdf_createfolder & bblock_filename, "compatibility=" & pdfcompatible)

        End If

        Dim startspread As Integer = 1
        If singlepageproduct Then startspread = 0

        For x As Integer = startspread To spreadLst.Count - 1

            Dim spread As XmlNode = spreadLst(x)
            Dim spreadID As String = spread.Attributes.GetNamedItem("spreadID").Value

            Dim pages As XmlNodeList = spread.SelectNodes("descendant::page")

            If isLayFlat = True Then

                Dim filename As String = exportfolder & spreadID & ".jpg"

                Dim width As Double = XmlConvert.ToDouble(spreadLst(2).Attributes.GetNamedItem("totalWidth").Value.ToString())
                Dim height As Double = XmlConvert.ToDouble(spread.Attributes.GetNamedItem("totalHeight").Value.ToString())
                Dim bleed As Double = XmlConvert.ToDouble(pages(0).Attributes.GetNamedItem("horizontalBleed").Value.ToString())

                Dim trimbox As String = "trimbox={" & bleed & " " & bleed & " " & width - bleed & " " & height - bleed & "}"

                p.begin_page_ext(width, height, currentpageorientation & " " & trimbox)

                numPages += 2
                'Now load the image in the PDF
                Dim img As Integer = -1
                img = p.load_image("auto", filename, "")
                If img = -1 Then
                    ExitOrder(p.get_errmsg)
                End If

                Dim layflatmargin As Double = 0
                Dim singlepageFirst As Boolean = (pages(0).Attributes.GetNamedItem("singlepageFirst").Value.ToString.ToLower = "true")
                If singlepageFirst = True Then
                    layflatmargin = XmlConvert.ToDouble(pages(0).Attributes.GetNamedItem("pageWidth").Value)
                End If

                Dim imageopt As String = "boxsize={" & width & " " & height & "} position={center} fitmethod=slice"
                p.fit_image(img, 0, height, imageopt)
                p.close_image(img)

                Dim elements As XmlNodeList = spread.SelectSingleNode("elements").SelectNodes("element")

                For Each element As XmlNode In elements

                    If element.Attributes.GetNamedItem("type").Value = "text" Then

                        'Find the textelement and place it
                        Dim posx As Double = layflatmargin + XmlConvert.ToDouble(element.Attributes.GetNamedItem("objectX").Value.ToString)
                        Dim posy As Double = XmlConvert.ToDouble(element.Attributes.GetNamedItem("objectY").Value.ToString)
                        Dim rotation As Double = XmlConvert.ToDouble(element.Attributes.GetNamedItem("rotation").Value.ToString)

                        Dim tw As Double = XmlConvert.ToDouble(element.Attributes.GetNamedItem("objectWidth").Value.ToString)
                        Dim th As Double = XmlConvert.ToDouble(element.Attributes.GetNamedItem("objectHeight").Value.ToString)

                        If CreateTextLines(element.Attributes.GetNamedItem("tfID").Value.ToString(), posx, posy, rotation) = False Then
                            ExitOrder("Problem with text")
                        End If

                    End If
                Next

                If DrawBBlockCropmarks(width, height, bleed) = True Then
                    p.end_page_ext("")
                End If

            Else 'isLayFlat = false

                Dim filename As String = exportfolder & spreadID & ".jpg"
                Dim xmargin As Double = 0

                For y As Integer = 0 To pages.Count - 1

                    Dim page As XmlNode = pages.Item(y)

                    Dim singlepage As Boolean = (page.Attributes.GetNamedItem("singlepage").Value.ToString.ToLower = "true")
                    Dim singlepageFirst As Boolean = (page.Attributes.GetNamedItem("singlepageFirst").Value.ToString.ToLower = "true")
                    Dim singlepageLast As Boolean = (page.Attributes.GetNamedItem("singlepageLast").Value.ToString.ToLower = "true")
                    Dim pageLeftRight As String = page.Attributes.GetNamedItem("pageLeftRight").Value.ToString.ToLower

                    Dim wrap As Double = 0
                    If singlepageproduct = True Then
                        If Not IsNothing(pages(0).Attributes.GetNamedItem("horizontalWrap")) Then
                            wrap = XmlConvert.ToDouble(pages(0).Attributes.GetNamedItem("horizontalWrap").Value.ToString())
                        End If
                    End If

                    Dim bleed As Double = XmlConvert.ToDouble(page.Attributes.GetNamedItem("horizontalBleed").Value.ToString())

                    Dim width As Double = XmlConvert.ToDouble(page.Attributes.GetNamedItem("pageWidth").Value.ToString()) + (2 * bleed) + (2 * wrap)
                    Dim height As Double = XmlConvert.ToDouble(page.Attributes.GetNamedItem("pageHeight").Value.ToString()) + (2 * bleed) + (2 * wrap)
                    xmargin = 0

                    Dim trimbox As String
                    If singlepageproduct = True Then
                        trimbox = "trimbox={" & bleed + wrap & " " & bleed + wrap & " " & width - bleed - wrap & " " & height - bleed - wrap & "}"
                        trimbox += " bleedbox={" & bleed & " " & bleed & " " & width - bleed & " " & height - bleed & "}"
                    Else
                        trimbox = "trimbox={" & bleed & " " & bleed & " " & width - bleed & " " & height - bleed & "}"
                    End If

                    p.begin_page_ext(width, height, currentpageorientation & " " & trimbox)

                    numPages += 1

                    'Now load the image in the PDF
                    Dim img As Integer = -1
                    img = p.load_image("auto", filename, "")
                    If img = -1 Then
                        ExitOrder(p.get_errmsg)
                    End If

                    Dim position As String = "left top"
                    If pageLeftRight = "right" Then
                        position = "right top"
                        xmargin = width - (2 * bleed)
                    End If

                    If singlepageFirst = True Then
                        xmargin = 0
                    End If

                    Dim imageopt As String = "boxsize={" & width & " " & height & "} position={" & position & "} fitmethod=slice"
                    p.fit_image(img, 0, height, imageopt)
                    p.close_image(img)

                    Dim elements As XmlNodeList = spread.SelectSingleNode("elements").SelectNodes("element")

                    For Each element As XmlNode In elements

                        If element.Attributes.GetNamedItem("type").Value = "text" Then

                            'Find the textelement and place it
                            Dim posx As Double = XmlConvert.ToDouble(element.Attributes.GetNamedItem("objectX").Value.ToString) - xmargin
                            Dim posy As Double = XmlConvert.ToDouble(element.Attributes.GetNamedItem("objectY").Value.ToString)
                            Dim rotation As Double = XmlConvert.ToDouble(element.Attributes.GetNamedItem("rotation").Value.ToString)

                            Dim tw As Double = XmlConvert.ToDouble(element.Attributes.GetNamedItem("objectWidth").Value.ToString)
                            Dim th As Double = XmlConvert.ToDouble(element.Attributes.GetNamedItem("objectHeight").Value.ToString)

                            If CreateTextLines(element.Attributes.GetNamedItem("tfID").Value.ToString(), posx, posy, rotation) = False Then
                                ExitOrder("Problem with text")
                            End If

                        End If
                    Next

                    If DrawBBlockCropmarks(width, height, bleed) = True Then
                        p.end_page_ext("")
                    End If

                Next

            End If

        Next

        If isLayFlat = True Then
            numPages -= 2
        End If

        p.end_document("destination={type=fixed zoom=.5}")

        p.Dispose()
        p = Nothing

        '==============================================================================
        ' Remove content or create the image directory
        '==============================================================================
        If Directory.Exists(exportfolder) Then
            Try
                Directory.Delete(exportfolder, True)
            Catch ex As Exception
                Debug.Print("Error deleting directory! " & ex.Message)
            End Try
        End If

        '==============================================================================
        ' Update the order with the new result and filenames etc
        '==============================================================================
        Dim sqlStr As String
        If singlepageproduct = True Then
            sqlStr = "UPDATE pdfengine_order_pdfs SET status = 'finished', path_bbloc='" & bblock_filename & "', path_cover='', nr_pages=" & numPages & " WHERE id = " & currentOrder.id
        Else
            sqlStr = "UPDATE pdfengine_order_pdfs SET status = 'finished', path_bbloc='" & bblock_filename & "', path_cover='" & cover_filename & "', nr_pages=" & numPages & " WHERE id = " & currentOrder.id
        End If

        Dim connStringSQL As New MySqlConnection(mySqlConnection)
        Dim myCommand As New MySqlCommand(sqlStr, connStringSQL)
        myCommand.Connection.Open()
        myCommand.ExecuteScalar()
        myCommand.Connection.Close()

        'Send an email to Maurice to check the PDF!
        Try

            Dim SmtpServer As New SmtpClient()
            Dim mail As New MailMessage()
            SmtpServer.Port = 25
            SmtpServer.Host = "217.195.125.26"
            SmtpServer.Credentials = New NetworkCredential("Administrator", "9ASIOsWC")
            mail = New MailMessage()
            mail.From = New MailAddress("pdfserver@vm1624.local")
            mail.To.Add("maurice@fotoalbum.nl")
            'mail.To.Add("ole@fotoalbum.nl")
            mail.Subject = "PDF FOR CHECK | pdf_engine_order: " & currentOrder.id
            If cover_filename <> "" Then
                mail.Body = "path_cover: http://www.fotoalbum.nl/pdfexports/" & cover_filename & vbCrLf & "path_bbloc: http://www.fotoalbum.nl/pdfexports/" & bblock_filename & vbCrLf & "nr_pages: " & numPages & vbCrLf & "platform: " & currentOrder.platform & vbCrLf & "/maak-nu/" & currentOrder.product_id & "/" & currentOrder.user_product_id
            Else
                mail.Body = "path_bbloc: http://ww.fotoalbum.nl/pdfexports/" & bblock_filename & vbCrLf & "nr_pages: " & numPages & vbCrLf & "platform: " & currentOrder.platform & vbCrLf & "/maak-nu/" & currentOrder.product_id & "/" & currentOrder.user_product_id
            End If
            If singlepageproduct = False Then
                'SmtpServer.Send(mail)
            End If

        Catch ex As Exception

            'Do nothing for now


        End Try

        'Remove the current order from the orderrows
        orders.RemoveAt(0)

        If orders.Count > 0 Then
            'Process the order
            currentOrder = orders.Item(0)
            exportfolder = original_exportfolder & currentOrder.id & "\"
            CreatePDFRender()
        Else
            ordertimer.Start()
        End If

    End Sub

    Public Function DrawBBlockCropmarks(width As Double, height As Double, bleed As Double) As Boolean

        If bleed > 0 Then

            'Draw the cropmarks etc
            p.setlinewidth(0.3)
            p.setcolor("stroke", "rgb", 1, 1, 1, 1)
            'left top vertical
            p.moveto(bleed, 0)
            p.lineto(bleed, bleed - 2)
            p.stroke()

            p.setlinewidth(0.1)
            p.setcolor("stroke", "rgb", 0, 0, 0, 0)
            'left top vertical
            p.moveto(bleed, 0)
            p.lineto(bleed, bleed - 2)
            p.stroke()

            p.setlinewidth(0.3)
            p.setcolor("stroke", "rgb", 1, 1, 1, 1)
            'left top horizontal
            p.moveto(0, bleed)
            p.lineto(bleed - 2, bleed)
            p.stroke()

            p.setlinewidth(0.1)
            p.setcolor("stroke", "rgb", 0, 0, 0, 0)
            'left top horizontal
            p.moveto(0, bleed)
            p.lineto(bleed - 2, bleed)
            p.stroke()

            p.setlinewidth(0.3)
            p.setcolor("stroke", "rgb", 1, 1, 1, 1)
            'right top vertical
            p.moveto(width - bleed, 0)
            p.lineto(width - bleed, bleed - 2)
            p.stroke()

            p.setlinewidth(0.1)
            p.setcolor("stroke", "rgb", 0, 0, 0, 0)
            'right top vertical
            p.moveto(width - bleed, 0)
            p.lineto(width - bleed, bleed - 2)
            p.stroke()

            p.setlinewidth(0.3)
            p.setcolor("stroke", "rgb", 1, 1, 1, 1)
            'right top horizontal
            p.moveto(width, bleed)
            p.lineto(width - bleed + 2, bleed)
            p.stroke()

            p.setlinewidth(0.1)
            p.setcolor("stroke", "rgb", 0, 0, 0, 0)
            'right top horizontal
            p.moveto(width, bleed)
            p.lineto(width - bleed + 2, bleed)
            p.stroke()

            p.setlinewidth(0.3)
            p.setcolor("stroke", "rgb", 1, 1, 1, 1)
            'right bottom vertical
            p.moveto(width - bleed, height)
            p.lineto(width - bleed, height - bleed + 2)
            p.stroke()

            p.setlinewidth(0.1)
            p.setcolor("stroke", "rgb", 0, 0, 0, 0)
            'right bottom vertical
            p.moveto(width - bleed, height)
            p.lineto(width - bleed, height - bleed + 2)
            p.stroke()

            p.setlinewidth(0.3)
            p.setcolor("stroke", "rgb", 1, 1, 1, 1)
            'right bottom horizontal
            p.moveto(width, height - bleed)
            p.lineto(width - bleed + 2, height - bleed)
            p.stroke()

            p.setlinewidth(0.1)
            p.setcolor("stroke", "rgb", 0, 0, 0, 0)
            'right bottom horizontal
            p.moveto(width, height - bleed)
            p.lineto(width - bleed + 2, height - bleed)
            p.stroke()

            p.setlinewidth(0.3)
            p.setcolor("stroke", "rgb", 1, 1, 1, 1)
            'left bottom vertical
            p.moveto(bleed, height)
            p.lineto(bleed, height - bleed + 2)
            p.stroke()

            p.setlinewidth(0.1)
            p.setcolor("stroke", "rgb", 0, 0, 0, 0)
            'left bottom vertical
            p.moveto(bleed, height)
            p.lineto(bleed, height - bleed + 2)
            p.stroke()

            p.setlinewidth(0.3)
            p.setcolor("stroke", "rgb", 1, 1, 1, 1)
            'left bottom horizontal
            p.moveto(0, height - bleed)
            p.lineto(bleed - 2, height - bleed)
            p.stroke()

            p.setlinewidth(0.1)
            p.setcolor("stroke", "rgb", 0, 0, 0, 0)
            'left bottom horizontal
            p.moveto(0, height - bleed)
            p.lineto(bleed - 2, height - bleed)
            p.stroke()

        End If

        Return True

    End Function

    Public Sub Pause()

        'do nothing really

    End Sub

    Public Sub SaveCover()

        'savecovertimer.Stop()
        'savecovertimer = Nothing

        Try

            Dim filename As String = exportfolder & spreadLst(renderindex).Attributes.GetNamedItem("spreadID").Value & ".jpg"

            Dim placeholder As UIElement = Me.Content.Children(0)
            placeholder.UpdateLayout()

            Dim b As Rect = VisualTreeHelper.GetDescendantBounds(Me)

            Dim bmpRen As New RenderTargetBitmap(Math.Ceiling(b.Width), Math.Ceiling(b.Height), 96, 96, PixelFormats.Pbgra32)

            placeholder.Measure(placeholder.RenderSize)
            placeholder.Arrange(New Rect(placeholder.RenderSize))

            bmpRen.Render(placeholder) 'render the viewport as 2D snapshot

            Dim encoder As New JpegBitmapEncoder
            encoder.QualityLevel = 100
            encoder.Frames.Add(BitmapFrame.Create(bmpRen))
            Dim fileStream As Stream = File.Open(filename, FileMode.Create)
            encoder.Save(fileStream)
            fileStream.Close()
            fileStream.Dispose()

            encoder = Nothing
            bmpRen = Nothing

            Me.Content = Nothing

            Me.InvalidateVisual()

            Me.UpdateLayout()

            Application.Current.Dispatcher.Invoke(New Action(AddressOf SavePDFCover), DispatcherPriority.ContextIdle)

        Catch ex As Exception
            ExitOrder(ex.Message)
        End Try

    End Sub

    Private Sub SaveBBlock(sender As Object, e As EventArgs)

        savebblocktimer.Stop()
        savebblocktimer = Nothing

        Dim filename As String = exportfolder & spreadLst(renderindex).Attributes.GetNamedItem("spreadID").Value & ".jpg"

        Try

            Dim placeholder As UIElement = Me.Content.Children(0)
            placeholder.UpdateLayout()

            Dim b As Rect = VisualTreeHelper.GetDescendantBounds(Me)

            Dim bmpRen As New RenderTargetBitmap(Math.Ceiling(b.Width), Math.Ceiling(b.Height), 96, 96, PixelFormats.Pbgra32)

            placeholder.Measure(placeholder.RenderSize)
            placeholder.Arrange(New Rect(placeholder.RenderSize))

            bmpRen.Render(placeholder) 'render the viewport as 2D snapshot

            Dim encoder As New JpegBitmapEncoder
            encoder.QualityLevel = 100
            encoder.Frames.Add(BitmapFrame.Create(bmpRen))
            Dim fileStream As Stream = File.Open(filename, FileMode.Create)
            encoder.Save(fileStream)
            fileStream.Close()

            fileStream.Dispose()

            encoder = Nothing
            fileStream = Nothing
            bmpRen = Nothing

            Me.Content.Children.Clear()

            Me.Content = Nothing

            Me.UpdateLayout()

            GC.Collect()

            If renderindex = spreadLst.Count - 1 Then

                Application.Current.Dispatcher.Invoke(New Action(AddressOf SavePDFBlock), DispatcherPriority.ContextIdle)

            Else

                'Move on to next spread!
                renderindex += 1

                Application.Current.Dispatcher.Invoke(New Action(AddressOf CreatePageBlock), DispatcherPriority.ContextIdle)


            End If

        Catch ex As Exception
            ExitOrder(ex.Message)
        End Try

    End Sub

    Public Sub CloseApplication()

        Application.Current.Shutdown()

    End Sub

    Public Function GetColorRgb(color As String) As ArrayList

        Dim rgb As New ArrayList

        For Each col As XmlNode In colorlist

            If col.Attributes.GetNamedItem("id").Value.ToString = color Then

                Dim rgbarr As Array = col.Attributes.GetNamedItem("rgb").Value.ToString.Split(";")
                rgb.Add(rgbarr(0))
                rgb.Add(rgbarr(1))
                rgb.Add(rgbarr(2))
                Exit For
            End If

        Next

        If rgb.Count = 0 Then
            rgb.Add(0)
            rgb.Add(0)
            rgb.Add(0)
        End If

        Return rgb

    End Function

    Public Function GetColorStringRgb(color As String) As String

        Dim c As Color = System.Drawing.Color.FromArgb(Integer.Parse(color))

        Dim rgb As String = " fillcolor={rgb " & c.R / 255 & " " & c.G / 255 & " " & c.B / 255 & "}"

        Return rgb

    End Function

    Public Function MakeGrayscale(ByVal original As BitmapImage) As FormatConvertedBitmap

        Dim newBitmap As New FormatConvertedBitmap()
        newBitmap.BeginInit()
        newBitmap.Source = original
        newBitmap.DestinationFormat = PixelFormats.Gray32Float
        newBitmap.EndInit()

        Return newBitmap

    End Function

    Public Function MakeSepia(ByVal original As System.Windows.Controls.Image) As System.Windows.Controls.Image

        original.Effect = New FAEffects.Effects.SepiaEffect

        Return original

    End Function

    Public Sub ExitOrder(message As String)

        If Not p Is Nothing Then
            p.Dispose()
            p = Nothing
        End If

        Me.Content = Nothing
        Me.UpdateLayout()

        '============================================================================
        ' Update the order with the new result and filenames etc
        '==============================================================================
        Dim err_order_id As String = "ERROR_" & currentOrder.order_id
        Dim sqlStr As String = "UPDATE pdfengine_order_pdfs SET status='error_check', path_bbloc='', path_cover='' WHERE id=" & currentOrder.id
        Dim connStringSQL As New MySqlConnection(mySqlConnection)
        Dim myCommand As New MySqlCommand(sqlStr, connStringSQL)
        myCommand.Connection.Open()
        myCommand.ExecuteScalar()
        myCommand.Connection.Close()

        'Remove the current order from the orderrows
        If orders.Count > 0 Then
            orders.RemoveAt(0)
        End If

        'Send an email to Maurice to check the PDF!
        Try

            Dim SmtpServer As New SmtpClient()
            Dim mail As New MailMessage()
            SmtpServer.Port = 25
            SmtpServer.Host = "217.195.125.26"
            SmtpServer.Credentials = New NetworkCredential("Administrator", "9ASIOsWC")
            mail = New MailMessage()
            mail.From = New MailAddress("pdfserver@vm1624.local")
            mail.To.Add("maurice@fotoalbum.nl")
            mail.Subject = "PDF FOUT | pdf_engine_order: " & currentOrder.id
            mail.Body = "Error:" & vbCrLf & vbCrLf & message & vbCrLf & "http://www.fotoalbum.nl/maak-nu/" & currentOrder.product_id & "/" & currentOrder.user_product_id & "?check_enabled=helpdesk"
            SmtpServer.Send(mail)

        Catch ex As Exception
            'Do nothing for now
        End Try

        If orders.Count > 0 Then
            'Process the order
            currentOrder = orders.Item(0)
            exportfolder = original_exportfolder & currentOrder.id & "\"
            CreatePDFRender()
        Else
            ordertimer.Start()
        End If

    End Sub

    Public Function GetFileUrlFromUpload(id As String) As ArrayList

        Dim result As New ArrayList

        Dim dt As DataTable = New DataTable

        Dim query As String = "SELECT url FROM xhibit_documents WHERE guid='" & id & "'"

        Dim connStringSQL As New MySqlConnection(mySqlConnection)
        Dim myAdapter As New MySqlDataAdapter(query, connStringSQL)
        myAdapter.Fill(dt)

        For Each row As DataRow In dt.Rows
            'Dim hires As String = row("hires")
            'Dim thumb As String = "thumb_" & hires
            'Dim url As String = row("url")
            'url = url.Replace(hires, thumb)
            'Debug.Print("GetFileUrl : " & url)
            'result.Add(url)

            result.Add(row("url"))
        Next

        Return result

    End Function

    Public Function mm2pt(pt As Double) As Double

        Return pt / 0.35277777777738178

    End Function

End Class



