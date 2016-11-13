Imports System.IO
Imports System.Text
Public Class Events
    Dim connectionString As String = Nothing
    Dim ds As DataSet = Nothing
    Dim html As StringBuilder = Nothing
    Dim i As Integer = 0 'Счетчик для плюсиков
    Dim margin As Integer = 0
    Dim numOfGraphic = 0
    Dim CounterOfGraphicsInSuperSchema = 1
    Dim MaxindexOfGraphicInSuperSchema = 0
    Public Sub New()

        InitializeComponent()
        ' Пытаемся прочесть txt файл, в котором записан путь к БД.
        ' Если файла нет или путь в нем неверный, то 
        ' Предоставляем пользователю указать путь к файлу, 
        ' И далее сохраняем его в текстовом документе.
        Try
            OpenConnection()
        Catch ex As Exception
            GetPath()
            OpenConnection()
        End Try

    End Sub

    Private Sub GetPath()
        ' Сохраняем путь к БД в текстовом файле.
        ' Если он существует, то перезаписываем его.
        Dim path As String = Nothing
        Dim FilePath As String = ".\connectionString.txt"
        Dim ofd = New OpenFileDialog()

        ofd.InitialDirectory = "c:\"
        ofd.Filter = "mdb files (*.mdb)|*.mdb|accdb files (*.accdb)|*.accdb|All files (*.*)|*.*"
        ofd.FilterIndex = 2
        ofd.RestoreDirectory = True

        If ofd.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            path = ofd.FileName
            ofd.Reset()
            Try

                If File.Exists(FilePath) Then
                    File.Delete(FilePath)
                End If
                Using fs As StreamWriter = New StreamWriter(FilePath)
                    fs.Write(path)
                End Using

            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        Else
            Environment.Exit(0)
        End If

    End Sub

    Private Sub OpenConnection()
        ' Получаем путь к БД из текстового файла.
        ' Далее в зависимости от версии создаем строку подключения
        ' Через схему получаем имена всех таблиц.
        ' Затем через адаптер получаем в DataSet все таблицы.
        Dim pathToDB As String = Nothing
        Using fr As StreamReader = New StreamReader(".\connectionString.txt")
            pathToDB = fr.ReadLine()
        End Using
        If pathToDB.EndsWith("accdb") Then
            connectionString =
            "Provider=Microsoft.ACE.OLEDB.12.0;data source=" + pathToDB
        Else
            connectionString = "provider=Microsoft.Jet.OLEDB.4.0;data source=" + pathToDB
        End If
        HtmlStart()
        ds = New DataSet()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            CallNewParagraph("Par1")
            AddText("SuperSchema")
            Call AddSuperSchema(6, ds.Tables(0))
            Call CallLineGraphic(ds.Tables(0), Nothing, 0, 1)
            Call CallLineGraphic(ds.Tables(0), Nothing, 0, 1, 2)

            Call CallPieGraphic(ds.Tables(0), 1, Nothing)
            Call CallPieGraphic(ds.Tables(0), 2, Nothing)
            Call CallColumnGraphic(ds.Tables(0), 0, 1, 2)
            Call CallColumnGraphic(ds.Tables(0), 0, 1)
            CloseParagraph()
            CallNewParagraph("Part2")
            CallNewParagraph("Line")
            CallLineGraphic(ds.Tables(0), "Some text", 0, 1, 2)
            CloseParagraph()
            CloseParagraph()
            CallNewParagraph("Donut")
            CallPieGraphic(ds.Tables(0), 1, "Donut")
            CloseParagraph()
            CallNewParagraph("Column")
            CallColumnGraphic(ds.Tables(0), 0, 1, 2)
            CloseParagraph()
            CallNewParagraph("Table")
            CallTable(ds.Tables(0))
            CloseParagraph()

            Call HtmlEnd()
            Call CreateHTML()
            Environment.Exit(0)

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub AddSuperSchema(countOfGraphics As Integer, table As DataTable)
        Dim doubl As Double = Nothing
        Dim strTable = New StringBuilder()
        Dim height As Integer
        If (countOfGraphics * 100) + 30 > (table.Rows.Count * 25) + 40 Then
            height = (countOfGraphics * 100) + 30
        Else
            height = (table.Rows.Count * 25) + 40
        End If
        strTable.AppendLine("<table style=""margin-top:15px;margin-bottom:15px;text-align:right;"">")
        strTable.AppendLine("<tr>")
        For i As Integer = 0 To table.Columns.Count - 1
            strTable.AppendLine("<th>" + table.Columns(i).ColumnName + "</th>")
        Next
        strTable.AppendLine("<th>rjger</th><th>grogr</th><th>romegoem</th>")
        strTable.AppendLine("</tr>")
        For i As Integer = 0 To table.Rows.Count - 1
            strTable.AppendLine("<tr>")
            For j As Integer = 0 To table.Columns.Count - 1
                If Double.TryParse(table.Rows(i)(j), doubl) AndAlso table.Rows(i)(j) Mod 1 <> 0 Then
                    strTable.AppendLine("<td>" + Double.Parse(table.Rows(i)(j)).ToString("N2") + "</td>")
                Else
                    strTable.AppendLine("<td>" + table.Rows(i)(j).ToString() + "</td>")
                End If
            Next
            strTable.AppendLine("</tr>")
        Next
        strTable.AppendLine("</table>")

        MaxindexOfGraphicInSuperSchema = countOfGraphics
        Dim index = 1
        html.AppendLine("<div style=""width:80%; height:" + height.ToString + "px;margin:auto"">
<div style=""width:25%;float:left"">")
        html.AppendLine("<div> </div>")
        For i As Integer = 0 To (countOfGraphics / 2) - 1
            html.AppendLine("<div id=""graph" + index.ToString + """></div>")
            index += 1
        Next
        html.AppendLine("</div><div style=""width:25%;float:right"">")
        html.AppendLine("<div> </div>")
        For i As Integer = countOfGraphics / 2 To countOfGraphics - 1
            html.AppendLine("<div id=""graph" + index.ToString + """></div>")
            index += 1
        Next
        html.AppendLine("</div>
<div style=""width:30%;"">" + strTable.ToString + "</div>
</div>")
    End Sub

    Private Sub CallPieGraphic(table As DataTable, column As Integer, name As String)
        Dim data = New StringBuilder()
        data.AppendLine("[")
        data.Append("[")
        data.Append("'" + table.Columns(0).ColumnName + "',")
        Data.Append("'" + table.Columns(column).ColumnName + "'")
        data.AppendLine("],")
        For i As Integer = 0 To table.Rows.Count - 1
            data.Append("[")
            data.Append("'" + table.Rows(i)(0).ToString + "',")
            data.Append(table.Rows(i)(column))
            data.Append("]")
            If i + 1 <> table.Rows.Count Then
                data.AppendLine(", ")
            End If
        Next
        data.AppendLine("]);")
        html.AppendLine("<script type = ""text/javascript"">
      google.charts.load('current', {'packages':['corechart']});
      google.charts.setOnLoadCallback(drawChart);

      function drawChart() {

var data = google.visualization.arrayToDataTable(" + data.ToString +
"var options = {
          title: '" + name + "',
          pieHole: 0.4,
        };")
        If MaxindexOfGraphicInSuperSchema > 0 AndAlso CounterOfGraphicsInSuperSchema <= MaxindexOfGraphicInSuperSchema Then
            html.AppendLine("var chart = new google.visualization.PieChart(document.getElementById('graph" +
                CounterOfGraphicsInSuperSchema.ToString + "'));
                chart.draw(data, options);
      }
    </script>")
            CounterOfGraphicsInSuperSchema += 1
        Else
            html.AppendLine("var chart = new google.visualization.PieChart(document.getElementById('donutchart" +
                            numOfGraphic.ToString + "'));
        chart.draw(data, options);
      }
    </script>
        <div id=""donutchart" + numOfGraphic.ToString + """ style=""width: 900px; height: 500px;""></div>")
            numOfGraphic += 1
        End If
    End Sub

    Private Sub CallLineGraphic(table As DataTable, name As String, ByVal ParamArray columns() As Integer)
        Dim data = New StringBuilder()
        data.AppendLine("[")
        data.Append("[")
        For i As Integer = 0 To columns.Count - 1
            data.Append("'" + table.Columns(columns(i)).ColumnName.ToString + "'")
            If i + 1 <> columns.Count Then
                data.Append(", ")
            End If
        Next
        data.AppendLine("],")
        For i As Integer = 0 To table.Rows.Count - 1
            data.Append("[")
            For j As Integer = 0 To columns.Count - 1
                data.Append(table.Rows(i)(j).ToString)
                If j + 1 <> columns.Count Then
                    data.Append(", ")
                End If
            Next
            data.Append("]")
            If i + 1 <> table.Rows.Count Then
                data.AppendLine(", ")
            End If
        Next
        data.AppendLine("]);")
        html.AppendLine("<script type = ""text/javascript"">
      google.charts.load('current', {'packages':['corechart', 'line']});
      google.charts.setOnLoadCallback(drawChart);

      function drawChart() {

var data = google.visualization.arrayToDataTable(" + data.ToString +
"var options = {
          title: '" + name + "',
          curveType: 'function',
          legend: { position: 'bottom' }
        };
")
        If MaxindexOfGraphicInSuperSchema > 0 AndAlso CounterOfGraphicsInSuperSchema <= MaxindexOfGraphicInSuperSchema Then
            html.AppendLine("var chart = new google.visualization.LineChart(document.getElementById('graph" +
                CounterOfGraphicsInSuperSchema.ToString + "'));
                chart.draw(data, options);
      }
    </script>")
            CounterOfGraphicsInSuperSchema += 1
        Else
            html.AppendLine("var chart = new google.visualization.LineChart(document.getElementById('curve_chart" +
                            numOfGraphic.ToString + "'));
        chart.draw(data, options);
      }
    </script>
        <div id =""curve_chart" + numOfGraphic.ToString +
        """ style=""margin: auto;""></div>")
            numOfGraphic += 1
        End If
    End Sub

    Private Sub CallNewParagraph(name As String)
        html.Append("<table style=""cursor: pointer;background-color: rgb(49, 133, 156);margin-bottom:5px;margin-left:" +
                    margin.ToString + "px;font-size:"" class=""group_title"" 
                    cellpadding=""0"" cellspacing=""0""; onclick=""toggle('t" + i.ToString + "');"" 
                    onmouseover=""this.style.cursor = 'hand'; this.style.backgroundColor = '#1d7188';"" 
                    onmouseout=""this.style.backgroundColor = '#31859c';"">
                                <tr style="""">
                                    <td style=""width: 0px;""><span id=""t" + i.ToString + "img"" class=""collapse""></span></td>
                                    <td style=""
                                    font-family: Arial;padding-left:5px;font-weight: bold; color: White;font-size:18px;"">")
        html.Append(name)
        html.AppendLine("</td>
                                </tr>
                            </table>
<div id = ""t" + i.ToString + """>")
        margin += 15
        i += 1
    End Sub

    Private Sub CallColumnGraphic(table As DataTable, ByVal ParamArray columns() As Integer)
        Dim data = New StringBuilder()
        data.AppendLine("[")
        data.Append("[")
        For i As Integer = 0 To columns.Count - 1
            data.Append("'" + table.Columns(columns(i)).ColumnName + "'")

            data.Append(", ")
            If i + 1 = columns.Count Then
                data.Append("{ role: 'annotation' }")
            End If
        Next
        data.AppendLine("], ")
        For i As Integer = 0 To table.Rows.Count - 1
            data.Append("[")
            For j As Integer = 0 To columns.Count - 1
                If j = 0 Then
                    data.Append("'" + table.Rows(i)(j).ToString + "'")
                Else
                    data.Append(table.Rows(i)(j))
                End If

                data.Append(", ")
                If j + 1 = columns.Count Then
                    data.Append("''")
                End If
            Next
            data.AppendLine("]")
            If i + 1 <> table.Rows.Count Then
                data.Append(", ")
            End If
        Next
        data.AppendLine("]);")
        html.AppendLine("<script type = ""text/javascript"">
      google.charts.load(""visualization"", ""1"" , {packages:[""corechart""]});
      google.charts.setOnLoadCallback(drawChart);
      function drawChart() {
        var data = google.visualization.arrayToDataTable(" + data.ToString + "
        var options = {
        legend: { position: 'top', maxLines: 3, textStyle: {color: 'black', fontSize: 16 } },
		isStacked: true, 
      };")
        If MaxindexOfGraphicInSuperSchema > 0 AndAlso CounterOfGraphicsInSuperSchema <= MaxindexOfGraphicInSuperSchema Then
            html.AppendLine("var chart = new google.visualization.ColumnChart(document.getElementById('graph" +
                CounterOfGraphicsInSuperSchema.ToString + "'));
                chart.draw(data, options);
      }
    </script>")
            CounterOfGraphicsInSuperSchema += 1
        Else
            html.AppendLine("var chart = new google.visualization.ColumnChart(document.getElementById('chart_div" + numOfGraphic.ToString + "'));
        chart.draw(data, options);
 	  }

    </script>
    <div id=""chart_div" + numOfGraphic.ToString + """ style=""margin: auto; max-width:700px; height:350px;max-height:400px;""></div>")
            numOfGraphic += 1
        End If
    End Sub

    Private Sub CreateHTML()
        Using fs = New FileStream("tables.htm", FileMode.Create)
            Using w = New StreamWriter(fs, Encoding.UTF8)
                w.WriteLine(html)
            End Using
        End Using
        html = Nothing
        Process.Start("tables.htm")
    End Sub

    Private Sub CallTable(table As DataTable)
        Dim doubl As Double = Nothing
        html.AppendLine("<table style=""margin-top:15px;margin-bottom:15px;"">")
        html.AppendLine("<tr>")
        For i As Integer = 0 To table.Columns.Count - 1
            html.AppendLine("<th>" + table.Columns(i).ColumnName + "</th>")
        Next
        html.AppendLine("</tr>")
        For i As Integer = 0 To table.Rows.Count - 1
            html.AppendLine("<tr>")
            For j As Integer = 0 To table.Columns.Count - 1
                If Double.TryParse(table.Rows(i)(j), doubl) AndAlso table.Rows(i)(j) Mod 1 <> 0 Then
                    html.AppendLine("<td>" + Double.Parse(table.Rows(i)(j)).ToString("N2") + "</td>")
                Else
                    html.AppendLine("<td>" + table.Rows(i)(j).ToString() + "</td>")
                End If
            Next
            html.AppendLine("</tr>")
        Next
        html.AppendLine("</table>")
    End Sub

    Private Sub HtmlStart()
        ' Тут все стили и затем скрипты...
        ' И мне почему-то стыдно за такой вид
        html = New StringBuilder()
        html.AppendLine("<html>")
        html.AppendLine("<head>")
        html.AppendLine("<script type=""text/javascript"" src=""https://www.gstatic.com/charts/loader.js""></script>")
        html.AppendLine("<style type=""text/css"">")
        html.AppendLine("table.group_title td {font-family:Calibri; font-size: 26px;
border: 0px solid white; border-top: 0px solid black; margin-top: 5px; width: 100%;}

span.collapse{display:inline-block;width:16px;height:16px;background: #fff url(data:image/gif;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAGXRFWHRTb2Z0d2FyZQBBZG9iZSBJbWFnZVJlYWR5ccllPAAAAdtJREFUeNqkUzsvBFEUPjNzZhY7az3CFh4RkfWqNKJFolDoyIaCTlR0SqGhEp0oNKqlEgkRCX6ARIdEZLWWwdrZsY95ufeuO3Zssgon+XLP6ztz7jl3BNd14T+Cc2tPIIpiGBFXiR0jiPzBSRLELctacRznAx3HBoLVgai8GBurhaY6qSL7JWVH4mfpxatbk5pLomEYQBCbGg1CveoCqVwRNGdypIZxKBd1XaeVIo1hkSWUysTSDTuPtvp9/sawBITHroqZTJo5yX3KCvDYbz8dPI+hYXwyxbZtBi6maQKPJRKPJFYsIkkIqqp6Mczlcl4HFEPTF2WDm1y+9/SDjSjk83ngPLQs12srm81De/uIl5xInLOzs3P0Z+/4TNYuAeehICBTstkCaasA++s1XvJgrBjbX/9ZraZVkQICcB55P9VMKRQckGUF0unPkq8VY6U+RQn4Yqgooe/hKNT03b2tbZCdsiyUzYXzMBBgSlJLiZGmesWXdLgpf6/T9L/lVxMIL8k+bGjX0NIz0/quw1B3RzWEgsiGVBwUMHCbQksB7B1n4OlV3L07nz2lvTUE66JdfcM7C2pD7ziA0Fz5X3KfM293J7eX89tG6v6BFqDTCJUN4G+hL0v/EmAAoNXlG97vnHoAAAAASUVORK5CYII=) 0px 0px no-repeat;}
span.expand{display:inline-block;width:16px;height:16px;background: #fff url(data:image/gif;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAGXRFWHRTb2Z0d2FyZQBBZG9iZSBJbWFnZVJlYWR5ccllPAAAAZ1JREFUeNqkkr1OAkEQx+fuhkPgkK8ghYZYGPzo7GjFxMJeQ7TQzljhGxhs9A2MhY0V2muMifoAJpaQGB+Aw0NPjuOA+3J34eg8ME7yT2Y3+/vv7OxwruvCfwL3T+rA83wMEctkXSTKjGFkooplWceO43yj49hAVF7NBUrFjWlIxwVf+kO1M5WHVumlatLlEa/rOhAVt9cjkJBcIM6+ome2CmHGUBY1TaNOmVSMZwcmiVRMAMKxp2K73WKb5D0TG9DGexzqeocltm0zeWGaJi0TOh2D5bY9MBcEBEmSwOOw2+2OKqDK7zz53n5zloNerwceh5bljsoyjB5kswX/f8cG+XYBPA45DlliGH1SVh+uT8O+BooyRQw48DgyPyGW9PsOBAIitFodXwNRDA7BAYeiGB02R6TLiUfY4zAYZImsqHwmnRAnguWmCYST2cW68gqzS7tzXxrkF+dDEI0ga9JvUlSAq9s21Jv8Ze1x754jJslIPLewsnZxKCWXNwG4mTFj1Gh/1u6qzwfnuvr2Tg1oN6J/asAg6GRpPwIMAFcAzawVzQR4AAAAAElFTkSuQmCC) 0px 0px no-repeat;}
body, th, td {font-family:Calibri; font-size: 16px;}
table { border-collapse: collapse; margin:auto;max-width:100%;}
div { margin:auto;}
th {word-break: break-all; max-width:75px;text-align:center}
h1 { font-family: Calibri; font-size: 32px; margin-bottom: 30px;}
th{border:2px solid #1d7188;font-weight: bold; padding: 3px 1px;white-space:pre-wrap; font-size: 20px; background-color: #EBEEEE;}
td{border:2px solid #1d7188; white-space: nowrap;font-size: 16px;}
</style></head><body><header style=""font-family: Arial; 
display: block;	
color: White; margin:center; margin-top: 5px; font-weight: bold;
			width: 100%; font-size: 32px; padding-top: 3%;
			height: 100px; text-align: center;
			background-color: #18424e; margin-bottom: 0px; 
			border-bottom: solid 1px darkgray;
"">Отчет</header>")
    End Sub

    Private Sub HtmlEnd()
        html.AppendLine("<script type=""text/javascript"">")
        For j = 0 To i - 1
            html.AppendLine("document.getElementById('t" + j.ToString + "').style.display = 'none';")
        Next
        html.AppendLine("function Init() {
	var arr = document.getElementsByTagName('DIV');
	for (el in arr) {
		var obj = arr[el];
		var id = arr[el].id;
		var node = arr[el].nodeName;
		var tag = arr[el].tagName;
		if (arr[el].nodeName == 'DIV' && arr[el].id.indexOf('t') == 0) {
			toggle(arr[el].id);
			arr[el].style.display = 'none';
		}
	}
	toggle('t0');
}

function toggle(objName) {
	var imgName = objName + 'img';

	var obj = document.getElementById(objName);
	var objImg = document.getElementById(imgName);

	if (obj.style.display != 'none') {
		obj.style.display = 'none';
		objImg.className = 'expand';
	}
	else {
		obj.style.display = 'block';
		objImg.className = 'collapse';
	}
}
</script>")
        html.AppendLine("<footer style=""font-size: 24px;
display: block;
			height: 80px;
			background-color: #18424e; padding-left: 1%;
			border-top: solid 1px darkgray;
font-weight: bold;padding-top: 1%; color: White;
"">" + DateTime.Today + "</footer></body></html>")
    End Sub

    Private Sub ButtonForQueries_Click(sender As Object, e As EventArgs) Handles ButtonForQueries.Click
        Dim sql = TextBoxForQueries.Text
        Try
            Dim dt = New DataTable()
            Using adapter = New OleDb.OleDbDataAdapter(sql, New OleDb.OleDbConnection(connectionString))
                adapter.Fill(dt)
                ds.Tables.Add(dt)
            End Using
            TextBoxForQueries.Text = ""
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub AddText(text As String)
        Try
            html.Append("<div style=""display:block;margin-left:" + margin.ToString + ";
font-family:Arial;
width: 35%;
margin: auto;
			height: 20px;
			padding: 5px;
			font-size: 18px;
			font-weight: bold;
			text-align:center;
			color: black;
			background-color: gainsboro;
			margin-top: 30px;"">" + text + "</div>")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub CloseParagraph()
        html.AppendLine("</div>")
        margin -= 15
    End Sub

    Private Sub ButtonForChangeDataBase_Click(sender As Object, e As EventArgs) Handles ButtonForChangeDataBase.Click
        GetPath()
        OpenConnection()
    End Sub
End Class

