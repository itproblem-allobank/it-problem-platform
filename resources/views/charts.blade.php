<!DOCTYPE html>
<html>

<head>
    <link href="https://cdn.datatables.net/1.10.23/css/jquery.dataTables.min.css" rel="stylesheet">
    <link href="https://cdn.datatables.net/1.10.23/css/dataTables.bootstrap4.min.css" rel="stylesheet">
    <script src="http://ajax.googleapis.com/ajax/libs/jquery/1.9.1/jquery.min.js"></script>
    <script src="https://cdn.datatables.net/1.10.23/js/jquery.dataTables.min.js" defer></script>
    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/css/bootstrap.min.css" integrity="sha384-ggOyR0iXCbMQv3Xipma34MD+dH/1fQ784/j6cY/iJTQUOhcWr7x9JvoRxT2MZw1T" crossorigin="anonymous">
    <style>
        table tr td,
        table tr th {
            font-size: 8pt;
        }

        table.fixed {
            table-layout: fixed;
            width: 100%;
        }

        table.fixed td {
            overflow: hidden;
        }

        body {
            background-color: #D6DBDF;
        }

        .letter {
            background-color: #FFFFFF;
            padding: 20px;
            height: auto;
            width: 700px;
            margin-left: auto;
            margin-right: auto;
        }

        .bordered {
            padding-top: 10px;
            padding-bottom: 10px;
            margin-bottom: 10px;
            border-style: solid;
            border-color: grey;
            border-width: 1px;
        }

        th {
            text-align: center;
        }

        tr {
            text-align: center;
        }
    </style>
</head>


<body>
    <div class="letter">
        <form action=" /chart/print" method="POST" enctype="multipart/form-data">
            @csrf
            <input type="hidden" name="weekly" id="weeklyData">
            <input type="hidden" name="total" id="totalData">
            <input type="hidden" name="priority" id="priorityData">
            <button type="submit" class="btn btn-danger" style="float: right;">Export to PDF</button>
        </form>
        <h1>Report IT Problem Weekly</h1>
        <p id="date"></p>
        <div class="mt-2">
            <div class="bordered" id="chart_priority"></div>
            <div class="bordered" id="chart_weekly"></div>
            <div class="bordered" id="chart_total"></div>
        </div>

        <!-- <div>{{$priority}}</div> -->

        <div class="row">
            <div class="col">
                <table class="table-bordered fixed">
                    <thead>
                        <tr>
                            <th>Paylater</th>
                        </tr>
                    </thead>
                </table>
                <table class="table-bordered fixed">
                    <thead>
                        <tr>
                            <th>H</th>
                            <th>M</th>
                            <th>L</th>
                        </tr>
                    </thead>
                    <tbody>
                        <tr>
                            <td>1</td>
                            <td>2</td>
                            <td>3</td>
                        </tr>
                        <tr>
                            <td>-</td>
                            <td>-</td>
                            <td>-</td>
                        </tr>
                    </tbody>
                </table>
            </div>
            <div class="col">
                <table class="table-bordered fixed">
                    <thead>
                        <tr>
                            <th>Onboarding</th>
                        </tr>
                    </thead>
                </table>
                <table class="table-bordered fixed">
                    <thead>
                        <tr>
                            <th>H</th>
                            <th>M</th>
                            <th>L</th>
                        </tr>
                    </thead>
                    <tbody>
                        <tr>
                            <td>1</td>
                            <td>2</td>
                            <td>3</td>
                        </tr>
                        <tr>
                            <td>-</td>
                            <td>-</td>
                            <td>-</td>
                        </tr>
                    </tbody>
                </table>
            </div>
        </div>

        <table class='table-bordered fixed mt-4'>
            <thead>
                <tr>
                    <th style="width: 5%">No</th>
                    <th style="width: 20%">Problem</th>
                    <th style="width: 30%">Root Cause</th>
                    <th style="width: 30%">Work Around</th>
                    <th style="width: 15%">Status</th>
                </tr>
            </thead>
            <tbody>
                @php $i=1 @endphp
                @foreach($table as $d)
                <tr>
                    <td>{{ $i++ }}</td>
                    <td>{{ $d->summary }}</td>
                    <td>{{ $d->root_cause }}</td>
                    <td>{{ $d->work_around }}</td>
                    <td>{{ $d->status }} - {{$d->pending_reason}}</td>
                </tr>
                @endforeach
            </tbody>
        </table>
    </div>
</body>

</html>





<script type="text/javascript" src="https://www.gstatic.com/charts/loader.js"></script>

<script type="text/javascript">
    var d = new Date();
    document.getElementById("date").innerHTML = d.toLocaleDateString("id-ID");
</script>

<script type="text/javascript">
    $(function() {
        google.charts.load("current", {
            packages: ["corechart"]
        });
        google.charts.setOnLoadCallback(drawChart);

        function drawChart() {
            $.ajax({
                url: "{{ route('chart.weekly') }}",
                dataType: "json",
                success: function(jsonData) {

                    var category = [];
                    category.push('Category');
                    jsonData.data.forEach(function(data) {
                        category.push(data.category);
                        category.push({
                            type: 'string',
                            role: 'annotation'
                        });
                    })

                    var value = [];
                    value.push('Last Week');
                    jsonData.data.forEach(function(data) {
                        value.push(data.count);
                        value.push(data.count.toString());
                    })

                    // console.log(category);
                    // console.log(value);

                    var data = google.visualization.arrayToDataTable([
                        category,
                        value,
                    ]);

                    var options = {
                        title: 'Ticket Weekly',
                        legend: {
                            position: 'top',
                            maxlines: 3,
                        },
                        bar: {
                            groupWidth: '85%'
                        },
                        // isStacked: true
                    };

                    let chart_div = document.getElementById('chart_weekly');
                    let chart = new google.visualization.BarChart(chart_div);

                    google.visualization.events.addListener(chart, 'ready', function() {
                        chart_div.innerHTML = '<img src="' + chart.getImageURI() + '"' + 'width="650">';
                    });

                    chart.draw(data, options);
                }
            });
        }

        setTimeout(function() {
            let chartsData = $("#chart_weekly").html();
            $("#weeklyData").val(chartsData);
        }, 5000);

    });
</script>


<!-- total -->
<script type="text/javascript">
    $(function() {
        google.charts.load("current", {
            packages: ["corechart"]
        });
        google.charts.setOnLoadCallback(drawChart);

        function drawChart() {
            $.ajax({
                url: "{{ route('chart.total') }}",
                dataType: "json",
                success: function(jsonData) {
                    var problem = [];
                    problem.push('Problem');
                    jsonData.total.forEach(function(data) {
                        problem.push(data.problem);
                        problem.push({
                            type: 'string',
                            role: 'annotation'
                        });
                    });

                    var total = [];
                    total.push('Total');
                    jsonData.total.forEach(function(data) {
                        total.push(data.count);
                        total.push(data.count.toString());
                    })
                    // console.log(jsonData.total);

                    var closed = [];
                    closed.push('Closed');
                    jsonData.closed.forEach(function(data) {
                        closed.push(data.count);
                        closed.push(data.count.toString());
                    })

                    var pending = [];
                    pending.push('Pending');
                    jsonData.pending.forEach(function(data) {
                        pending.push(data.count);
                        pending.push(data.count.toString());
                    })
                    // console.log(jsonData.pending);

                    var data = google.visualization.arrayToDataTable([
                        problem,
                        total,
                        closed,
                        pending,
                    ])

                    var options = {
                        title: 'Total Ticket',
                        chartArea: {height: '55%' },
                        legend: {
                            position: 'top',
                            maxLines: 2
                        },
                        bar: {
                            groupWidth: '80%'
                        },

                    };

                    let chart_div = document.getElementById('chart_total');
                    let chart = new google.visualization.ColumnChart(chart_div);

                    google.visualization.events.addListener(chart, 'ready', function() {
                        chart_div.innerHTML = '<img src="' + chart.getImageURI() + '"' + 'width="650">';
                    });

                    chart.draw(data, options);
                }
            });
        }

        setTimeout(function() {
            let chartsData = $("#chart_total").html();
            $("#totalData").val(chartsData);
        }, 5000);

    });
</script>


<!-- piecharts -->
<script type="text/javascript">
    $(function() {
        google.charts.load('current', {
            'packages': ['corechart']
        });
        google.charts.setOnLoadCallback(drawChart);

        function drawChart() {

            var highest = <?php echo $highest; ?>;
            var high = <?php echo $high; ?>;
            var medium = <?php echo $medium; ?>;
            var low = <?php echo $low; ?>;
            var lowest = <?php echo $lowest; ?>;

            var datajira = {
                'high': highest + high,
                'medium': medium,
                'low': low + lowest
            };

            var data = google.visualization.arrayToDataTable([
                ['Priority', 'Total'],
                ['High', highest + high],
                ['Medium', medium],
                ['Low', low + lowest]
            ]);

            var options = {
                title: 'Ticket Priority',
                pieSliceText: 'value'
            };

            let chart_div = document.getElementById('chart_priority');
            let chart = new google.visualization.PieChart(chart_div);

            google.visualization.events.addListener(chart, 'ready', function() {
                chart_div.innerHTML = '<img src="' + chart.getImageURI() + '"' + 'width="650">';
            });

            chart.draw(data, options);
        }
        setTimeout(function() {
            let chartsData = $("#chart_priority").html();
            $("#priorityData").val(chartsData);
        }, 5000);
    });
</script>