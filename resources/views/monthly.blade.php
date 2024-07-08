@extends('layouts.admin')
@section('main-content')
<!-- Page Heading -->
<h1 class="h3 mb-4 text-gray-800">{{ __('Data') }}</h1>

@if (session('success'))
<div class="alert alert-success border-left-success alert-dismissible fade show" role="alert">
    {{ session('success') }}
    <button type="button" class="close" data-dismiss="alert" aria-label="Close">
        <span aria-hidden="true">&times;</span>
    </button>
</div>
@endif

@if ($errors->any())
<div class="alert alert-danger border-left-danger" role="alert">
    <ul class="pl-4 my-2">
        @foreach ($errors->all() as $error)
        <li>{{ $error }}</li>
        @endforeach
    </ul>
</div>
@endif


<body>
    <div style="margin-left: 25px; margin-bottom: 15px">
        <button type=" button" class="btn btn-primary" data-toggle="modal" data-target="#import">Import Data</button>
        <a type=" button" class="btn btn-primary" href="{{ route('monthly.pptx') }}">Generate PPTX</a>
        @if($data == '[]')
        <button type="button" class="btn btn-warning" data-toggle="modal" data-target="#delete" disabled>Delete Data</button>
        @else
        <button type="button" class="btn btn-warning" data-toggle="modal" data-target="#delete">Delete Data</button>
        <form action="{{ route('monthly.export') }}" method="POST" enctype="multipart/form-data">
            @csrf
            <input type="hidden" name="total" id="url_total">
            <input type="hidden" name="pending" id="url_pending">
            <input type="hidden" name="closed" id="url_closed">
            <button type="submit" class="btn btn-danger mt-2">Export to Word</button>
        </form>
        @endif
    </div>

    
    @if($data == '[]')
    <br>
    @else
    <div class="card shadow p-4 mb-2">
        <div class="row mb-4">
            @foreach ($total as $data)
            <div class="col">
                <table border="1" style="border-radius: 10px">
                    <tr>
                        <th colspan="3" style="text-align: center; padding-left: 10px; padding-right: 10px">{{ $data['total']}}<br>{{ $data['problem']}}</td>
                    </tr>
                    <tr>
                        <td style="text-align: center;padding-left: 10px; padding-right: 10px">High</td>
                        <td style="text-align: center;padding-left: 10px; padding-right: 10px">Medium</td>
                        <td style="text-align: center;padding-left: 10px; padding-right: 10px">Low</td>
                    </tr>
                    <tr>
                        <td style="text-align: center;padding-left: 10px; padding-right: 10px">{{ $data['high'] }}</td>
                        <td style="text-align: center;padding-left: 10px; padding-right: 10px">{{ $data['medium'] }}</td>
                        <td style="text-align: center;padding-left: 10px; padding-right: 10px">{{ $data['low'] }}</td>
                    </tr>
                    <tr>
                        <td style="text-align: center;padding-left: 10px; padding-right: 10px">{{ $data['highmonthly'] }}</td>
                        <td style="text-align: center;padding-left: 10px; padding-right: 10px">{{ $data['mediummonthly'] }}</td>
                        <td style="text-align: center;padding-left: 10px; padding-right: 10px">{{ $data['lowmonthly'] }}</td>
                    </tr>
                </table>
            </div>
            @endforeach
        </div>

        <div class="row">
            <div class="col-6" id="tsp_category"></div>
            <div class="col-6" id="ticket_yearly"></div>
        </div>
    </div>
    @endif

    <!-- <div id="container_chart" class="card shadow p-4 mb-2">
        <div class="row">
            <div class="col-4" id="chart_total"></div>
            <div class="col-4" id="chart_pending"></div>
            <div class="col-4" id="chart_closed"></div>
        </div>
    </div> -->

    <div class="card shadow p-4">
        <table id="getTables" class="stripe" style="width:100%">
            <thead>
                <tr>
                    <th>No</th>
                    <th>Problem</th>
                    <th>Summary</th>
                    <th>Priority</th>
                    <th>Status</th>
                    <th>Impact Analyst</th>
                    <th>Root Cause</th>
                    <th>Work Around</th>
                    <th>Assignee</th>
                    <th>Updated</th>
                </tr>
            </thead>
        </table>
    </div>

</body>





<link href="https://cdn.datatables.net/1.10.23/css/jquery.dataTables.min.css" rel="stylesheet">
<link href="https://cdn.datatables.net/1.10.23/css/dataTables.bootstrap4.min.css" rel="stylesheet">
<script src="http://ajax.googleapis.com/ajax/libs/jquery/1.9.1/jquery.min.js"></script>
<script src="https://cdn.datatables.net/1.10.23/js/jquery.dataTables.min.js" defer></script>

<!-- New Charts -->
<script type="text/javascript" src="https://www.gstatic.com/charts/loader.js"></script>
<script type="text/javascript">
    google.charts.load('current', {
        packages: ['corechart', 'bar']
    });
    google.charts.setOnLoadCallback(drawTSPCategory);

    function drawTSPCategory() {
        $.ajax({
            url: "{{ route('monthly.chartcategory') }}",
            dataType: "json",
            success: function(jsonData) {

                category = [];
                category.push("Status");
                jsonData.closed.forEach(function(data) {
                    category.push(data.problem);
                    category.push({
                        type: "string",
                        role: "annotation"
                    });
                })

                //declare data closed
                ticketclosed = [];
                ticketclosed.push("Closed");
                jsonData.closed.forEach(function(data) {
                    ticketclosed.push(data.total);
                    ticketclosed.push(data.total.toString());
                })

                //declare data pending
                ticketpending = [];
                ticketpending.push("Pending");
                jsonData.pending.forEach(function(data) {
                    ticketpending.push(data.total);
                    ticketpending.push(data.total.toString());
                })

                //create charts
                var data = google.visualization.arrayToDataTable([
                    category,
                    ticketclosed,
                    ticketpending,
                ]);

                var options = {
                    height: 400,
                    title: 'Ticket By Category',
                    annotations: {
                        textStyle: {
                            fontSize: 10,
                        },
                    },

                };
                let chart_div = document.getElementById('tsp_category');
                var chart = new google.visualization.ColumnChart(chart_div);
                chart.draw(data, options);
            },
        });
    }

    google.charts.load('current', {
        packages: ['corechart', 'bar']
    });
    google.charts.setOnLoadCallback(drawTicketYearly);

    function drawTicketYearly() {
        $.ajax({
            url: "{{ route('monthly.chartyearly') }}",
            dataType: "json",
            success: function(jsonData) {
                var data = google.visualization.arrayToDataTable([
                    ['Status', 'Total', {type: "string",role: "annotation"}, 'Closed', {type: "string",role: "annotation"},'Pending', {type: "string",role: "annotation"},'Work In Progress',{type: "string",role: "annotation"}],
                    ['2024', jsonData.total2024, jsonData.total2024.toString(), jsonData.closed2024, jsonData.closed2024.toString(), jsonData.pending2024, jsonData.pending2024.toString(), jsonData.wip2024, jsonData.wip2024.toString()],
                    ['2023', jsonData.total2023, jsonData.total2023.toString(), jsonData.closed2023, jsonData.closed2023.toString(), jsonData.pending2023, jsonData.pending2023.toString(), jsonData.wip2023, jsonData.wip2023.toString()],
                ]);

                var options = {
                    title: 'Ticket By Yearly',
                    height: 400,
                    legend: {
                        position: 'top',
                        // maxLines: 3
                    },
                    bar: {
                        groupWidth: '30%'
                    },
                    isStacked: true
                };

                //create charts
                let chart_div = document.getElementById('ticket_yearly');
                var chart = new google.visualization.BarChart(chart_div);
                chart.draw(data, options);
            },
        });
    }
</script>


<!-- total -->
<!-- <script type="text/javascript" src="https://www.gstatic.com/charts/loader.js"></script>
<script type="text/javascript">
    var image_url = [];

    google.charts.load('current', {
        packages: ['corechart', 'bar']
    });
    google.charts.setOnLoadCallback(drawTotal);

    function drawTotal() {
        $.ajax({
            url: "{{ route('monthly.chart') }}",
            dataType: "json",
            success: function(jsonData) {
                var data = new google.visualization.DataTable();
                data.addColumn('string', 'Category');
                data.addColumn('number', 'High');
                data.addColumn({
                    type: 'string',
                    role: 'annotation'
                })
                data.addColumn('number', 'Med');
                data.addColumn({
                    type: 'string',
                    role: 'annotation'
                })
                data.addColumn('number', 'Low');
                data.addColumn({
                    type: 'string',
                    role: 'annotation'
                })

                tempdata = [];

                jsonData.total.forEach(function(data) {
                    tempdata.push([data.problem + ': ' + data.total, data.high, data.high.toString(), data.medium, data.medium.toString(), data.low, data.low.toString()])
                })

                data.addRows(tempdata);

                var options = {
                    height: 400,
                    title: 'Ticket Total',
                    colors: ['#B22222', '#FFA500', '#9ACD32'],
                    isStacked: true,
                    annotations: {
                        textStyle: {
                            fontSize: 10,
                        },
                    },

                };
                let chart_div = document.getElementById('chart_total');
                var chart = new google.visualization.ColumnChart(chart_div);
                google.visualization.events.addListener(chart, 'ready', function() {
                    chart_div.innerHTML = '<img src="' + chart.getImageURI() + '"' + '>';

                    $("#url_total").val(chart.getImageURI());
                });
                chart.draw(data, options);
            },
        });
    }


    // pending
    google.charts.load('current', {
        packages: ['corechart', 'bar']
    });
    google.charts.setOnLoadCallback(drawPending);

    function drawPending() {
        $.ajax({
            url: "{{ route('monthly.chart') }}",
            dataType: "json",
            success: function(jsonData) {
                var data = new google.visualization.DataTable();
                data.addColumn('string', 'Category');
                data.addColumn('number', 'High');
                data.addColumn({
                    type: 'string',
                    role: 'annotation'
                })
                data.addColumn('number', 'Med');
                data.addColumn({
                    type: 'string',
                    role: 'annotation'
                })
                data.addColumn('number', 'Low');
                data.addColumn({
                    type: 'string',
                    role: 'annotation'
                })

                tempdata = [];

                jsonData.pending.forEach(function(data) {
                    tempdata.push([data.problem + ': ' + data.total, data.high, data.high.toString(), data.medium, data.medium.toString(), data.low, data.low.toString()])
                })
                data.addRows(tempdata);

                var options = {
                    height: 400,
                    title: 'Ticket Pending',
                    colors: ['#B22222', '#FFA500', '#9ACD32'],
                    isStacked: true,
                    annotations: {
                        textStyle: {
                            fontSize: 10,
                        },
                    },

                };

                let chart_div = document.getElementById('chart_pending');
                var chart = new google.visualization.ColumnChart(chart_div);
                google.visualization.events.addListener(chart, 'ready', function() {
                    chart_div.innerHTML = '<img src="' + chart.getImageURI() + '"' + '>';

                    $("#url_pending").val(chart.getImageURI());
                });
                chart.draw(data, options);
            },
        });
    }

    //closed
    google.charts.load('current', {
        packages: ['corechart', 'bar']
    });
    google.charts.setOnLoadCallback(drawClosed);

    function drawClosed() {
        $.ajax({
            url: "{{ route('monthly.chart') }}",
            dataType: "json",
            success: function(jsonData) {
                var data = new google.visualization.DataTable();
                data.addColumn('string', 'Category');
                data.addColumn('number', 'High');
                data.addColumn({
                    type: 'string',
                    role: 'annotation'
                })
                data.addColumn('number', 'Med');
                data.addColumn({
                    type: 'string',
                    role: 'annotation'
                })
                data.addColumn('number', 'Low');
                data.addColumn({
                    type: 'string',
                    role: 'annotation'
                })

                tempdata = [];

                jsonData.closed.forEach(function(data) {
                    tempdata.push([data.problem + ': ' + data.total, data.high, data.high.toString(), data.medium, data.medium.toString(), data.low, data.low.toString()])
                })

                data.addRows(tempdata);

                var options = {
                    height: 400,
                    title: 'Ticket Closed',
                    colors: ['#B22222', '#FFA500', '#9ACD32'],
                    isStacked: true,
                    annotations: {
                        textStyle: {
                            fontSize: 10,
                        },
                    },

                };

                let chart_div = document.getElementById('chart_closed');
                var chart = new google.visualization.ColumnChart(chart_div);
                google.visualization.events.addListener(chart, 'ready', function() {
                    chart_div.innerHTML = '<img src="' + chart.getImageURI() + '"' + '>';

                    $("#url_closed").val(chart.getImageURI());
                });
                chart.draw(data, options);
            },
        });
    }
</script> -->


<script>
    $(document).ready(function() {
        let i = 1;
        $('#getTables').DataTable({
            processing: true,
            serverSide: true,
            ajax: "{{ route('monthly') }}",
            columns: [{
                    data: 'id',
                    name: 'id'
                },
                {
                    data: 'problem',
                    name: 'problem'
                },
                {
                    data: 'summary',
                    name: 'summary'
                },
                {
                    data: 'priority',
                    name: 'priority'
                },
                {
                    data: 'status',
                    name: 'status'
                },
                {
                    data: 'impact_analyst',
                    name: 'impact_analyst'
                },
                {
                    data: 'root_cause',
                    name: 'root_cause'
                },
                {
                    data: 'work_around',
                    name: 'work_around'
                },
                {
                    data: 'assignee_to',
                    name: 'assignee_to'
                },
                {
                    data: 'updated',
                    name: 'updated'
                },
            ],
            responsive: true
        });
    });
</script>

<!-- modal import -->
<div class="modal fade" id="import" tabindex="-1" role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
    <div class="modal-dialog" role="document">
        <div class="modal-content">
            <div class="modal-header">
                <h5 class="modal-title">Import Data</h5>
                <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                    <span aria-hidden="true">&times;</span>
                </button>
            </div>
            <form action="{{ route('monthly.import') }}" method="POST" enctype="multipart/form-data">
                @csrf
                <div class="modal-body">
                    <div class="form-group">
                        <label>PILIH FILE</label>
                        <input type="file" name="file" class="form-control" required>
                    </div>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-secondary" data-dismiss="modal">TUTUP</button>
                    <button type="submit" class="btn btn-success">IMPORT</button>
                </div>
            </form>
        </div>
    </div>
</div>

<!-- modal delete -->
<div class="modal fade" id="delete" tabindex="-1" role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
    <div class="modal-dialog" role="document">
        <div class="modal-content">
            <div class="modal-header">
                <h5 class="modal-title">Delete Data</h5>
                <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                    <span aria-hidden="true">&times;</span>
                </button>
            </div>
            <form action="{{ route('monthly.delete') }}" method="POST">
                @csrf
                <div class="modal-body">
                    <div class="form-group">
                        <label>Apakah Anda yakin ingin menghapus semua data ini?</label>
                    </div>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-secondary" data-dismiss="modal">Batalkan</button>
                    <button type="submit" class="btn btn-success">Yakin</button>
                </div>
            </form>
        </div>
    </div>
</div>




@endsection