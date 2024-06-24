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
    <div style="margin-left: 25px; margin-bottom: 15px"">
        <form action=" /chart/print" method="POST" enctype="multipart/form-data">
        @csrf
        <input type="hidden" name="weekly" id="weeklyData">
        <input type="hidden" name="total" id="totalData">
        <input type="hidden" name="priority" id="priorityData">
        <button type="button" class="btn btn-primary" data-toggle="modal" data-target="#import">Import Data</button>
        <button type="button" class="btn btn-secondary" data-toggle="modal" data-target="#delete">Delete Data</button>
        <button type="submit" class="btn btn-primary">Export to PDF</button>
        </form>
    </div>


    @if($ticket_weekly == '[]')
    <br>
    @else
    <div class="container-fluid" style="margin-bottom: 25px;">
        <div class="row align-items-start">
            <div class="col-8">
                <div id="chart_weekly"></div>
            </div>
            <div class="col-4">
                <div id="chart_priority"></div>
            </div>
        </div>
        <div class="row align-items-start mt-4">
            <div class="col-12">
                <div id="chart_total"></div>
            </div>
        </div>
    </div>
    @endif

    <div class="container-fluid text-center">
        <div class="card">
            <div class="card-body">
                <div class="table-responsive">
                    <table id="get_data" class="table table-bordered">
                        <thead>
                            <tr>
                                <th>No</th>
                                <th>Env</th>
                                <th>Problem Category</th>
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
            </div>
        </div>
    </div>
</body>


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
            <form action="{{ route('data.import') }}" method="POST" enctype="multipart/form-data">
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
            <form action="{{ route('data.delete') }}" method="POST">
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



<link href="https://cdn.datatables.net/1.10.23/css/jquery.dataTables.min.css" rel="stylesheet">
<link href="https://cdn.datatables.net/1.10.23/css/dataTables.bootstrap4.min.css" rel="stylesheet">
<script src="http://ajax.googleapis.com/ajax/libs/jquery/1.9.1/jquery.min.js"></script>
<script src="https://cdn.datatables.net/1.10.23/js/jquery.dataTables.min.js" defer></script>


<script type="text/javascript">
    $(function() {
        let i = 1;
        var table = $('#get_data').DataTable({
            processing: true,
            serverSide: true,
            ajax: "{{ route('data.getdata') }}",
            columns: [{
                    "render": function() {
                        return i++;
                    }
                },
                {
                    data: 'environment',
                    name: 'environment'
                },
                {
                    data: 'problem_category',
                    name: 'problem_category'
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
            ]
        });
    });
</script>

<script type="text/javascript" src="https://www.gstatic.com/charts/loader.js"></script>
<script type="text/javascript">
    $(function() {
        $("#chart_weekly").append("<div id= 'chart_weekly'></div>");
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
                        category.push(data.problem_category);
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
                            position: 'bottom',
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
                        chart_div.innerHTML = '<img src="' + chart.getImageURI()  + '"' + 'width="650">';
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
        $("#chart_total").append("<div id= 'chart_total'></div>");
        google.charts.load("current", {
            packages: ["corechart"]
        });
        google.charts.setOnLoadCallback(drawChart);

        function drawChart() {
            $.ajax({
                url: "{{ route('chart.total') }}",
                dataType: "json",
                success: function(jsonData) {
                    var category = [];
                    category.push('Category');
                    jsonData.total.forEach(function(data) {
                        category.push(data.problem_category);
                    });

                    var total = [];
                    total.push('Total');
                    jsonData.total.forEach(function(data) {
                        total.push(data.count);
                    })
                    // console.log(jsonData.total);

                    var closed = [];
                    closed.push('Closed');
                    jsonData.total.forEach(function(data) {
                        closed.push(data.count);
                    })

                    var pending = [];
                    pending.push('Pending');
                    jsonData.total.forEach(function(data) {
                        pending.push(data.count);
                    })
                    // console.log(jsonData.pending);

                    var data = google.visualization.arrayToDataTable([
                        category,
                        total,
                        closed,
                        pending,
                    ])

                    var options = {
                        title: "Total Ticket Problem",
                        legend: {
                            position: "bottom",
                            maxlines: 3,
                        },
                        bar: {
                            groupWidth: '80%'
                        },

                    };

                    let chart_div = document.getElementById('chart_total');
                    let chart = new google.visualization.ColumnChart(chart_div);

                    google.visualization.events.addListener(chart, 'ready', function() {
                        chart_div.innerHTML = '<img src="' + chart.getImageURI() + '"' + 'width="700">';
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
        $("#chart_priority").append("<div id= 'chart_priority'></div>");
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
                chart_div.innerHTML = '<img src="' + chart.getImageURI() + '"' + 'width="400">';
            });

            chart.draw(data, options);
        }
        setTimeout(function() {
            let chartsData = $("#chart_priority").html();
            $("#priorityData").val(chartsData);
        }, 5000);
    });
</script>

@endsection