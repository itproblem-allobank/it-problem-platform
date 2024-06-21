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
        <button type=" button" class="btn btn-primary" data-toggle="modal" data-target="#import">
        Import Data
        </button>
        <!-- <a href="/data/cetak_pdf" class="btn btn-primary" target="_blank">CETAK PDF</a> -->
        <form action="/chart/print" method="POST" enctype="multipart/form-data">
            @csrf
            <input type="hidden" name="weekly" id="weeklyData">
            <input type="hidden" name="total" id="totalData">
            <input type="hidden" name="priority" id="priorityData">
            <input type="submit"  class="btn btn-primary" value="Print Chart">
        </form>
    </div>


    <div class="container-fluid" style="margin-bottom: 25px;">
        <div class="row align-items-start">
            <div class="col">
            <div id="chart_weekly"  style="width: 450px; height: 450px;"></div>
            </div>
            <div class="col">
            <div id="chart_total"  style="width: 450px; height: 450px;"></div>
            </div>
            <div class="col">
                <div id="chart_priority" style="width: 450px; height: 450px;"></div>
            </div>
        </div>
    </div>

    <!-- <div id="draw-charts"></div> -->

    <!-- {{ $ticket_weekly }} -->
    <!-- @if($ticket_weekly == '[]')
    <br>
    @else
    <div class="container-fluid" style="margin-bottom: 25px;">
        <div class="row align-items-start">
            <div class="col">
                <div id="columnchart_material" style="width: 450px; height: 450px;"></div>
            </div>
            <div class="col">
                <div id="barchart_values" style="width: 450px; height: 450px;"></div>
            </div>
            <div class="col">
                <div id="piechart" style="width: 450px; height: 450px;"></div>
            </div>
        </div>
    </div>
    @endif -->

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

<!-- modal export -->
<div class="modal fade" id="export" tabindex="-1" role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
    <div class="modal-dialog" role="document">
        <div class="modal-content">
            <div class="modal-header">
                <h5 class="modal-title">Export Data</h5>
                <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                    <span aria-hidden="true">&times;</span>
                </button>
            </div>
            <form action="{{ route('data.export') }}" method="POST">
                <div class="modal-footer">
                    <button type="button" class="btn btn-secondary" data-dismiss="modal">TUTUP</button>
                    <button type="submit" class="btn btn-success">EXPORT</button>
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
                    })
                    category.push({
                        role: 'annotation'
                    });

                    var value = [];
                    value.push('Last Week');
                    jsonData.data.forEach(function(data) {
                        value.push(data.count);
                    })
                    value.push('');

                    var data = google.visualization.arrayToDataTable([
                        category,
                        value,
                    ]);

                    var options = {
                        title: 'Ticket Weekly',
                        legend: {
                            position: 'bottom',
                            maxlines: 2,
                        },
                        bar: {
                            groupWidth: '80%'
                        },
                        // isStacked: true
                    };

                    let chart_div = document.getElementById('chart_weekly');
                    let chart = new google.visualization.BarChart(chart_div);

                    google.visualization.events.addListener(chart, 'ready', function() {
                        chart_div.innerHTML = '<img src="' + chart.getImageURI() + '">';
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
                    console.log(jsonData.pending);

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
                            maxlines: 4,
                        },
                        bar: {
                            groupWidth: '100%'
                        },

                    };

                    let chart_div = document.getElementById('chart_total');
                    let chart = new google.visualization.ColumnChart(chart_div);

                    google.visualization.events.addListener(chart, 'ready', function() {
                        chart_div.innerHTML = '<img src="' + chart.getImageURI() + '">';
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
                chart_div.innerHTML = '<img src="' + chart.getImageURI() + '">' ;
            });

            chart.draw(data, options);
        }
        setTimeout(function() {
            let chartsData = $("#chart_priority").html();
            $("#priorityData").val(chartsData);
        }, 5000);
    });
</script>



<!-- <script type="text/javascript">
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

        var chart = new google.visualization.PieChart(document.getElementById('piechart'));

        chart.draw(data, options);
    }
</script>

<script type="text/javascript">
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
                })
                category.push({
                    role: 'annotation'
                });

                var value = [];
                value.push('Last Week');
                jsonData.data.forEach(function(data) {
                    value.push(data.count);
                })
                value.push('');

                var data = google.visualization.arrayToDataTable([
                    category,
                    value,
                ]);

                var options = {
                    title: 'Ticket Weekly',
                    legend: {
                        position: 'bottom',
                        maxlines: 2,
                    },
                    bar: {
                        groupWidth: '80%'
                    },
                    // isStacked: true
                };
                var chart = new google.visualization.BarChart(document.getElementById("barchart_values"));
                chart.draw(data, options);
            }
        });
    }
</script>

<script type="text/javascript">
    google.charts.load('current', {
        'packages': ['bar']
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
                console.log(jsonData.pending);

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
                        maxlines: 2,
                    },
                    bar: {
                        groupWidth: '100%'
                    },

                };

                var chart = new google.visualization.ColumnChart(document.getElementById('columnchart_material'));
                chart.draw(data, google.charts.Bar.convertOptions(options));

            }
        });
    }
</script> -->


@endsection