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


<div class="container-fluid text-center">
    <div class="card">
        <div class="card-body">
            <div class="table-responsive">
                <table id="get_data" class="table table-bordered">
                    <thead>
                        <tr>
                            <th>Key</th>
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

<link href="https://cdn.datatables.net/1.10.23/css/jquery.dataTables.min.css" rel="stylesheet">
<link href="https://cdn.datatables.net/1.10.23/css/dataTables.bootstrap4.min.css" rel="stylesheet">
<script src="http://ajax.googleapis.com/ajax/libs/jquery/1.9.1/jquery.min.js"></script>
<script src="https://cdn.datatables.net/1.10.23/js/jquery.dataTables.min.js" defer></script>
<script src="https://cdn.datatables.net/1.10.23/js/dataTables.bootstrap4.min.js"></script>


<script type="text/javascript">
    $(function () {
          var table = $('#get_data').DataTable({
              processing: true,
              serverSide: true,
              ajax: "{{ route('data.getdata') }}",
              columns: [
                  {data: 'code_jira', name: 'code_jira'},
                  {data: 'environment', name: 'environment'},
                  {data: 'problem_category', name: 'problem_category'},
                  {data: 'summary', name: 'summary'},
                  {data: 'priority', name: 'priority'},
                  {data: 'status', name: 'status'},
                  {data: 'impact_analyst', name: 'impact_analyst'},
                  {data: 'root_cause', name: 'root_cause'},
                  {data: 'work_around', name: 'work_around'},
                  {data: 'assignee_to', name: 'assignee_to'},
                  {data: 'updated', name: 'updated'},
              ]
          });
        });
</script>

<!-- <script>
    $(document).ready(function () {

        var dataTable = $('#get_data').DataTable({
            'processing': true,
            'serverSide': true,
            'serverMethod': 'get',
            'ajax': {
                'url': '{{route("data.getdata")}}',
            },
            'aaSorting': [],
            'columns': [
                { data: 'code_jira' },
                { data: 'environment' },
                { data: 'problem_category' },
                { data: 'summary' },
                { data: 'priority' },
                { data: 'status' },
                { data: 'impact_analyst' },
                { data: 'root_cause' },
                { data: 'work_around' },
                { data: 'assignee_to' },
                { data: 'updated' }
            ]
        });

    });
</script> -->


@endsection