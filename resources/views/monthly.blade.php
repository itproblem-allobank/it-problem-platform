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
        @if($data == '[]')
            <button type="button" class="btn btn-warning" data-toggle="modal" data-target="#delete" disabled>Delete
                Data</button>
            <button type="button" class="btn btn-danger" disabled>Export to PDF</button>
        @else
            <button type="button" class="btn btn-warning" data-toggle="modal" data-target="#delete">Delete Data</button>
            <a href="{{ route('chart.index') }}" class="btn btn-danger">Export to PDF</a>
        @endif
    </div>


    <div class="card shadow p-4">
        <table id="getTables" class="table table-striped">
            <thead>
                <tr>
                    <th>No</th>
                    <th>Env</th>
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

<script>
    $(document).ready(function () {
        let i = 1;
        $('#getTables').DataTable({
            processing: true,
            serverSide: true,
            ajax: '{{ route('monthly') }}',
            columns: [
                { "render": function () { return i++; } },
                { data: 'environment', name: 'environment' },
                { data: 'problem', name: 'problem' },
                { data: 'summary', name: 'summary' },
                { data: 'priority', name: 'priority' },
                { data: 'status', name: 'status' },
                { data: 'impact_analyst', name: 'impact_analyst' },
                { data: 'root_cause', name: 'root_cause' },
                { data: 'work_around', name: 'work_around' },
                { data: 'assignee_to', name: 'assignee_to' },
                { data: 'updated', name: 'updated' },
            ],
            responsive: true
        });
    });
</script>

@endsection