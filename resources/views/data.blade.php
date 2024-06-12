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

<script>
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
</script>


@endsection