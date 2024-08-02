@extends('layouts.admin')
@section('main-content')
<!-- Page Heading -->
<h1 class="h3 ml-4 mb-4 text-gray-800">{{ __('Service Request') }}</h1>

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

<!-- Display Page -->

<body>
    <div style="margin-left: 25px; margin-bottom: 15px">
        <button type=" button" class="btn btn-primary" data-toggle="modal" data-target="#import">Import Data</button>
        @if($data == '[]')

        @else
        <button type="button" class="btn btn-warning" data-toggle="modal" data-target="#delete">Delete Data</button>
        @endif
    </div>

    <div class="card shadow p-4">
        <table id="getTables" class="stripe" style="width:100%">
            <thead>
                <tr>
                    <th>No</th>
                    <th>Issue Type</th>
                    <th>Summary</th>
                    <th>Assignee</th>
                    <th>Reporter</th>
                    <th>Status</th>
                    <th>Created</th>
                    <th>Updated</th>
                    <th>Priority</th>
                    <th>Sub Category</th>
                    <th>Ticket Number</th>
                </tr>
            </thead>
        </table>
    </div>
</body>

<!-- Library -->
<link href="https://cdn.datatables.net/1.10.23/css/jquery.dataTables.min.css" rel="stylesheet">
<link href="https://cdn.datatables.net/1.10.23/css/dataTables.bootstrap4.min.css" rel="stylesheet">
<script src="http://ajax.googleapis.com/ajax/libs/jquery/1.9.1/jquery.min.js"></script>
<script src="https://cdn.datatables.net/1.10.23/js/jquery.dataTables.min.js" defer></script>

<!-- Script Tables -->
<script>
    $(document).ready(function() {
        let i = 1;
        $('#getTables').DataTable({
            processing: true,
            serverSide: true,
            ajax: "{{ route('service.index') }}",
            columns: [{
                    data: 'id',
                    name: 'id'
                },
                {
                    data: 'issue_type',
                    name: 'issue_type'
                },
                {
                    data: 'summary',
                    name: 'summary'
                },
                {
                    data: 'assignee',
                    name: 'assignee'
                },
                {
                    data: 'reporter',
                    name: 'reporter'
                },
                {
                    data: 'status',
                    name: 'status'
                },
                {
                    data: 'created',
                    name: 'created'
                },
                {
                    data: 'updated',
                    name: 'updated'
                },
                {
                    data: 'priority',
                    name: 'priority'
                },
                {
                    data: 'sub_category',
                    name: 'sub_category'
                },
                {
                    data: 'ticket_number',
                    name: 'ticket_number'
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
            <form action="{{ route('service.import') }}" method="POST" enctype="multipart/form-data">
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
            <form action="{{ route('service.delete') }}" method="POST">
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