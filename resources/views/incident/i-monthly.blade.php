@extends('layouts.admin')
@section('main-content')
<!-- Page Heading -->
<h1 class="h3 ml-4 mb-4 text-gray-800">{{ __('Monthly Report Incident') }}</h1>

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

<div class="card shadow p-4 mb-2">
    <form method="GET" action="{{ route('i-monthly.download') }}">
        <input type="hidden" name="_token" value="{{ csrf_token() }}">
        <div class="pl-lg-4">
            <div class="row">
                <div class="col-lg-3">
                    <div class="form-group">
                        <label class="form-control-label">Start Date</label>
                        <input type="date" id="start_date" class="form-control" name="start_date" onchange="setEndDate()" required>
                    </div>
                </div>
            </div>
            <div class="row">
                <div class="col-lg-3">
                    <div class="form-group">
                        <label class="form-control-label">End Date</label>
                        <input type="date" id="end_date" class="form-control" name="end_date"required>
                    </div>
                </div>
            </div>
        </div>

        <!-- Button -->
        <div class="pl-lg-4">
            <div class="row">
                <div class="col-lg-3">
                    <button type="submit" class="btn btn-primary">Generate</button>
                </div>
            </div>
        </div>
    </form>
</div>

<script>
    function setEndDate() {
        const startDateInput = document.getElementById('start_date');
        const endDateInput = document.getElementById('end_date');

        if (startDateInput.value) {
            const startDate = new Date(startDateInput.value);
            const startDay = startDate.getDate();

            // Check if the start date is the first day of the month
            if (startDay === 1) {
                const endDate = new Date(startDate.getFullYear(), startDate.getMonth() + 1, 1);
                endDateInput.value = endDate.toISOString().split('T')[0];
            }
        }
    }
</script>

@endsection