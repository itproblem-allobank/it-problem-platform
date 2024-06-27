<!DOCTYPE html>
<html>
<link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/css/bootstrap.min.css" integrity="sha384-ggOyR0iXCbMQv3Xipma34MD+dH/1fQ784/j6cY/iJTQUOhcWr7x9JvoRxT2MZw1T" crossorigin="anonymous">
<style type="text/css">
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

    .letter {
        background-color: #FFFFFF;
        padding: 20px;
        height: auto;
        width: 700px;
        margin-left: auto;
        margin-right: auto;
    }

    .bordered {
        padding: 10px;
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

<body>
    <div>
        <h1>Report IT Problem Weekly</h1>
        <p>{!! $today->toFormattedDateString() !!}</p>

        <div class="mt-2">
            <div class="bordered">
                {!! $priority !!}</div>
            <div class="bordered">
                {!! $weekly !!}</div>
            <div class="bordered">
                {!! $total !!}</div>
        </div>

        <div class="row">
            <div class="col-3">
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
            <div class="col-3">
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