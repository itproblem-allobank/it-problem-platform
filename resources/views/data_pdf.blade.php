<!DOCTYPE html>
<html>
<head>
	<title>Membuat Laporan PDF Dengan DOMPDF Laravel</title>
	<link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/css/bootstrap.min.css" integrity="sha384-ggOyR0iXCbMQv3Xipma34MD+dH/1fQ784/j6cY/iJTQUOhcWr7x9JvoRxT2MZw1T" crossorigin="anonymous">
</head>
<body>
	<style type="text/css">
		table tr td,
		table tr th{
			font-size: 6pt;
		}
		table.fixed {
      table-layout: fixed;
      width: 100%;
    	}
		table.fixed td {
		overflow: hidden;
		}
	</style>
	<center>
		<h5>Report Weekly IT Problem</h4>
	</center>
	
	<table class='table-bordered fixed'>
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
			@foreach($data as $d)
			<tr>
				<td >{{ $i++ }}</td>
                <td >{{ $d->summary }}</td>
                <td>{{ $d->root_cause }}</td>
                <td>{{ $d->work_around }}</td>
                <td>{{ $d->status }} - {{$d->pending_reason}}</td>
			</tr>
			@endforeach
		</tbody>
	</table>
 
</body>
</html>