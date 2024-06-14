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
	</style>
	<center>
		<h5>Report Weekly IT Problem</h4>
	</center>
 
	<table class='table table-bordered'>
		<thead>
			<tr>
                <th>No</th>
                <th>Problem</th>
                <th>Root Cause</th>
                <th>Work Around</th>
                <th>Status</th>
			</tr>
		</thead>
		<tbody>
			@php $i=1 @endphp
			@foreach($data as $p)
			<tr>
				<td>{{ $i++ }}</td>
                <td>{{ $p->summary }}</td>
                <td>{{ $p->root_cause }}</td>
                <td>{{ $p->work_around }}</td>
                <td>{{ $p->status }}</td>
			</tr>
			@endforeach
		</tbody>
	</table>
 
</body>
</html>