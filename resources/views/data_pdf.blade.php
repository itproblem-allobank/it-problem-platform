<h1>Report IT Problem Weekly</h1>
<link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/css/bootstrap.min.css" integrity="sha384-ggOyR0iXCbMQv3Xipma34MD+dH/1fQ784/j6cY/iJTQUOhcWr7x9JvoRxT2MZw1T" crossorigin="anonymous">

<body>
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
	</style>


	<div id="piechart" style="width: 900px; height: 500px;"></div>


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
				<td>{{ $i++ }}</td>
				<td>{{ $d->summary }}</td>
				<td>{{ $d->root_cause }}</td>
				<td>{{ $d->work_around }}</td>
				<td>{{ $d->status }} - {{$d->pending_reason}}</td>
			</tr>
			@endforeach
		</tbody>
	</table>
</body>

<script type="text/javascript" src="https://www.gstatic.com/charts/loader.js"></script>
<script type="text/javascript">
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
		console.log(datajira);

		var data = google.visualization.arrayToDataTable([
			['Priority', 'Total'],
			['High', highest + high],
			['Medium', medium],
			['Low', low + lowest]
		]);

		var options = {
			title: 'My Daily Activities'
		};

		var chart = new google.visualization.PieChart(document.getElementById('piechart'));

		chart.draw(data, options);
	}
</script>