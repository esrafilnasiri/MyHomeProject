// Please see documentation at https://docs.microsoft.com/aspnet/core/client-side/bundling-and-minification
// for details on configuring this project to bundle and minify static web assets.

// Write your JavaScript code.

    $('#btnTsetmc').on('click', function (e) {
        e.preventDefault(); 
        $.ajax({
            url: "/Home/FromTseTmc",
            type: "POST",
            data: JSON.stringify({ 'Options': 'null' }),
            dataType: "json",
            traditional: true,
            contentType: "application/json; charset=utf-8",
            success: function (data) {
                if (data.success) {
                    alert("Done");
                } else {
                    alert("Error occurs on the Database level!");
                }
            },
            error: function () {
                alert("An error has occured!!!");
            }
        });
    });

$('#btncharts').on('click', function (e) {
    e.preventDefault();
    $.ajax({
        url: "/Home/CreateChart",
        type: "POST",
        data: JSON.stringify({ 'Options': 'null' }),
        dataType: "json",
        traditional: true,
        contentType: "application/json; charset=utf-8",
        success: function (data) {
            if (data.success)
            {
                console.log(data);
                console.log(data.chartData);
                Highcharts.chart('container', {
                    title: data.chartData.title,
                    subtitle: data.chartData.subtitle,
                    yAxis: data.chartData.yAxis,
                    xAxis: data.chartData.xAxis,
                    legend: data.chartData.legend,
                    plotOptions: data.chartData.plotOptions,
                    series: data.chartData.series,
                    //responsive: {
                    //    rules: [{
                    //        condition: {
                    //            maxWidth: 500
                    //        },
                    //        chartOptions: {
                    //            legend: {
                    //                layout: 'horizontal',
                    //                align: 'center',
                    //                verticalAlign: 'bottom'
                    //            }
                    //        }
                    //    }]
                    //}
                });

                Highcharts.chart('maxZarar3day', {
                    title: data.maxZarar3Day.title,
                    subtitle: data.maxZarar3Daysubtitle,
                    yAxis: data.maxZarar3Day.yAxis,
                    xAxis: data.maxZarar3Day.xAxis,
                    legend: data.maxZarar3Day.legend,
                    plotOptions: data.maxZarar3Day.plotOptions,
                    series: data.maxZarar3Day.series,
                });

                Highcharts.chart('maxZarar7day', {
                    title: data.maxZarar7Day.title,
                    subtitle: data.maxZarar7Daysubtitle,
                    yAxis: data.maxZarar7Day.yAxis,
                    xAxis: data.maxZarar7Day.xAxis,
                    legend: data.maxZarar7Day.legend,
                    plotOptions: data.maxZarar7Day.plotOptions,
                    series: data.maxZarar7Day.series,
                });

                Highcharts.chart('maxZarar14day', {
                    title: data.maxZarar14Day.title,
                    subtitle: data.maxZarar14Daysubtitle,
                    yAxis: data.maxZarar14Day.yAxis,
                    xAxis: data.maxZarar14Day.xAxis,
                    legend: data.maxZarar14Day.legend,
                    plotOptions: data.maxZarar14Day.plotOptions,
                    series: data.maxZarar14Day.series,
                });

                Highcharts.chart('maxZarar30day', {
                    title: data.maxZarar30Day.title,
                    subtitle: data.maxZarar30Daysubtitle,
                    yAxis: data.maxZarar30Day.yAxis,
                    xAxis: data.maxZarar30Day.xAxis,
                    legend: data.maxZarar30Day.legend,
                    plotOptions: data.maxZarar30Day.plotOptions,
                    series: data.maxZarar30Day.series,
                });

                Highcharts.chart('maxZarar30daySabe', {
                    title: data.maxZarar30DaySabe.title,
                    subtitle: data.maxZarar30DaySabe.subtitle,
                    yAxis: data.maxZarar30DaySabe.yAxis,
                    xAxis: data.maxZarar30DaySabe.xAxis,
                    legend: data.maxZarar30DaySabe.legend,
                    plotOptions: data.maxZarar30DaySabe.plotOptions,
                    series: data.maxZarar30DaySabe.series,
                });

                Highcharts.chart('maxSood3day', {
                    title: data.maxSood3Day.title,
                    subtitle: data.maxSood3Daysubtitle,
                    yAxis: data.maxSood3Day.yAxis,
                    xAxis: data.maxSood3Day.xAxis,
                    legend: data.maxSood3Day.legend,
                    plotOptions: data.maxSood3Day.plotOptions,
                    series: data.maxSood3Day.series,
                });

                Highcharts.chart('maxSood7day', {
                    title: data.maxSood7Day.title,
                    subtitle: data.maxSood7Daysubtitle,
                    yAxis: data.maxSood7Day.yAxis,
                    xAxis: data.maxSood7Day.xAxis,
                    legend: data.maxSood7Day.legend,
                    plotOptions: data.maxSood7Day.plotOptions,
                    series: data.maxSood7Day.series,
                });

                Highcharts.chart('maxSood14day', {
                    title: data.maxSood14Day.title,
                    subtitle: data.maxSood14Daysubtitle,
                    yAxis: data.maxSood14Day.yAxis,
                    xAxis: data.maxSood14Day.xAxis,
                    legend: data.maxSood14Day.legend,
                    plotOptions: data.maxSood14Day.plotOptions,
                    series: data.maxSood14Day.series,
                });

                Highcharts.chart('maxSood30day', {
                    title: data.maxSood30Day.title,
                    subtitle: data.maxSood30Daysubtitle,
                    yAxis: data.maxSood30Day.yAxis,
                    xAxis: data.maxSood30Day.xAxis,
                    legend: data.maxSood30Day.legend,
                    plotOptions: data.maxSood30Day.plotOptions,
                    series: data.maxSood30Day.series,
                });

                Highcharts.chart('maxSood30daySabe', {
                    title: data.maxSood30DaySabe.title,
                    subtitle: data.maxSood30DaySabe.subtitle,
                    yAxis: data.maxSood30DaySabe.yAxis,
                    xAxis: data.maxSood30DaySabe.xAxis,
                    legend: data.maxSood30DaySabe.legend,
                    plotOptions: data.maxSood30DaySabe.plotOptions,
                    series: data.maxSood30DaySabe.series,
                });
            } else {
                alert(data.message);
            }
        },
        error: function () {
            alert("An error has occured!!!");
        }
    });
});


$('#btnSahamyab').on('click', function (e) {

    e.preventDefault();
    $.ajax({
        url: "/Home/FromSahamyab",
        type: "POST",
        data: JSON.stringify({ 'Options': 'null' }),
        dataType: "json",
        traditional: true,
        contentType: "application/json; charset=utf-8",
        success: function (data) {
            if (data.success) {
                alert("Done");
            } else {
                alert("Error occurs on the Database level!");
            }
        },
        error: function () {
            alert("An error has occured!!!");
        }
    });
});



$('#btnoldDays').on('click', function (e) {

    e.preventDefault();
    var date = $('#txtdate').val();
    $.ajax({
        url: "/Home/FromTseTmcOldDate",
        type: "POST",
        //data: JSON.stringify( { 'Option': date }),
        data:{ 'Option': date },
        //dataType: "json",
        //traditional: true,
        //contentType: "application/json; charset=utf-8",

        success: function (data) {
            if (data.success) {
                alert("Done");
            } else {
                alert("Error occurs on the Database level!");
            }
        },
        error: function () {
            alert("An error has occured!!!");
        }
    });
});

