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

