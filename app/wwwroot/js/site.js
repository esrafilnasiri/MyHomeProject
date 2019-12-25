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