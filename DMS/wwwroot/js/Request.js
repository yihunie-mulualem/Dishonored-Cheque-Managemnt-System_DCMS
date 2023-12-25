$(document).ready(function () {
    GetRequests();
});

function GetRequests() {
    $.ajax({
        URL:'/Requests/GetRequests',
        type:'get',
        datatype:'json',
        contentType: 'application/json;charset=utf-8',

        success: function (response) {
            if (response == null || response == undefined || response.length == 0) {
                var object = '';
                object += '<tr>';
                object += '<td colspan="8"> ' + 'Request Not Found' + '</td>';
                object += '</tr>';
                $('#tblBody').html(object);
            } else {
                var object = '';
                $.each(res, function (index, item) {
                    /*var ind = index + 1;*/
                    
                    object += '<tr>';
                   /* object += '<td>' + ind + '</td>';*/
                    object += '<td>' + item.ApplicationName + '</td>';
                    object += '<td>' + item.RequestedBy + '</td>';
                    object += '<td>' + item.Department + '</td>';
                    object += '<td>' + item.ResonForChange + '</td>';
                    object += '<td>' + item.Description + '</td>';
                    object += '<td>' + item.Remark + '</td>';
                    object += '<td>' + item.CreatedBy + '</td>';
                    object += '<td>' + item.Status + '</td>';
                    //object += '<td> <a href="#" class="btn btn-primary btn-sm" onclick="Edit(' + item.Id + ')">Edit</a> <a href="#" class="btn btn-danger btn-sm" onclick="Delete(' + item.Id + ')">Edit</a> </td>';
                    object += '<tr>';
                });
                $('#tblBody').html(object);
            }
        },
        Error: function () {
            alert('Unable to read !!!')
        }



    });

}
