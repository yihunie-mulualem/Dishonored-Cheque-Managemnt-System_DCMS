


$(function () {
    $("#loaderbody").addClass('hide');
    $(document).bind('ajaxStart', function () {
        $("#loaderbody").removeClass("hide");
    }).bind('ajaxStop', function () {
        $("#loaderbody").addClass('hide');
    });
});

showInPopUp = (url, title) => {
    $.ajax({
        type: "GET",
        url: url,
        success: function (response) {
            $("#form-modal .modal-title").html(title);
            $("#form-modal .modal-body").html(response);
            $("#form-modal").modal('show');
        }
    });
}
JQueryAjaxPost = form => {
    try {
        $.ajax({
            type: 'POST',
            url: form.action,
            data: new FromData(form),
            contentType: false,
            processData: false,
            success: function (res) {
                if (res.isValid) {
                    $("#view-all").html(res.html);
                    $("#form-modal .modal-title").html('');
                    $("#form-modal .modal-body").html('');
                    $("#form-modal").modal('hide');
                    $.notify("Submited Successfully", "success");
                    $.notify('Submited Successfully', { globalPosition: 'top center', className: 'success' });

                } else {
                    $("#form-modal .modal-body").html(res.html);
                    $.notify('Something is Wrong', { globalPosition: 'top center', className: 'warn' });
                }
            },
            error: function (err) {
                console.log(err);
            }

        });
    } catch (e) {
        console.log(e);
    }
}





/**
$(document).ready(function () {
    $("#button1").click(function () {
            $.ajax({
                url: '@Url.Action("AuthorizeReject", "DishonoredCheques")',
                type: 'POST',
                data: $("#Id").value,
                success: function (result) {
                    alert("wellcome bekele");
                },
                error: function () {
                    // Handle the error here
                    console.log('An error occurred.');
                }
            });
        });
    });

    */
function Instancecheck()
{
    var Acc = $("#number").val();
    var url = "../../DishonoredCheques/InsertCheques/";

    $.ajax({
        url: "../../DishonoredCheques/Instancecheckvariable/", // the url of the controller action
        type: "GET", // the http method
        data: { id: Acc }, // the data to send
        success: function (result) {
            // Check if the result contains a redirect URL
            if (result.redirectUrl) {
                // Redirect to the specified URL
                window.location.href = result.redirectUrl;
            } else {
                // Handle other cases if needed
            }
        },
       // error: function () {
      //     alert("well");
       // }
        
    });
    // for signof
}
    
showInPopUp = (url, title) => {
    $.ajax({
        type: "GET",
        url: url,
        success: function (response) {
            $("#form-modal .modal-title").html(title);
            $("#form-modal .modal-body").html(response);
            $("#form-modal").modal('show');
        }
    });
}
JQueryAjaxPost = form => {
    try
    {
        $.ajax({
            type: 'POST',
            url: form.action,
            data: new FromData(form),
            contentType: false,
            processData: false,
            success: function (res) {
                if (res.isValid) {
                    $("#view-all").html(res.html);
                    $("#form-modal .modal-title").html('');
                    $("#form-modal .modal-body").html('');
                    $("#form-modal").modal('hide');
                    $.notify("Submited Successfully", "success");
                    $.notify('Submited Successfully', { globalPosition: 'top center', className: 'success' });

                } else {
                    $("#form-modal .modal-body").html(res.html);
                    $.notify('Something is Wrong', { globalPosition: 'top center', className: 'warn' });
                }
            },
            error: function (err) {
                console.log(err);
            }

        });
    } catch (e) {
        console.log(e);
    }
}

