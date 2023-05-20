function validateEntry(inside_field, optional_field) {
    var error = 0;
    var vvl = document.getElementById(inside_field).getElementsByClassName("grpreq");
    for (var i = 0; i < vvl.length; i++) {
        var val = $(".grpreq")[i].value;
        var curr_Id = $(".grpreq")[i].id;
        if (val === '') {
            if (optional_field === curr_Id) {

            } else {
                document.getElementById(curr_Id).style.border = "1px solid red";
                document.getElementById(curr_Id + "_error").innerHTML = " * This Field Is Required";
                error = 1;
            }

        } else {
            document.getElementById(curr_Id).style.border = "1px solid green";
            document.getElementById(curr_Id + "_error").innerHTML = "";
        }
    }
    return error;
}


function ShowNotify(title, txt, tp) {
    new PNotify({
        title: title,
        text: txt,
        type: tp,
        hide: true,
        styling: 'bootstrap3'
    });
}

