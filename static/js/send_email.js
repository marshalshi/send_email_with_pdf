$(function(){

    $('#all_toggle').click(function(){
        var _this = $(this),
            row_input_toggle = $('.row_toggle'),
            row_toggle = $('tbody tr');

        if (_this.prop('checked')) {
            row_input_toggle.prop('checked', true);
            row_toggle.data('checked', true);
        } else {
            row_input_toggle.prop('checked', false);
            row_toggle.data('checked', false);
        }
    });

    $('.toggle_ae_send').click(function(e){
        e.preventDefault();

        var tr = $($(this).parents('tr'));
            rows = $('tbody tr[data-ae_send='+tr.data('ae_send')+']');

        if ($(tr).data('checked')) {
            rows.data('checked', false);
            rows.find('.row_toggle').prop('checked', false);
        } else {
            rows.data('checked', true);
            rows.find('.row_toggle').prop('checked', true);
        }        
    });

    $('#submit').click(function(){

        var subject = $('#subject').val().trim(),
            content = $('textarea#content').val().trim(),
            send_type = $('input[name=send_type]:checked').val();

        var error_p = $('.error');
        error_p.text('');
        
        if (!subject) {
            error_p.text('Please input subject');
            return false;
        }
        if (!content) {
            error_p.text('Please input content');
            return false;
        }
        if (!send_type) {
            error_p.text('Please select send type');
            return false;
        }

        var selected_rows = $('.row_toggle:checked');
        if (!selected_rows) {
            error_p.text('Please select target sending persons.')
            return false;
        }

        var that = this;
        $(this).prop('disabled', true);
        $('#loader-bar').removeClass('hidden');

        
        var trs = selected_rows.parents('tr:not(.no_file)');

        // get all pdf_path and relevant email address
        var file_email_map = {};
        for (var i=0; i<trs.length; i++) {
            var tr = $(trs[i]);
            var tds = tr.find('td');
            for (var j=0; j<tds.length; j++){
                var td = $(tds[j]);
                // second cell is pdf path
                if (j==1) {
                    var pdf_path = pdf_dict[td.data('value')];
                    file_email_map[pdf_path] = [];
                }
                // get all emails 
                if (j >= 4) {
                    var email = td.data('value');
                    file_email_map[pdf_dict[$(tds[1]).data('value')]].push(email);
                }
            }
        }

        $.ajax({
            url: send_email_json_url,
            type: 'POST',
            dataType: 'json',
            data: {
                'subject': subject,
                'content': content,
                'send_type': send_type,
                'file_email_map': JSON.stringify(file_email_map),
                'csrfmiddlewaretoken': csrf_token
            },
            success: function(data, textStatus, jqXHR){
                alert('Emails already sent.');
            },
            error: function(){
                alert('SERVER ERROR!');
            },
            complete: function(){
                $('#loader-bar').addClass('hidden');
                $(that).prop('disabled', false);
            }
        });
        
    });
    
});
