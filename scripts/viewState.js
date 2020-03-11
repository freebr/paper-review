function restoreViewState($form, data, callback) {
    $form.find(':input[name]:not([type=button],[type=submit],[type=radio],[type=checkbox],[readonly])').val('');
    $form.find(':input[type=radio]:not([readonly]),:input[type=checkbox]:not([readonly])').attr('checked', false);
    for(var field in data) {
        var state = data[field];
        var $el = $form.find(":input[name="+field+"]");
        if (Array.isArray(state)) {
            state.forEach(function(s, index) {
                var $sub = $el.eq(index);
                if (!$sub.size()) return;
                if (['radio', 'checkbox'].indexOf($sub.attr('type')) != -1) {
                    $sub.prop({'checked': s.checked, 'disabled': s.disabled});
                } else {
                    $sub.val(s.value).attr('disabled', s.disabled);
                }
                if ($sub[0].tagName.toLowerCase() == 'select') $sub.change();
            });
        } else {
            $el = $el.eq(0);
            if (['radio', 'checkbox'].indexOf($el.attr('type')) != -1) {
                $el.prop('checked', state.checked);
            } else {
                $el.val(state.value);
            }
            $el.attr('disabled', state.disabled);
            if ($el[0].tagName.toLowerCase() == 'select') $el.change();
        }
    }
    typeof callback === "function" && callback(data);
}

function bundleViewState($form) {
    $form = $($form);
    var bundle = {};
    $.each($.makeArray($form.find(':input[name]:not([type=button],[type=submit],[readonly])')),
        function(index, el) {
            var name = el.name;
            var state = {};
            var type = el.tagName.toLowerCase() == 'input' ? el.type : el.tagName.toLowerCase();
            switch(type) {
                case "select":
                    state.value = el.selectedIndex === -1 ? null : el.value;
                    break;
                case "radio":
                case "checkbox":
                    state.checked = el.checked;
                    break;
                case "textarea":
                default:
                    if (!el.value) return;
                    state.value = el.value;
            }
            if (el.disabled) state.disabled = true;
            if (!bundle[name]) {
                bundle[name] = state;
            } else if (Array.isArray(bundle[name])) {
                bundle[name].push(state);
            } else {
                bundle[name] = [bundle[name], state];
            }
        }
    );
    return JSON.stringify(bundle);
}

function initViewState($form, init_data, callback) {
    var path_api=location.origin+"/PaperReview/api/";
    $form = $($form);
    var $hidden=$("<input name='view_state' type='hidden' />");
    $form.submit(function() {
        $(":hidden[name=view_state]").val(bundleViewState($form));
    }).append($hidden);
    $form.find('#btnsavedraft').click(function() {
        $.post(path_api+"save-view-state",
            {
                user_id: init_data.user_id,
                user_type: init_data.user_type,
                view_name: init_data.view_name,
                view_state: bundleViewState($form)
            },
            function (data) {
                if (data.status==='ok') {
                    $.messager.show({ msg: "操作成功", title: "保存草稿", timeout: 2000 });
                }
            }
        );
    });
    $form.find('#btnloaddraft').click(function() {
        $.post(path_api+"get-view-state",
            {
                user_id: init_data.user_id,
                user_type: init_data.user_type,
                view_name: init_data.view_name
            },
            function (data) {
                if (data.status==='ok') {
                    if (!data.data) {
                        $.messager.show({ msg: "当前没有可读取的草稿。", title: "读取草稿", timeout: 2000 });
                        return;
                    }
                    if (!confirm("确定读取草稿吗？这将覆盖当前所做的修改。")) return;
                    restoreViewState($form, JSON.parse(data.data), callback);
                    $.messager.show({ msg: "读取完成", title: "读取草稿", timeout: 2000 });
                }
            }
        );
    });
    init_data.view_state && restoreViewState($form, init_data.view_state, callback);
}