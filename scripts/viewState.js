function restoreViewState($form, data, callback) {
    $form.find(':input[name]:not([type=button],[type=submit])').val('');
    for(var field in data) {
        var state = data[field];
        var $el = $form.find(":input[name="+field+"]");
        if (Array.isArray(state)) {
            state.forEach(function(s, index) {
                $el.eq(index).val(s.value).attr('disabled', s.disabled);
            });
        } else {
            $el = $el.eq(0);
            if ($el.type == 'radio' || $el.type == 'checkbox') {
                $el.attr('checked', state.value);
            } else {
                $el.val(state.value);
            }
            if ($el[0].tagName.toLowerCase() == 'select') $el.change();
            $el.attr('disabled', state.disabled);
        }
    }
    typeof callback === "function" && callback(data);
}

function bundleViewState($form) {
    $form = $($form);
    var bundle = {};
    $.each($.makeArray($form.find(':input[name]:not([type=button],[type=submit])')),
        function(index, el) {
            if (el.readOnly) return;
            var name = el.name;
            var state = {};
            var type = el.tagName == 'input' ? el.type : el.tagName;
            switch(type) {
                case "select":
                    state.value = el.selectedIndex === -1 ? null : el.options[el.selectedIndex];
                    break;
                case "radio":
                case "checkbox":
                    state.value = el.checked;
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
    $form = $($form);
    var $hidden=$("<input name='view_state' type='hidden' />");
    $form.submit(function() {
        $hidden.val(bundleViewState($form));
    }).append($hidden);
    $form.find('#btnsavedraft').click(function() {
        $.post("../api/save-view-state",
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
        $.post("../api/get-view-state",
            {
                user_id: init_data.user_id,
                user_type: init_data.user_type,
                view_name: init_data.view_name
            },
            function (data) {
                if (data.status==='ok') {
                    if (!data.data) {
                        $.messager.show({ msg: "当前没有已保存的草稿可读取。", title: "读取草稿", timeout: 2000 });
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