<%
Sub SetEditor(cont)
    '新建对象
    dim editor
    set editor = New CKEditor
    
    '设置安装路径
    editor.basePath = "../inc/ckeditor/"
    
		editor.config("toolbar") = Array( _
		Array("Source","-","Save","NewPage","Preview","-","Templates"), _
		Array("Cut","Copy","Paste","PasteText","PasteFromWord","-","Print", "SpellChecker", "Scayt"), _
		Array("Undo","Redo","-","Find","Replace","-","SelectAll","RemoveFormat"), _
		Array("Form", "Checkbox", "Radio", "TextField", "Textarea", "Select", "Button", "ImageButton", "HiddenField"), _
		"/", _
		Array("Bold","Italic","Underline","Strike","-","Subscript","Superscript"), _
		Array("NumberedList","BulletedList","-","Outdent","Indent","Blockquote"), _
		Array("JustifyLeft","JustifyCenter","JustifyRight","JustifyBlock"), _
		Array("Link","Unlink","Anchor"), _
		Array("Image","Flash","Video","Table","HorizontalRule","Smiley","SpecialChar","PageBreak"), _
		"/", _
		Array("Styles","Format","Font","FontSize","lineheight"), _
		Array("TextColor","BGColor"), _
		Array("Maximize", "ShowBlocks","-","About") _
		)
		
    editor.config("height") = 300
    editor.config("extraPlugins") = "video,lineheight"
    editor.config("filebrowserVideoBrowseUrl") = "../inc/ckfinder/ckfinder.html?Type=Videos"
    editor.instanceConfig("skin") = "office2013"
    editor.instanceConfig("font_names") = "Arial/Arial;Georgia/Georgia;宋体/宋体;黑体/黑体;仿宋/仿宋_GB2312;楷体/楷体_GB2312;隶书/隶书;幼圆/幼圆;微软雅黑/微软雅黑;"
    '嵌入图片上传模块
    CKFinder_SetupCKEditor editor, "../inc/ckfinder/", empty, empty
    
    '替换textarea样式
    editor.editor "content1", cont
End Sub

Sub SetEditorWithName(name,cont,height)
    '新建对象
    dim editor
    set editor = New CKEditor
    
    '设置安装路径
    editor.basePath = "../inc/ckeditor/"
    
    If IsNull(height) Then height=300
    If Not IsNumeric(height) Then height=300
    editor.config("height") = height
    editor.instanceConfig("skin") = "office2013"
    editor.instanceConfig("font_names") = "Arial/Arial;Georgia/Georgia;宋体/宋体;黑体/黑体;仿宋/仿宋_GB2312;楷体/楷体_GB2312;隶书/隶书;幼圆/幼圆;微软雅黑/微软雅黑;"
    
    '嵌入图片上传模块
    CKFinder_SetupCKEditor editor, "../inc/ckfinder/", empty, empty
    
    '替换textarea样式
    editor.editor name, cont
End Sub
%>