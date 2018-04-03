/*
 * CKFinder
 * ========
 * http://ckfinder.com
 * Copyright (C) 2007-2012, CKSource - Frederico Knabben. All rights reserved.
 *
 * The software, this file, and its contents are subject to the CKFinder
 * License. Please read the license.txt file before using, installing, copying,
 * modifying, or distributing this file or part of its contents. The contents of
 * this file is part of the Source Code of CKFinder.
 *
 */

/**
 * @fileOverview Defines the {@link CKFinder.lang} object for the Chinese (Taiwan)
 *		language.
 */

/**
 * Contains the dictionary of language entries.
 * @namespace
 */
CKFinder.lang['zh-tw'] =
{
	appTitle : 'CKFinder',

	// Common messages and labels.
	common :
	{
		// Put the voice-only part of the label in the span.
		unavailable		: '%1<span class="cke_accessibility">, unavailable</span>', // MISSING
		confirmCancel	: 'Some of the options were changed. Are you sure you want to close the dialog window?', // MISSING
		ok				: 'OK', // MISSING
		cancel			: 'Cancel', // MISSING
		confirmationTitle	: 'Confirmation', // MISSING
		messageTitle	: 'Information', // MISSING
		inputTitle		: 'Question', // MISSING
		undo			: 'Undo', // MISSING
		redo			: 'Redo', // MISSING
		skip			: 'Skip', // MISSING
		skipAll			: 'Skip all', // MISSING
		makeDecision	: 'What action should be taken?', // MISSING
		rememberDecision: 'Remember my decision' // MISSING
	},


	// Language direction, 'ltr' or 'rtl'.
	dir : 'ltr',
	HelpLang : 'zh-tw',
	LangCode : 'zh-tw',

	// Date Format
	//		d    : Day
	//		dd   : Day (padding zero)
	//		m    : Month
	//		mm   : Month (padding zero)
	//		yy   : Year (two digits)
	//		yyyy : Year (four digits)
	//		h    : Hour (12 hour clock)
	//		hh   : Hour (12 hour clock, padding zero)
	//		H    : Hour (24 hour clock)
	//		HH   : Hour (24 hour clock, padding zero)
	//		M    : Minute
	//		MM   : Minute (padding zero)
	//		a    : Firt char of AM/PM
	//		aa   : AM/PM
	DateTime : 'mm/dd/yyyy HH:MM',
	DateAmPm : ['上午', '下午'],

	// Folders
	FoldersTitle	: '目录',
	FolderLoading	: '载入中...',
	FolderNew		: '请输入新目录名称: ',
	FolderRename	: '请输入新目录名称: ',
	FolderDelete	: '确定删除 "%1" 这个目录吗?',
	FolderRenaming	: ' (修改目录...)',
	FolderDeleting	: ' (删除目录...)',
	DestinationFolder	: 'Destination Folder', // MISSING

	// Files
	FileRename		: '请输入新档案名称: ',
	FileRenameExt	: '确定变更这个档案的副档名吗? 变更後 , 此档案可能会无法使用 !',
	FileRenaming	: '修改档案名称...',
	FileDelete		: '确定要删除这个档案 "%1"?',
	FilesDelete	: 'Are you sure you want to delete %1 files?', // MISSING
	FilesLoading	: '载入中...',
	FilesEmpty		: 'The folder is empty.', // MISSING
	DestinationFile	: 'Destination File', // MISSING
	SkippedFiles	: 'List of skipped files:', // MISSING

	// Basket
	BasketFolder		: 'Basket', // MISSING
	BasketClear			: 'Clear Basket', // MISSING
	BasketRemove		: 'Remove from Basket', // MISSING
	BasketOpenFolder	: 'Open Parent Folder', // MISSING
	BasketTruncateConfirm : 'Do you really want to remove all files from the basket?', // MISSING
	BasketRemoveConfirm	: 'Do you really want to remove the file "%1" from the basket?', // MISSING
	BasketRemoveConfirmMultiple	: 'Do you really want to remove %1 files from the basket?', // MISSING
	BasketEmpty			: 'No files in the basket, drag and drop some.', // MISSING
	BasketCopyFilesHere	: 'Copy Files from Basket', // MISSING
	BasketMoveFilesHere	: 'Move Files from Basket', // MISSING

	// Global messages
	OperationCompletedSuccess	: 'Operation completed successfully.', // MISSING
	OperationCompletedErrors		: 'Operation completed with errors.', // MISSING
	FileError				: '%s: %e', // MISSING

	// Move and Copy files
	MovedFilesNumber		: 'Number of files moved: %s.', // MISSING
	CopiedFilesNumber	: 'Number of files copied: %s.', // MISSING
	MoveFailedList		: 'The following files could not be moved:<br />%s', // MISSING
	CopyFailedList		: 'The following files could not be copied:<br />%s', // MISSING

	// Toolbar Buttons (some used elsewhere)
	Upload		: '上传档案',
	UploadTip	: '上传一个新档案',
	Refresh		: '重新整理',
	Settings	: '偏好设定',
	Help		: '说明',
	HelpTip		: '说明',

	// Context Menus
	Select			: '选择',
	SelectThumbnail : 'Select Thumbnail', // MISSING
	View			: '浏览',
	Download		: '下载',

	NewSubFolder	: '建立新子目录',
	Rename			: '重新命名',
	Delete			: '删除',
	DeleteFiles		: 'Delete Files', // MISSING

	CopyDragDrop	: 'Copy Here', // MISSING
	MoveDragDrop	: 'Move Here', // MISSING

	// Dialogs
	RenameDlgTitle		: 'Rename', // MISSING
	NewNameDlgTitle		: 'New Name', // MISSING
	FileExistsDlgTitle	: 'File Already Exists', // MISSING
	SysErrorDlgTitle : 'System Error', // MISSING

	FileOverwrite	: 'Overwrite', // MISSING
	FileAutorename	: 'Auto-rename', // MISSING
	ManuallyRename	: 'Manually rename', // MISSING

	// Generic
	OkBtn		: '确定',
	CancelBtn	: '取消',
	CloseBtn	: '关闭',

	// Upload Panel
	UploadTitle			: '上传新档案',
	UploadSelectLbl		: '请选择要上传的档案',
	UploadProgressLbl	: '(档案上传中 , 请稍候...)',
	UploadBtn			: '将档案上传到伺服器',
	UploadBtnCancel		: '取消',

	UploadNoFileMsg		: '请从你的电脑选择一个档案.',
	UploadNoFolder		: 'Please select a folder before uploading.', // MISSING
	UploadNoPerms		: 'File upload not allowed.', // MISSING
	UploadUnknError		: 'Error sending the file.', // MISSING
	UploadExtIncorrect	: 'File extension not allowed in this folder.', // MISSING

	// Flash Uploads
	UploadLabel			: 'Files to Upload', // MISSING
	UploadTotalFiles	: 'Total Files:', // MISSING
	UploadTotalSize		: 'Total Size:', // MISSING
	UploadSend			: '上传档案',
	UploadAddFiles		: 'Add Files', // MISSING
	UploadClearFiles	: 'Clear Files', // MISSING
	UploadCancel		: 'Cancel Upload', // MISSING
	UploadRemove		: 'Remove', // MISSING
	UploadRemoveTip		: 'Remove !f', // MISSING
	UploadUploaded		: 'Uploaded !n%', // MISSING
	UploadProcessing	: 'Processing...', // MISSING

	// Settings Panel
	SetTitle		: '设定',
	SetView			: '浏览方式:',
	SetViewThumb	: '缩图预览',
	SetViewList		: '清单列表',
	SetDisplay		: '显示栏位:',
	SetDisplayName	: '档案名称',
	SetDisplayDate	: '档案日期',
	SetDisplaySize	: '档案大小',
	SetSort			: '排序方式:',
	SetSortName		: '依 档案名称',
	SetSortDate		: '依 档案日期',
	SetSortSize		: '依 档案大小',
	SetSortExtension		: 'by Extension', // MISSING

	// Status Bar
	FilesCountEmpty : '<此目录没有任何档案>',
	FilesCountOne	: '1 个档案',
	FilesCountMany	: '%1 个档案',

	// Size and Speed
	Kb				: '%1 KB',
	Mb				: '%1 MB', // MISSING
	Gb				: '%1 GB', // MISSING
	SizePerSecond	: '%1/s', // MISSING

	// Connector Error Messages.
	ErrorUnknown	: '无法连接到伺服器 ! (错误代码 %1)',
	Errors :
	{
	 10 : '不合法的指令.',
	 11 : '连接过程中 , 未指定资源形态 !',
	 12 : '连接过程中出现不合法的资源形态 !',
	102 : '不合法的档案或目录名称 !',
	103 : '无法连接：可能是使用者权限设定错误 !',
	104 : '无法连接：可能是伺服器档案权限设定错误 !',
	105 : '无法上传：不合法的副档名 !',
	109 : '不合法的请求 !',
	110 : '不明错误 !',
	111 : 'It was not possible to complete the request due to resulting file size.', // MISSING
	115 : '档案或目录名称重复 !',
	116 : '找不到目录 ! 请先重新整理 , 然後再试一次 !',
	117 : '找不到档案 ! 请先重新整理 , 然後再试一次 !',
	118 : 'Source and target paths are equal.', // MISSING
	201 : '伺服器上已有相同的档案名称 ! 您上传的档案名称将会自动更改为 "%1".',
	202 : '不合法的档案 !',
	203 : '不合法的档案 ! 档案大小超过预设值 !',
	204 : '您上传的档案已经损毁 !',
	205 : '伺服器上没有预设的暂存目录 !',
	206 : '档案上传程序因为安全因素已被系统自动取消 ! 可能是上传的档案内容包含 HTML 码 !',
	207 : '您上传的档案名称将会自动更改为 "%1".',
	300 : 'Moving file(s) failed.', // MISSING
	301 : 'Copying file(s) failed.', // MISSING
	500 : '因为安全因素 , 档案浏览器已被停用 ! 请联络您的系统管理者并检查 CKFinder 的设定档 config.php !',
	501 : '缩图预览功能已被停用 !'
	},

	// Other Error Messages.
	ErrorMsg :
	{
		FileEmpty		: '档案名称不能空白 !',
		FileExists		: 'File %s already exists.', // MISSING
		FolderEmpty		: '目录名称不能空白 !',
		FolderExists	: 'Folder %s already exists.', // MISSING
		FolderNameExists	: 'Folder already exists.', // MISSING

		FileInvChar		: '档案名称不能包含以下字元： \n\\ / : * ? " < > |',
		FolderInvChar	: '目录名称不能包含以下字元： \n\\ / : * ? " < > |',

		PopupBlockView	: '无法在新视窗开启档案 ! 请检查浏览器的设定并且针对这个网站 关闭 <封锁弹跳视窗> 这个功能 !',
		XmlError		: 'It was not possible to properly load the XML response from the web server.', // MISSING
		XmlEmpty		: 'It was not possible to load the XML response from the web server. The server returned an empty response.', // MISSING
		XmlRawResponse	: 'Raw response from the server: %s' // MISSING
	},

	// Imageresize plugin
	Imageresize :
	{
		dialogTitle		: 'Resize %s', // MISSING
		sizeTooBig		: 'Cannot set image height or width to a value bigger than the original size (%size).', // MISSING
		resizeSuccess	: 'Image resized successfully.', // MISSING
		thumbnailNew	: 'Create a new thumbnail', // MISSING
		thumbnailSmall	: 'Small (%s)', // MISSING
		thumbnailMedium	: 'Medium (%s)', // MISSING
		thumbnailLarge	: 'Large (%s)', // MISSING
		newSize			: 'Set a new size', // MISSING
		width			: 'Width', // MISSING
		height			: 'Height', // MISSING
		invalidHeight	: 'Invalid height.', // MISSING
		invalidWidth	: 'Invalid width.', // MISSING
		invalidName		: 'Invalid file name.', // MISSING
		newImage		: 'Create a new image', // MISSING
		noExtensionChange : 'File extension cannot be changed.', // MISSING
		imageSmall		: 'Source image is too small.', // MISSING
		contextMenuName	: 'Resize', // MISSING
		lockRatio		: 'Lock ratio', // MISSING
		resetSize		: 'Reset size' // MISSING
	},

	// Fileeditor plugin
	Fileeditor :
	{
		save			: 'Save', // MISSING
		fileOpenError	: 'Unable to open file.', // MISSING
		fileSaveSuccess	: 'File saved successfully.', // MISSING
		contextMenuName	: 'Edit', // MISSING
		loadingFile		: 'Loading file, please wait...' // MISSING
	},

	Maximize :
	{
		maximize : 'Maximize', // MISSING
		minimize : 'Minimize' // MISSING
	},

	Gallery :
	{
		current : 'Image {current} of {total}' // MISSING
	},

	Zip :
	{
		extractHereLabel	: 'Extract here', // MISSING
		extractToLabel		: 'Extract to...', // MISSING
		downloadZipLabel	: 'Download as zip', // MISSING
		compressZipLabel	: 'Compress to zip', // MISSING
		removeAndExtract	: 'Remove existing and extract', // MISSING
		extractAndOverwrite	: 'Extract overwriting existing files', // MISSING
		extractSuccess		: 'File extracted successfully.' // MISSING
	}
};
