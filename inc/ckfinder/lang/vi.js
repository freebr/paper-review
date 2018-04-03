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
 * @fileOverview Defines the {@link CKFinder.lang} object for the Vietnamese
 *		language.
 */

/**
 * Contains the dictionary of language entries.
 * @namespace
 */
CKFinder.lang['vi'] =
{
	appTitle : 'CKFinder',

	// Common messages and labels.
	common :
	{
		// Put the voice-only part of the label in the span.
		unavailable		: '%1<span class="cke_accessibility">, kh?ng kh? d?ng</span>',
		confirmCancel	: 'Vài tùy ch?n ?? thay ??i. B?n có mu?n ?óng h?p tho?i?',
		ok				: 'OK',
		cancel			: 'H?y',
		confirmationTitle	: 'Xác nh?n',
		messageTitle	: 'Th?ng tin',
		inputTitle		: 'Cau h?i',
		undo			: 'Hoàn tác',
		redo			: 'Làm l?i',
		skip			: 'B? qua',
		skipAll			: 'B? qua t?t c?',
		makeDecision	: 'Ch?n hành ??ng nào?',
		rememberDecision: 'Ghi nh? quy?t ??nh này'
	},


	// Language direction, 'ltr' or 'rtl'.
	dir : 'ltr',
	HelpLang : 'en',
	LangCode : 'vi',

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
	DateTime : 'd/m/yyyy h:MM aa',
	DateAmPm : ['SA', 'CH'],

	// Folders
	FoldersTitle	: 'Th? m?c',
	FolderLoading	: '?ang t?i...',
	FolderNew		: 'Xin ch?n tên cho th? m?c m?i: ',
	FolderRename	: 'Xin ch?n tên m?i cho th? m?c: ',
	FolderDelete	: 'B?n có ch?c mu?n xóa th? m?c "%1"?',
	FolderRenaming	: ' (?ang ??i tên...)',
	FolderDeleting	: ' (?ang xóa...)',
	DestinationFolder	: 'Destination Folder', // MISSING

	// Files
	FileRename		: 'Xin nh?p tên t?p tin m?i: ',
	FileRenameExt	: 'B?n có ch?c mu?n ??i ph?n m? r?ng? T?p tin có th? s? kh?ng dùng ???c.',
	FileRenaming	: '?ang ??i tên...',
	FileDelete		: 'B?n có ch?c mu?n xóa t?p tin "%1"?',
	FilesDelete	: 'Are you sure you want to delete %1 files?', // MISSING
	FilesLoading	: '?ang t?i...',
	FilesEmpty		: 'Th? m?c tr?ng.',
	DestinationFile	: 'Destination File', // MISSING
	SkippedFiles	: 'List of skipped files:', // MISSING

	// Basket
	BasketFolder		: 'R?',
	BasketClear			: 'D?n r?',
	BasketRemove		: 'Xóa kh?i r?',
	BasketOpenFolder	: 'M? th? m?c cha',
	BasketTruncateConfirm : 'B?n có ch?c mu?n b? t?t c? t?p tin trong r??',
	BasketRemoveConfirm	: 'B?n có ch?c mu?n b? t?p tin "%1" kh?i r??',
	BasketRemoveConfirmMultiple	: 'Do you really want to remove %1 files from the basket?', // MISSING
	BasketEmpty			: 'Kh?ng có t?p tin trong r?, h?y kéo và th? t?p tin vào r?.',
	BasketCopyFilesHere	: 'Chép t?p tin t? r?',
	BasketMoveFilesHere	: 'Chuy?n t?p tin t? r?',

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
	Upload		: 'T?i lên',
	UploadTip	: 'T?i t?p tin m?i',
	Refresh		: 'Làm t??i',
	Settings	: 'Thi?t l?p',
	Help		: 'H??ng d?n',
	HelpTip		: 'H??ng d?n',

	// Context Menus
	Select			: 'Ch?n',
	SelectThumbnail : 'Ch?n ?nh m?u',
	View			: 'Xem',
	Download		: 'T?i v?',

	NewSubFolder	: 'T?o th? m?c con',
	Rename			: '??i tên',
	Delete			: 'Xóa',
	DeleteFiles		: 'Delete Files', // MISSING

	CopyDragDrop	: 'Sao chép ? ?ay',
	MoveDragDrop	: 'Di chuy?n ? ?ay',

	// Dialogs
	RenameDlgTitle		: '??i tên',
	NewNameDlgTitle		: 'Tên m?i',
	FileExistsDlgTitle	: 'T?p tin ?? t?n t?i',
	SysErrorDlgTitle : 'L?i h? th?ng',

	FileOverwrite	: 'Ghi ?è',
	FileAutorename	: 'T? ??i tên',
	ManuallyRename	: 'Manually rename', // MISSING

	// Generic
	OkBtn		: 'OK',
	CancelBtn	: 'H?y b?',
	CloseBtn	: '?óng',

	// Upload Panel
	UploadTitle			: 'T?i t?p tin m?i',
	UploadSelectLbl		: 'Ch?n t?p tin t?i lên',
	UploadProgressLbl	: '(?ang t?i lên, vui lòng ch?...)',
	UploadBtn			: 'T?i t?p tin ?? ch?n',
	UploadBtnCancel		: 'H?y b?',

	UploadNoFileMsg		: 'Xin ch?n m?t t?p tin trong máy tính.',
	UploadNoFolder		: 'Xin ch?n th? m?c tr??c khi t?i lên.',
	UploadNoPerms		: 'Kh?ng ???c phép t?i lên.',
	UploadUnknError		: 'L?i khi t?i t?p tin.',
	UploadExtIncorrect	: 'Ki?u t?p tin kh?ng ???c ch?p nh?n trong th? m?c này.',

	// Flash Uploads
	UploadLabel			: 'T?p tin s? t?i:',
	UploadTotalFiles	: 'T?ng s? t?p tin:',
	UploadTotalSize		: 'Dung l??ng t?ng c?ng:',
	UploadSend			: 'T?i lên',
	UploadAddFiles		: 'Thêm t?p tin',
	UploadClearFiles	: 'Xóa t?p tin',
	UploadCancel		: 'H?y t?i',
	UploadRemove		: 'Xóa',
	UploadRemoveTip		: 'Xóa !f',
	UploadUploaded		: '?? t?i !n%',
	UploadProcessing	: '?ang x? lí...',

	// Settings Panel
	SetTitle		: 'Thi?t l?p',
	SetView			: 'Xem:',
	SetViewThumb	: '?nh m?u',
	SetViewList		: 'Danh sách',
	SetDisplay		: 'Hi?n th?:',
	SetDisplayName	: 'Tên t?p tin',
	SetDisplayDate	: 'Ngày',
	SetDisplaySize	: 'Dung l??ng',
	SetSort			: 'S?p x?p:',
	SetSortName		: 'theo tên',
	SetSortDate		: 'theo ngày',
	SetSortSize		: 'theo dung l??ng',
	SetSortExtension		: 'theo ph?n m? r?ng',

	// Status Bar
	FilesCountEmpty : '<Th? m?c r?ng>',
	FilesCountOne	: '1 t?p tin',
	FilesCountMany	: '%1 t?p tin',

	// Size and Speed
	Kb				: '%1 KB',
	Mb				: '%1 MB',
	Gb				: '%1 GB',
	SizePerSecond	: '%1/s',

	// Connector Error Messages.
	ErrorUnknown	: 'Kh?ng th? hoàn t?t yêu c?u. (L?i %1)',
	Errors :
	{
	 10 : 'L?nh kh?ng h?p l?.',
	 11 : 'Ki?u tài nguyên kh?ng ???c ch? ??nh trong yêu c?u.',
	 12 : 'Ki?u tài nguyên yêu c?u kh?ng h?p l?.',
	102 : 'Tên t?p tin hay th? m?c kh?ng h?p l?.',
	103 : 'Kh?ng th? hoàn t?t yêu c?u vì gi?i h?n quy?n.',
	104 : 'Kh?ng th? hoàn t?t yêu c?u vì gi?i h?n quy?n c?a h? th?ng t?p tin.',
	105 : 'Ph?n m? r?ng t?p tin kh?ng h?p l?.',
	109 : 'Yêu c?u kh?ng h?p l?.',
	110 : 'L?i kh?ng xác ??nh.',
	111 : 'It was not possible to complete the request due to resulting file size.', // MISSING
	115 : 'T?p tin ho?c th? m?c cùng tên ?? t?n t?i.',
	116 : 'Kh?ng th?y th? m?c. H?y làm t??i và th? l?i.',
	117 : 'Kh?ng th?y t?p tin. H?y làm t??i và th? l?i.',
	118 : '???ng d?n ngu?n và ?ích gi?ng nhau.',
	201 : 'T?p tin cùng tên ?? t?n t?i. T?p tin v?a t?i lên ???c ??i tên thành "%1".',
	202 : 'T?p tin kh?ng h?p l?.',
	203 : 'T?p tin kh?ng h?p l?. Dung l??ng quá l?n.',
	204 : 'T?p tin t?i lên b? h?ng.',
	205 : 'Kh?ng có th? m?c t?m ?? t?i t?p tin.',
	206 : 'Hu? t?i lên vì lí do b?o m?t. T?p tin ch?a d? li?u gi?ng HTML.',
	207 : 'T?p tin ???c ??i tên thành "%1".',
	300 : 'Di chuy?n t?p tin th?t b?i.',
	301 : 'Chép t?p tin th?t b?i.',
	500 : 'Trình duy?t t?p tin b? v? hi?u vì lí do b?o m?t. Xin liên h? qu?n tr? h? th?ng và ki?m tra t?p tin c?u hình CKFinder.',
	501 : 'Ch?c n?ng h? tr? ?nh m?u b? v? hi?u.'
	},

	// Other Error Messages.
	ErrorMsg :
	{
		FileEmpty		: 'Kh?ng th? ?? tr?ng tên t?p tin.',
		FileExists		: 'T?p tin %s ?? t?n t?i.',
		FolderEmpty		: 'Kh?ng th? ?? tr?ng tên th? m?c.',
		FolderExists	: 'Folder %s already exists.', // MISSING
		FolderNameExists	: 'Folder already exists.', // MISSING

		FileInvChar		: 'Tên t?p tin kh?ng th? ch?a các kí t?: \n\\ / : * ? " < > |',
		FolderInvChar	: 'Tên th? m?c kh?ng th? ch?a các kí t?: \n\\ / : * ? " < > |',

		PopupBlockView	: 'Kh?ng th? m? t?p tin trong c?a s? m?i. H?y ki?m tra trình duy?t và t?t ch?c n?ng ch?n popup trên trang web này.',
		XmlError		: 'Kh?ng th? n?p h?i ?áp XML t? máy ch? web.',
		XmlEmpty		: 'Kh?ng th? n?p h?i ?áp XML t? máy ch? web. D? li?u r?ng.',
		XmlRawResponse	: 'H?i ?áp th? t? máy ch?: %s'
	},

	// Imageresize plugin
	Imageresize :
	{
		dialogTitle		: '??i kích th??c %s',
		sizeTooBig		: 'Kh?ng th? ??t chi?u cao ho?c r?ng to h?n kích th??c g?c (%size).',
		resizeSuccess	: '??i kích th??c ?nh thành c?ng.',
		thumbnailNew	: 'T?o ?nh m?u m?i',
		thumbnailSmall	: 'Nh? (%s)',
		thumbnailMedium	: 'V?a (%s)',
		thumbnailLarge	: 'L?n (%s)',
		newSize			: 'Ch?n kích th??c m?i',
		width			: 'R?ng',
		height			: 'Cao',
		invalidHeight	: 'Chi?u cao kh?ng h?p l?.',
		invalidWidth	: 'Chi?u r?ng kh?ng h?p l?.',
		invalidName		: 'Tên t?p tin kh?ng h?p l?.',
		newImage		: 'T?o ?nh m?i',
		noExtensionChange : 'Kh?ng th? thay ??i ph?n m? r?ng.',
		imageSmall		: '?nh ngu?n quá nh?.',
		contextMenuName	: '??i kích th??c',
		lockRatio		: 'Khoá t? l?',
		resetSize		: '??t l?i kích th??c'
	},

	// Fileeditor plugin
	Fileeditor :
	{
		save			: 'L?u',
		fileOpenError	: 'Kh?ng th? m? t?p tin.',
		fileSaveSuccess	: 'L?u t?p tin thành c?ng.',
		contextMenuName	: 'S?a',
		loadingFile		: '?ang t?i t?p tin, xin ch?...'
	},

	Maximize :
	{
		maximize : 'C?c ??i hóa',
		minimize : 'C?c ti?u hóa'
	},

	Gallery :
	{
		current : 'Hình th? {current} trên {total}'
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
