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
 * @fileOverview Defines the {@link CKFinder.lang} object, for the Turkish
 *		language.
 *
 *	Turkish translation by Abdullah M CEYLAN a.k.a. Kenan Balamir. Updated.
 */

/**
 * Contains the dictionary of language entries.
 * @namespace
 */
CKFinder.lang['tr'] =
{
	appTitle : 'CKFinder',

	// Common messages and labels.
	common :
	{
		// Put the voice-only part of the label in the span.
		unavailable		: '%1<span class="cke_accessibility"> ??esi, mevcut de?il</span>',
		confirmCancel	: 'Baz? se?enekler de?i?tirildi. Pencereyi kapatmak istiyor musunuz?',
		ok				: 'Tamam',
		cancel			: 'Vazge?',
		confirmationTitle	: 'Onay',
		messageTitle	: 'Bilgi',
		inputTitle		: 'Soru',
		undo			: 'Geri Al',
		redo			: 'Yinele',
		skip			: 'Atla',
		skipAll			: 'Tümünü Atla',
		makeDecision	: 'Hangi i?lem yap?ls?n?',
		rememberDecision: 'Karar?m? hat?rla'
	},


	// Language direction, 'ltr' or 'rtl'.
	dir : 'ltr',
	HelpLang : 'en',
	LangCode : 'tr',

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
	DateAmPm : ['GN', 'GC'],

	// Folders
	FoldersTitle	: 'Klas?rler',
	FolderLoading	: 'Yükleniyor...',
	FolderNew		: 'Lütfen yeni klas?r ad?n? yaz?n: ',
	FolderRename	: 'Lütfen yeni klas?r ad?n? yaz?n: ',
	FolderDelete	: '"%1" klas?rünü silmek istedi?inizden emin misiniz?',
	FolderRenaming	: ' (Yeniden adland?r?l?yor...)',
	FolderDeleting	: ' (Siliniyor...)',
	DestinationFolder	: 'Destination Folder', // MISSING

	// Files
	FileRename		: 'Lütfen yeni dosyan?n ad?n? yaz?n: ',
	FileRenameExt	: 'Dosya uzant?s?n? de?i?tirmek istiyor musunuz? Bu, dosyay? kullan?lamaz hale getirebilir.',
	FileRenaming	: 'Yeniden adland?r?l?yor...',
	FileDelete		: '"%1" dosyas?n? silmek istedi?inizden emin misiniz?',
	FilesDelete	: 'Are you sure you want to delete %1 files?', // MISSING
	FilesLoading	: 'Yükleniyor...',
	FilesEmpty		: 'Klas?r bo?',
	DestinationFile	: 'Destination File', // MISSING
	SkippedFiles	: 'List of skipped files:', // MISSING

	// Basket
	BasketFolder		: 'Sepet',
	BasketClear			: 'Sepeti temizle',
	BasketRemove		: 'Sepetten sil',
	BasketOpenFolder	: 'üst klas?rü a?',
	BasketTruncateConfirm : 'Sepetteki tüm dosyalar? silmek istedi?inizden emin misiniz?',
	BasketRemoveConfirm	: 'Sepetteki %1% dosyas?n? silmek istedi?inizden emin misiniz?',
	BasketRemoveConfirmMultiple	: 'Do you really want to remove %1 files from the basket?', // MISSING
	BasketEmpty			: 'Sepette hi? dosya yok, birka? tane sürükleyip b?rakabilirsiniz',
	BasketCopyFilesHere	: 'Sepetten Dosya Kopyala',
	BasketMoveFilesHere	: 'Sepetten Dosya Ta??',

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
	Upload		: 'Yükle',
	UploadTip	: 'Yeni Dosya Yükle',
	Refresh		: 'Yenile',
	Settings	: 'Ayarlar',
	Help		: 'Yard?m',
	HelpTip		: 'Yard?m',

	// Context Menus
	Select			: 'Se?',
	SelectThumbnail : '?nizleme Olarak Se?',
	View			: 'G?rüntüle',
	Download		: '?ndir',

	NewSubFolder	: 'Yeni Altklas?r',
	Rename			: 'Yeniden Adland?r',
	Delete			: 'Sil',
	DeleteFiles		: 'Delete Files', // MISSING

	CopyDragDrop	: 'Buraya kopyala',
	MoveDragDrop	: 'Buraya ta??',

	// Dialogs
	RenameDlgTitle		: 'Yeniden Adland?r',
	NewNameDlgTitle		: 'Yeni Ad?',
	FileExistsDlgTitle	: 'Dosya zaten var',
	SysErrorDlgTitle : 'Sistem hatas?',

	FileOverwrite	: 'üzerine yaz',
	FileAutorename	: 'Oto-Yeniden Adland?r',
	ManuallyRename	: 'Manually rename', // MISSING

	// Generic
	OkBtn		: 'Tamam',
	CancelBtn	: 'Vazge?',
	CloseBtn	: 'Kapat',

	// Upload Panel
	UploadTitle			: 'Yeni Dosya Yükle',
	UploadSelectLbl		: 'Yüklenecek dosyay? se?in',
	UploadProgressLbl	: '(Yükleniyor, lütfen bekleyin...)',
	UploadBtn			: 'Se?ili Dosyay? Yükle',
	UploadBtnCancel		: 'Vazge?',

	UploadNoFileMsg		: 'Lütfen bilgisayar?n?zdan dosya se?in',
	UploadNoFolder		: 'Lütfen yüklemeden ?nce klas?r se?in.',
	UploadNoPerms		: 'Dosya yüklemeye izin verilmiyor.',
	UploadUnknError		: 'Dosya g?nderme hatas?.',
	UploadExtIncorrect	: 'Bu dosya uzant?s?na, bu klas?rde izin verilmiyor.',

	// Flash Uploads
	UploadLabel			: 'G?nderilecek Dosyalar',
	UploadTotalFiles	: 'Toplam Dosyalar:',
	UploadTotalSize		: 'Toplam Büyüklük:',
	UploadSend			: 'Yükle',
	UploadAddFiles		: 'Dosyalar? Ekle',
	UploadClearFiles	: 'Dosyalar? Temizle',
	UploadCancel		: 'G?ndermeyi ?ptal Et',
	UploadRemove		: 'Sil',
	UploadRemoveTip		: '!f sil',
	UploadUploaded		: '!n% g?nderildi',
	UploadProcessing	: 'G?nderiliyor...',

	// Settings Panel
	SetTitle		: 'Ayarlar',
	SetView			: 'G?rünüm:',
	SetViewThumb	: '?nizlemeler',
	SetViewList		: 'Liste',
	SetDisplay		: 'G?sterim:',
	SetDisplayName	: 'Dosya ad?',
	SetDisplayDate	: 'Tarih',
	SetDisplaySize	: 'Dosya boyutu',
	SetSort			: 'S?ralama:',
	SetSortName		: 'Dosya ad?na g?re',
	SetSortDate		: 'Tarihe g?re',
	SetSortSize		: 'Boyuta g?re',
	SetSortExtension		: 'Uzant?s?na g?re',

	// Status Bar
	FilesCountEmpty : '<Klas?rde Dosya Yok>',
	FilesCountOne	: '1 dosya',
	FilesCountMany	: '%1 dosya',

	// Size and Speed
	Kb				: '%1 KB',
	Mb				: '%1 MB',
	Gb				: '%1 GB',
	SizePerSecond	: '%1/sn',

	// Connector Error Messages.
	ErrorUnknown	: '?ste?inizi yerine getirmek mümkün de?il. (Hata %1)',
	Errors :
	{
	 10 : 'Ge?ersiz komut.',
	 11 : '?stekte kaynak türü belirtilmemi?.',
	 12 : 'Talep edilen kaynak türü ge?ersiz.',
	102 : 'Ge?ersiz dosya ya da klas?r ad?.',
	103 : 'Kimlik do?rulama k?s?tlamalar? nedeni ile talebinizi yerine getiremiyoruz.',
	104 : 'Dosya sistemi k?s?tlamalar? nedeni ile talebinizi yerine getiremiyoruz.',
	105 : 'Ge?ersiz dosya uzant?s?.',
	109 : 'Ge?ersiz istek.',
	110 : 'Bilinmeyen hata.',
	111 : 'It was not possible to complete the request due to resulting file size.', // MISSING
	115 : 'Ayn? isimde bir dosya ya da klas?r zaten var.',
	116 : 'Klas?r bulunamad?. Lütfen yenileyin ve tekrar deneyin.',
	117 : 'Dosya bulunamad?. Lütfen dosya listesini yenileyin ve tekrar deneyin.',
	118 : 'Kaynak ve hedef yol ayn?!',
	201 : 'Ayn? ada sahip bir dosya zaten var. Yüklenen dosyan?n ad? "%1" olarak de?i?tirildi.',
	202 : 'Ge?ersiz dosya',
	203 : 'Ge?ersiz dosya. Dosya boyutu ?ok büyük.',
	204 : 'Yüklenen dosya bozuk.',
	205 : 'Dosyalar? yüklemek i?in gerekli ge?ici klas?r sunucuda bulunamad?.',
	206 : 'Güvenlik nedeni ile yükleme iptal edildi. Dosya HTML benzeri veri i?eriyor.',
	207 : 'Yüklenen dosyan?n ad? "%1" olarak de?i?tirildi.',
	300 : 'Dosya ta??ma i?lemi ba?ar?s?z.',
	301 : 'Dosya kopyalama i?lemi ba?ar?s?z.',
	500 : 'Güvenlik nedeni ile dosya gezgini devred??? b?rak?ld?. Lütfen sistem y?neticiniz ile irtibata ge?in ve CKFinder yap?land?rma dosyas?n? kontrol edin.',
	501 : '?nizleme deste?i devred???.'
	},

	// Other Error Messages.
	ErrorMsg :
	{
		FileEmpty		: 'Dosya ad? bo? olamaz',
		FileExists		: '%s dosyas? zaten var',
		FolderEmpty		: 'Klas?r ad? bo? olamaz',
		FolderExists	: 'Folder %s already exists.', // MISSING
		FolderNameExists	: 'Folder already exists.', // MISSING

		FileInvChar		: 'Dosya ad?n?n i?ermesi mümkün olmayan karakterler: \n\\ / : * ? " < > |',
		FolderInvChar	: 'Klas?r ad?n?n i?ermesi mümkün olmayan karakterler: \n\\ / : * ? " < > |',

		PopupBlockView	: 'Dosyay? yeni pencerede a?mak i?in, taray?c? ayarlar?ndan bu sitenin a??l?r pencerelerine izin vermeniz gerekiyor.',
		XmlError		: 'Web sunucusundan XML yan?t? düzgün bir ?ekilde yüklenemedi.',
		XmlEmpty		: 'Web sunucusundan XML yan?t? düzgün bir ?ekilde yüklenemedi. Sunucudan bo? cevap d?ndü.',
		XmlRawResponse	: 'Sunucudan gelen ham mesaj: %s'
	},

	// Imageresize plugin
	Imageresize :
	{
		dialogTitle		: 'Boyutland?r: %s',
		sizeTooBig		: 'Yükseklik ve geni?lik de?eri orijinal boyuttan büyük oldu?undan, i?lem ger?ekle?tirilemedi (%size).',
		resizeSuccess	: 'Resim ba?ar?yla yeniden boyutland?r?ld?.',
		thumbnailNew	: 'Yeni ?nizleme olu?tur',
		thumbnailSmall	: 'Kü?ük (%s)',
		thumbnailMedium	: 'Orta (%s)',
		thumbnailLarge	: 'Büyük (%s)',
		newSize			: 'Yeni boyutu ayarla',
		width			: 'Geni?lik',
		height			: 'Yükseklik',
		invalidHeight	: 'Ge?ersiz yükseklik.',
		invalidWidth	: 'Ge?ersiz geni?lik.',
		invalidName		: 'Ge?ersiz dosya ad?.',
		newImage		: 'Yeni resim olu?tur',
		noExtensionChange : 'Dosya uzant?s? de?i?tirilemedi.',
		imageSmall		: 'Kaynak resim ?ok kü?ük',
		contextMenuName	: 'Boyutland?r',
		lockRatio		: 'Oran? kilitle',
		resetSize		: 'Büyüklü?ü s?f?rla'
	},

	// Fileeditor plugin
	Fileeditor :
	{
		save			: 'Kaydet',
		fileOpenError	: 'Dosya a??lamad?.',
		fileSaveSuccess	: 'Dosya ba?ar?yla kaydedildi.',
		contextMenuName	: 'Düzenle',
		loadingFile		: 'Dosya yükleniyor, lütfen bekleyin...'
	},

	Maximize :
	{
		maximize : 'Büyült',
		minimize : 'Kü?ült'
	},

	Gallery :
	{
		current : '{current} / {total} resim'
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
