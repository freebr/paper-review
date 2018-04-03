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
 * @fileOverview Defines the {@link CKFinder.lang} object for the Czech
 *		language.
 */

/**
 * Contains the dictionary of language entries.
 * @namespace
 */
CKFinder.lang['cs'] =
{
	appTitle : 'CKFinder',

	// Common messages and labels.
	common :
	{
		// Put the voice-only part of the label in the span.
		unavailable		: '%1<span class="cke_accessibility">, nedostupné</span>',
		confirmCancel	: 'Některá z nastavení byla změněna. Skute?ně chcete dialogové okno zav?ít?',
		ok				: 'OK',
		cancel			: 'Zru?it',
		confirmationTitle	: 'Potvrzení',
		messageTitle	: 'Informace',
		inputTitle		: 'Otázka',
		undo			: 'Zpět',
		redo			: 'Znovu',
		skip			: 'P?esko?it',
		skipAll			: 'P?esko?it v?e',
		makeDecision	: 'Co by se mělo provést?',
		rememberDecision: 'Zapamatovat si mé rozhodnutí'
	},


	// Language direction, 'ltr' or 'rtl'.
	dir : 'ltr',
	HelpLang : 'cs',
	LangCode : 'cs',

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
	DateTime : 'd/m/yyyy H:MM',
	DateAmPm : ['AM', 'PM'],

	// Folders
	FoldersTitle	: 'Slo?ky',
	FolderLoading	: 'Na?ítání...',
	FolderNew		: 'Zadejte název nové slo?ky: ',
	FolderRename	: 'Zadejte novy název slo?ky: ',
	FolderDelete	: 'Opravdu chcete slo?ku "%1" smazat?',
	FolderRenaming	: ' (P?ejmenovávání...)',
	FolderDeleting	: ' (Mazání...)',
	DestinationFolder	: 'Destination Folder', // MISSING

	// Files
	FileRename		: 'Zadejte novy název souboru: ',
	FileRenameExt	: 'Opravdu chcete změnit p?íponu souboru? Soubor se m??e stát nepou?itelnym.',
	FileRenaming	: 'P?ejmenovávání...',
	FileDelete		: 'Opravdu chcete smazat soubor "%1"?',
	FilesDelete	: 'Are you sure you want to delete %1 files?', // MISSING
	FilesLoading	: 'Na?ítání...',
	FilesEmpty		: 'Prázdná slo?ka.',
	DestinationFile	: 'Destination File', // MISSING
	SkippedFiles	: 'List of skipped files:', // MISSING

	// Basket
	BasketFolder		: 'Ko?ík',
	BasketClear			: 'Vy?istit Ko?ík',
	BasketRemove		: 'Odstranit z Ko?íku',
	BasketOpenFolder	: 'Otev?ít nad?azenou slo?ku',
	BasketTruncateConfirm : 'Opravdu chcete z Ko?íku odstranit v?echny soubory?',
	BasketRemoveConfirm	: 'Opravdu chcete odstranit soubor "%1" z Ko?íku?',
	BasketRemoveConfirmMultiple	: 'Do you really want to remove %1 files from the basket?', // MISSING
	BasketEmpty			: 'V Ko?íku nejsou ?ádné soubory, tak sem některé p?etáhněte.',
	BasketCopyFilesHere	: 'Kopírovat soubory z Ko?íku',
	BasketMoveFilesHere	: 'P?esunout soubory z Ko?íku',

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
	Upload		: 'Nahrát',
	UploadTip	: 'Nahrát novy soubor',
	Refresh		: 'Znovu na?íst',
	Settings	: 'Nastavení',
	Help		: 'Nápověda',
	HelpTip		: 'Nápověda',

	// Context Menus
	Select			: 'Vybrat',
	SelectThumbnail : 'Vybrat náhled',
	View			: 'Zobrazit',
	Download		: 'Ulo?it jako',

	NewSubFolder	: 'Nová podslo?ka',
	Rename			: 'P?ejmenovat',
	Delete			: 'Smazat',
	DeleteFiles		: 'Delete Files', // MISSING

	CopyDragDrop	: 'Zkopírovat sem',
	MoveDragDrop	: 'P?esunout sem',

	// Dialogs
	RenameDlgTitle		: 'P?ejmenovat',
	NewNameDlgTitle		: 'Novy název',
	FileExistsDlgTitle	: 'Soubor ji? existuje',
	SysErrorDlgTitle : 'Chyba systému',

	FileOverwrite	: 'P?epsat',
	FileAutorename	: 'Automaticky p?ejmenovat',
	ManuallyRename	: 'Manually rename', // MISSING

	// Generic
	OkBtn		: 'OK',
	CancelBtn	: 'Zru?it',
	CloseBtn	: 'Zav?ít',

	// Upload Panel
	UploadTitle			: 'Nahrát novy soubor',
	UploadSelectLbl		: 'Zvolit soubor k nahrání',
	UploadProgressLbl	: '(Probíhá nahrávání, ?ekejte...)',
	UploadBtn			: 'Nahrát zvoleny soubor',
	UploadBtnCancel		: 'Zru?it',

	UploadNoFileMsg		: 'Vyberte prosím soubor z Va?eho po?íta?e.',
	UploadNoFolder		: 'P?ed nahráváním vyberte slo?ku prosím.',
	UploadNoPerms		: 'Nahrávání soubor? není povoleno.',
	UploadUnknError		: 'Chyba p?i posílání souboru.',
	UploadExtIncorrect	: 'P?ípona souboru není v této slo?ce povolena.',

	// Flash Uploads
	UploadLabel			: 'Soubory k nahrání',
	UploadTotalFiles	: 'Celkem soubor?:',
	UploadTotalSize		: 'Celková velikost:',
	UploadSend			: 'Nahrát',
	UploadAddFiles		: 'P?idat soubory',
	UploadClearFiles	: 'Vy?istit soubory',
	UploadCancel		: 'Zru?it nahrávání',
	UploadRemove		: 'Odstranit',
	UploadRemoveTip		: 'Odstranit !f',
	UploadUploaded		: 'Nahráno !n%',
	UploadProcessing	: 'Zpracovávání...',

	// Settings Panel
	SetTitle		: 'Nastavení',
	SetView			: 'Zobrazení:',
	SetViewThumb	: 'Náhled',
	SetViewList		: 'Seznam',
	SetDisplay		: 'Zobrazit:',
	SetDisplayName	: 'Název',
	SetDisplayDate	: 'Datum',
	SetDisplaySize	: 'Velikost',
	SetSort			: 'Se?azení:',
	SetSortName		: 'Podle názvu',
	SetSortDate		: 'Podle data',
	SetSortSize		: 'Podle velikosti',
	SetSortExtension		: 'Podle p?ípony',

	// Status Bar
	FilesCountEmpty : '<Prázdná slo?ka>',
	FilesCountOne	: '1 soubor',
	FilesCountMany	: '%1 soubor?',

	// Size and Speed
	Kb				: '%1 KB',
	Mb				: '%1 MB',
	Gb				: '%1 GB',
	SizePerSecond	: '%1/s',

	// Connector Error Messages.
	ErrorUnknown	: 'P?íkaz nebylo mo?né dokon?it. (Chyba %1)',
	Errors :
	{
	 10 : 'Neplatny p?íkaz.',
	 11 : 'Typ zdroje nebyl v po?adavku ur?en.',
	 12 : 'Po?adovany typ zdroje není platny.',
	102 : '?patné název souboru, nebo slo?ky.',
	103 : 'Nebylo mo?né p?íkaz dokon?it kv?li omezení oprávnění.',
	104 : 'Nebylo mo?né p?íkaz dokon?it kv?li omezení oprávnění souborového systému.',
	105 : 'Neplatná p?ípona souboru.',
	109 : 'Neplatny po?adavek.',
	110 : 'Neznámá chyba.',
	111 : 'It was not possible to complete the request due to resulting file size.', // MISSING
	115 : 'Soubor nebo slo?ka se stejnym názvem ji? existuje.',
	116 : 'Slo?ka nenalezena, prosím obnovte a zkuste znovu.',
	117 : 'Soubor nenalezen, prosím obnovte seznam soubor? a zkuste znovu.',
	118 : 'Cesty zdroje a cíle jsou stejné.',
	201 : 'Soubor se stejnym názvem je ji? dostupny, nahrany soubor byl p?ejmenován na "%1".',
	202 : 'Neplatny soubor.',
	203 : 'Neplatny soubor. Velikost souboru je p?íli? velká.',
	204 : 'Nahrany soubor je po?kozen.',
	205 : 'Na serveru není dostupná do?asná slo?ka pro nahrávání.',
	206 : 'Nahrávání zru?eno z bezpe?nostních d?vod?. Soubor obsahuje data podobná HTML.',
	207 : 'Nahrany soubor byl p?ejmenován na "%1".',
	300 : 'P?esunování souboru(?) selhalo.',
	301 : 'Kopírování souboru(?) selhalo.',
	500 : 'Pr?zkumník soubor? je z bezpe?nostních d?vod? zakázán. Zdělte to prosím správci systému a zkontrolujte soubor nastavení CKFinder.',
	501 : 'Podpora náhled? je zakázána.'
	},

	// Other Error Messages.
	ErrorMsg :
	{
		FileEmpty		: 'Název souboru nem??e byt prázdny.',
		FileExists		: 'Soubor %s ji? existuje.',
		FolderEmpty		: 'Název slo?ky nem??e byt prázdny.',
		FolderExists	: 'Folder %s already exists.', // MISSING
		FolderNameExists	: 'Folder already exists.', // MISSING

		FileInvChar		: 'Název souboru nesmí obsahovat následující znaky: \n\\ / : * ? " < > |',
		FolderInvChar	: 'Název slo?ky nesmí obsahovat následující znaky: \n\\ / : * ? " < > |',

		PopupBlockView	: 'Soubor nebylo mo?né otev?ít do nového okna. Prosím nastavte si Vá? prohlí?e? a zaka?te ve?keré blokování vyskakovacích oken.',
		XmlError		: 'Nebylo mo?né správně na?íst XML odpově? z internetového serveru.',
		XmlEmpty		: 'Nebylo mo?né na?íst XML odpově? z internetového serveru. Server vrátil prázdnou odpově?.',
		XmlRawResponse	: '?istá odpově? od serveru: %s'
	},

	// Imageresize plugin
	Imageresize :
	{
		dialogTitle		: 'Změnit velikost %s',
		sizeTooBig		: 'Nelze nastavit ?í?ku ?i vy?ku obrázku na hodnotu vy??í ne? p?vodní velikost (%size).',
		resizeSuccess	: 'úspě?ně změněna velikost obrázku.',
		thumbnailNew	: 'Vytvo?it novy náhled',
		thumbnailSmall	: 'Maly (%s)',
		thumbnailMedium	: 'St?ední (%s)',
		thumbnailLarge	: 'Velky (%s)',
		newSize			: 'Nastavit novou velikost',
		width			: '?í?ka',
		height			: 'Vy?ka',
		invalidHeight	: 'Neplatná vy?ka.',
		invalidWidth	: 'Neplatná ?í?ka.',
		invalidName		: 'Neplatny název souboru.',
		newImage		: 'Vytvo?it novy obrázek',
		noExtensionChange : 'P?íponu souboru nelze změnit.',
		imageSmall		: 'Zdrojovy obrázek je p?íli? maly.',
		contextMenuName	: 'Změnit velikost',
		lockRatio		: 'Uzamknout poměr',
		resetSize		: 'P?vodní velikost'
	},

	// Fileeditor plugin
	Fileeditor :
	{
		save			: 'Ulo?it',
		fileOpenError	: 'Soubor nelze otev?ít.',
		fileSaveSuccess	: 'Soubor úspě?ně ulo?en.',
		contextMenuName	: 'Upravit',
		loadingFile		: 'Na?ítání souboru, ?ekejte prosím...'
	},

	Maximize :
	{
		maximize : 'Maximalizovat',
		minimize : 'Minimalizovat'
	},

	Gallery :
	{
		current : 'Obrázek {current} z {total}'
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
