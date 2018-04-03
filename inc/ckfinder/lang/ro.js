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
 * @fileOverview Defines the {@link CKFinder.lang} object for the Romanian
 *		language.
 */

/**
 * Contains the dictionary of language entries.
 * @namespace
 */
CKFinder.lang['ro'] =
{
	appTitle : 'CKFinder',

	// Common messages and labels.
	common :
	{
		// Put the voice-only part of the label in the span.
		unavailable		: '%1<span class="cke_accessibility">, indisponibil</span>',
		confirmCancel	: 'Unele op?iuni au fost schimbate. E?ti sigur c? vrei s? ?nchizi fereastra de dialog?',
		ok				: 'OK',
		cancel			: 'Anuleaz?',
		confirmationTitle	: 'Confirm?',
		messageTitle	: 'Informa?ii',
		inputTitle		: '?ntreab?',
		undo			: 'Starea anterioar?',
		redo			: 'Starea ulterioar?(redo)',
		skip			: 'S?ri',
		skipAll			: 'S?ri peste toate',
		makeDecision	: 'Ce ac?iune trebuie luat??',
		rememberDecision: 'Re?ine ac?iunea pe viitor'
	},


	// Language direction, 'ltr' or 'rtl'.
	dir : 'ltr',
	HelpLang : 'en',
	LangCode : 'ro',

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
	DateTime : 'dd/mm/yyyy HH:MM',
	DateAmPm : ['AM', 'PM'],

	// Folders
	FoldersTitle	: 'Dosare',
	FolderLoading	: '?nc?rcare...',
	FolderNew		: 'Te rug?m s? introduci numele dosarului nou: ',
	FolderRename	: 'Te rug?m s? introduci numele nou al dosarului: ',
	FolderDelete	: 'E?ti sigur c? vrei s? ?tergi dosarul "%1"?',
	FolderRenaming	: ' (Redenumire...)',
	FolderDeleting	: ' (?tergere...)',
	DestinationFolder	: 'Destination Folder', // MISSING

	// Files
	FileRename		: 'Te rug?m s? introduci numele nou al fi?ierului: ',
	FileRenameExt	: 'E?ti sigur c? vrei s? schimbi extensia fi?ierului? Fi?ierul poate deveni inutilizabil.',
	FileRenaming	: 'Redenumire...',
	FileDelete		: 'E?ti sigur c? vrei s? ?tergi fi?ierul "%1"?',
	FilesDelete	: 'Are you sure you want to delete %1 files?', // MISSING
	FilesLoading	: '?nc?rcare...',
	FilesEmpty		: 'Dosarul este gol.',
	DestinationFile	: 'Destination File', // MISSING
	SkippedFiles	: 'List of skipped files:', // MISSING

	// Basket
	BasketFolder		: 'Co?',
	BasketClear			: 'Gole?te co?',
	BasketRemove		: 'Elimin? din co?',
	BasketOpenFolder	: 'Deschide dosarul p?rinte',
	BasketTruncateConfirm : 'Sigur dore?ti s? elimini toate fi?ierele din co??',
	BasketRemoveConfirm	: 'Sigur dore?ti s? elimini fi?ierul "%1" din co??',
	BasketRemoveConfirmMultiple	: 'Do you really want to remove %1 files from the basket?', // MISSING
	BasketEmpty			: 'Niciun fi?ier ?n co?, trage ?i a?eaz? cu mouse-ul.',
	BasketCopyFilesHere	: 'Copiaz? fi?iere din co?',
	BasketMoveFilesHere	: 'Mut? fi?iere din co?',

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
	Upload		: '?ncarc?',
	UploadTip	: '?ncarc? un fi?ier nou',
	Refresh		: 'Re?mprosp?tare',
	Settings	: 'Set?ri',
	Help		: 'Ajutor',
	HelpTip		: 'Ajutor',

	// Context Menus
	Select			: 'Selecteaz?',
	SelectThumbnail : 'Selecteaz? Thumbnail',
	View			: 'Vizualizeaz?',
	Download		: 'Descarc?',

	NewSubFolder	: 'Subdosar nou',
	Rename			: 'Redenume?te',
	Delete			: '?terge',
	DeleteFiles		: 'Delete Files', // MISSING

	CopyDragDrop	: 'Copiaz? aici',
	MoveDragDrop	: 'Mut? aici',

	// Dialogs
	RenameDlgTitle		: 'Redenume?te',
	NewNameDlgTitle		: 'Nume nou',
	FileExistsDlgTitle	: 'Fi?ierul exist? deja',
	SysErrorDlgTitle : 'Eroare de sistem',

	FileOverwrite	: 'Suprascriere',
	FileAutorename	: 'Auto-redenumire',
	ManuallyRename	: 'Manually rename', // MISSING

	// Generic
	OkBtn		: 'OK',
	CancelBtn	: 'Anuleaz?',
	CloseBtn	: '?nchide',

	// Upload Panel
	UploadTitle			: '?ncarc? un fi?ier nou',
	UploadSelectLbl		: 'Selecteaz? un fi?ier de ?nc?rcat',
	UploadProgressLbl	: '(?nc?rcare ?n progres, te rog a?teapt?...)',
	UploadBtn			: '?ncarc? fi?ierul selectat',
	UploadBtnCancel		: 'Anuleaz?',

	UploadNoFileMsg		: 'Te rug?m s? selectezi un fi?ier din computer.',
	UploadNoFolder		: 'Te rug?m s? selectezi un dosar ?nainte de a ?nc?rca.',
	UploadNoPerms		: '?nc?rcare fi?ier nepermis?.',
	UploadUnknError		: 'Eroare la trimiterea fi?ierului.',
	UploadExtIncorrect	: 'Extensie fi?ier nepermis? ?n acest dosar.',

	// Flash Uploads
	UploadLabel			: 'Fi?iere de ?nc?rcat',
	UploadTotalFiles	: 'Total fi?iere:',
	UploadTotalSize		: 'Total m?rime:',
	UploadSend			: '?ncarc?',
	UploadAddFiles		: 'Adaug? fi?iere',
	UploadClearFiles	: 'Renun?? la toate',
	UploadCancel		: 'Anuleaz? ?nc?rcare',
	UploadRemove		: 'Elimin?',
	UploadRemoveTip		: 'Elimin? !f',
	UploadUploaded		: '?ncarc? !n%',
	UploadProcessing	: 'Prelucrare...',

	// Settings Panel
	SetTitle		: 'Set?ri',
	SetView			: 'Vizualizeaz?:',
	SetViewThumb	: 'Thumbnails',
	SetViewList		: 'List?',
	SetDisplay		: 'Afi?eaz?:',
	SetDisplayName	: 'Nume fi?ier',
	SetDisplayDate	: 'Dat?',
	SetDisplaySize	: 'M?rime fi?ier',
	SetSort			: 'Sortare:',
	SetSortName		: 'dup? nume fi?ier',
	SetSortDate		: 'dup? dat?',
	SetSortSize		: 'dup? m?rime',
	SetSortExtension		: 'dup? extensie',

	// Status Bar
	FilesCountEmpty : '<Dosar Gol>',
	FilesCountOne	: '1 fi?ier',
	FilesCountMany	: '%1 fi?iere',

	// Size and Speed
	Kb				: '%1 KB',
	Mb				: '%1 MB',
	Gb				: '%1 GB',
	SizePerSecond	: '%1/s',

	// Connector Error Messages.
	ErrorUnknown	: 'Nu a fost posibil? finalizarea cererii. (Eroare %1)',
	Errors :
	{
	 10 : 'Comand? invalid?.',
	 11 : 'Tipul de resurs? nu a fost specificat ?n cerere.',
	 12 : 'Tipul de resurs? cerut nu este valid.',
	102 : 'Nume fi?ier sau nume dosar invalid.',
	103 : 'Nu a fost posibili? finalizarea cererii din cauza restric?iilor de autorizare.',
	104 : 'Nu a fost posibili? finalizarea cererii din cauza restric?iilor de permisiune la sistemul de fi?iere.',
	105 : 'Extensie fi?ier invalid?.',
	109 : 'Cerere invalid?.',
	110 : 'Eroare necunoscut?.',
	111 : 'It was not possible to complete the request due to resulting file size.', // MISSING
	115 : 'Exist? deja un fi?ier sau un dosar cu acela?i nume.',
	116 : 'Dosar neg?sit. Te rog ?mprosp?teaz? ?i ?ncearc? din nou.',
	117 : 'Fi?ier neg?sit. Te rog ?mprosp?teaz? lista de fi?iere ?i ?ncearc? din nou.',
	118 : 'Calea sursei ?i a ?intei sunt egale.',
	201 : 'Un fi?ier cu acela?i nume este deja disponibil. Fi?ierul ?nc?rcat a fost redenumit cu "%1".',
	202 : 'Fi?ier invalid.',
	203 : 'Fi?ier invalid. M?rimea fi?ierului este prea mare.',
	204 : 'Fi?ierul ?nc?rcat este corupt.',
	205 : 'Niciun dosar temporar nu este disponibil pentru ?nc?rcarea pe server.',
	206 : '?nc?rcare anulat? din motive de securitate. Fi?ierul con?ine date asem?n?toare cu HTML.',
	207 : 'Fi?ierul ?nc?rcat a fost redenumit cu "%1".',
	300 : 'Mutare fi?ier(e) e?uat?.',
	301 : 'Copiere fi?ier(e) e?uat?.',
	500 : 'Browser-ul de fi?iere este dezactivat din motive de securitate. Te rog contacteaz? administratorul de sistem ?i verific? configurarea de fi?iere CKFinder.',
	501 : 'Func?ionalitatea de creat thumbnails este dezactivat?.'
	},

	// Other Error Messages.
	ErrorMsg :
	{
		FileEmpty		: 'Numele fi?ierului nu poate fi gol.',
		FileExists		: 'Fi?ierul %s exist? deja.',
		FolderEmpty		: 'Numele dosarului nu poate fi gol.',
		FolderExists	: 'Folder %s already exists.', // MISSING
		FolderNameExists	: 'Folder already exists.', // MISSING

		FileInvChar		: 'Numele fi?ierului nu poate con?ine niciunul din urm?toarele caractere: \n\\ / : * ? " < > |',
		FolderInvChar	: 'Numele dosarului nu poate con?ine niciunul din urm?toarele caractere: \n\\ / : * ? " < > |',

		PopupBlockView	: 'Nu a fost posibil? deschiderea fi?ierului ?ntr-o fereastr? nou?. Te rug?m s? configurezi browser-ul ?i s? dezactivezi toate popup-urile blocate pentru acest site.',
		XmlError		: 'Nu a fost posibil? ?nc?rcarea ?n mod corespunz?tor a r?spunsului XML de pe serverul web.',
		XmlEmpty		: 'Nu a fost posibil? ?nc?rcarea r?spunsului XML de pe serverul web. Serverul a returnat un r?spuns gol.',
		XmlRawResponse	: 'R?spuns brut de la server: %s'
	},

	// Imageresize plugin
	Imageresize :
	{
		dialogTitle		: 'Redimensioneaz? %s',
		sizeTooBig		: 'Nu se pot seta ?n?l?imea sau l??imea unei imagini la o valoare mai mare decat dimesiunea original? (%size).',
		resizeSuccess	: 'Imagine redimensionat? cu succes.',
		thumbnailNew	: 'Creaz? un thumbnail nou',
		thumbnailSmall	: 'Mic (%s)',
		thumbnailMedium	: 'Mediu (%s)',
		thumbnailLarge	: 'Mare (%s)',
		newSize			: 'Seteaz? o dimensiune nou?',
		width			: 'L??ime',
		height			: '?n?l?ime',
		invalidHeight	: '?n?l?ime invalid?.',
		invalidWidth	: 'L??ime invalid?.',
		invalidName		: 'Nume fi?ier invalid.',
		newImage		: 'Creeaz? o imagine nou?',
		noExtensionChange : 'Extensia fi?ierului nu poate fi schimbat?.',
		imageSmall		: 'Imaginea surs? este prea mic?.',
		contextMenuName	: 'Redimensioneaz?',
		lockRatio		: 'Blocheaz? raport',
		resetSize		: 'Reseteaz? dimensiunea'
	},

	// Fileeditor plugin
	Fileeditor :
	{
		save			: 'Salveaz?',
		fileOpenError	: 'Fi?ierul nu a putut fi deschis.',
		fileSaveSuccess	: 'Fi?ier salvat cu succes.',
		contextMenuName	: 'Editeaz?',
		loadingFile		: '?nc?rcare fi?ier, te rog a?teapt?...'
	},

	Maximize :
	{
		maximize : 'Maximizare',
		minimize : 'Minimizare'
	},

	Gallery :
	{
		current : 'Imaginea {current} din {total}'
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
