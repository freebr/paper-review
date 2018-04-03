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
 * @fileOverview Defines the {@link CKFinder.lang} object for the Slovak
 *		language.
 */

/**
 * Contains the dictionary of language entries.
 * @namespace
 */
CKFinder.lang['sk'] =
{
	appTitle : 'CKFinder',

	// Common messages and labels.
	common :
	{
		// Put the voice-only part of the label in the span.
		unavailable		: '%1<span class="cke_accessibility">, nedostupné</span>',
		confirmCancel	: 'Niektoré mo?nosti boli zmenené. Naozaj chcete zavrie? okno?',
		ok				: 'OK',
		cancel			: 'Zru?i?',
		confirmationTitle	: 'Potvrdenie',
		messageTitle	: 'Informácia',
		inputTitle		: 'Otázka',
		undo			: 'Sp??',
		redo			: 'Znovu',
		skip			: 'Presko?i?',
		skipAll			: 'Presko?i? v?etko',
		makeDecision	: 'Aky úkon sa má vykona??',
		rememberDecision: 'Pam?ta? si rozhodnutie'
	},


	// Language direction, 'ltr' or 'rtl'.
	dir : 'ltr',
	HelpLang : 'en',
	LangCode : 'sk',

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
	DateAmPm : ['AM', 'PM'],

	// Folders
	FoldersTitle	: 'Adresáre',
	FolderLoading	: 'Nahrávam...',
	FolderNew		: 'Zadajte prosím meno nového adresára: ',
	FolderRename	: 'Zadajte prosím meno nového adresára: ',
	FolderDelete	: 'Skuto?ne zmaza? adresár "%1"?',
	FolderRenaming	: ' (Prebieha premenovanie adresára...)',
	FolderDeleting	: ' (Prebieha zmazanie adresára...)',
	DestinationFolder	: 'Destination Folder', // MISSING

	// Files
	FileRename		: 'Zadajte prosím meno nového súboru: ',
	FileRenameExt	: 'Skuto?ne chcete zmeni? príponu súboru? Upozornenie: zmenou prípony sa súbor m??e sta? nepou?ite?nym, pokia? prípona nie je podporovaná.',
	FileRenaming	: 'Prebieha premenovanie súboru...',
	FileDelete		: 'Skuto?ne chcete odstráni? súbor "%1"?',
	FilesDelete	: 'Are you sure you want to delete %1 files?', // MISSING
	FilesLoading	: 'Nahrávam...',
	FilesEmpty		: 'Adresár je prázdny.',
	DestinationFile	: 'Destination File', // MISSING
	SkippedFiles	: 'List of skipped files:', // MISSING

	// Basket
	BasketFolder		: 'Ko?ík',
	BasketClear			: 'Vyprázdni? ko?ík',
	BasketRemove		: 'Odstráni? z ko?íka',
	BasketOpenFolder	: 'Otvori? nadradeny adresár',
	BasketTruncateConfirm : 'Naozaj chcete odstráni? v?etky súbory z ko?íka?',
	BasketRemoveConfirm	: 'Naozaj chcete odstráni? súbor "%1" z ko?íka?',
	BasketRemoveConfirmMultiple	: 'Do you really want to remove %1 files from the basket?', // MISSING
	BasketEmpty			: 'V ko?íku nie sú ?iadne súbory, potiahnite a vlo?te nejaky.',
	BasketCopyFilesHere	: 'Prekopírova? súbory z ko?íka',
	BasketMoveFilesHere	: 'Presunú? súbory z ko?íka',

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
	Upload		: 'Prekopírova? na server (Upload)',
	UploadTip	: 'Prekopírova? novy súbor',
	Refresh		: 'Znovuna?íta? (Refresh)',
	Settings	: 'Nastavenia',
	Help		: 'Pomoc',
	HelpTip		: 'Pomoc',

	// Context Menus
	Select			: 'Vybra?',
	SelectThumbnail : 'Zvo?te miniatúru',
	View			: 'Náh?ad',
	Download		: 'Stiahnu?',

	NewSubFolder	: 'Novy podadresár',
	Rename			: 'Premenova?',
	Delete			: 'Zmaza?',
	DeleteFiles		: 'Delete Files', // MISSING

	CopyDragDrop	: 'Prekopírova? sem',
	MoveDragDrop	: 'Presunú? sem',

	// Dialogs
	RenameDlgTitle		: 'Premenova?',
	NewNameDlgTitle		: 'Nové meno',
	FileExistsDlgTitle	: 'Súbor u? existuje',
	SysErrorDlgTitle : 'Systémová chyba',

	FileOverwrite	: 'Prepísa?',
	FileAutorename	: 'Auto-premenovanie',
	ManuallyRename	: 'Manually rename', // MISSING

	// Generic
	OkBtn		: 'OK',
	CancelBtn	: 'Zru?i?',
	CloseBtn	: 'Zatvori?',

	// Upload Panel
	UploadTitle			: 'Nahra? novy súbor',
	UploadSelectLbl		: 'Vyberte súbor, ktory chcete prekopírova? na server',
	UploadProgressLbl	: '(Prebieha kopírovanie, ?akajte prosím...)',
	UploadBtn			: 'Prekopírova? vybraty súbor',
	UploadBtnCancel		: 'Zru?i?',

	UploadNoFileMsg		: 'Vyberte prosím súbor na Va?om po?íta?i!',
	UploadNoFolder		: 'Pred náhrávaním zvo?te adresár, prosím',
	UploadNoPerms		: 'Nahratie súboru nie je povolené.',
	UploadUnknError		: 'V priebehu posielania súboru sa vyskytla chyba.',
	UploadExtIncorrect	: 'V tomto adresári nie je povoleny tento formát súboru.',

	// Flash Uploads
	UploadLabel			: 'Súbory k nahratiu',
	UploadTotalFiles	: 'V?etky súbory:',
	UploadTotalSize		: 'Celková ve?kos?:',
	UploadSend			: 'Prekopírova? na server',
	UploadAddFiles		: 'Prida? súbory',
	UploadClearFiles	: 'Vy?isti? súbory',
	UploadCancel		: 'Zru?i? nahratie',
	UploadRemove		: 'Odstráni?',
	UploadRemoveTip		: 'Odstráni? !f',
	UploadUploaded		: 'Nahraté !n%',
	UploadProcessing	: 'Spracováva sa ...',

	// Settings Panel
	SetTitle		: 'Nastavenia',
	SetView			: 'Náh?ad:',
	SetViewThumb	: 'Miniobrázky',
	SetViewList		: 'Zoznam',
	SetDisplay		: 'Zobrazi?:',
	SetDisplayName	: 'Názov súboru',
	SetDisplayDate	: 'Dátum',
	SetDisplaySize	: 'Ve?kos? súboru',
	SetSort			: 'Zoradenie:',
	SetSortName		: 'pod?a názvu súboru',
	SetSortDate		: 'pod?a dátumu',
	SetSortSize		: 'pod?a ve?kosti',
	SetSortExtension		: 'pod?a formátu',

	// Status Bar
	FilesCountEmpty : '<Prázdny adresár>',
	FilesCountOne	: '1 súbor',
	FilesCountMany	: '%1 súborov',

	// Size and Speed
	Kb				: '%1 KB',
	Mb				: '%1 MB',
	Gb				: '%1 GB',
	SizePerSecond	: '%1/s',

	// Connector Error Messages.
	ErrorUnknown	: 'Server nemohol dokon?i? spracovanie po?iadavky. (Chyba %1)',
	Errors :
	{
	 10 : 'Neplatny príkaz.',
	 11 : 'V po?iadavke nebol ?pecifikovany typ súboru.',
	 12 : 'Nepodporovany typ súboru.',
	102 : 'Neplatny názov súboru alebo adresára.',
	103 : 'Nebolo mo?né dokon?i? spracovanie po?iadavky kv?li neposta?ujúcej úrovni oprávnení.',
	104 : 'Nebolo mo?né dokon?i? spracovanie po?iadavky kv?li obmedzeniam v prístupovych právach k súborom.',
	105 : 'Neplatná prípona súboru.',
	109 : 'Neplatná po?iadavka.',
	110 : 'Neidentifikovaná chyba.',
	111 : 'It was not possible to complete the request due to resulting file size.', // MISSING
	115 : 'Zadany súbor alebo adresár u? existuje.',
	116 : 'Adresár nebol nájdeny. Aktualizujte obsah adresára (Znovuna?íta?) a skúste znovu.',
	117 : 'Súbor nebol nájdeny. Aktualizujte obsah adresára (Znovuna?íta?) a skúste znovu.',
	118 : 'Zdrojové a cie?ové cesty sú rovnaké.',
	201 : 'Súbor so zadanym názvom u? existuje. Prekopírovany súbor bol premenovany na "%1".',
	202 : 'Neplatny súbor.',
	203 : 'Neplatny súbor - súbor presahuje maximálnu povolenú ve?kos?.',
	204 : 'Kopírovany súbor je po?kodeny.',
	205 : 'Server nemá ?pecifikovany do?asny adresár pre kopírované súbory.',
	206 : 'Kopírovanie preru?ené kv?li nedostato?nému zabezpe?eniu. Súbor obsahuje HTML data.',
	207 : 'Prekopírovany súbor bol premenovany na "%1".',
	300 : 'Presunutie súborov zlyhalo.',
	301 : 'Kopírovanie súborov zlyhalo.',
	500 : 'Prehliadanie súborov je zakázané kv?li bezpe?nosti. Kontaktujte prosím administrátora a overte nastavenia v konfigura?nom súbore pre CKFinder.',
	501 : 'Momentálne nie je zapnutá podpora pre generáciu miniobrázkov.'
	},

	// Other Error Messages.
	ErrorMsg :
	{
		FileEmpty		: 'Názov súboru nesmie by? prázdne.',
		FileExists		: 'Súbor %s u? existuje.',
		FolderEmpty		: 'Názov adresára nesmie by? prázdny.',
		FolderExists	: 'Folder %s already exists.', // MISSING
		FolderNameExists	: 'Folder already exists.', // MISSING

		FileInvChar		: 'Súbor nesmie obsahova? ?iadny z nasledujúcich znakov: \n\\ / : * ? " < > |',
		FolderInvChar	: 'Adresár nesmie obsahova? ?iadny z nasledujúcich znakov: \n\\ / : * ? " < > |',

		PopupBlockView	: 'Nebolo mo?né otvori? súbor v novom okne. Overte nastavenia Vá?ho prehliada?a a zaká?te v?etky blokova?e popup okien pre túto webstránku.',
		XmlError		: 'Nebolo mo?né korektne na?íta? XML odozvu z web serveu.',
		XmlEmpty		: 'Nebolo mo?né korektne na?íta? XML odozvu z web serveu. Server vrátil prázdnu odpove? (odozvu).',
		XmlRawResponse	: 'Neupravená odpove? zo servera: %s'
	},

	// Imageresize plugin
	Imageresize :
	{
		dialogTitle		: 'Zmeni? ve?kos? %s',
		sizeTooBig		: 'Nie je mo?né nastavi? vy?ku alebo ?írku obrázku na v???ie hodnoty ako originálnu ve?kos? (%size).',
		resizeSuccess	: 'Zmena v?kosti obrázku bola úspe?ne vykonaná.',
		thumbnailNew	: 'Vytvori? novú miniatúru obrázku',
		thumbnailSmall	: 'Maly (%s)',
		thumbnailMedium	: 'Stredny (%s)',
		thumbnailLarge	: 'Ve?ky (%s)',
		newSize			: 'Nastavi? novú ve?kos?',
		width			: '?írka',
		height			: 'Vy?ka',
		invalidHeight	: 'Neplatná vy?ka.',
		invalidWidth	: 'Neplatná ?írka.',
		invalidName		: 'Neplatny názov súboru.',
		newImage		: 'Vytvori? novy obrázok',
		noExtensionChange : 'Nie je mo?né zmeni? formát súboru.',
		imageSmall		: 'Zdrojovy obrázok je ve?mi maly.',
		contextMenuName	: 'Zmeni? ve?kos?',
		lockRatio		: 'Zámok',
		resetSize		: 'P?vodná ve?kos?'
	},

	// Fileeditor plugin
	Fileeditor :
	{
		save			: 'Ulo?i?',
		fileOpenError	: 'Nie je mo?né otvori? súbor.',
		fileSaveSuccess	: 'Súbor bol úspe?ne ulo?eny.',
		contextMenuName	: 'Upravi?',
		loadingFile		: 'Súbor sa nahráva, prosím ?aka?...'
	},

	Maximize :
	{
		maximize : 'Maximalizova?',
		minimize : 'Minimalizova?'
	},

	Gallery :
	{
		current : 'Obrázok {current} z {total}'
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
