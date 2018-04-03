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
 * @fileOverview Defines the {@link CKFinder.lang} object for the Lithuanian
 *		language.
 */

/**
 * Contains the dictionary of language entries.
 * @namespace
 */
CKFinder.lang['lt'] =
{
	appTitle : 'CKFinder',

	// Common messages and labels.
	common :
	{
		// Put the voice-only part of the label in the span.
		unavailable		: '%1<span class="cke_accessibility">, n?ra</span>',
		confirmCancel	: 'Kai kurie nustatymai buvo pakeisti. Ar tikrai norite u?daryti ?? lang??',
		ok				: 'Gerai',
		cancel			: 'At?aukti',
		confirmationTitle	: 'Patvirtinimas',
		messageTitle	: 'Informacija',
		inputTitle		: 'Klausimas',
		undo			: 'Veiksmas atgal',
		redo			: 'Veiksmas pirmyn',
		skip			: 'Praleisti',
		skipAll			: 'Praleisti visk?',
		makeDecision	: 'K? pasirinksite?',
		rememberDecision: 'Atsiminti mano pasirinkim?'
	},


	// Language direction, 'ltr' or 'rtl'.
	dir : 'ltr',
	HelpLang : 'lt',
	LangCode : 'lt',

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
	DateTime : 'yyyy.mm.dd H:MM',
	DateAmPm : ['AM', 'PM'],

	// Folders
	FoldersTitle	: 'Segtuvai',
	FolderLoading	: 'Pra?au palaukite...',
	FolderNew		: 'Pra?au ?ra?ykite naujo segtuvo pavadinim?: ',
	FolderRename	: 'Pra?au ?ra?ykite naujo segtuvo pavadinim?: ',
	FolderDelete	: 'Ar tikrai norite i?trinti "%1" segtuv??',
	FolderRenaming	: ' (Pervadinama...)',
	FolderDeleting	: ' (Trinama...)',
	DestinationFolder	: 'Destination Folder', // MISSING

	// Files
	FileRename		: 'Pra?au ?ra?ykite naujo failo pavadinim?: ',
	FileRenameExt	: 'Ar tikrai norite pakeisti ?io failo pl?tin?? Failas gali būti nebepanaudojamas',
	FileRenaming	: 'Pervadinama...',
	FileDelete		: 'Ar tikrai norite i?trinti fail? "%1"?',
	FilesDelete	: 'Are you sure you want to delete %1 files?', // MISSING
	FilesLoading	: 'Pra?au palaukite...',
	FilesEmpty		: 'Tu??ias segtuvas',
	DestinationFile	: 'Destination File', // MISSING
	SkippedFiles	: 'List of skipped files:', // MISSING

	// Basket
	BasketFolder		: 'Krep?elis',
	BasketClear			: 'I?tu?tinti krep?el?',
	BasketRemove		: 'I?trinti krep?el?',
	BasketOpenFolder	: 'Atidaryti failo segtuv?',
	BasketTruncateConfirm : 'Ar tikrai norite i?trinti visus failus i? krep?elio?',
	BasketRemoveConfirm	: 'Ar tikrai norite i?trinti fail? "%1" i? krep?elio?',
	BasketRemoveConfirmMultiple	: 'Do you really want to remove %1 files from the basket?', // MISSING
	BasketEmpty			: 'Krep?elyje fail? n?ra, nuvilkite ir ?meskite juos ? krep?el?.',
	BasketCopyFilesHere	: 'Kopijuoti failus i? krep?elio',
	BasketMoveFilesHere	: 'Perkelti failus i? krep?elio',

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
	Upload		: '?kelti',
	UploadTip	: '?kelti nauj? fail?',
	Refresh		: 'Atnaujinti',
	Settings	: 'Nustatymai',
	Help		: 'Pagalba',
	HelpTip		: 'Patarimai',

	// Context Menus
	Select			: 'Pasirinkti',
	SelectThumbnail : 'Pasirinkti miniatiūr?',
	View			: 'Per?iūr?ti',
	Download		: 'Atsisi?sti',

	NewSubFolder	: 'Naujas segtuvas',
	Rename			: 'Pervadinti',
	Delete			: 'I?trinti',
	DeleteFiles		: 'Delete Files', // MISSING

	CopyDragDrop	: 'Nukopijuoti ?ia',
	MoveDragDrop	: 'Perkelti ?ia',

	// Dialogs
	RenameDlgTitle		: 'Pervadinti',
	NewNameDlgTitle		: 'Naujas pavadinimas',
	FileExistsDlgTitle	: 'Toks failas jau egzistuoja',
	SysErrorDlgTitle : 'Sistemos klaida',

	FileOverwrite	: 'U?ra?yti ant vir?aus',
	FileAutorename	: 'Automati?kai pervadinti',
	ManuallyRename	: 'Manually rename', // MISSING

	// Generic
	OkBtn		: 'Gerai',
	CancelBtn	: 'At?aukti',
	CloseBtn	: 'U?daryti',

	// Upload Panel
	UploadTitle			: '?kelti nauj? fail?',
	UploadSelectLbl		: 'Pasirinkite fail? ?k?limui',
	UploadProgressLbl	: '(Vykdomas ?k?limas, pra?au palaukite...)',
	UploadBtn			: '?kelti pasirinkt? fail?',
	UploadBtnCancel		: 'At?aukti',

	UploadNoFileMsg		: 'Pasirinkite fail? i? savo kompiuterio',
	UploadNoFolder		: 'Pasirinkite segtuv? prie? ?keliant.',
	UploadNoPerms		: 'Fail? ?k?limas u?draustas.',
	UploadUnknError		: '?vyko klaida siun?iant fail?.',
	UploadExtIncorrect	: '?iame segtuve toks fail? pl?tinys yra u?draustas.',

	// Flash Uploads
	UploadLabel			: '?keliami failai',
	UploadTotalFiles	: 'I? viso fail?:',
	UploadTotalSize		: 'Visa apimtis:',
	UploadSend			: '?kelti',
	UploadAddFiles		: 'Prid?ti failus',
	UploadClearFiles	: 'I?valyti failus',
	UploadCancel		: 'At?aukti nusiuntim?',
	UploadRemove		: 'Pa?alinti',
	UploadRemoveTip		: 'Pa?alinti !f',
	UploadUploaded		: '?keltas !n%',
	UploadProcessing	: 'Apdorojama...',

	// Settings Panel
	SetTitle		: 'Nustatymai',
	SetView			: 'Per?iūr?ti:',
	SetViewThumb	: 'Miniatiūros',
	SetViewList		: 'S?ra?as',
	SetDisplay		: 'Rodymas:',
	SetDisplayName	: 'Failo pavadinimas',
	SetDisplayDate	: 'Data',
	SetDisplaySize	: 'Failo dydis',
	SetSort			: 'Rū?iavimas:',
	SetSortName		: 'pagal failo pavadinim?',
	SetSortDate		: 'pagal dat?',
	SetSortSize		: 'pagal apimt?',
	SetSortExtension		: 'pagal pl?tin?',

	// Status Bar
	FilesCountEmpty : '<Tu??ias segtuvas>',
	FilesCountOne	: '1 failas',
	FilesCountMany	: '%1 failai',

	// Size and Speed
	Kb				: '%1 KB',
	Mb				: '%1 MB',
	Gb				: '%1 GB',
	SizePerSecond	: '%1/s',

	// Connector Error Messages.
	ErrorUnknown	: 'U?klausos ?vykdyti nepavyko. (Klaida %1)',
	Errors :
	{
	 10 : 'Neteisinga komanda.',
	 11 : 'Resurso rū?is nenurodyta u?klausoje.',
	 12 : 'Neteisinga resurso rū?is.',
	102 : 'Netinkamas failas arba segtuvo pavadinimas.',
	103 : 'Nepavyko ?vykdyti u?klausos d?l autorizavimo apribojim?.',
	104 : 'Nepavyko ?vykdyti u?klausos d?l fail? sistemos leidim? apribojim?.',
	105 : 'Netinkamas failo pl?tinys.',
	109 : 'Netinkama u?klausa.',
	110 : 'Ne?inoma klaida.',
	111 : 'It was not possible to complete the request due to resulting file size.', // MISSING
	115 : 'Failas arba segtuvas su tuo pa?iu pavadinimu jau yra.',
	116 : 'Segtuvas nerastas. Pabandykite atnaujinti.',
	117 : 'Failas nerastas. Pabandykite atnaujinti fail? s?ra??.',
	118 : '?altinio ir nurodomos vietos nuorodos yra vienodos.',
	201 : 'Failas su tuo pa?iu pavadinimu jau tra. ?keltas failas buvo pervadintas ? "%1"',
	202 : 'Netinkamas failas',
	203 : 'Netinkamas failas. Failo apimtis yra per didel?.',
	204 : '?keltas failas yra pa?eistas.',
	205 : 'N?ra laikinojo segtuvo skirto failams ?kelti.',
	206 : '?k?limas bus nutrauktas d?l saugumo sumetim?. ?iame faile yra HTML duomenys.',
	207 : '?keltas failas buvo pervadintas ? "%1"',
	300 : 'Fail? perk?limas nepavyko.',
	301 : 'Fail? kopijavimas nepavyko.',
	500 : 'Fail? nar?ykl? yra i?jungta d?l saugumo nustaym?. Pra?au susisiekti su sistem? administratoriumi ir patikrinkite CKFinder konfigūracin? fail?.',
	501 : 'Miniatiūr? palaikymas i?jungtas.'
	},

	// Other Error Messages.
	ErrorMsg :
	{
		FileEmpty		: 'Failo pavadinimas negali būti tu??ias',
		FileExists		: 'Failas %s jau egzistuoja',
		FolderEmpty		: 'Segtuvo pavadinimas negali būti tu??ias',
		FolderExists	: 'Folder %s already exists.', // MISSING
		FolderNameExists	: 'Folder already exists.', // MISSING

		FileInvChar		: 'Failo pavadinimas negali tur?ti bent vieno i? ?i? simboli?: \n\\ / : * ? " < > |',
		FolderInvChar	: 'Segtuvo pavadinimas negali tur?ti bent vieno i? ?i? simboli?: \n\\ / : * ? " < > |',

		PopupBlockView	: 'Nepavyko atidaryti failo naujame lange. Pra?au pakeiskite savo nar?ykl?s nustatymus, kad būt? leid?iami i?kylantys langai ?iame tinklapyje.',
		XmlError		: 'Nepavyko ?krauti XML atsako i? web serverio.',
		XmlEmpty		: 'Nepavyko ?krauti XML atsako i? web serverio. Serveris gra?ino tu??i? u?klaus?.',
		XmlRawResponse	: 'Vientisas atsakas i? serverio: %s'
	},

	// Imageresize plugin
	Imageresize :
	{
		dialogTitle		: 'Keisti matmenis %s',
		sizeTooBig		: 'Negalima nustatyti auk??io ir plo?io ? didesnius nei originalaus paveiksliuko (%size).',
		resizeSuccess	: 'Paveiksliuko matmenys pakeisti.',
		thumbnailNew	: 'Sukurti nauj? miniatiūr?',
		thumbnailSmall	: 'Ma?as (%s)',
		thumbnailMedium	: 'Vidutinis (%s)',
		thumbnailLarge	: 'Didelis (%s)',
		newSize			: 'Nustatyti naujus matmenis',
		width			: 'Plotis',
		height			: 'Auk?tis',
		invalidHeight	: 'Neteisingas auk?tis.',
		invalidWidth	: 'Neteisingas plotis.',
		invalidName		: 'Neteisingas pavadinimas.',
		newImage		: 'Sukurti nauj? paveiksliuk?',
		noExtensionChange : 'Failo pl?tinys negali būti pakeistas.',
		imageSmall		: '?altinio paveiksliukas yra per ma?as',
		contextMenuName	: 'Pakeisti matmenis',
		lockRatio		: 'I?laikyti matmen? santyk?',
		resetSize		: 'Nustatyti dyd? i? naujo'
	},

	// Fileeditor plugin
	Fileeditor :
	{
		save			: 'I?saugoti',
		fileOpenError	: 'Nepavyko atidaryti failo.',
		fileSaveSuccess	: 'Failas s?kmingai i?saugotas.',
		contextMenuName	: 'Redaguoti',
		loadingFile		: '?kraunamas failas, pra?au palaukite...'
	},

	Maximize :
	{
		maximize : 'Padidinti',
		minimize : 'Suma?inti'
	},

	Gallery :
	{
		current : 'Nuotrauka {current} i? {total}'
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
