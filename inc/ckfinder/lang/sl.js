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
* @fileOverview Defines the {@link CKFinder.lang} object for the Slovenian
*		language.
*/

/**
 * Contains the dictionary of language entries.
 * @namespace
 */
CKFinder.lang['sl'] =
{
	appTitle : 'CKFinder',

	// Common messages and labels.
	common :
	{
		// Put the voice-only part of the label in the span.
		unavailable		: '%1<span class="cke_accessibility">, nedostopen</span>',
		confirmCancel	: 'Nekatere opcije so bile spremenjene. Ali res ?elite zapreti pogovorno okno?',
		ok				: 'Potrdi',
		cancel			: 'Prekli?i',
		confirmationTitle	: 'Potrditev',
		messageTitle	: 'Informacija',
		inputTitle		: 'Vpra?anje',
		undo			: 'Razveljavi',
		redo			: 'Obnovi',
		skip			: 'Presko?i',
		skipAll			: 'Presko?i vse',
		makeDecision	: 'Katera aktivnost naj se izvede?',
		rememberDecision: 'Zapomni si mojo izbiro'
	},


	// Language direction, 'ltr' or 'rtl'.
	dir : 'ltr',
	HelpLang : 'en',
	LangCode : 'sl',

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
	DateTime : 'd.m.yyyy H:MM',
	DateAmPm : ['AM', 'PM'],

	// Folders
	FoldersTitle	: 'Mape',
	FolderLoading	: 'Nalagam...',
	FolderNew		: 'Vnesite ime za novo mapo: ',
	FolderRename	: 'Vnesite ime nove mape: ',
	FolderDelete	: 'Ali ste prepri?ani, da ?elite zbrisati mapo "%1"?',
	FolderRenaming	: ' (Preimenujem...)',
	FolderDeleting	: ' (Bri?em...)',
	DestinationFolder	: 'Destination Folder', // MISSING

	// Files
	FileRename		: 'Vnesite novo ime datoteke: ',
	FileRenameExt	: 'Ali ste prepri?ani, da ?elite spremeniti kon?nico datoteke? Mo?no je, da potem datoteka ne bo uporabna.',
	FileRenaming	: 'Preimenujem...',
	FileDelete		: 'Ali ste prepri?ani, da ?elite izbrisati datoteko "%1"?',
	FilesDelete	: 'Are you sure you want to delete %1 files?', // MISSING
	FilesLoading	: 'Nalagam...',
	FilesEmpty		: 'Prazna mapa',
	DestinationFile	: 'Destination File', // MISSING
	SkippedFiles	: 'List of skipped files:', // MISSING

	// Basket
	BasketFolder		: 'Ko?',
	BasketClear			: 'Izprazni ko?',
	BasketRemove		: 'Odstrani iz ko?a',
	BasketOpenFolder	: 'Odpri izvorno mapo',
	BasketTruncateConfirm : 'Ali res ?elite odstraniti vse datoteke iz ko?a?',
	BasketRemoveConfirm	: 'Ali res ?elite odstraniti datoteko "%1" iz ko?a?',
	BasketRemoveConfirmMultiple	: 'Do you really want to remove %1 files from the basket?', // MISSING
	BasketEmpty			: 'V ko?u ni datotek. Lahko jih povle?ete in spustite.',
	BasketCopyFilesHere	: 'Kopiraj datoteke iz ko?a',
	BasketMoveFilesHere	: 'Premakni datoteke iz ko?a',

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
	Upload		: 'Nalo?i na stre?nik',
	UploadTip	: 'Nalo?i novo datoteko na stre?nik',
	Refresh		: 'Osve?i',
	Settings	: 'Nastavitve',
	Help		: 'Pomo?',
	HelpTip		: 'Pomo?',

	// Context Menus
	Select			: 'Izberi',
	SelectThumbnail : 'Izberi malo sli?ico (predogled)',
	View			: 'Predogled',
	Download		: 'Prenesi na svoj ra?unalnik',

	NewSubFolder	: 'Nova podmapa',
	Rename			: 'Preimenuj',
	Delete			: 'Zbri?i',
	DeleteFiles		: 'Delete Files', // MISSING

	CopyDragDrop	: 'Kopiraj',
	MoveDragDrop	: 'Premakni',

	// Dialogs
	RenameDlgTitle		: 'Preimenuj',
	NewNameDlgTitle		: 'Novo ime',
	FileExistsDlgTitle	: 'Datoteka ?e obstaja',
	SysErrorDlgTitle : 'Sistemska napaka',

	FileOverwrite	: 'Prepi?i',
	FileAutorename	: 'Avtomatsko preimenuj',
	ManuallyRename	: 'Manually rename', // MISSING

	// Generic
	OkBtn		: 'Potrdi',
	CancelBtn	: 'Prekli?i',
	CloseBtn	: 'Zapri',

	// Upload Panel
	UploadTitle			: 'Nalo?i novo datoteko na stre?nik',
	UploadSelectLbl		: 'Izberi datoteko za prenos na stre?nik',
	UploadProgressLbl	: '(Prenos na stre?nik poteka, prosimo po?akajte...)',
	UploadBtn			: 'Prenesi izbrano datoteko na stre?nik',
	UploadBtnCancel		: 'Prekli?i',

	UploadNoFileMsg		: 'Prosimo izberite datoteko iz svojega ra?unalnika za prenos na stre?nik.',
	UploadNoFolder		: 'Izberite mapo v katero se bo nalo?ilo datoteko!',
	UploadNoPerms		: 'Nalaganje datotek ni dovoljeno.',
	UploadUnknError		: 'Napaka pri po?iljanju datoteke.',
	UploadExtIncorrect	: 'V tej mapi ta vrsta datoteke ni dovoljena.',

	// Flash Uploads
	UploadLabel			: 'Datoteke za prenos',
	UploadTotalFiles	: 'Skupaj datotek:',
	UploadTotalSize		: 'Skupaj velikost:',
	UploadSend			: 'Nalo?i na stre?nik',
	UploadAddFiles		: 'Dodaj datoteke',
	UploadClearFiles	: 'Po?isti datoteke',
	UploadCancel		: 'Prekli?i prenos',
	UploadRemove		: 'Odstrani',
	UploadRemoveTip		: 'Odstrani !f',
	UploadUploaded		: 'Prene?eno !n%',
	UploadProcessing	: 'Delam...',

	// Settings Panel
	SetTitle		: 'Nastavitve',
	SetView			: 'Pogled:',
	SetViewThumb	: 'majhne sli?ice',
	SetViewList		: 'seznam',
	SetDisplay		: 'Prikaz:',
	SetDisplayName	: 'ime datoteke',
	SetDisplayDate	: 'datum',
	SetDisplaySize	: 'velikost datoteke',
	SetSort			: 'Razvr??anje:',
	SetSortName		: 'po imenu datoteke',
	SetSortDate		: 'po datumu',
	SetSortSize		: 'po velikosti',
	SetSortExtension		: 'po kon?nici',

	// Status Bar
	FilesCountEmpty : '<Prazna mapa>',
	FilesCountOne	: '1 datoteka',
	FilesCountMany	: '%1 datotek(e)',

	// Size and Speed
	Kb				: '%1 KB',
	Mb				: '%1 MB',
	Gb				: '%1 GB',
	SizePerSecond	: '%1/s',

	// Connector Error Messages.
	ErrorUnknown	: 'Pri?lo je do napake. (Napaka %1)',
	Errors :
	{
	 10 : 'Napa?en ukaz.',
	 11 : 'V poizvedbi ni bil jasen tip (resource type).',
	 12 : 'Tip datoteke ni primeren.',
	102 : 'Napa?no ime mape ali datoteke.',
	103 : 'Va?ega ukaza se ne da izvesti zaradi te?av z avtorizacijo.',
	104 : 'Va?ega ukaza se ne da izvesti zaradi te?av z nastavitvami pravic v datote?nem sistemu.',
	105 : 'Napa?na kon?nica datoteke.',
	109 : 'Napa?na zahteva.',
	110 : 'Neznana napaka.',
	111 : 'It was not possible to complete the request due to resulting file size.', // MISSING
	115 : 'Datoteka ali mapa s tem imenom ?e obstaja.',
	116 : 'Mapa ni najdena. Prosimo osve?ite okno in poskusite znova.',
	117 : 'Datoteka ni najdena. Prosimo osve?ite seznam datotek in poskusite znova.',
	118 : 'Za?etna in kon?na pot je ista.',
	201 : 'Datoteka z istim imenom ?e obstaja. Nalo?ena datoteka je bila preimenovana v "%1".',
	202 : 'Neprimerna datoteka.',
	203 : 'Datoteka je prevelika in zasede preve? prostora.',
	204 : 'Nalo?ena datoteka je okvarjena.',
	205 : 'Na stre?niku ni na voljo za?asna mapa za prenos datotek.',
	206 : 'Nalaganje je bilo prekinjeno zaradi varnostnih razlogov. Datoteka vsebuje podatke, ki spominjajo na HTML kodo.',
	207 : 'Nalo?ena datoteka je bila preimenovana v "%1".',
	300 : 'Premikanje datotek(e) ni uspelo.',
	301 : 'Kopiranje datotek(e) ni uspelo.',
	500 : 'Brskalnik je onemogo?en zaradi varnostnih razlogov. Prosimo kontaktirajte upravljalca spletnih strani.',
	501 : 'Ni podpore za majhne sli?ice (predogled).'
	},

	// Other Error Messages.
	ErrorMsg :
	{
		FileEmpty		: 'Ime datoteke ne more biti prazno.',
		FileExists		: 'Datoteka %s ?e obstaja.',
		FolderEmpty		: 'Mapa ne more biti prazna.',
		FolderExists	: 'Folder %s already exists.', // MISSING
		FolderNameExists	: 'Folder already exists.', // MISSING

		FileInvChar		: 'Ime datoteke ne sme vsebovati naslednjih znakov: \n\\ / : * ? " < > |',
		FolderInvChar	: 'Ime mape ne sme vsebovati naslednjih znakov: \n\\ / : * ? " < > |',

		PopupBlockView	: 'Datoteke ni mo?no odpreti v novem oknu. Prosimo nastavite svoj brskalnik tako, da bo dopu??al odpiranje oken (popups) oz. izklopite filtre za blokado odpiranja oken.',
		XmlError		: 'Nalaganje XML odgovora iz stre?nika ni uspelo.',
		XmlEmpty		: 'Nalaganje XML odgovora iz stre?nika ni uspelo. Stre?nik je vrnil prazno sporo?ilo.',
		XmlRawResponse	: 'Surov odgovor iz stre?nika je: %s'
	},

	// Imageresize plugin
	Imageresize :
	{
		dialogTitle		: 'Spremeni velikost slike %s',
		sizeTooBig		: '?irina ali vi?ina slike ne moreta biti ve?ji kot je originalna velikost (%size).',
		resizeSuccess	: 'Velikost slike je bila uspe?no spremenjena.',
		thumbnailNew	: 'Kreiraj novo majhno sli?ico',
		thumbnailSmall	: 'majhna (%s)',
		thumbnailMedium	: 'srednja (%s)',
		thumbnailLarge	: 'velika (%s)',
		newSize			: 'Dolo?ite novo velikost',
		width			: '?irina',
		height			: 'Vi?ina',
		invalidHeight	: 'Nepravilna vi?ina.',
		invalidWidth	: 'Nepravilna ?irina.',
		invalidName		: 'Nepravilno ime datoteke.',
		newImage		: 'Kreiraj novo sliko',
		noExtensionChange : 'Kon?nica datoteke se ne more spremeniti.',
		imageSmall		: 'Izvorna slika je premajhna.',
		contextMenuName	: 'Spremeni velikost',
		lockRatio		: 'Zakleni razmerje',
		resetSize		: 'Ponastavi velikost'
	},

	// Fileeditor plugin
	Fileeditor :
	{
		save			: 'Shrani',
		fileOpenError	: 'Datoteke ni mogo?e odpreti.',
		fileSaveSuccess	: 'Datoteka je bila shranjena.',
		contextMenuName	: 'Uredi',
		loadingFile		: 'Nalaganje datoteke, prosimo po?akajte ...'
	},

	Maximize :
	{
		maximize : 'Maksimiraj',
		minimize : 'Minimiraj'
	},

	Gallery :
	{
		current : 'Slika {current} od {total}'
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
