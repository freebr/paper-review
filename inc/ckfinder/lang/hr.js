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
 * @fileOverview Defines the {@link CKFinder.lang} object for the Croatian
 *		language.
 */

/**
 * Contains the dictionary of language entries.
 * @namespace
 */
CKFinder.lang['hr'] =
{
	appTitle : 'CKFinder',

	// Common messages and labels.
	common :
	{
		// Put the voice-only part of the label in the span.
		unavailable		: '%1<span class="cke_accessibility">, nedostupno</span>',
		confirmCancel	: 'Neke od opcija su promjenjene. Sigurno ?elite zatvoriti prozor??',
		ok				: 'U redu',
		cancel			: 'Poni?ti',
		confirmationTitle	: 'Potvrda',
		messageTitle	: 'Informacija',
		inputTitle		: 'Pitanje',
		undo			: 'Poni?ti',
		redo			: 'Preuredi',
		skip			: 'Presko?i',
		skipAll			: 'Presko?i sve',
		makeDecision	: '?to bi trebali napraviti?',
		rememberDecision: 'Zapamti moj izbor'
	},


	// Language direction, 'ltr' or 'rtl'.
	dir : 'ltr',
	HelpLang : 'en',
	LangCode : 'hr',

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
	DateTime : 'm/d/yyyy h:MM aa',
	DateAmPm : ['AM', 'PM'],

	// Folders
	FoldersTitle	: 'Direktoriji',
	FolderLoading	: 'U?itavam...',
	FolderNew		: 'Unesite novo ime direktorija: ',
	FolderRename	: 'Unesite novo ime direktorija: ',
	FolderDelete	: 'Sigurno ?elite obrisati direktorij "%1"?',
	FolderRenaming	: ' (Mijenjam ime...)',
	FolderDeleting	: ' (Bri?em...)',
	DestinationFolder	: 'Destination Folder', // MISSING

	// Files
	FileRename		: 'Unesite novo ime datoteke: ',
	FileRenameExt	: 'Sigurno ?elite promijeniti vrstu datoteke? Datoteka mo?e postati neiskoristiva.',
	FileRenaming	: 'Mijenjam ime...',
	FileDelete		: 'Sigurno ?elite obrisati datoteku "%1"?',
	FilesDelete	: 'Are you sure you want to delete %1 files?', // MISSING
	FilesLoading	: 'U?itavam...',
	FilesEmpty		: 'Direktorij je prazan.',
	DestinationFile	: 'Destination File', // MISSING
	SkippedFiles	: 'List of skipped files:', // MISSING

	// Basket
	BasketFolder		: 'Ko?ara',
	BasketClear			: 'Isprazni ko?aru',
	BasketRemove		: 'Ukloni iz ko?are',
	BasketOpenFolder	: 'Otvori nadre?eni direktorij',
	BasketTruncateConfirm : 'Sigurno ?elite obrisati sve datoteke iz ko?are?',
	BasketRemoveConfirm	: 'Sigurno ?elite obrisati datoteku "%1" iz ko?are?',
	BasketRemoveConfirmMultiple	: 'Do you really want to remove %1 files from the basket?', // MISSING
	BasketEmpty			: 'Nema niti jedne datoteke, ubacite koju.',
	BasketCopyFilesHere	: 'Kopiraj datoteke iz ko?are',
	BasketMoveFilesHere	: 'Premjesti datoteke iz ko?are',

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
	Upload		: 'Po?alji',
	UploadTip	: 'Po?alji nove datoteke na server',
	Refresh		: 'Osvje?i',
	Settings	: 'Postavke',
	Help		: 'Pomo?',
	HelpTip		: 'Pomo?',

	// Context Menus
	Select			: 'Odaberi',
	SelectThumbnail : 'Odaberi manju sliku',
	View			: 'Pogledaj',
	Download		: 'Skini',

	NewSubFolder	: 'Novi poddirektorij',
	Rename			: 'Promijeni naziv',
	Delete			: 'Obri?i',
	DeleteFiles		: 'Delete Files', // MISSING

	CopyDragDrop	: 'Kopiraj ovdje',
	MoveDragDrop	: 'Premjesti ovdje',

	// Dialogs
	RenameDlgTitle		: 'Promijeni naziv',
	NewNameDlgTitle		: 'Novi naziv',
	FileExistsDlgTitle	: 'Datoteka ve? postoji',
	SysErrorDlgTitle : 'Gre?ka sustava',

	FileOverwrite	: 'Prepi?i',
	FileAutorename	: 'Automatska promjena naziva',
	ManuallyRename	: 'Manually rename', // MISSING

	// Generic
	OkBtn		: 'U redu',
	CancelBtn	: 'Poni?ti',
	CloseBtn	: 'Zatvori',

	// Upload Panel
	UploadTitle			: 'Po?alji novu datoteku',
	UploadSelectLbl		: 'Odaberi datoteku za slanje',
	UploadProgressLbl	: '(Slanje u tijeku, molimo pri?ekajte...)',
	UploadBtn			: 'Po?alji odabranu datoteku',
	UploadBtnCancel		: 'Poni?ti',

	UploadNoFileMsg		: 'Odaberite datoteku na Va?em ra?unalu.',
	UploadNoFolder		: 'Odaberite direktorije prije slanja.',
	UploadNoPerms		: 'Slanje datoteka nije dozvoljeno.',
	UploadUnknError		: 'Gre?ka kod slanja datoteke.',
	UploadExtIncorrect	: 'Vrsta datoteka nije dozvoljena.',

	// Flash Uploads
	UploadLabel			: 'Datoteka za slanje:',
	UploadTotalFiles	: 'Ukupno datoteka:',
	UploadTotalSize		: 'Ukupna veli?ina:',
	UploadSend			: 'Po?alji',
	UploadAddFiles		: 'Dodaj datoteke',
	UploadClearFiles	: 'Izbaci datoteke',
	UploadCancel		: 'Poni?ti slanje',
	UploadRemove		: 'Ukloni',
	UploadRemoveTip		: 'Ukloni !f',
	UploadUploaded		: 'Poslano !n%',
	UploadProcessing	: 'Obra?ujem...',

	// Settings Panel
	SetTitle		: 'Postavke',
	SetView			: 'Pregled:',
	SetViewThumb	: 'Mala slika',
	SetViewList		: 'Lista',
	SetDisplay		: 'Prikaz:',
	SetDisplayName	: 'Naziv datoteke',
	SetDisplayDate	: 'Datum',
	SetDisplaySize	: 'Veli?ina datoteke',
	SetSort			: 'Sortiranje:',
	SetSortName		: 'po nazivu',
	SetSortDate		: 'po datumu',
	SetSortSize		: 'po veli?ini',
	SetSortExtension		: 'po vrsti datoteke',

	// Status Bar
	FilesCountEmpty : '<Prazan direktorij>',
	FilesCountOne	: '1 datoteka',
	FilesCountMany	: '%1 datoteka',

	// Size and Speed
	Kb				: '%1 KB',
	Mb				: '%1 MB',
	Gb				: '%1 GB',
	SizePerSecond	: '%1/s',

	// Connector Error Messages.
	ErrorUnknown	: 'Nije mogu?e zavr?iti zahtjev. (Gre?ka %1)',
	Errors :
	{
	 10 : 'Nepoznata naredba.',
	 11 : 'Nije navedena vrsta u zahtjevu.',
	 12 : 'Zatra?ena vrsta nije va?e?a.',
	102 : 'Neispravno naziv datoteke ili direktoija.',
	103 : 'Nije mogu?e izvr?iti zahtjev zbog ograni?enja pristupa.',
	104 : 'Nije mogu?e izvr?iti zahtjev zbog ograni?enja postavka sustava.',
	105 : 'Nedozvoljena vrsta datoteke.',
	109 : 'Nedozvoljen zahtjev.',
	110 : 'Nepoznata gre?ka.',
	111 : 'It was not possible to complete the request due to resulting file size.', // MISSING
	115 : 'Datoteka ili direktorij s istim nazivom ve? postoji.',
	116 : 'Direktorij nije prona?en. Osvje?ite stranicu i poku?ajte ponovo.',
	117 : 'Datoteka nije prona?ena. Osvje?ite listu datoteka i poku?ajte ponovo.',
	118 : 'Putanje izvora i odredi?ta su jednake.',
	201 : 'Datoteka s istim nazivom ve? postoji. Poslana datoteka je promjenjena u "%1".',
	202 : 'Neispravna datoteka.',
	203 : 'Neispravna datoteka. Veli?ina datoteke je prevelika.',
	204 : 'Poslana datoteka je neispravna.',
	205 : 'Ne postoji privremeni direktorij za slanje na server.',
	206 : 'Slanje je poni?teno zbog sigurnosnih postavki. Naziv datoteke sadr?i HTML podatke.',
	207 : 'Poslana datoteka je promjenjena u "%1".',
	300 : 'Premje?tanje datoteke(a) nije uspjelo.',
	301 : 'Kopiranje datoteke(a) nije uspjelo.',
	500 : 'Pretra?ivanje datoteka nije dozvoljeno iz sigurnosnih razloga. Molimo kontaktirajte administratora sustava kako bi provjerili postavke CKFinder konfiguracijske datoteke.',
	501 : 'The thumbnails support is disabled.'
	},

	// Other Error Messages.
	ErrorMsg :
	{
		FileEmpty		: 'Naziv datoteke ne mo?e biti prazan.',
		FileExists		: 'Datoteka %s ve? postoji.',
		FolderEmpty		: 'Naziv direktorija ne mo?e biti prazan.',
		FolderExists	: 'Folder %s already exists.', // MISSING
		FolderNameExists	: 'Folder already exists.', // MISSING

		FileInvChar		: 'Naziv datoteke ne smije sadr?avati niti jedan od sljede?ih znakova: \n\\ / : * ? " < > |',
		FolderInvChar	: 'Naziv direktorija ne smije sadr?avati niti jedan od sljede?ih znakova: \n\\ / : * ? " < > |',

		PopupBlockView	: 'Nije mogu?e otvoriti datoteku u novom prozoru. Promijenite postavke svog Internet preglednika i isklju?ite sve popup blokere za ove web stranice.',
		XmlError		: 'Nije mogu?e u?itati XML odgovor od web servera.',
		XmlEmpty		: 'Nije mogu?e u?itati XML odgovor od web servera. Server je vratio prazan odgovor.',
		XmlRawResponse	: 'Odgovor servera: %s'
	},

	// Imageresize plugin
	Imageresize :
	{
		dialogTitle		: 'Promijeni veli?inu %s',
		sizeTooBig		: 'Nije mogu?e postaviti veli?inu ve?u od originala (%size).',
		resizeSuccess	: 'Slika je uspje?no promijenjena.',
		thumbnailNew	: 'Napravi malu sliku',
		thumbnailSmall	: 'Mala (%s)',
		thumbnailMedium	: 'Srednja (%s)',
		thumbnailLarge	: 'Velika (%s)',
		newSize			: 'Postavi novu veli?inu',
		width			: '?irina',
		height			: 'Visina',
		invalidHeight	: 'Neispravna visina.',
		invalidWidth	: 'Neispravna ?irina.',
		invalidName		: 'Neispravan naziv datoteke.',
		newImage		: 'Napravi novu sliku',
		noExtensionChange : 'Vrsta datoteke se ne smije mijenjati.',
		imageSmall		: 'Izvorna slika je premala.',
		contextMenuName	: 'Promijeni veli?inu',
		lockRatio		: 'Zaklju?aj odnose',
		resetSize		: 'Vrati veli?inu'
	},

	// Fileeditor plugin
	Fileeditor :
	{
		save			: 'Snimi',
		fileOpenError	: 'Nije mogu?e otvoriti datoteku.',
		fileSaveSuccess	: 'Datoteka je uspje?no snimljena.',
		contextMenuName	: 'Promjeni',
		loadingFile		: 'U?itavam, molimo pri?ekajte...'
	},

	Maximize :
	{
		maximize : 'Pove?aj',
		minimize : 'Smanji'
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
