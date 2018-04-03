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
 * @fileOverview Defines the {@link CKFinder.lang} object for the German
 *		language.
 */

/**
 * Contains the dictionary of language entries.
 * @namespace
 */
CKFinder.lang['de'] =
{
	appTitle : 'CKFinder',

	// Common messages and labels.
	common :
	{
		// Put the voice-only part of the label in the span.
		unavailable		: '%1<span class="cke_accessibility">, nicht verfügbar</span>',
		confirmCancel	: 'Einige Optionen wurden ge?ndert. Wollen Sie den Dialog dennoch schlie?en?',
		ok				: 'OK',
		cancel			: 'Abbrechen',
		confirmationTitle	: 'Best?tigung',
		messageTitle	: 'Information',
		inputTitle		: 'Frage',
		undo			: 'Rückg?ngig',
		redo			: 'Wiederherstellen',
		skip			: 'überspringen',
		skipAll			: 'Alle überspringen',
		makeDecision	: 'Bitte Auswahl treffen.',
		rememberDecision: 'Entscheidung merken'
	},


	// Language direction, 'ltr' or 'rtl'.
	dir : 'ltr',
	HelpLang : 'en',
	LangCode : 'de',

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
	FoldersTitle	: 'Verzeichnisse',
	FolderLoading	: 'Laden...',
	FolderNew		: 'Bitte geben Sie den neuen Verzeichnisnamen an: ',
	FolderRename	: 'Bitte geben Sie den neuen Verzeichnisnamen an: ',
	FolderDelete	: 'Wollen Sie wirklich den Ordner "%1" l?schen?',
	FolderRenaming	: ' (Umbenennen...)',
	FolderDeleting	: ' (L?schen...)',
	DestinationFolder	: 'Destination Folder', // MISSING

	// Files
	FileRename		: 'Bitte geben Sie den neuen Dateinamen an: ',
	FileRenameExt	: 'Wollen Sie wirklich die Dateierweiterung ?ndern? Die Datei k?nnte unbrauchbar werden!',
	FileRenaming	: 'Umbennenen...',
	FileDelete		: 'Wollen Sie wirklich die Datei "%1" l?schen?',
	FilesDelete	: 'Are you sure you want to delete %1 files?', // MISSING
	FilesLoading	: 'Laden...',
	FilesEmpty		: 'Verzeichnis ist leer.',
	DestinationFile	: 'Destination File', // MISSING
	SkippedFiles	: 'List of skipped files:', // MISSING

	// Basket
	BasketFolder		: 'Korb',
	BasketClear			: 'Korb l?schen',
	BasketRemove		: 'Aus dem Korb entfernen',
	BasketOpenFolder	: 'übergeordneten Ordner ?ffnen',
	BasketTruncateConfirm : 'Wollen Sie wirklich alle Dateien aus dem Korb entfernen?',
	BasketRemoveConfirm	: 'Wollen Sie wirklich die Datei "%1" aus dem Korb entfernen?',
	BasketRemoveConfirmMultiple	: 'Do you really want to remove %1 files from the basket?', // MISSING
	BasketEmpty			: 'Keine Dateien im Korb, einfach welche reinziehen.',
	BasketCopyFilesHere	: 'Dateien aus dem Korb kopieren',
	BasketMoveFilesHere	: 'Dateien aus dem Korb verschieben',

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
	Upload		: 'Hochladen',
	UploadTip	: 'Neue Datei hochladen',
	Refresh		: 'Aktualisieren',
	Settings	: 'Einstellungen',
	Help		: 'Hilfe',
	HelpTip		: 'Hilfe',

	// Context Menus
	Select			: 'Ausw?hlen',
	SelectThumbnail : 'Miniatur ausw?hlen',
	View			: 'Ansehen',
	Download		: 'Herunterladen',

	NewSubFolder	: 'Neues Unterverzeichnis',
	Rename			: 'Umbenennen',
	Delete			: 'L?schen',
	DeleteFiles		: 'Delete Files', // MISSING

	CopyDragDrop	: 'Hierher kopieren',
	MoveDragDrop	: 'Hierher verschieben',

	// Dialogs
	RenameDlgTitle		: 'Umbenennen',
	NewNameDlgTitle		: 'Neuer Name',
	FileExistsDlgTitle	: 'Datei existiert bereits',
	SysErrorDlgTitle : 'Systemfehler',

	FileOverwrite	: 'überschreiben',
	FileAutorename	: 'Automatisch umbenennen',
	ManuallyRename	: 'Manually rename', // MISSING

	// Generic
	OkBtn		: 'OK',
	CancelBtn	: 'Abbrechen',
	CloseBtn	: 'Schlie?en',

	// Upload Panel
	UploadTitle			: 'Neue Datei hochladen',
	UploadSelectLbl		: 'Bitte w?hlen Sie die Datei aus',
	UploadProgressLbl	: '(Die Daten werden übertragen, bitte warten...)',
	UploadBtn			: 'Ausgew?hlte Datei hochladen',
	UploadBtnCancel		: 'Abbrechen',

	UploadNoFileMsg		: 'Bitte w?hlen Sie eine Datei auf Ihrem Computer aus.',
	UploadNoFolder		: 'Bitte ein Verzeichnis vor dem Hochladen w?hlen.',
	UploadNoPerms		: 'Datei hochladen nicht erlaubt.',
	UploadUnknError		: 'Fehler bei Dateitragung.',
	UploadExtIncorrect	: 'Dateinamekürzel nicht in diesem Verzeichnis erlaubt.',

	// Flash Uploads
	UploadLabel			: 'Dateien zum Hochladen',
	UploadTotalFiles	: 'Gesamtanzahl Dateien:',
	UploadTotalSize		: 'Gesamtgr??e:',
	UploadSend			: 'Hochladen',
	UploadAddFiles		: 'Datei hinzufügen',
	UploadClearFiles	: 'Dateiliste l?schen',
	UploadCancel		: 'Upload abbrechen',
	UploadRemove		: 'Entfernen',
	UploadRemoveTip		: 'Entfernen !f',
	UploadUploaded		: 'Hochgeladen !n%',
	UploadProcessing	: 'In Arbeit...',

	// Settings Panel
	SetTitle		: 'Einstellungen',
	SetView			: 'Ansicht:',
	SetViewThumb	: 'Miniaturansicht',
	SetViewList		: 'Liste',
	SetDisplay		: 'Anzeige:',
	SetDisplayName	: 'Dateiname',
	SetDisplayDate	: 'Datum',
	SetDisplaySize	: 'Dateigr??e',
	SetSort			: 'Sortierung:',
	SetSortName		: 'nach Dateinamen',
	SetSortDate		: 'nach Datum',
	SetSortSize		: 'nach Gr??e',
	SetSortExtension		: 'nach Dateiendung',

	// Status Bar
	FilesCountEmpty : '<Leeres Verzeichnis>',
	FilesCountOne	: '1 Datei',
	FilesCountMany	: '%1 Datei',

	// Size and Speed
	Kb				: '%1 KB',
	Mb				: '%1 MB',
	Gb				: '%1 GB',
	SizePerSecond	: '%1/s',

	// Connector Error Messages.
	ErrorUnknown	: 'Ihre Anfrage konnte nicht bearbeitet werden. (Fehler %1)',
	Errors :
	{
	 10 : 'Unbekannter Befehl.',
	 11 : 'Der Ressourcentyp wurde nicht spezifiziert.',
	 12 : 'Der Ressourcentyp ist nicht gültig.',
	102 : 'Ungültiger Datei oder Verzeichnisname.',
	103 : 'Ihre Anfrage konnte wegen Authorisierungseinschr?nkungen nicht durchgeführt werden.',
	104 : 'Ihre Anfrage konnte wegen Dateisystemeinschr?nkungen nicht durchgeführt werden.',
	105 : 'Invalid file extension.',
	109 : 'Unbekannte Anfrage.',
	110 : 'Unbekannter Fehler.',
	111 : 'It was not possible to complete the request due to resulting file size.', // MISSING
	115 : 'Es existiert bereits eine Datei oder ein Ordner mit dem gleichen Namen.',
	116 : 'Verzeichnis nicht gefunden. Bitte aktualisieren Sie die Anzeige und versuchen es noch einmal.',
	117 : 'Datei nicht gefunden. Bitte aktualisieren Sie die Dateiliste und versuchen es noch einmal.',
	118 : 'Quell- und Zielpfad sind gleich.',
	201 : 'Es existiert bereits eine Datei unter gleichem Namen. Die hochgeladene Datei wurde unter "%1" gespeichert.',
	202 : 'Ungültige Datei.',
	203 : 'ungültige Datei. Die Dateigr??e ist zu gro?.',
	204 : 'Die hochgeladene Datei ist korrupt.',
	205 : 'Es existiert kein temp. Ordner für das Hochladen auf den Server.',
	206 : 'Das Hochladen wurde aus Sicherheitsgründen abgebrochen. Die Datei enth?lt HTML-Daten.',
	207 : 'Die hochgeladene Datei wurde unter "%1" gespeichert.',
	300 : 'Verschieben der Dateien fehlgeschlagen.',
	301 : 'Kopieren der Dateien fehlgeschlagen.',
	500 : 'Der Dateibrowser wurde aus Sicherheitsgründen deaktiviert. Bitte benachrichtigen Sie Ihren Systemadministrator und prüfen Sie die Konfigurationsdatei.',
	501 : 'Die Miniaturansicht wurde deaktivert.'
	},

	// Other Error Messages.
	ErrorMsg :
	{
		FileEmpty		: 'Der Dateinamen darf nicht leer sein.',
		FileExists		: 'Datei %s existiert bereits.',
		FolderEmpty		: 'Der Verzeichnisname darf nicht leer sein.',
		FolderExists	: 'Folder %s already exists.', // MISSING
		FolderNameExists	: 'Folder already exists.', // MISSING

		FileInvChar		: 'Der Dateinamen darf nicht eines der folgenden Zeichen enthalten: \n\\ / : * ? " < > |',
		FolderInvChar	: 'Der Verzeichnisname darf nicht eines der folgenden Zeichen enthalten: \n\\ / : * ? " < > |',

		PopupBlockView	: 'Die Datei konnte nicht in einem neuen Fenster ge?ffnet werden. Bitte deaktivieren Sie in Ihrem Browser alle Popup-Blocker für diese Seite.',
		XmlError		: 'Es war nicht m?glich die XML-Antwort von dem Server herunterzuladen.',
		XmlEmpty		: 'Es war nicht m?glich die XML-Antwort von dem Server herunterzuladen. Der Server hat eine leere Nachricht zurückgeschickt.',
		XmlRawResponse	: 'Raw-Antwort vom Server: %s'
	},

	// Imageresize plugin
	Imageresize :
	{
		dialogTitle		: 'Gr??en?nderung %s',
		sizeTooBig		: 'Bildgr??e kann nicht gr??er als das Originalbild werden (%size).',
		resizeSuccess	: 'Bildgr??e erfolgreich ge?ndert.',
		thumbnailNew	: 'Neues Vorschaubild erstellen',
		thumbnailSmall	: 'Klein (%s)',
		thumbnailMedium	: 'Mittel (%s)',
		thumbnailLarge	: 'Gro? (%s)',
		newSize			: 'Eine neue Gr??e setzen',
		width			: 'Breite',
		height			: 'H?he',
		invalidHeight	: 'Ungültige H?he.',
		invalidWidth	: 'Ungültige Breite.',
		invalidName		: 'Ungültiger Name.',
		newImage		: 'Neues Bild erstellen',
		noExtensionChange : 'Dateierweiterung kann nicht ge?ndert werden.',
		imageSmall		: 'Bildgr??e zu klein.',
		contextMenuName	: 'Gr??en?nderung',
		lockRatio		: 'Gr??enverh?ltnis beibehalten',
		resetSize		: 'Gr??e zurücksetzen'
	},

	// Fileeditor plugin
	Fileeditor :
	{
		save			: 'Speichern',
		fileOpenError	: 'Datei kann nicht ge?ffnet werden.',
		fileSaveSuccess	: 'Datei erfolgreich gespeichert.',
		contextMenuName	: 'Bearbeitung',
		loadingFile		: 'Datei wird geladen, einen Moment noch...'
	},

	Maximize :
	{
		maximize : 'Maximieren',
		minimize : 'Minimieren'
	},

	Gallery :
	{
		current : 'Bild {current} von {total}'
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
