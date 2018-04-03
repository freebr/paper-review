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
 * @fileOverview Defines the {@link CKFinder.lang} object for the Polish
 *		language.
 */

/**
 * Contains the dictionary of language entries.
 * @namespace
 */
CKFinder.lang['pl'] =
{
	appTitle : 'CKFinder',

	// Common messages and labels.
	common :
	{
		// Put the voice-only part of the label in the span.
		unavailable		: '%1<span class="cke_accessibility">, wy??czone</span>',
		confirmCancel	: 'Pewne opcje zosta?y zmienione. Czy na pewno zamkn?? okno dialogowe?',
		ok				: 'OK',
		cancel			: 'Anuluj',
		confirmationTitle	: 'Potwierdzenie',
		messageTitle	: 'Informacja',
		inputTitle		: 'Pytanie',
		undo			: 'Cofnij',
		redo			: 'Ponów',
		skip			: 'Pomiń',
		skipAll			: 'Pomiń wszystkie',
		makeDecision	: 'Wybierz jedn? z opcji:',
		rememberDecision: 'Zapami?taj mój wybór'
	},


	// Language direction, 'ltr' or 'rtl'.
	dir : 'ltr',
	HelpLang : 'pl',
	LangCode : 'pl',

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
	DateTime : 'yyyy-mm-dd HH:MM',
	DateAmPm : ['AM', 'PM'],

	// Folders
	FoldersTitle	: 'Foldery',
	FolderLoading	: '?adowanie...',
	FolderNew		: 'Podaj nazw? nowego folderu: ',
	FolderRename	: 'Podaj now? nazw? folderu: ',
	FolderDelete	: 'Czy na pewno chcesz usun?? folder "%1"?',
	FolderRenaming	: ' (Zmieniam nazw?...)',
	FolderDeleting	: ' (Kasowanie...)',
	DestinationFolder	: 'Folder docelowy',

	// Files
	FileRename		: 'Podaj now? nazw? pliku: ',
	FileRenameExt	: 'Czy na pewno chcesz zmieni? rozszerzenie pliku? Mo?e to spowodowa? problemy z otwieraniem pliku przez innych u?ytkowników.',
	FileRenaming	: 'Zmieniam nazw?...',
	FileDelete		: 'Czy na pewno chcesz usun?? plik "%1"?',
	FilesDelete	: 'Czy na pewno chcesz usun?? pliki (razem: %1)?',
	FilesLoading	: '?adowanie...',
	FilesEmpty		: 'Folder jest pusty',
	DestinationFile	: 'Plik docelowy',
	SkippedFiles	: 'Lista pomini?tych plików:',

	// Basket
	BasketFolder		: 'Koszyk',
	BasketClear			: 'Wyczy?? koszyk',
	BasketRemove		: 'Usuń z koszyka',
	BasketOpenFolder	: 'Otwórz folder z plikiem',
	BasketTruncateConfirm : 'Czy naprawd? chcesz usun?? wszystkie pliki z koszyka?',
	BasketRemoveConfirm	: 'Czy naprawd? chcesz usun?? plik "%1" z koszyka?',
	BasketRemoveConfirmMultiple	: 'Czy naprawd? chcesz usun?? pliki (razem: %1) z koszyka?',
	BasketEmpty			: 'Brak plików w koszyku. Aby doda? plik, przeci?gnij i upu?? (drag\'n\'drop) dowolny plik do koszyka.',
	BasketCopyFilesHere	: 'Skopiuj pliki z koszyka',
	BasketMoveFilesHere	: 'Przenie? pliki z koszyka',

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
	Upload		: 'Wy?lij',
	UploadTip	: 'Wy?lij plik',
	Refresh		: 'Od?wie?',
	Settings	: 'Ustawienia',
	Help		: 'Pomoc',
	HelpTip		: 'Wskazówka',

	// Context Menus
	Select			: 'Wybierz',
	SelectThumbnail : 'Wybierz miniaturk?',
	View			: 'Zobacz',
	Download		: 'Pobierz',

	NewSubFolder	: 'Nowy podfolder',
	Rename			: 'Zmień nazw?',
	Delete			: 'Usuń',
	DeleteFiles		: 'Usuń pliki',

	CopyDragDrop	: 'Skopiuj tutaj',
	MoveDragDrop	: 'Przenie? tutaj',

	// Dialogs
	RenameDlgTitle		: 'Zmiana nazwy',
	NewNameDlgTitle		: 'Nowa nazwa',
	FileExistsDlgTitle	: 'Plik ju? istnieje',
	SysErrorDlgTitle : 'B??d systemu',

	FileOverwrite	: 'Nadpisz',
	FileAutorename	: 'Zmień automatycznie nazw?',
	ManuallyRename	: 'Zmień nazw? r?cznie',

	// Generic
	OkBtn		: 'OK',
	CancelBtn	: 'Anuluj',
	CloseBtn	: 'Zamknij',

	// Upload Panel
	UploadTitle			: 'Wy?lij plik',
	UploadSelectLbl		: 'Wybierz plik',
	UploadProgressLbl	: '(Trwa wysy?anie pliku, prosz? czeka?...)',
	UploadBtn			: 'Wy?lij wybrany plik',
	UploadBtnCancel		: 'Anuluj',

	UploadNoFileMsg		: 'Wybierz plik ze swojego komputera.',
	UploadNoFolder		: 'Wybierz folder przed wys?aniem pliku.',
	UploadNoPerms		: 'Wysy?anie plików nie jest dozwolone.',
	UploadUnknError		: 'B??d podczas wysy?ania pliku.',
	UploadExtIncorrect	: 'Rozszerzenie pliku nie jest dozwolone w tym folderze.',

	// Flash Uploads
	UploadLabel			: 'Pliki do wys?ania',
	UploadTotalFiles	: 'Ilo?? razem:',
	UploadTotalSize		: 'Rozmiar razem:',
	UploadSend			: 'Wy?lij',
	UploadAddFiles		: 'Dodaj pliki',
	UploadClearFiles	: 'Wyczy?? wszystko',
	UploadCancel		: 'Anuluj wysy?anie',
	UploadRemove		: 'Usuń',
	UploadRemoveTip		: 'Usuń !f',
	UploadUploaded		: 'Wys?ano: !n%',
	UploadProcessing	: 'Przetwarzanie...',

	// Settings Panel
	SetTitle		: 'Ustawienia',
	SetView			: 'Widok:',
	SetViewThumb	: 'Miniaturki',
	SetViewList		: 'Lista',
	SetDisplay		: 'Wy?wietlanie:',
	SetDisplayName	: 'Nazwa pliku',
	SetDisplayDate	: 'Data',
	SetDisplaySize	: 'Rozmiar pliku',
	SetSort			: 'Sortowanie:',
	SetSortName		: 'wg nazwy pliku',
	SetSortDate		: 'wg daty',
	SetSortSize		: 'wg rozmiaru',
	SetSortExtension		: 'wg rozszerzenia',

	// Status Bar
	FilesCountEmpty : '<Pusty folder>',
	FilesCountOne	: '1 plik',
	FilesCountMany	: 'Ilo?? plików: %1',

	// Size and Speed
	Kb				: '%1 KB',
	Mb				: '%1 MB',
	Gb				: '%1 GB',
	SizePerSecond	: '%1/s',

	// Connector Error Messages.
	ErrorUnknown	: 'Wykonanie operacji zakończy?o si? niepowodzeniem. (B??d %1)',
	Errors :
	{
	 10 : 'Nieprawid?owe polecenie (command).',
	 11 : 'Brak wymaganego parametru: typ danych (resource type).',
	 12 : 'Nieprawid?owy typ danych (resource type).',
	102 : 'Nieprawid?owa nazwa pliku lub folderu.',
	103 : 'Wykonanie operacji nie jest mo?liwe: brak uprawnień.',
	104 : 'Wykonanie operacji nie powiod?o si? z powodu niewystarczaj?cych uprawnień do systemu plików.',
	105 : 'Nieprawid?owe rozszerzenie.',
	109 : 'Nieprawi?owe ??danie.',
	110 : 'Niezidentyfikowany b??d.',
	111 : 'Wykonanie operacji nie powiod?o si? z powodu zbyt du?ego rozmiaru pliku wynikowego.',
	115 : 'Plik lub folder o podanej nazwie ju? istnieje.',
	116 : 'Nie znaleziono folderu. Od?wie? panel i spróbuj ponownie.',
	117 : 'Nie znaleziono pliku. Od?wie? list? plików i spróbuj ponownie.',
	118 : '?cie?ki ?ród?owa i docelowa s? jednakowe.',
	201 : 'Plik o podanej nazwie ju? istnieje. Nazwa przes?anego pliku zosta?a zmieniona na "%1".',
	202 : 'Nieprawid?owy plik.',
	203 : 'Nieprawid?owy plik. Plik przekracza dozwolony rozmiar.',
	204 : 'Przes?any plik jest uszkodzony.',
	205 : 'Brak folderu tymczasowego na serwerze do przesy?ania plików.',
	206 : 'Przesy?anie pliku zakończy?o si? niepowodzeniem z powodów bezpieczeństwa. Plik zawiera dane przypominaj?ce HTML.',
	207 : 'Nazwa przes?anego pliku zosta?a zmieniona na "%1".',
	300 : 'Przenoszenie nie powiod?o si?.',
	301 : 'Kopiowanie nie powiodo si?.',
	500 : 'Mened?er plików jest wy??czony z powodów bezpieczeństwa. Skontaktuj si? z administratorem oraz sprawd? plik konfiguracyjny CKFindera.',
	501 : 'Tworzenie miniaturek jest wy??czone.'
	},

	// Other Error Messages.
	ErrorMsg :
	{
		FileEmpty		: 'Nazwa pliku nie mo?e by? pusta.',
		FileExists		: 'Plik %s ju? istnieje.',
		FolderEmpty		: 'Nazwa folderu nie mo?e by? pusta.',
		FolderExists	: 'Folder %s ju? istnieje.',
		FolderNameExists	: 'Folder ju? istnieje.',

		FileInvChar		: 'Nazwa pliku nie mo?e zawiera? ?adnego z podanych znaków: \n\\ / : * ? " < > |',
		FolderInvChar	: 'Nazwa folderu nie mo?e zawiera? ?adnego z podanych znaków: \n\\ / : * ? " < > |',

		PopupBlockView	: 'Otwarcie pliku w nowym oknie nie powiod?o si?. Nale?y zmieni? konfiguracj? przegl?darki i wy??czy? wszelkie blokady okienek popup dla tej strony.',
		XmlError		: 'Nie mo?na poprawnie za?adowa? odpowiedzi XML z serwera WWW.',
		XmlEmpty		: 'Nie mo?na za?adowa? odpowiedzi XML z serwera WWW. Serwer zwróci? pust? odpowied?.',
		XmlRawResponse	: 'Odpowied? serwera: %s'
	},

	// Imageresize plugin
	Imageresize :
	{
		dialogTitle		: 'Zmiana rozmiaru %s',
		sizeTooBig		: 'Nie mo?esz zmieni? wysoko?ci lub szeroko?ci na warto?? wi?ksz? od oryginalnego rozmiaru (%size).',
		resizeSuccess	: 'Obrazek zosta? pomy?lnie przeskalowany.',
		thumbnailNew	: 'Utwórz now? miniaturk?',
		thumbnailSmall	: 'Ma?a (%s)',
		thumbnailMedium	: '?rednia (%s)',
		thumbnailLarge	: 'Du?a (%s)',
		newSize			: 'Podaj nowe wymiary',
		width			: 'Szeroko??',
		height			: 'Wysoko??',
		invalidHeight	: 'Nieprawid?owa wysoko??.',
		invalidWidth	: 'Nieprawid?owa szeroko??.',
		invalidName		: 'Nieprawid?owa nazwa pliku.',
		newImage		: 'Utwórz nowy obrazek',
		noExtensionChange : 'Rozszerzenie pliku nie mo?e zostac zmienione.',
		imageSmall		: 'Plik ?ród?owy jest zbyt ma?y.',
		contextMenuName	: 'Zmień rozmiar',
		lockRatio		: 'Zablokuj proporcje',
		resetSize		: 'Przywró? rozmiar'
	},

	// Fileeditor plugin
	Fileeditor :
	{
		save			: 'Zapisz',
		fileOpenError	: 'Nie uda?o si? otworzy? pliku.',
		fileSaveSuccess	: 'Plik zosta? zapisany pomy?lnie.',
		contextMenuName	: 'Edytuj',
		loadingFile		: 'Trwa ?adowanie pliku, prosz? czeka?...'
	},

	Maximize :
	{
		maximize : 'Maksymalizuj',
		minimize : 'Minimalizuj'
	},

	Gallery :
	{
		current : 'Obrazek {current} z {total}'
	},

	Zip :
	{
		extractHereLabel	: 'Wypakuj tutaj',
		extractToLabel		: 'Wypakuj do...',
		downloadZipLabel	: 'Pobierz jako zip',
		compressZipLabel	: 'Kompresuj do zip',
		removeAndExtract	: 'Usuń poprzedni i wypakuj',
		extractAndOverwrite	: 'Wypakuj do bie??cego nadpisuj?c istniej?ce pliki',
		extractSuccess		: 'Plik zosta? pomy?lnie wypakowany.'
	}
};
