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
 * @fileOverview Defines the {@link CKFinder.lang} object for the Swedish
 *		language.
*/

/**
 * Contains the dictionary of language entries.
 * @namespace
 */
CKFinder.lang['sv'] =
{
	appTitle : 'CKFinder',

	// Common messages and labels.
	common :
	{
		// Put the voice-only part of the label in the span.
		unavailable		: '%1<span class="cke_accessibility">, Ej tillg?nglig</span>',
		confirmCancel	: 'N?gra av alternativen har ?ndrats. ?r du s?ker p? att du vill st?nga dialogrutan?',
		ok				: 'OK',
		cancel			: 'Avbryt',
		confirmationTitle	: 'Bekr?ftelse',
		messageTitle	: 'Information',
		inputTitle		: 'Fr?ga',
		undo			: '?ngra',
		redo			: 'G?r om',
		skip			: 'Hoppa ?ver',
		skipAll			: 'Hoppa ?ver alla',
		makeDecision	: 'Vilken ?tg?rd ska utf?ras?',
		rememberDecision: 'Kom ih?g mitt val'
	},


	// Language direction, 'ltr' or 'rtl'.
	dir : 'ltr',
	HelpLang : 'en',
	LangCode : 'sv',

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
	FoldersTitle	: 'Mappar',
	FolderLoading	: 'Laddar...',
	FolderNew		: 'Skriv namnet p? den nya mappen: ',
	FolderRename	: 'Skriv det nya namnet p? mappen: ',
	FolderDelete	: '?r du s?ker p? att du vill radera mappen "%1"?',
	FolderRenaming	: ' (Byter mappens namn...)',
	FolderDeleting	: ' (Raderar...)',
	DestinationFolder	: 'Destination Folder', // MISSING

	// Files
	FileRename		: 'Skriv det nya filnamnet: ',
	FileRenameExt	: '?r du s?ker p? att du vill ?ndra fil?ndelsen? Filen kan bli oanv?ndbar.',
	FileRenaming	: 'Byter filnamn...',
	FileDelete		: '?r du s?ker p? att du vill radera filen "%1"?',
	FilesDelete	: 'Are you sure you want to delete %1 files?', // MISSING
	FilesLoading	: 'Laddar...',
	FilesEmpty		: 'Mappen ?r tom.',
	DestinationFile	: 'Destination File', // MISSING
	SkippedFiles	: 'List of skipped files:', // MISSING

	// Basket
	BasketFolder		: 'Filkorg',
	BasketClear			: 'Rensa filkorgen',
	BasketRemove		: 'Ta bort fr?n korgen',
	BasketOpenFolder	: '?ppna ?verliggande mapp',
	BasketTruncateConfirm : 'Vill du verkligen ta bort alla filer fr?n korgen?',
	BasketRemoveConfirm	: 'Vill du verkligen ta bort filen "%1" fr?n korgen?',
	BasketRemoveConfirmMultiple	: 'Do you really want to remove %1 files from the basket?', // MISSING
	BasketEmpty			: 'Inga filer i korgen, dra och sl?pp n?gra.',
	BasketCopyFilesHere	: 'Kopiera filer fr?n korgen',
	BasketMoveFilesHere	: 'Flytta filer fr?n korgen',

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
	Upload		: 'Ladda upp',
	UploadTip	: 'Ladda upp en ny fil',
	Refresh		: 'Uppdatera',
	Settings	: 'Inst?llningar',
	Help		: 'Hj?lp',
	HelpTip		: 'Hj?lp',

	// Context Menus
	Select			: 'Infoga bild',
	SelectThumbnail : 'Infoga som tumnagel',
	View			: 'Visa',
	Download		: 'Ladda ner',

	NewSubFolder	: 'Ny Undermapp',
	Rename			: 'Byt namn',
	Delete			: 'Radera',
	DeleteFiles		: 'Delete Files', // MISSING

	CopyDragDrop	: 'Kopiera hit',
	MoveDragDrop	: 'Flytta hit',

	// Dialogs
	RenameDlgTitle		: 'Byt namn',
	NewNameDlgTitle		: 'Nytt namn',
	FileExistsDlgTitle	: 'Filen finns redan',
	SysErrorDlgTitle : 'Systemfel',

	FileOverwrite	: 'Skriv ?ver',
	FileAutorename	: 'Auto-namn?ndring',
	ManuallyRename	: 'Manually rename', // MISSING

	// Generic
	OkBtn		: 'OK',
	CancelBtn	: 'Avbryt',
	CloseBtn	: 'St?ng',

	// Upload Panel
	UploadTitle			: 'Ladda upp en ny fil',
	UploadSelectLbl		: 'V?lj fil att ladda upp',
	UploadProgressLbl	: '(Laddar upp filen, var god v?nta...)',
	UploadBtn			: 'Ladda upp den valda filen',
	UploadBtnCancel		: 'Avbryt',

	UploadNoFileMsg		: 'V?lj en fil fr?n din dator.',
	UploadNoFolder		: 'V?lj en mapp f?re uppladdning.',
	UploadNoPerms		: 'Filuppladdning ej till?ten.',
	UploadUnknError		: 'Fel vid filuppladdning.',
	UploadExtIncorrect	: 'Fil?ndelsen ?r inte till?ten i denna mapp.',

	// Flash Uploads
	UploadLabel			: 'Filer att ladda upp',
	UploadTotalFiles	: 'Totalt antal filer:',
	UploadTotalSize		: 'Total storlek:',
	UploadSend			: 'Ladda upp',
	UploadAddFiles		: 'L?gg till filer',
	UploadClearFiles	: 'Rensa filer',
	UploadCancel		: 'Avbryt uppladdning',
	UploadRemove		: 'Ta bort',
	UploadRemoveTip		: 'Ta bort !f',
	UploadUploaded		: 'Uppladdat !n%',
	UploadProcessing	: 'Bearbetar...',

	// Settings Panel
	SetTitle		: 'Inst?llningar',
	SetView			: 'Visa:',
	SetViewThumb	: 'Tumnaglar',
	SetViewList		: 'Lista',
	SetDisplay		: 'Visa:',
	SetDisplayName	: 'Filnamn',
	SetDisplayDate	: 'Datum',
	SetDisplaySize	: 'Storlek',
	SetSort			: 'Sortering:',
	SetSortName		: 'Filnamn',
	SetSortDate		: 'Datum',
	SetSortSize		: 'Storlek',
	SetSortExtension		: 'Fil?ndelse',

	// Status Bar
	FilesCountEmpty : '<Tom Mapp>',
	FilesCountOne	: '1 fil',
	FilesCountMany	: '%1 filer',

	// Size and Speed
	Kb				: '%1 KB',
	Mb				: '%1 MB',
	Gb				: '%1 GB',
	SizePerSecond	: '%1/s',

	// Connector Error Messages.
	ErrorUnknown	: 'Beg?ran kunde inte utf?ras eftersom ett fel uppstod. (Fel %1)',
	Errors :
	{
	 10 : 'Ogiltig beg?ran.',
	 11 : 'Resursens typ var inte specificerad i f?rfr?gan.',
	 12 : 'Den efterfr?gade resurstypen ?r inte giltig.',
	102 : 'Ogiltigt fil- eller mappnamn.',
	103 : 'Beg?ran kunde inte utf?ras p.g.a. restriktioner av r?ttigheterna.',
	104 : 'Beg?ran kunde inte utf?ras p.g.a. restriktioner av r?ttigheter i filsystemet.',
	105 : 'Ogiltig fil?ndelse.',
	109 : 'Ogiltig beg?ran.',
	110 : 'Ok?nt fel.',
	111 : 'It was not possible to complete the request due to resulting file size.', // MISSING
	115 : 'En fil eller mapp med aktuellt namn finns redan.',
	116 : 'Mappen kunde inte hittas. Var god uppdatera sidan och f?rs?k igen.',
	117 : 'Filen kunde inte hittas. Var god uppdatera sidan och f?rs?k igen.',
	118 : 'S?kv?g till k?lla och m?l ?r identisk.',
	201 : 'En fil med aktuellt namn fanns redan. Den uppladdade filen har d?pts om till "%1".',
	202 : 'Ogiltig fil.',
	203 : 'Ogiltig fil. Filen var f?r stor.',
	204 : 'Den uppladdade filen var korrupt.',
	205 : 'En tillf?llig mapp f?r uppladdning ?r inte tillg?nglig p? servern.',
	206 : 'Uppladdningen stoppades av s?kerhetssk?l. Filen inneh?ller HTML-liknande data.',
	207 : 'Den uppladdade filen har d?pts om till "%1".',
	300 : 'Flytt av fil(er) misslyckades.',
	301 : 'Kopiering av fil(er) misslyckades.',
	500 : 'Filhanteraren har stoppats av s?kerhetssk?l. Var god kontakta administrat?ren f?r att kontrollera konfigurationsfilen f?r CKFinder.',
	501 : 'St?d f?r tumnaglar har st?ngts av.'
	},

	// Other Error Messages.
	ErrorMsg :
	{
		FileEmpty		: 'Filnamnet f?r inte vara tomt.',
		FileExists		: 'Filen %s finns redan.',
		FolderEmpty		: 'Mappens namn f?r inte vara tomt.',
		FolderExists	: 'Folder %s already exists.', // MISSING
		FolderNameExists	: 'Folder already exists.', // MISSING

		FileInvChar		: 'Filnamnet f?r inte inneh?lla n?got av f?ljande tecken: \n\\ / : * ? " < > |',
		FolderInvChar	: 'Mappens namn f?r inte inneh?lla n?got av f?ljande tecken: \n\\ / : * ? " < > |',

		PopupBlockView	: 'Det gick inte att ?ppna filen i ett nytt f?nster. ?ndra inst?llningarna i din webbl?sare s? att den till?ter popup-f?nster p? den h?r webbplatsen.',
		XmlError		: 'Det gick inte att ladda XML-svaret fr?n webbservern ordentligt.',
		XmlEmpty		: 'Det gick inte att ladda XML-svaret fr?n webbservern. Servern returnerade ett tomt svar.',
		XmlRawResponse	: 'Svar fr?n servern: %s'
	},

	// Imageresize plugin
	Imageresize :
	{
		dialogTitle		: 'Storleks?ndra %s',
		sizeTooBig		: 'Bildens h?jd eller bredd kan inte vara st?rre ?n originalfilens storlek (%size).',
		resizeSuccess	: 'Storleks?ndring lyckades.',
		thumbnailNew	: 'Skapa en ny tumnagel',
		thumbnailSmall	: 'Liten (%s)',
		thumbnailMedium	: 'Mellan (%s)',
		thumbnailLarge	: 'Stor (%s)',
		newSize			: 'V?lj en ny storlek',
		width			: 'Bredd',
		height			: 'H?jd',
		invalidHeight	: 'Ogiltig h?jd.',
		invalidWidth	: 'Ogiltig bredd.',
		invalidName		: 'Ogiltigt filnamn.',
		newImage		: 'Skapa en ny bild',
		noExtensionChange : 'Fil?ndelsen kan inte ?ndras.',
		imageSmall		: 'Originalbilden ?r f?r liten.',
		contextMenuName	: '?ndra storlek',
		lockRatio		: 'L?s h?jd/bredd f?rh?llanden',
		resetSize		: '?terst?ll storlek'
	},

	// Fileeditor plugin
	Fileeditor :
	{
		save			: 'Spara',
		fileOpenError	: 'Kan inte ?ppna filen.',
		fileSaveSuccess	: 'Filen sparades.',
		contextMenuName	: 'Redigera',
		loadingFile		: 'Laddar fil, var god v?nta...'
	},

	Maximize :
	{
		maximize : 'Maximera',
		minimize : 'Minimera'
	},

	Gallery :
	{
		current : 'Bild {current} av {total}'
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
