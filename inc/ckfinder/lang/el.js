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
 * @fileOverview Defines the {@link CKFinder.lang} object for the Greek
 *		language.
 */

/**
 * Contains the dictionary of language entries.
 * @namespace
 */
CKFinder.lang['el'] =
{
	appTitle : 'CKFinder',

	// Common messages and labels.
	common :
	{
		// Put the voice-only part of the label in the span.
		unavailable		: '%1<span class="cke_accessibility">, μη διαθ?σιμο</span>',
		confirmCancel	: 'Κ?ποιε? απ? τι? επιλογ?? ?χουν αλλ?ξει. Θ?λετε σ?γουρα να κλε?σετε το παρ?θυρο διαλ?γου;',
		ok				: 'OK',
		cancel			: 'Ακ?ρωση',
		confirmationTitle	: 'Επιβεβα?ωση',
		messageTitle	: 'Πληροφορ?ε?',
		inputTitle		: 'Ερ?τηση',
		undo			: 'Ανα?ρεση',
		redo			: 'Επαναφορ?',
		skip			: 'Παρ?βλεψη',
		skipAll			: 'Παρ?βλεψη ?λων',
		makeDecision	: 'Ποια εν?ργεια πρ?πει να ληφθε?;',
		rememberDecision: 'Να θυμ?σαι την απ?φασ? μου'
	},


	// Language direction, 'ltr' or 'rtl'.
	dir : 'ltr',
	HelpLang : 'en',
	LangCode : 'el',

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
	DateAmPm : ['ΜΜ', 'ΠΜ'],

	// Folders
	FoldersTitle	: 'Φ?κελοι',
	FolderLoading	: 'Φ?ρτωση...',
	FolderNew		: 'Παρακαλο?με πληκτρολογ?στε την ονομασ?α του ν?ου φακ?λου: ',
	FolderRename	: 'Παρακαλο?με πληκτρολογ?στε την ν?α ονομασ?α του φακ?λου: ',
	FolderDelete	: 'Ε?στε σ?γουροι ?τι θ?λετε να διαγρ?ψετε το φ?κελο "%1";',
	FolderRenaming	: ' (Μετονομασ?α...)',
	FolderDeleting	: ' (Διαγραφ?...)',
	DestinationFolder	: 'Destination Folder', // MISSING

	// Files
	FileRename		: 'Παρακαλο?με πληκτρολογ?στε την ν?α ονομασ?α του αρχε?ου: ',
	FileRenameExt	: 'Ε?στε σ?γουροι ?τι θ?λετε να αλλ?ξετε την επ?κταση του αρχε?ου; Μετ? απ? αυτ? την εν?ργεια το αρχε?ο ε?ναι δυνατ?ν να μην μπορε? να χρησιμοποιηθε?',
	FileRenaming	: 'Μετονομασ?α...',
	FileDelete		: 'Ε?στε σ?γουροι ?τι θ?λετε να διαγρ?ψετε το αρχε?ο "%1"?',
	FilesDelete	: 'Are you sure you want to delete %1 files?', // MISSING
	FilesLoading	: 'Φ?ρτωση...',
	FilesEmpty		: 'Ο φ?κελο? ε?ναι κεν??.',
	DestinationFile	: 'Destination File', // MISSING
	SkippedFiles	: 'List of skipped files:', // MISSING

	// Basket
	BasketFolder		: 'Καλ?θι',
	BasketClear			: 'Καθαρισμ?? καλαθιο?',
	BasketRemove		: 'Αφα?ρεση απ? το καλ?θι',
	BasketOpenFolder	: '?νοιγμα γονικο? φακ?λου',
	BasketTruncateConfirm : 'Θ?λετε σ?γουρα να αφαιρ?σετε ?λα τα αρχε?α απ? το καλ?θι;',
	BasketRemoveConfirm	: 'Θ?λετε σ?γουρα να αφαιρ?σετε το αρχε?ο "%1" απ? το καλ?θι;',
	BasketRemoveConfirmMultiple	: 'Do you really want to remove %1 files from the basket?', // MISSING
	BasketEmpty			: 'Δεν υπ?ρχουν αρχε?α στο καλ?θι, μεταφ?ρετε κ?ποια με drag and drop.',
	BasketCopyFilesHere	: 'Αντιγραφ? αρχε?ων απ? το καλ?θι',
	BasketMoveFilesHere	: 'Μετακ?νηση αρχε?ων απ? το καλ?θι',

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
	Upload		: 'Μεταφ?ρτωση',
	UploadTip	: 'Μεταφ?ρτωση ν?ου αρχε?ου',
	Refresh		: 'Αναν?ωση',
	Settings	: 'Ρυθμ?σει?',
	Help		: 'Βο?θεια',
	HelpTip		: 'Βο?θεια',

	// Context Menus
	Select			: 'Επιλογ?',
	SelectThumbnail : 'Επιλογ? μικρογραφ?α?',
	View			: 'Προβολ?',
	Download		: 'Λ?ψη αρχε?ου',

	NewSubFolder	: 'Ν?ο? υποφ?κελο?',
	Rename			: 'Μετονομασ?α',
	Delete			: 'Διαγραφ?',
	DeleteFiles		: 'Delete Files', // MISSING

	CopyDragDrop	: 'Αντ?γραψε εδ?',
	MoveDragDrop	: 'Μετακ?νησε εδ?',

	// Dialogs
	RenameDlgTitle		: 'Μετονομασ?α',
	NewNameDlgTitle		: 'Ν?α ονομασ?α',
	FileExistsDlgTitle	: 'Το αρχε?ο υπ?ρχει ?δη',
	SysErrorDlgTitle : 'Σφ?λμα συστ?ματο?',

	FileOverwrite	: 'Αντικατ?σταση αρχε?ου',
	FileAutorename	: 'Αυτ?ματη-μετονομασ?α',
	ManuallyRename	: 'Manually rename', // MISSING

	// Generic
	OkBtn		: 'OK',
	CancelBtn	: 'Ακ?ρωση',
	CloseBtn	: 'Κλε?σιμο',

	// Upload Panel
	UploadTitle			: 'Μεταφ?ρτωση ν?ου αρχε?ου',
	UploadSelectLbl		: 'επιλ?ξτε το αρχε?ο που θ?λετε να μεταφερθε? κ?νοντα? κλ?κ στο κουμπ?',
	UploadProgressLbl	: '(Η μεταφ?ρτωση εκτελε?ται, παρακαλο?με περιμ?νετε...)',
	UploadBtn			: 'Μεταφ?ρτωση επιλεγμ?νου αρχε?ου',
	UploadBtnCancel		: 'Ακ?ρωση',

	UploadNoFileMsg		: 'Παρακαλο?με επιλ?ξτε ?να αρχε?ο απ? τον υπολογιστ? σα?.',
	UploadNoFolder		: 'Παρακαλο?με επιλ?ξτε ?να φ?κελο πριν εκκιν?σετε την διαδικασ?α τη? μεταφ?ρτωση?.',
	UploadNoPerms		: 'Η μεταφ?ρτωση των αρχε?ων δεν επιτρ?πεται.',
	UploadUnknError		: 'Παρουσι?στηκε σφ?λμα κατ? την αποστολ? του αρχε?ου.',
	UploadExtIncorrect	: 'Η επ?κταση του αρχε?ου δεν επιτρ?πεται σε αυτ?ν τον φ?κελο.',

	// Flash Uploads
	UploadLabel			: 'Αρχε?α προ? μεταφ?ρτωση',
	UploadTotalFiles	: 'Συνολικ? αρχε?α:',
	UploadTotalSize		: 'Συνολικ? μ?γεθο?:',
	UploadSend			: 'Μεταφ?ρτωση',
	UploadAddFiles		: 'Προσθ?κη αρχε?ων',
	UploadClearFiles	: 'Αφα?ρεση αρχε?ων',
	UploadCancel		: 'Ακ?ρωση μεταφ?ρτωση?',
	UploadRemove		: 'Αφα?ρεση',
	UploadRemoveTip		: 'Αφα?ρεση !f',
	UploadUploaded		: 'Μεταφορτ?θηκε !n%',
	UploadProcessing	: 'Επεξεργασ?α...',

	// Settings Panel
	SetTitle		: 'Ρυθμ?σει?',
	SetView			: 'Προβολ?:',
	SetViewThumb	: 'Μικρογραφ?ε?',
	SetViewList		: 'Λ?στα',
	SetDisplay		: 'Εμφ?νιση:',
	SetDisplayName	: '?νομα αρχε?ου',
	SetDisplayDate	: 'Ημερομην?α',
	SetDisplaySize	: 'Μ?γεθο? αρχε?ου',
	SetSort			: 'Ταξιν?μηση:',
	SetSortName		: 'β?σει ?νοματο? αρχε?ου',
	SetSortDate		: 'β?σει Ημερομ?νια?',
	SetSortSize		: 'β?σει Μεγ?θου?',
	SetSortExtension		: 'β?σει Επ?κταση?',

	// Status Bar
	FilesCountEmpty : '<Κεν?? Φ?κελο?>',
	FilesCountOne	: '1 αρχε?ο',
	FilesCountMany	: '%1 αρχε?α',

	// Size and Speed
	Kb				: '%1 KB',
	Mb				: '%1 MB',
	Gb				: '%1 GB',
	SizePerSecond	: '%1/s',

	// Connector Error Messages.
	ErrorUnknown	: 'Η εν?ργεια δεν ?ταν δυνατ?ν να εκτελεστε?. (Σφ?λμα %1)',
	Errors :
	{
	 10 : 'Λανθασμ?νη Εντολ?.',
	 11 : 'Το resource type δεν ?ταν δυνατ?ν να προσδιοριστε?.',
	 12 : 'Το resource type δεν ε?ναι ?γκυρο.',
	102 : 'Το ?νομα αρχε?ου ? φακ?λου δεν ε?ναι ?γκυρο.',
	103 : 'Δεν ?ταν δυνατ? η εκτ?λεση τη? εν?ργεια? λ?γω ?λλειψη? δικαιωμ?των ασφαλε?α?.',
	104 : 'Δεν ?ταν δυνατ? η εκτ?λεση τη? εν?ργεια? λ?γω περιορισμ?ν του συστ?ματο? αρχε?ων.',
	105 : 'Λανθασμ?νη επ?κταση αρχε?ου.',
	109 : 'Λανθασμ?νη εν?ργεια.',
	110 : '?γνωστο λ?θο?.',
	111 : 'It was not possible to complete the request due to resulting file size.', // MISSING
	115 : 'Το αρχε?ο ? φ?κελο? υπ?ρχει ?δη.',
	116 : 'Ο φ?κελο? δεν βρ?θηκε. Παρακαλο?με ανανε?στε τη σελ?δα και προσπαθ?στε ξαν?.',
	117 : 'Το αρχε?ο δεν βρ?θηκε. Παρακαλο?με ανανε?στε τη σελ?δα και προσπαθ?στε ξαν?.',
	118 : 'Η αρχικ? και τελικ? διαδρομ? ε?ναι ?διε?.',
	201 : '?να αρχε?ο με την ?δια ονομασ?α υπ?ρχει ?δη. Το μεταφορτωμ?νο αρχε?ο μετονομ?στηκε σε "%1".',
	202 : 'Λανθασμ?νο αρχε?ο.',
	203 : 'Λανθασμ?νο αρχε?ο. Το μ?γεθο? του αρχε?ου ε?ναι πολ? μεγ?λο.',
	204 : 'Το μεταφορτωμ?νο αρχε?ο ε?ναι χαλασμ?νο.',
	205 : 'Δεν υπ?ρχει προσωριν?? φ?κελο? για να χρησιμοποιηθε? για τι? μεταφορτ?σει? των αρχε?ων.',
	206 : 'Η μεταφ?ρτωση ακυρ?θηκε για λ?γου? ασφαλε?α?. Το αρχε?ο περι?χει δεδομ?να μορφ?? HTML.',
	207 : 'Το μεταφορτωμ?νο αρχε?ο μετονομ?στηκε σε "%1".',
	300 : 'Η μετακ?νηση των αρχε?ων απ?τυχε.',
	301 : 'Η αντιγραφ? των αρχε?ων απ?τυχε.',
	500 : 'Ο πλοηγ?? αρχε?ων ?χει απενεργοποιηθε? για λ?γου? ασφαλε?α?. Παρακαλο?με επικοινων?στε με τον διαχειριστ? τη? ιστοσελ?δα? και ελ?γξτε το αρχε?ο ρυθμ?σεων του πλοηγο? (CKFinder).',
	501 : 'Η υποστ?ριξη των μικρογραφι?ν ?χει απενεργοποιηθε?.'
	},

	// Other Error Messages.
	ErrorMsg :
	{
		FileEmpty		: 'Η ονομασ?α του αρχε?ου δεν μπορε? να ε?ναι κεν?.',
		FileExists		: 'Το αρχε?ο %s υπ?ρχει ?δη.',
		FolderEmpty		: 'Η ονομασ?α του φακ?λου δεν μπορε? να ε?ναι κεν?.',
		FolderExists	: 'Folder %s already exists.', // MISSING
		FolderNameExists	: 'Folder already exists.', // MISSING

		FileInvChar		: 'Η ονομασ?α του αρχε?ου δεν μπορε? να περι?χει του? ακ?λουθου? χαρακτ?ρε?: \n\\ / : * ? " < > |',
		FolderInvChar	: 'Η ονομασ?α του φακ?λου δεν μπορε? να περι?χει του? ακ?λουθου? χαρακτ?ρε?: \n\\ / : * ? " < > |',

		PopupBlockView	: 'Δεν ?ταν εφικτ? να ανο?ξει το αρχε?ο σε ν?ο παρ?θυρο. Παρακαλ?, ελ?γξτε τι? ρυθμ?σει? του? πλοηγο? σα? και απενεργοποι?στε ?λου? του? popup blockers για αυτ? την ιστοσελ?δα.',
		XmlError		: 'Δεν ?ταν εφικτ? η σωστ? αν?γνωση του XML response απ? τον διακομιστ?.',
		XmlEmpty		: 'Δεν ?ταν εφικτ? η φ?ρτωση του XML response απ? τον διακομιστ?. Ο διακομιστ?? επ?στρεψε ?να κεν? response.',
		XmlRawResponse	: 'Raw response απ? τον διακομιστ?: %s'
	},

	// Imageresize plugin
	Imageresize :
	{
		dialogTitle		: 'Αλλαγ? διαστ?σεων τη? εικ?να? %s',
		sizeTooBig		: 'Το πλ?το? ? το ?ψο? τη? εικ?να? δεν μπορε? να ε?ναι μεγαλ?τερα των αρχικ?ν διαστ?σεων (%size).',
		resizeSuccess	: 'Οι διαστ?σει? τη? εικ?να? ?λλαξαν επιτυχ??.',
		thumbnailNew	: 'Δημιουργ?α ν?α? μικρογραφ?α?',
		thumbnailSmall	: 'Μικρ? (%s)',
		thumbnailMedium	: 'Μεσα?α (%s)',
		thumbnailLarge	: 'Μεγ?λη (%s)',
		newSize			: 'Ορισμ?? ν?ου μεγ?θου?',
		width			: 'Πλ?το?',
		height			: '?ψο?',
		invalidHeight	: 'Μη ?γκυρο ?ψο?.',
		invalidWidth	: 'Μη ?γκυρο πλ?το?.',
		invalidName		: 'Μη ?γκυρο ?νομα αρχε?ου.',
		newImage		: 'Δημιουργ?α ν?α? εικ?να?',
		noExtensionChange : 'Η επ?κταση του αρχε?ου δεν μπορε? να αλλ?ξει.',
		imageSmall		: 'Η αρχικ? εικ?να ε?ναι πολ? μικρ?.',
		contextMenuName	: 'Αλλαγ? διαστ?σεων',
		lockRatio		: 'Κλε?δωμα αναλογ?α?',
		resetSize		: 'Επαναφορ? αρχικο? μεγ?θου?'
	},

	// Fileeditor plugin
	Fileeditor :
	{
		save			: 'Αποθ?κευση',
		fileOpenError	: 'Δεν ?ταν εφικτ? το ?νοιγμα του αρχε?ου.',
		fileSaveSuccess	: 'Το αρχε?ο αποθηκε?τηκε επιτυχ??.',
		contextMenuName	: 'Επεξεργασ?α',
		loadingFile		: 'Φ?ρτωση αρχε?ου, παρακαλ? περιμ?νετε...'
	},

	Maximize :
	{
		maximize : 'Μεγιστοπο?ηση',
		minimize : 'Ελαχιστοπο?ηση'
	},

	Gallery :
	{
		current : 'Εικ?να {current} απ? {total}'
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
