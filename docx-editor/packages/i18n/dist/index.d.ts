/**
 * Shared locale data, types, and runtime helpers for the @npde
 * docx-editor adapters.
 *
 * Import everything from the package root. `sideEffects: false` lets
 * consumer bundlers tree-shake unused locales.
 *
 * ```ts
 * import {
 *   en, de, pl, tr, he, ptBR, zhCN,    // typed locale data
 *   locales,                            // record keyed by BCP-47 tag
 *   deepMerge, createT,                 // build a t() for custom hosts
 *   type LocaleStrings,                 // shape of `en` (source of truth)
 *   type Translations,                  // shape of a community partial
 *   type TranslationKey,                // every valid `t()` key
 *   type LocaleCode,                    // 'en' | 'de' | 'pt-BR' | ...
 * } from '@npde/docx-editor-i18n';
 * ```
 *
 * The React and Vue adapters wrap `createT` in framework-native bindings
 * (`useTranslation`, `LocaleProvider`, etc.); use those for app code.
 * Reach for `createT` directly when building a non-React/Vue host.
 *
 * @packageDocumentation
 * @public
 */
var _lang = "en";
var common = {
	cancel: "Cancel",
	insert: "Insert",
	apply: "Apply",
	close: "Close",
	"delete": "Delete",
	update: "Update",
	save: "Save",
	retry: "Retry",
	send: "Send",
	edit: "Edit",
	comment: "Comment",
	reply: "Reply",
	accept: "Accept",
	reject: "Reject",
	dismiss: "Dismiss",
	change: "Change",
	clear: "Clear",
	px: "px",
	closeDialog: "Close dialog"
};
var toolbar = {
	ariaLabel: "Formatting toolbar",
	file: "File",
	format: "Format",
	insert: "Insert",
	open: "Open",
	openShortcut: "Ctrl+O",
	save: "Save",
	saveShortcut: "Ctrl+S",
	print: "Print",
	printShortcut: "Ctrl+P",
	pageSetup: "Page setup",
	leftToRight: "Left-to-right text",
	rightToLeft: "Right-to-left text",
	image: "Image",
	table: "Table",
	"break": "Break",
	pageBreak: "Page break",
	sectionBreakNextPage: "Section break (next page)",
	sectionBreakContinuous: "Section break (continuous)",
	tableOfContents: "Table of contents",
	symbol: "Symbol",
	watermark: "Watermark",
	help: "Help",
	reportIssue: "Report issue"
};
var formattingBar = {
	groups: {
		history: "History",
		zoom: "Zoom",
		styles: "Styles",
		font: "Font",
		textFormatting: "Text formatting",
		script: "Script",
		alignment: "Alignment",
		listFormatting: "List formatting",
		image: "Image",
		table: "Table"
	},
	undo: "Undo",
	undoShortcut: "Undo (Ctrl+Z)",
	redo: "Redo",
	redoShortcut: "Redo (Ctrl+Y)",
	bold: "Bold",
	boldShortcut: "Bold (Ctrl+B)",
	italic: "Italic",
	italicShortcut: "Italic (Ctrl+I)",
	underline: "Underline",
	underlineShortcut: "Underline (Ctrl+U)",
	strikethrough: "Strikethrough",
	fontColor: "Font Color",
	highlightColor: "Text Highlight Color",
	insertLink: "Insert link",
	insertLinkShortcut: "Insert link (Ctrl+K)",
	superscript: "Superscript",
	superscriptShortcut: "Superscript (Ctrl+Shift+=)",
	subscript: "Subscript",
	subscriptShortcut: "Subscript (Ctrl+=)",
	imageProperties: "Image properties",
	imagePropertiesShortcut: "Image properties (alt text, border)...",
	clearFormatting: "Clear formatting",
	commentsAndChanges: "Comments & Changes",
	moreItems: "{count, plural, one {# more item} other {# more items}}"
};
var alignment = {
	alignLeft: "Align Left",
	alignLeftShortcut: "Ctrl+L",
	center: "Center",
	centerShortcut: "Ctrl+E",
	alignRight: "Align Right",
	alignRightShortcut: "Ctrl+R",
	justify: "Justify",
	justifyShortcut: "Ctrl+J"
};
var lists = {
	ariaLabel: "List formatting",
	typeAriaLabel: "List type",
	indentationAriaLabel: "List indentation",
	bulletList: "Bullet List",
	numberedList: "Numbered List",
	decreaseIndent: "Decrease Indent",
	increaseIndent: "Increase Indent"
};
var lineSpacing = {
	label: "Line spacing",
	single: "Single",
	double: "Double",
	lineSpacingTitle: "Line spacing: {label}",
	paragraphSpacing: "Paragraph spacing"
};
var styles = {
	selectAriaLabel: "Select paragraph style",
	normalText: "Normal text",
	title: "Title",
	subtitle: "Subtitle",
	heading1: "Heading 1",
	heading2: "Heading 2",
	heading3: "Heading 3"
};
var font = {
	selectAriaLabel: "Select font family",
	sansSerif: "Sans Serif",
	serif: "Serif",
	monospace: "Monospace",
	other: "Other",
	documentFonts: "Document fonts"
};
var fontSize = {
	decrease: "Decrease font size",
	increase: "Increase font size",
	label: "Font size",
	listLabel: "Font sizes"
};
var zoom = {
	ariaLabel: "Zoom: {label}",
	zoomOut: "Zoom out",
	zoomIn: "Zoom in",
	zoomLevel: "Zoom level"
};
var colorPicker = {
	ariaLabel: "{type} color picker",
	highlightColors: "Highlight Colors",
	customColor: "Custom Color",
	noColor: "No Color",
	automatic: "Automatic",
	themeColors: "Theme Colors",
	standardColors: "Standard Colors",
	colors: {
		black: "Black",
		darkRed: "Dark Red",
		red: "Red",
		orange: "Orange",
		yellow: "Yellow",
		darkYellow: "Dark Yellow",
		green: "Green",
		darkGreen: "Dark Green",
		teal: "Teal",
		darkTeal: "Dark Teal",
		blue: "Blue",
		darkBlue: "Dark Blue",
		purple: "Purple",
		darkPurple: "Dark Purple",
		brown: "Brown",
		grey50: "Grey 50%",
		grey25: "Grey 25%",
		grey10: "Grey 10%",
		white: "White",
		lightRed: "Light Red",
		lightOrange: "Light Orange",
		lightYellow: "Light Yellow",
		lightGreen: "Light Green",
		lightBlue: "Light Blue",
		lightPurple: "Light Purple",
		pink: "Pink",
		rose: "Rose",
		gold: "Gold",
		aqua: "Aqua",
		lavender: "Lavender",
		silver: "Silver",
		darkOrange: "Dark Orange",
		darkGray: "Dark Gray",
		gray: "Gray",
		cyan: "Cyan",
		magenta: "Magenta",
		paleGreen: "Pale Green",
		lightCyan: "Light Cyan",
		skyBlue: "Sky Blue",
		lightBlue2: "Light Blue 2",
		lightMagenta: "Light Magenta",
		brightGreen: "Bright Green",
		violet: "Violet"
	}
};
var dialogs = {
	findReplace: {
		titleFind: "Find",
		titleFindReplace: "Find and Replace",
		findLabel: "Find:",
		findPlaceholder: "Enter text to find...",
		findAriaLabel: "Find text",
		findPrevious: "Find previous",
		findPreviousTitle: "Find Previous (Shift+Enter)",
		findNext: "Find next",
		findNextTitle: "Find Next (Enter)",
		matchCount: "{current} of {total} matches",
		noResults: "No results found",
		replaceLabel: "Replace:",
		replacePlaceholder: "Enter replacement text...",
		replaceAriaLabel: "Replace text",
		replaceButton: "Replace",
		replaceCurrentTitle: "Replace current match",
		replaceAllButton: "Replace All",
		replaceAllTitle: "Replace all matches",
		matchCase: "Match case",
		wholeWords: "Whole words",
		matchesFound: "{total} matches",
		toggleReplace: "+ Replace"
	},
	hyperlink: {
		titleEdit: "Edit Hyperlink",
		titleInsert: "Insert Hyperlink",
		tabWebAddress: "Web Address",
		tabBookmark: "Bookmark",
		urlLabel: "URL",
		urlPlaceholder: "https://example.com",
		urlHint: "Enter a web address, email (mailto:), or phone (tel:)",
		bookmarkLabel: "Bookmark",
		bookmarkPlaceholder: "Select a bookmark...",
		displayTextLabel: "Display Text",
		displayTextPlaceholder: "Text to display (optional)",
		displayTextHint: "Leave empty to use the selected text",
		tooltipLabel: "Tooltip (optional)",
		tooltipPlaceholder: "Text shown on hover",
		removeLink: "Remove Link",
		invalidUrl: "Please enter a valid URL",
		urlRequired: "URL is required"
	},
	insertTable: {
		title: "Insert Table",
		hoverToSelect: "Hover to select size",
		tableSize: "{cols} x {rows} Table",
		orSpecifySize: "or specify size",
		rowsLabel: "Rows:",
		columnsLabel: "Columns:",
		insertButton: "Insert Table",
		sizeSelector: "Table size selector",
		columnWidthLabel: "Column width",
		fixedWidth: "Fixed",
		autofit: "Auto-fit to contents",
		tableStyleLabel: "Table style",
		validationHint: "Rows: {minRows}-{maxRows}, Columns: {minCols}-{maxCols}"
	},
	watermark: {
		title: "Watermark",
		noWatermark: "No watermark",
		picture: "Picture watermark",
		text: "Text watermark",
		selectPicture: "Select Picture...",
		scale: "Scale",
		washout: "Washout",
		textLabel: "Text",
		fontLabel: "Font",
		sizeLabel: "Size",
		sizeAuto: "Auto",
		colorLabel: "Color",
		layoutLabel: "Layout",
		diagonal: "Diagonal",
		horizontal: "Horizontal",
		semitransparent: "Semitransparent",
		presetLabel: "Preset",
		applyButton: "Apply",
		cancelButton: "Cancel"
	},
	splitCell: {
		title: "Split Cell",
		description: "Set how many rows and columns to split the selected cell into.",
		rowsLabel: "Rows:",
		columnsLabel: "Columns:",
		currentMinimum: "Minimum from current span: {rows} row(s) x {cols} column(s)",
		minValue: "Use at least {rows} row(s) and {cols} column(s).",
		notOneByOne: "Choose at least two resulting cells."
	},
	insertImage: {
		title: "Insert Image",
		uploadAriaLabel: "Click or drag to upload image",
		uploadText: "Click to select or drag and drop an image",
		uploadSubtext: "PNG, JPG, GIF up to 10MB",
		dimensions: "Dimensions",
		widthLabel: "Width:",
		heightLabel: "Height:",
		aspectRatioLocked: "Aspect ratio locked",
		aspectRatioUnlocked: "Aspect ratio unlocked",
		altTextLabel: "Alt Text (optional)",
		altTextPlaceholder: "Describe the image for accessibility",
		insertButton: "Insert Image",
		invalidFile: "Please select a valid image file",
		fileTooLarge: "Image file is too large (max 10MB)",
		readFailed: "Failed to read image file",
		loadFailed: "Failed to load image",
		preview: "Preview"
	},
	insertSymbol: {
		title: "Insert Symbol",
		searchPlaceholder: "Search symbols (character or Unicode)...",
		noResultsEmpty: "No symbols found",
		noResults: "No symbols found for \"{query}\"",
		decimal: "Decimal: {value}",
		recent: "Recent:",
		categories: {
			common: "Common",
			arrows: "Arrows",
			math: "Math",
			greek: "Greek",
			shapes: "Shapes",
			punctuation: "Punctuation",
			currency: "Currency",
			music: "Music",
			emoji: "Emoji"
		}
	},
	imageProperties: {
		title: "Image Properties",
		altText: "Alt Text",
		altTextPlaceholder: "Describe this image for accessibility...",
		border: "Border",
		width: "Width",
		style: "Style",
		color: "Color",
		preview: "Preview",
		textWrapping: "Text wrapping",
		dimensions: "Dimensions",
		widthLabel: "Width:",
		heightLabel: "Height:",
		lockAspectRatio: "Lock aspect ratio",
		wrapOptions: {
			inline: "Inline with text",
			wrapRight: "Wrap text, float left",
			wrapLeft: "Wrap text, float right",
			topAndBottom: "Top and bottom",
			behind: "Behind text",
			inFront: "In front of text"
		},
		borderStyles: {
			solid: "Solid",
			dashed: "Dashed",
			dotted: "Dotted",
			double: "Double",
			groove: "Groove",
			ridge: "Ridge",
			inset: "Inset",
			outset: "Outset"
		}
	},
	imagePosition: {
		title: "Image Position",
		horizontal: "Horizontal",
		vertical: "Vertical",
		position: "Position",
		alignment: "Alignment",
		offset: "Offset",
		offsetPx: "Offset (px)",
		align: "Align",
		relativeTo: "Relative to",
		alignOptions: {
			left: "Left",
			center: "Center",
			right: "Right",
			top: "Top",
			bottom: "Bottom"
		},
		relativeOptions: {
			page: "Page",
			column: "Column",
			margin: "Margin",
			character: "Character",
			paragraph: "Paragraph",
			line: "Line"
		}
	},
	pageSetup: {
		title: "Page Setup",
		pageSize: "PAGE SIZE",
		sizeLabel: "Size",
		custom: "Custom",
		orientation: "Orientation",
		portrait: "Portrait",
		landscape: "Landscape",
		margins: "MARGINS",
		top: "Top",
		bottom: "Bottom",
		left: "Left",
		right: "Right",
		pageSizes: {
			letter: "Letter (8.5\" × 11\")",
			a4: "A4 (8.27\" × 11.69\")",
			legal: "Legal (8.5\" × 14\")",
			a3: "A3 (11.69\" × 16.54\")",
			a5: "A5 (5.83\" × 8.27\")",
			b5: "B5 (6.93\" × 9.84\")",
			executive: "Executive (7.25\" × 10.5\")"
		}
	},
	tableProperties: {
		title: "Table Properties",
		widthType: "Width type",
		widthLabel: "Width",
		alignmentLabel: "Alignment",
		widthTypes: {
			auto: "Auto",
			fixed: "Fixed (twips)",
			percentage: "Percentage"
		},
		units: {
			fiftiethsPercent: "(50ths of %)",
			twips: "tw"
		},
		alignOptions: {
			left: "Left",
			center: "Center",
			right: "Right"
		}
	},
	pasteSpecial: {
		title: "Paste Special",
		keepFormatting: "Keep Source Formatting",
		keepFormattingDescription: "Paste with original formatting",
		keepFormattingShortcut: "Ctrl+V",
		plainText: "Paste as Plain Text",
		plainTextDescription: "Paste without any formatting",
		plainTextShortcut: "Ctrl+Shift+V",
		readingClipboard: "Reading clipboard...",
		preview: "Preview:",
		noContent: "No content available to paste",
		clipboardError: "Unable to read clipboard. Please use Ctrl+V to paste."
	},
	footnoteProperties: {
		title: "Footnote & Endnote Properties",
		footnotes: "Footnotes",
		endnotes: "Endnotes",
		position: "Position",
		numberFormat: "Number format",
		startAt: "Start at",
		numbering: "Numbering",
		footnotePositions: {
			bottomOfPage: "Bottom of page",
			belowText: "Below text"
		},
		endnotePositions: {
			endOfDocument: "End of document",
			endOfSection: "End of section"
		},
		numberingOptions: {
			continuous: "Continuous",
			restartSection: "Restart each section",
			restartPage: "Restart each page"
		},
		formats: {
			decimal: "1, 2, 3, ...",
			lowerRoman: "i, ii, iii, ...",
			upperRoman: "I, II, III, ...",
			lowerAlpha: "a, b, c, ...",
			upperAlpha: "A, B, C, ...",
			symbols: "*, †, ‡, ..."
		}
	},
	keyboardShortcuts: {
		ariaLabel: "Keyboard Shortcuts",
		searchPlaceholder: "Search shortcuts...",
		categories: {
			editing: "Editing",
			formatting: "Formatting",
			navigation: "Navigation",
			clipboard: "Clipboard",
			selection: "Selection",
			view: "View",
			file: "File",
			other: "Other"
		},
		shortcuts: {
			save: "Save",
			saveDescription: "Save document",
			print: "Print",
			printDescription: "Print document",
			undo: "Undo",
			undoDescription: "Undo last action",
			redo: "Redo",
			redoDescription: "Redo last action",
			"delete": "Delete",
			deleteDescription: "Delete selected text",
			find: "Find",
			findDescription: "Find text in document",
			findReplace: "Find & Replace",
			findReplaceDescription: "Find and replace text",
			cut: "Cut",
			cutDescription: "Cut selected text",
			copy: "Copy",
			copyDescription: "Copy selected text",
			paste: "Paste",
			pasteDescription: "Paste from clipboard",
			pastePlainText: "Paste as Plain Text",
			pastePlainTextDescription: "Paste without formatting",
			bold: "Bold",
			boldDescription: "Toggle bold formatting",
			italic: "Italic",
			italicDescription: "Toggle italic formatting",
			underline: "Underline",
			underlineDescription: "Toggle underline formatting",
			strikethrough: "Strikethrough",
			strikethroughDescription: "Toggle strikethrough",
			subscript: "Subscript",
			subscriptDescription: "Toggle subscript",
			superscript: "Superscript",
			superscriptDescription: "Toggle superscript",
			alignLeft: "Align Left",
			alignLeftDescription: "Left align paragraph",
			alignCenter: "Align Center",
			alignCenterDescription: "Center align paragraph",
			alignRight: "Align Right",
			alignRightDescription: "Right align paragraph",
			justify: "Justify",
			justifyDescription: "Justify paragraph",
			increaseIndent: "Increase Indent",
			increaseIndentDescription: "Increase paragraph indent",
			decreaseIndent: "Decrease Indent",
			decreaseIndentDescription: "Decrease paragraph indent",
			selectAll: "Select All",
			selectAllDescription: "Select all content",
			selectWord: "Select Word",
			selectWordDescription: "Select current word",
			selectParagraph: "Select Paragraph",
			selectParagraphDescription: "Select current paragraph",
			extendSelectionByWord: "Extend Selection by Word",
			extendSelectionByWordDescription: "Extend selection to next/previous word",
			extendSelectionToLineEdge: "Extend Selection to Line Edge",
			extendSelectionToLineEdgeDescription: "Extend selection to line start/end",
			moveByWord: "Move by Word",
			moveByWordDescription: "Move cursor to next/previous word",
			moveToLineStart: "Move to Line Start",
			moveToLineStartDescription: "Move cursor to start of line",
			moveToLineEnd: "Move to Line End",
			moveToLineEndDescription: "Move cursor to end of line",
			moveToDocumentStart: "Move to Document Start",
			moveToDocumentStartDescription: "Move cursor to start of document",
			moveToDocumentEnd: "Move to Document End",
			moveToDocumentEndDescription: "Move cursor to end of document",
			pageUp: "Page Up",
			pageUpDescription: "Scroll up one page",
			pageDown: "Page Down",
			pageDownDescription: "Scroll down one page",
			zoomIn: "Zoom In",
			zoomInDescription: "Increase zoom level",
			zoomOut: "Zoom Out",
			zoomOutDescription: "Decrease zoom level",
			resetZoom: "Reset Zoom",
			resetZoomDescription: "Reset zoom to 100%",
			keyboardShortcuts: "Keyboard Shortcuts",
			keyboardShortcutsDescription: "Show this help dialog",
			insertLink: "Insert Link",
			insertLinkDescription: "Insert or edit hyperlink"
		},
		noResults: "No shortcuts found matching \"{query}\"",
		pressEscToClose: "Press {key} to close",
		or: "or"
	}
};
var comments = {
	resolved: "Resolved",
	resolve: "Resolve",
	reopen: "Reopen",
	moreOptions: "More options",
	unknown: "Unknown",
	addComment: "Add a comment...",
	replyPlaceholder: "Reply or add others with @",
	replyCount: "{count, plural, one {# reply} other {# replies}}"
};
var trackedChanges = {
	unknown: "Unknown",
	replaced: "Replaced",
	"with": "with",
	added: "Added",
	deleted: "Deleted"
};
var revisions = {
	paragraphMarkInserted: "Inserted paragraph break",
	paragraphMarkDeleted: "Deleted paragraph break",
	paragraphPropertiesChanged: "Changed paragraph properties",
	rowInserted: "Inserted row",
	rowDeleted: "Deleted row",
	rowPropertiesChanged: "Changed row properties",
	cellInserted: "Inserted cell",
	cellDeleted: "Deleted cell",
	cellMerged: "Merged cells",
	cellPropertiesChanged: "Changed cell properties",
	tablePropertiesChanged: "Changed table properties",
	tableInserted: "Inserted table",
	tableDeleted: "Deleted table"
};
var contextMenu = {
	ariaLabel: "AI actions menu",
	textMenuAriaLabel: "Text editing menu",
	customPromptPlaceholder: "Enter custom prompt...",
	cut: "Cut",
	cutShortcut: "Ctrl+X",
	copy: "Copy",
	copyShortcut: "Ctrl+C",
	paste: "Paste",
	pasteShortcut: "Ctrl+V",
	pastePlainText: "Paste as Plain Text",
	pastePlainTextShortcut: "Ctrl+Shift+V",
	"delete": "Delete",
	deleteShortcut: "Del",
	selectAll: "Select All",
	selectAllShortcut: "Ctrl+A",
	selected: "Selected:",
	aiActions: {
		askAi: "Ask AI",
		rewrite: "Rewrite",
		expand: "Expand",
		summarize: "Summarize",
		translate: "Translate",
		explain: "Explain",
		fixGrammar: "Fix Grammar",
		makeFormal: "Make Formal",
		makeCasual: "Make Casual",
		custom: "Custom"
	}
};
var documentOutline = {
	ariaLabel: "Document outline",
	closeAriaLabel: "Close outline",
	closeTitle: "Close outline",
	title: "Outline",
	noHeadings: "No headings found. Add headings to your document to see them here."
};
var sidebar = {
	ariaLabel: "Annotations sidebar"
};
var viewer = {
	pageIndicator: "{current} of {total}"
};
var titleBar = {
	untitled: "Untitled",
	documentNameAriaLabel: "Document name",
	menuBarAriaLabel: "Menu bar"
};
var errors = {
	loadingDocument: "Loading document...",
	noDocumentLoaded: "No document loaded",
	failedToLoad: "Failed to Load Document",
	unableToParse: "Unable to Parse Document",
	somethingWentWrong: "Something went wrong",
	errorDescription: "An error occurred while rendering this component. Please try again or contact support if the problem persists.",
	errorLabel: "Error:",
	componentStack: "Component Stack:",
	tryAgain: "Try Again",
	showDetails: "Show details",
	hideDetails: "Hide details",
	unsavedChanges: "You have unsaved changes. Are you sure you want to leave?"
};
var table = {
	insertRowAbove: "Insert row above",
	insertRowBelow: "Insert row below",
	insertColumnLeft: "Insert column left",
	insertColumnRight: "Insert column right",
	deleteRow: "Delete row",
	deleteColumn: "Delete column",
	deleteTable: "Delete table",
	selectTable: "Select entire table",
	mergeCells: "Merge cells",
	splitCell: "Split cell",
	editingTools: "Table editing tools",
	label: "Table:",
	cellFillColor: "Cell Fill Color",
	borderColor: "Border Color",
	borderWidth: "Border width",
	unknownAction: "Unknown Action",
	borders: {
		all: "All borders",
		outside: "Outside borders",
		inside: "Inside borders",
		none: "No borders",
		remove: "Remove borders",
		top: "Top border",
		bottom: "Bottom border",
		left: "Left border",
		right: "Right border",
		styleAriaLabel: "Border style",
		tooltip: "Borders"
	},
	moreOptions: "More table options",
	styles: {
		title: "Table Styles",
		label: "Styles",
		normalTable: "Normal Table",
		tableGrid: "Table Grid",
		gridTableLight: "Grid Table Light",
		plainTable1: "Plain Table 1",
		plainTable2: "Plain Table 2",
		plainTable3: "Plain Table 3",
		plainTable4: "Plain Table 4",
		gridTable1Light: "Grid Table 1 Light",
		gridTable4Accent1: "Grid Table 4 Accent 1",
		gridTable5Dark: "Grid Table 5 Dark",
		listTable3Accent2: "List Table 3 Accent 2",
		listTable4Accent3: "List Table 4 Accent 3",
		gridTable4Accent5: "Grid Table 4 Accent 5",
		gridTable4Accent6: "Grid Table 4 Accent 6"
	}
};
var tableAdvanced = {
	verticalAlignment: "Vertical alignment",
	top: "Top",
	middle: "Middle",
	bottom: "Bottom",
	cellMargins: "Cell margins",
	textDirection: "Text direction",
	textDirections: {
		horizontal: "Horizontal (LR)",
		verticalRL: "Vertical (top-bottom, RL)",
		verticalLR: "Vertical (bottom-top, LR)"
	},
	toggleNoWrap: "Toggle no-wrap",
	rowHeight: "Row height",
	heightRules: {
		auto: "Auto",
		atLeast: "At least",
		exact: "Exact"
	},
	rule: "Rule",
	height: "Height",
	toggleHeaderRow: "Toggle header row",
	distributeColumns: "Distribute columns evenly",
	autoFit: "Auto-fit to contents",
	tableProperties: "Table properties...",
	tableAlignment: "Table alignment",
	alignTableLeft: "Align table left",
	alignTableCenter: "Align table center",
	alignTableRight: "Align table right",
	tableOptionsMenu: "Table options menu",
	tableOptions: "Table options"
};
var imageTransform = {
	tooltip: "Transform",
	rotateClockwise: "Rotate clockwise",
	rotateCounterClockwise: "Rotate counter-clockwise",
	flipHorizontal: "Flip horizontal",
	flipVertical: "Flip vertical"
};
var imageWrap = {
	inline: "In Line with Text",
	floatLeft: "Square Left",
	floatRight: "Square Right",
	topAndBottom: "Top and Bottom",
	behindText: "Behind Text",
	inFrontOfText: "In Front of Text",
	tooltipPrefix: "Wrap: {label}",
	menu: {
		inLineWithText: "In Line with Text",
		squareLeft: "Square Left",
		squareRight: "Square Right",
		behindText: "Behind Text",
		inFrontOfText: "In Front of Text",
		ariaLabel: "Image layout options",
		imageProperties: "Image properties…"
	},
	menuDesc: {
		inLineWithText: "Image flows in the line as a glyph",
		squareLeft: "Image floats left, text wraps on the right",
		squareRight: "Image floats right, text wraps on the left",
		behindText: "Image paints behind body text",
		inFrontOfText: "Image paints over body text"
	}
};
var responsePreview = {
	loading: "{action}...",
	result: "{action} Result",
	closeEsc: "Close (Esc)",
	editPrompt: "Edit the result before accepting:",
	changes: "Changes:",
	original: "Original:",
	"new": "New:",
	cancelEdit: "Cancel Edit"
};
var commentMarkers = {
	resolvedComment: "Resolved comment",
	comment: "Comment"
};
var editor = {
	toggleCommentsSidebar: "Toggle comments sidebar",
	showDocumentOutline: "Show document outline",
	editing: "Editing",
	editingDescription: "Edit document directly",
	suggesting: "Suggesting",
	suggestingDescription: "Edits become suggestions",
	viewing: "Viewing",
	viewingDescription: "Read-only, no edits",
	failedToParse: "Failed to parse document",
	linkRemoved: "Link removed",
	linkCopied: "Link copied to clipboard",
	failedToSave: "Failed to save document"
};
var hyperlinkPopup = {
	displayTextPlaceholder: "Display text",
	urlPlaceholder: "https://example.com",
	copyLink: "Copy link",
	editLink: "Edit link",
	removeLink: "Remove link"
};
var headerFooter = {
	header: "Header",
	footer: "Footer",
	options: "Options",
	insertPageNumber: "Insert current page number",
	insertTotalPages: "Insert total page count",
	remove: "Remove {label}",
	closeEditing: "Close {label} editing"
};
var image = {
	placeholder: "Image placeholder",
	placeholderText: "[Image]",
	editableAriaLabel: "Editable image"
};
var imageOverlay = {
	rotate: "Rotate",
	imageProperties: "Image properties",
	deleteImage: "Delete image",
	replaceImage: "Replace image…"
};
var ruler = {
	horizontal: "Horizontal ruler",
	vertical: "Vertical ruler",
	firstLineIndent: "First line indent",
	leftIndent: "Left indent",
	rightIndent: "Right indent",
	topMargin: "Top margin",
	bottomMargin: "Bottom margin"
};
var print = {
	label: "Print",
	allPages: "All ({totalPages} pages)",
	singlePage: "Page {start}",
	pageRange: "Pages {start}-{end}"
};
var unsaved = {
	unsaved: "Unsaved",
	saved: "Saved",
	unsavedTitle: "Document has unsaved changes",
	savedTitle: "All changes saved",
	unsavedAriaLabel: "Unsaved changes",
	savedAriaLabel: "All changes saved"
};
var loading = {
	label: "Loading"
};
var agentPanel = {
	defaultTitle: "Assistant",
	toggle: "Open assistant",
	close: "Close panel",
	resizeHandle: "Resize agent panel",
	thinking: "Assistant is thinking",
	composerPlaceholder: "Ask the assistant…",
	send: "Send",
	timeline: {
		working: "Working… {count, plural, one {# step} other {# steps}}",
		summary: "{count, plural, one {# step} other {# steps}}",
		earlier: "+ {count, plural, one {# earlier step} other {# earlier steps}}"
	}
};
var enJson = {
	_lang: _lang,
	common: common,
	toolbar: toolbar,
	formattingBar: formattingBar,
	alignment: alignment,
	lists: lists,
	lineSpacing: lineSpacing,
	styles: styles,
	font: font,
	fontSize: fontSize,
	zoom: zoom,
	colorPicker: colorPicker,
	dialogs: dialogs,
	comments: comments,
	trackedChanges: trackedChanges,
	revisions: revisions,
	contextMenu: contextMenu,
	documentOutline: documentOutline,
	sidebar: sidebar,
	viewer: viewer,
	titleBar: titleBar,
	errors: errors,
	table: table,
	tableAdvanced: tableAdvanced,
	imageTransform: imageTransform,
	imageWrap: imageWrap,
	responsePreview: responsePreview,
	commentMarkers: commentMarkers,
	editor: editor,
	hyperlinkPopup: hyperlinkPopup,
	headerFooter: headerFooter,
	image: image,
	imageOverlay: imageOverlay,
	ruler: ruler,
	print: print,
	unsaved: unsaved,
	loading: loading,
	agentPanel: agentPanel
};

/**
 * Shared locale data, types, and runtime helpers for the @npde
 * docx-editor adapters.
 *
 * Import everything from the package root. `sideEffects: false` lets
 * consumer bundlers tree-shake unused locales.
 *
 * ```ts
 * import {
 *   en, de, pl, tr, he, ptBR, zhCN,    // typed locale data
 *   locales,                            // record keyed by BCP-47 tag
 *   deepMerge, createT,                 // build a t() for custom hosts
 *   type LocaleStrings,                 // shape of `en` (source of truth)
 *   type Translations,                  // shape of a community partial
 *   type TranslationKey,                // every valid `t()` key
 *   type LocaleCode,                    // 'en' | 'de' | 'pt-BR' | ...
 * } from '@npde/docx-editor-i18n';
 * ```
 *
 * The React and Vue adapters wrap `createT` in framework-native bindings
 * (`useTranslation`, `LocaleProvider`, etc.); use those for app code.
 * Reach for `createT` directly when building a non-React/Vue host.
 *
 * @packageDocumentation
 * @public
 */

/**
 * Full locale string set, auto-derived from `en.json` (the source of truth).
 * Every other locale is a `PartialLocaleStrings` against this shape.
 *
 * @public
 */
type LocaleStrings = typeof enJson;
/**
 * Every locale code shipped from this package. Pass to `locales[code]`
 * for runtime lookup; assign to `_lang` to drive `Intl.PluralRules`.
 *
 * Custom codes are accepted at runtime ({@link PartialLocaleStrings._lang}
 * widens to any string), but the shipped union is the IDE-completion list.
 *
 * @public
 */
type LocaleCode = 'en' | 'de' | 'fr' | 'he' | 'hi' | 'id' | 'pl' | 'pt-BR' | 'tr' | 'zh-CN';
/** English (`en`) — the source of truth, 100% covered. @public */
declare const en: LocaleStrings;
/** German (`de`). Community-maintained; null leaves fall back to English. @public */
declare const de: PartialLocaleStrings;
/** French (`fr`). Community-maintained; null leaves fall back to English. @public */
declare const fr: PartialLocaleStrings;
/** Hebrew (`he`). Community-maintained; null leaves fall back to English. @public */
declare const he: PartialLocaleStrings;
/** Hindi (`hi`). Community-maintained; null leaves fall back to English. @public */
declare const hi: PartialLocaleStrings;
/** Indonesian (`id`). Community-maintained; null leaves fall back to English. @public */
declare const id: PartialLocaleStrings;
/** Polish (`pl`). Community-maintained; null leaves fall back to English. @public */
declare const pl: PartialLocaleStrings;
/** Portuguese (Brazil) (`pt-BR`). Community-maintained; null leaves fall back to English. @public */
declare const ptBR: PartialLocaleStrings;
/** Turkish (`tr`). Community-maintained; null leaves fall back to English. @public */
declare const tr: PartialLocaleStrings;
/** Simplified Chinese (`zh-CN`). Community-maintained; null leaves fall back to English. @public */
declare const zhCN: PartialLocaleStrings;
/**
 * Every shipped locale, keyed by BCP-47 tag. Use for runtime locale
 * pickers and "look up the locale matching this user preference" code:
 *
 * ```ts
 * <DocxEditor i18n={locales[userLocale]} />
 * ```
 *
 * Importing `locales` defeats the per-locale tree-shake — the bundler
 * sees a static reference to every locale. If you only need one or two,
 * import them by name (`import { en, de } from '...'`) instead.
 *
 * @public
 */
declare const locales: Record<LocaleCode, PartialLocaleStrings>;
/**
 * Recursive Partial that allows `null` at leaves to signal "not yet
 * translated, fall back to English." Community translations use this shape;
 * `bun run i18n:fix` keeps every locale aligned to `en.json` with `null`
 * placeholders for missing keys.
 *
 * @public
 */
type DeepPartial<T> = {
    [K in keyof T]?: T[K] extends object ? DeepPartial<T[K]> : T[K] | null;
};
/**
 * Partial locale strings — what consumers pass to the editor's `i18n` prop.
 * Missing keys fall back to English. Optional `_lang` carries the BCP-47
 * tag used by `Intl.PluralRules`; shipped codes autocomplete but custom
 * strings are accepted.
 *
 * @public
 */
type PartialLocaleStrings = DeepPartial<LocaleStrings> & {
    _lang?: LocaleCode | (string & {});
};
/**
 * Alias for `PartialLocaleStrings`. Prefer this name when typing the
 * consumer-facing `i18n` prop or function parameter.
 *
 * @public
 */
type Translations = PartialLocaleStrings;
type DotPath<T, Prefix extends string = ''> = {
    [K in keyof T & string]: T[K] extends Record<string, unknown> ? DotPath<T[K], `${Prefix}${K}.`> : `${Prefix}${K}`;
}[keyof T & string];
/**
 * Every valid dot-notation key into `LocaleStrings`, e.g. `'toolbar.bold'`
 * or `'dialogs.findReplace.matchCount'`. Pass to `t(key, vars?)` for
 * compile-time-checked translation lookup.
 *
 * @public
 */
type TranslationKey = DotPath<LocaleStrings>;
type AnyRecord = Record<string, unknown>;
/**
 * Deep-merge a partial locale over a base locale. Null leaves in the
 * override are treated as "not translated" and fall back to the base.
 * Adapters call this once when the `i18n` prop changes, then hand the
 * result to {@link createT}.
 *
 * @public
 */
declare function deepMerge(base: AnyRecord, override: AnyRecord | undefined): AnyRecord;
/**
 * The signature of `t()`: look up a translation by dot-notation key,
 * interpolate `{vars}`, and resolve ICU plurals.
 *
 * @public
 */
type TFunction = (key: TranslationKey, vars?: Record<string, string | number>) => string;
/**
 * Build a typed `t(key, vars?)` function from a merged locale.
 *
 * - **Lookup**: dot-notation paths against the locale tree
 *   (`'toolbar.bold'`, `'dialogs.findReplace.matchCount'`).
 * - **Interpolation**: `{name}` placeholders read from `vars`.
 * - **Plurals**: ICU `{count, plural, =0 {none} one {# item} other {# items}}`
 *   with `Intl.PluralRules` for CLDR categories and `=N` for exact matches.
 * - **Fallback**: missing keys return the key string itself, useful for
 *   spotting un-translated UI in development.
 *
 * The React/Vue adapters wrap this in `useTranslation()`; use it directly
 * when building a non-React/Vue host (server-rendered docs, CLI, etc.).
 *
 * @example
 * ```ts
 * import { deepMerge, createT, en, de } from '@npde/docx-editor-i18n';
 * const merged = deepMerge(en, de) as LocaleStrings;
 * const t = createT(merged, 'de');
 * t('toolbar.bold');                          // → 'Fett'
 * t('dialogs.findReplace.matchCount', { current: 3, total: 15 });
 * ```
 *
 * @public
 */
declare function createT(strings: LocaleStrings, lang?: string): TFunction;

export { type DeepPartial, type LocaleCode, type LocaleStrings, type PartialLocaleStrings, type TFunction, type TranslationKey, type Translations, createT, de, deepMerge, en, fr, he, hi, id, locales, pl, ptBR, tr, zhCN };
