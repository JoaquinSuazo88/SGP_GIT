Attribute VB_Name = "Const_ge"
Option Explicit

' declaración de todas las contantes  para este sistema
'Nombre del proyecto
Global Const AppName = "Sodexo Chile, Minutas"

'constantes para el puntero del mouse
Global Const DEFAULT = 0        ' 0 - Default
Global Const ARROW = 1          ' 1 - Arrow
Global Const CROSSHAIR = 2      ' 2 - Cross
Global Const IBEAM = 3          ' 3 - I-Beam
Global Const ICON_POINTER = 4   ' 4 - Icon
Global Const SIZE_POINTER = 5   ' 5 - Size
Global Const SIZE_NE_SW = 6     ' 6 - Size NE SW
Global Const SIZE_N_S = 7       ' 7 - Size N S
Global Const SIZE_NW_SE = 8     ' 8 - Size NW SE
Global Const SIZE_W_E = 9       ' 9 - Size W E
Global Const UP_ARROW = 10      ' 10 - Up Arrow
Global Const HOURGLASS = 11     ' 11 - Hourglass
Global Const NO_DROP = 12       ' 12 - No drop

'constantes para el manejo de los objetos recordset
Global Const dbDenyWrite = 1           '  Other users can't change dynaset records.
Global Const dbDenyRead = 2            '  Other users can't read dynaset records.
Global Const dbReadOnly = 4            '  Open dynaset as read-only.
Global Const dbAppendOnly = 8          '  You can add new records to the dynaset, but you can 't read existing records.
Global Const dbInconsistent = 16       '  Updates apply to all dynaset fields, even if other records affected.
Global Const dbConsistent = 32         '  Updates apply only to those fields that will not affect other records in the dynaset.
Global Const dbSQLPassThrough = 64     '  Sends an SQL statement to an OdbC database.
Global Const dbFailOnError = 128       '  Roll back changes if error occurs.
Global Const dbForwardOnly = 256       '  Create forward-only scrolling snapshot-type Recordset.
Global Const dbOptionINIPath = 1       '  Set application initialization (.INI) filename and path .


'constantes para crear objetso recordset
Global Const dbOpenTable = 1     ' Open table-type Recordset.
Global Const dbOpenDynaset = 2   ' Open dynaset-type Recordset.
Global Const dbOpenSnapshot = 4  ' Open snapshot-type Recordset.


'constantes para respuestas de MsgBox
Global Const vbOKOnly = 0           ' Display OK button only.
Global Const vbOKCancel = 1         ' Display OK and Cancel buttons.
Global Const vbAbortRetryIgnore = 2 ' Display Abort, Retry, and Ignore buttons.
Global Const vbYesNoCancel = 3      ' Display Yes, No, and Cancel buttons.
Global Const vbYesNo = 4            ' Display Yes and No buttons.
Global Const vbRetryCancel = 5      ' Display Retry and Cancel buttons.
Global Const vbCritical = 16        ' Display Critical Message icon.
Global Const vbQuestion = 32        ' Display Warning Query icon.
Global Const vbExclamation = 48     ' Display Warning Message icon.
Global Const vbInformation = 64     ' Display Information Message icon.
Global Const vbDefaultButton1 = 0   ' First button is default.
Global Const vbDefaultButton2 = 256 ' Second button is default.
Global Const vbDefaultButton3 = 512 ' Third button is default.
Global Const vbApplicationModal = 0 ' Application modal; the user must respond to the message box before continuing work in the current application.
Global Const vbSystemModal = 4096   ' System modal; all applications are suspended until the user responds to the message box.

Global Const vbOK = 1     ' OK
Global Const vbCancel = 2 ' Cancel
Global Const vbAbort = 3  ' Abort
Global Const vbRetry = 4  ' Retry
Global Const vbIgnore = 5 ' Ignore
Global Const vbYes = 6    ' Yes
Global Const vbNo = 7     ' No


'constantes para uso de recursos externos
Global Const vbResBitmap = 0 'Bitmap resource
Global Const vbResIcon = 1   'Icon resource
Global Const vbResCursor = 2 'Cursor resource

'constantes para uso de focos de apertura de ventana en utilitarios
Global Const vbNormalFocus = 1      ' Window has focus and is restored to its original size and position.
Global Const vbMinimizedFocus = 2   ' Window is displayed as an icon with focus.
Global Const vbMaximizedFocus = 3   ' Window is maximized with focus.
Global Const vbNormalNoFocus = 4    ' Window is restored to its most recent size and position. The currently active window remains active.
Global Const vbMinimizedNoFocus = 6 ' Window is displayed as an icon. The currently active window remains active.

'Constantes para seteo de Botones del mouse
Global Const vbPopupMenuLeftButton = 1   '(Default) An item on the pop-up menu reacts to a mouse click only when you use the left mouse button.
Global Const vbPopupMenuRightButton = 2  'An item on the pop-up menu reacts to a mouse click when you use either the right or the left mouse button.  This flag can only be used in the MouseDown event.

'constantes para uso de la ayuda
Global Const cdlHelpContext = 1       'Displays Help for a particular context.  When using this setting, you must also specify a context using the HelpContext property.
Global Const cdlHelpHelpOnHelp = 4     'Displays Help for using the Help application itself.
Global Const cdlHelpQuit = 2           'Notifies the Help application that the specified Help file is no longer in use.
Global Const cdlHelpPartialKey = 105 'Calls the search engine in Windows Help.

Global Const vbYellow = &HC0FFFF

