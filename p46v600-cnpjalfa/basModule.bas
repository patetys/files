Attribute VB_Name = "basModule"
'------------------------------------------------------------------------------------------
' Programa cliente para acionamento de transações Nortix via web services
'
' Todos os direitos reservados.
' Uso exclusivo para os clientes e parceiros com contrato vigente.
'
' Versao 1.0 de 01-Fev-2008
'
' Nortix Informática S/S Ltda.
' fone (11) 2901-3515
' nortix@nortix.com.br
'-------------------------------------------------------------------------------------------

Option Explicit
Global global_rstAux As Object


'-- Microsoft ADO
'-- Copyright (c) 1996-1998 Microsoft Corporation.
'-- ADO Global Constants

'---- CursorTypeEnum Values ----
Global Const adOpenForwardOnly = 0
Global Const adOpenKeyset = 1
Global Const adOpenDynamic = 2
Global Const adOpenStatic = 3

'---- CursorOptionEnum Values ----
Global Const adHoldRecords = &H100
Global Const adMovePrevious = &H200
Global Const adAddNew = &H1000400
Global Const adDelete = &H1000800
Global Const adUpdate = &H1008000
Global Const adBookmark = &H2000
Global Const adApproxPosition = &H4000
Global Const adUpdateBatch = &H10000
Global Const adResync = &H20000
Global Const adNotify = &H40000
Global Const adFind = &H80000
Global Const adSeek = &H400000
Global Const adIndex = &H800000

'---- LockTypeEnum Values ----
Global Const adLockReadOnly = 1
Global Const adLockPessimistic = 2
Global Const adLockOptimistic = 3
Global Const adLockBatchOptimistic = 4

'---- ExecuteOptionEnum Values ----
Global Const adAsyncExecute = &H10
Global Const adAsyncFetch = &H20
Global Const adAsyncFetchNonBlocking = &H40
Global Const adExecuteNoRecords = &H80
Global Const adExecuteStream = &H400

'---- ConnectOptionEnum Values ----
Global Const adAsyncConnect = &H10

'---- ObjectStateEnum Values ----
Global Const adStateClosed = &H0
Global Const adStateOpen = &H1
Global Const adStateConnecting = &H2
Global Const adStateExecuting = &H4
Global Const adStateFetching = &H8

'---- CursorLocationEnum Values ----
Global Const adUseServer = 2
Global Const adUseClient = 3

'---- DataTypeEnum Values ----
Global Const adEmpty = 0
Global Const adTinyInt = 16
Global Const adSmallInt = 2
Global Const adInteger = 3
Global Const adBigInt = 20
Global Const adUnsignedTinyInt = 17
Global Const adUnsignedSmallInt = 18
Global Const adUnsignedInt = 19
Global Const adUnsignedBigInt = 21
Global Const adSingle = 4
Global Const adDouble = 5
Global Const adCurrency = 6
Global Const adDecimal = 14
Global Const adNumeric = 131
Global Const adBoolean = 11
Global Const adError = 10
Global Const adUserDefined = 132
Global Const adVariant = 12
Global Const adIDispatch = 9
Global Const adIUnknown = 13
Global Const adGUID = 72
Global Const adDate = 7
Global Const adDBDate = 133
Global Const adDBTime = 134
Global Const adDBTimeStamp = 135
Global Const adBSTR = 8
Global Const adChar = 129
Global Const adVarChar = 200
Global Const adLongVarChar = 201
Global Const adWChar = 130
Global Const adVarWChar = 202
Global Const adLongVarWChar = 203
Global Const adBinary = 128
Global Const adVarBinary = 204
Global Const adLongVarBinary = 205
Global Const adChapter = 136
Global Const adFileTime = 64
Global Const adPropVariant = 138
Global Const adVarNumeric = 139
Global Const adArray = &H2000

'---- FieldAttributeEnum Values ----
Global Const adFldMayDefer = &H2
Global Const adFldUpdatable = &H4
Global Const adFldUnknownUpdatable = &H8
Global Const adFldFixed = &H10
Global Const adFldIsNullable = &H20
Global Const adFldMayBeNull = &H40
Global Const adFldLong = &H80
Global Const adFldRowID = &H100
Global Const adFldRowVersion = &H200
Global Const adFldCacheDeferred = &H1000
Global Const adFldIsChapter = &H2000
Global Const adFldNegativeScale = &H4000
Global Const adFldKeyColumn = &H8000
Global Const adFldIsRowURL = &H10000
Global Const adFldIsDefaultStream = &H20000
Global Const adFldIsCollection = &H40000

'---- EditModeEnum Values ----
Global Const adEditNone = &H0
Global Const adEditInProgress = &H1
Global Const adEditAdd = &H2
Global Const adEditDelete = &H4

'---- RecordStatusEnum Values ----
Global Const adRecOK = &H0
Global Const adRecNew = &H1
Global Const adRecModified = &H2
Global Const adRecDeleted = &H4
Global Const adRecUnmodified = &H8
Global Const adRecInvalid = &H10
Global Const adRecMultipleChanges = &H40
Global Const adRecPendingChanges = &H80
Global Const adRecCanceled = &H100
Global Const adRecCantRelease = &H400
Global Const adRecConcurrencyViolation = &H800
Global Const adRecIntegrityViolation = &H1000
Global Const adRecMaxChangesExceeded = &H2000
Global Const adRecObjectOpen = &H4000
Global Const adRecOutOfMemory = &H8000
Global Const adRecPermissionDenied = &H10000
Global Const adRecSchemaViolation = &H20000
Global Const adRecDBDeleted = &H40000

'---- GetRowsOptionEnum Values ----
Global Const adGetRowsRest = -1

'---- PositionEnum Values ----
Global Const adPosUnknown = -1
Global Const adPosBOF = -2
Global Const adPosEOF = -3

'---- BookmarkEnum Values ----
Global Const adBookmarkCurrent = 0
Global Const adBookmarkFirst = 1
Global Const adBookmarkLast = 2

'---- MarshalOptionsEnum Values ----
Global Const adMarshalAll = 0
Global Const adMarshalModifiedOnly = 1

'---- AffectEnum Values ----
Global Const adAffectCurrent = 1
Global Const adAffectGroup = 2
Global Const adAffectAllChapters = 4

'---- ResyncEnum Values ----
Global Const adResyncUnderlyingValues = 1
Global Const adResyncAllValues = 2

'---- CompareEnum Values ----
Global Const adCompareLessThan = 0
Global Const adCompareEqual = 1
Global Const adCompareGreaterThan = 2
Global Const adCompareNotEqual = 3
Global Const adCompareNotComparable = 4

'---- FilterGroupEnum Values ----
Global Const adFilterNone = 0
Global Const adFilterPendingRecords = 1
Global Const adFilterAffectedRecords = 2
Global Const adFilterFetchedRecords = 3
Global Const adFilterConflictingRecords = 5

'---- SearchDirectionEnum Values ----
Global Const adSearchForward = 1
Global Const adSearchBackward = -1

'---- PersistFormatEnum Values ----
Global Const adPersistADTG = 0
Global Const adPersistXML = 1

'---- StringFormatEnum Values ----
Global Const adClipString = 2

'---- ConnectPromptEnum Values ----
Global Const adPromptAlways = 1
Global Const adPromptComplete = 2
Global Const adPromptCompleteRequired = 3
Global Const adPromptNever = 4

'---- ConnectModeEnum Values ----
Global Const adModeUnknown = 0
Global Const adModeRead = 1
Global Const adModeWrite = 2
Global Const adModeReadWrite = 3
Global Const adModeShareDenyRead = 4
Global Const adModeShareDenyWrite = 8
Global Const adModeShareExclusive = &HC
Global Const adModeShareDenyNone = &H10
Global Const adModeRecursive = &H400000

'---- RecordCreateOptionsEnum Values ----
Global Const adCreateCollection = &H2000
Global Const adCreateStructDoc = &H80000000
Global Const adCreateNonCollection = &H0
Global Const adOpenIfExists = &H2000000
Global Const adCreateOverwrite = &H4000000
Global Const adFailIfNotExists = -1

'---- RecordOpenOptionsEnum Values ----
Global Const adOpenRecordUnspecified = -1
Global Const adOpenOutput = &H800000
Global Const adOpenAsync = &H1000
Global Const adDelayFetchStream = &H4000
Global Const adDelayFetchFields = &H8000
Global Const adOpenExecuteCommand = &H10000

'---- IsolationLevelEnum Values ----
Global Const adXactUnspecified = &HFFFFFFFF
Global Const adXactChaos = &H10
Global Const adXactReadUncommitted = &H100
Global Const adXactBrowse = &H100
Global Const adXactCursorStability = &H1000
Global Const adXactReadCommitted = &H1000
Global Const adXactRepeatableRead = &H10000
Global Const adXactSerializable = &H100000
Global Const adXactIsolated = &H100000

'---- XactAttributeEnum Values ----
Global Const adXactCommitRetaining = &H20000
Global Const adXactAbortRetaining = &H40000

'---- PropertyAttributesEnum Values ----
Global Const adPropNotSupported = &H0
Global Const adPropRequired = &H1
Global Const adPropOptional = &H2
Global Const adPropRead = &H200
Global Const adPropWrite = &H400

'---- ErrorValueEnum Values ----
Global Const adErrProviderFailed = &HBB8
Global Const adErrInvalidArgument = &HBB9
Global Const adErrOpeningFile = &HBBA
Global Const adErrReadFile = &HBBB
Global Const adErrWriteFile = &HBBC
Global Const adErrNoCurrentRecord = &HBCD
Global Const adErrIllegalOperation = &HC93
Global Const adErrCantChangeProvider = &HC94
Global Const adErrInTransaction = &HCAE
Global Const adErrFeatureNotAvailable = &HCB3
Global Const adErrItemNotFound = &HCC1
Global Const adErrObjectInCollection = &HD27
Global Const adErrObjectNotSet = &HD5C
Global Const adErrDataConversion = &HD5D
Global Const adErrObjectClosed = &HE78
Global Const adErrObjectOpen = &HE79
Global Const adErrProviderNotFound = &HE7A
Global Const adErrBoundToCommand = &HE7B
Global Const adErrInvalidParamInfo = &HE7C
Global Const adErrInvalidConnection = &HE7D
Global Const adErrNotReentrant = &HE7E
Global Const adErrStillExecuting = &HE7F
Global Const adErrOperationCancelled = &HE80
Global Const adErrStillConnecting = &HE81
Global Const adErrInvalidTransaction = &HE82
Global Const adErrUnsafeOperation = &HE84
Global Const adwrnSecurityDialog = &HE85
Global Const adwrnSecurityDialogHeader = &HE86
Global Const adErrIntegrityViolation = &HE87
Global Const adErrPermissionDenied = &HE88
Global Const adErrDataOverflow = &HE89
Global Const adErrSchemaViolation = &HE8A
Global Const adErrSignMismatch = &HE8B
Global Const adErrCantConvertvalue = &HE8C
Global Const adErrCantCreate = &HE8D
Global Const adErrColumnNotOnThisRow = &HE8E
Global Const adErrURLIntegrViolSetColumns = &HE8F
Global Const adErrURLDoesNotExist = &HE8F
Global Const adErrTreePermissionDenied = &HE90
Global Const adErrInvalidURL = &HE91
Global Const adErrResourceLocked = &HE92
Global Const adErrResourceExists = &HE93
Global Const adErrCannotComplete = &HE94
Global Const adErrVolumeNotFound = &HE95
Global Const adErrOutOfSpace = &HE96
Global Const adErrResourceOutOfScope = &HE97
Global Const adErrUnavailable = &HE98
Global Const adErrURLNamedRowDoesNotExist = &HE99
Global Const adErrDelResOutOfScope = &HE9A
Global Const adErrPropInvalidColumn = &HE9B
Global Const adErrPropInvalidOption = &HE9C
Global Const adErrPropInvalidValue = &HE9D
Global Const adErrPropConflicting = &HE9E
Global Const adErrPropNotAllSettable = &HE9F
Global Const adErrPropNotSet = &HEA0
Global Const adErrPropNotSettable = &HEA1
Global Const adErrPropNotSupported = &HEA2
Global Const adErrCatalogNotSet = &HEA3
Global Const adErrCantChangeConnection = &HEA4
Global Const adErrFieldsUpdateFailed = &HEA5
Global Const adErrDenyNotSupported = &HEA6
Global Const adErrDenyTypeNotSupported = &HEA7
Global Const adErrProviderNotSpecified = &HEA9
Global Const adErrConnectionStringTooLong = &HEAA

'---- ParameterAttributesEnum Values ----
Global Const adParamSigned = &H10
Global Const adParamNullable = &H40
Global Const adParamLong = &H80

'---- ParameterDirectionEnum Values ----
Global Const adParamUnknown = &H0
Global Const adParamInput = &H1
Global Const adParamOutput = &H2
Global Const adParamInputOutput = &H3
Global Const adParamReturnValue = &H4

'---- CommandTypeEnum Values ----
Global Const adCmdUnknown = &H8
Global Const adCmdText = &H1
Global Const adCmdTable = &H2
Global Const adCmdStoredProc = &H4
Global Const adCmdFile = &H100
Global Const adCmdTableDirect = &H200

'---- EventStatusEnum Values ----
Global Const adStatusOK = &H1
Global Const adStatusErrorsOccurred = &H2
Global Const adStatusCantDeny = &H3
Global Const adStatusCancel = &H4
Global Const adStatusUnwantedEvent = &H5

'---- EventReasonEnum Values ----
Global Const adRsnAddNew = 1
Global Const adRsnDelete = 2
Global Const adRsnUpdate = 3
Global Const adRsnUndoUpdate = 4
Global Const adRsnUndoAddNew = 5
Global Const adRsnUndoDelete = 6
Global Const adRsnRequery = 7
Global Const adRsnResynch = 8
Global Const adRsnClose = 9
Global Const adRsnMove = 10
Global Const adRsnFirstChange = 11
Global Const adRsnMoveFirst = 12
Global Const adRsnMoveNext = 13
Global Const adRsnMovePrevious = 14
Global Const adRsnMoveLast = 15

'---- SchemaEnum Values ----
Global Const adSchemaProviderSpecific = -1
Global Const adSchemaAsserts = 0
Global Const adSchemaCatalogs = 1
Global Const adSchemaCharacterSets = 2
Global Const adSchemaCollations = 3
Global Const adSchemaColumns = 4
Global Const adSchemaCheckConstraints = 5
Global Const adSchemaConstraintColumnUsage = 6
Global Const adSchemaConstraintTableUsage = 7
Global Const adSchemaKeyColumnUsage = 8
Global Const adSchemaReferentialConstraints = 9
Global Const adSchemaTableConstraints = 10
Global Const adSchemaColumnsDomainUsage = 11
Global Const adSchemaIndexes = 12
Global Const adSchemaColumnPrivileges = 13
Global Const adSchemaTablePrivileges = 14
Global Const adSchemaUsagePrivileges = 15
Global Const adSchemaProcedures = 16
Global Const adSchemaSchemata = 17
Global Const adSchemaSQLLanguages = 18
Global Const adSchemaStatistics = 19
Global Const adSchemaTables = 20
Global Const adSchemaTranslations = 21
Global Const adSchemaProviderTypes = 22
Global Const adSchemaViews = 23
Global Const adSchemaViewColumnUsage = 24
Global Const adSchemaViewTableUsage = 25
Global Const adSchemaProcedureParameters = 26
Global Const adSchemaForeignKeys = 27
Global Const adSchemaPrimaryKeys = 28
Global Const adSchemaProcedureColumns = 29
Global Const adSchemaDBInfoKeywords = 30
Global Const adSchemaDBInfoLiterals = 31
Global Const adSchemaCubes = 32
Global Const adSchemaDimensions = 33
Global Const adSchemaHierarchies = 34
Global Const adSchemaLevels = 35
Global Const adSchemaMeasures = 36
Global Const adSchemaProperties = 37
Global Const adSchemaMembers = 38
Global Const adSchemaTrustees = 39
Global Const adSchemaFunctions = 40
Global Const adSchemaActions = 41
Global Const adSchemaCommands = 42
Global Const adSchemaSets = 43

'---- FieldStatusEnum Values ----
Global Const adFieldOK = 0
Global Const adFieldCantConvertValue = 2
Global Const adFieldIsNull = 3
Global Const adFieldTruncated = 4
Global Const adFieldSignMismatch = 5
Global Const adFieldDataOverflow = 6
Global Const adFieldCantCreate = 7
Global Const adFieldUnavailable = 8
Global Const adFieldPermissionDenied = 9
Global Const adFieldIntegrityViolation = 10
Global Const adFieldSchemaViolation = 11
Global Const adFieldBadStatus = 12
Global Const adFieldDefault = 13
Global Const adFieldIgnore = 15
Global Const adFieldDoesNotExist = 16
Global Const adFieldInvalidURL = 17
Global Const adFieldResourceLocked = 18
Global Const adFieldResourceExists = 19
Global Const adFieldCannotComplete = 20
Global Const adFieldVolumeNotFound = 21
Global Const adFieldOutOfSpace = 22
Global Const adFieldCannotDeleteSource = 23
Global Const adFieldReadOnly = 24
Global Const adFieldResourceOutOfScope = 25
Global Const adFieldAlreadyExists = 26
Global Const adFieldPendingInsert = &H10000
Global Const adFieldPendingDelete = &H20000
Global Const adFieldPendingChange = &H40000
Global Const adFieldPendingUnknown = &H80000
Global Const adFieldPendingUnknownDelete = &H100000

'---- SeekEnum Values ----
Global Const adSeekFirstEQ = &H1
Global Const adSeekLastEQ = &H2
Global Const adSeekAfterEQ = &H4
Global Const adSeekAfter = &H8
Global Const adSeekBeforeEQ = &H10
Global Const adSeekBefore = &H20

'---- ADCPROP_UPDATECRITERIA_ENUM Values ----
Global Const adCriteriaKey = 0
Global Const adCriteriaAllCols = 1
Global Const adCriteriaUpdCols = 2
Global Const adCriteriaTimeStamp = 3

'---- ADCPROP_ASYNCTHREADPRIORITY_ENUM Values ----
Global Const adPriorityLowest = 1
Global Const adPriorityBelowNormal = 2
Global Const adPriorityNormal = 3
Global Const adPriorityAboveNormal = 4
Global Const adPriorityHighest = 5

'---- ADCPROP_AUTORECALC_ENUM Values ----
Global Const adRecalcUpFront = 0
Global Const adRecalcAlways = 1

'---- ADCPROP_UPDATERESYNC_ENUM Values ----
Global Const adResyncNone = 0
Global Const adResyncAutoIncrement = 1
Global Const adResyncConflicts = 2
Global Const adResyncUpdates = 4
Global Const adResyncInserts = 8
Global Const adResyncAll = 15

'---- MoveRecordOptionsEnum Values ----
Global Const adMoveUnspecified = -1
Global Const adMoveOverWrite = 1
Global Const adMoveDontUpdateLinks = 2
Global Const adMoveAllowEmulation = 4

'---- CopyRecordOptionsEnum Values ----
Global Const adCopyUnspecified = -1
Global Const adCopyOverWrite = 1
Global Const adCopyAllowEmulation = 4
Global Const adCopyNonRecursive = 2

'---- StreamTypeEnum Values ----
Global Const adTypeBinary = 1
Global Const adTypeText = 2

'---- LineSeparatorEnum Values ----
Global Const adLF = 10
Global Const adCR = 13
Global Const adCRLF = -1

'---- StreamOpenOptionsEnum Values ----
Global Const adOpenStreamUnspecified = -1
Global Const adOpenStreamAsync = 1
Global Const adOpenStreamFromRecord = 4

'---- StreamWriteEnum Values ----
Global Const adWriteChar = 0
Global Const adWriteLine = 1

'---- SaveOptionsEnum Values ----
Global Const adSaveCreateNotExist = 1
Global Const adSaveCreateOverWrite = 2

'---- FieldEnum Values ----
Global Const adDefaultStream = -1
Global Const adRecordURL = -2

'---- StreamReadEnum Values ----
Global Const adReadAll = -1
Global Const adReadLine = -2

'---- RecordTypeEnum Values ----
Global Const adSimpleRecord = 0
Global Const adCollectionRecord = 1
Global Const adStructDoc = 2


Public gTipSYASEMIS     As String       'Tipo da base de dados: SYASEMIS.
Public gPwdSYASEMIS     As String       'Parâmetro PWS para base de dados: SYASEMIS.
Public gSerSYASEMIS     As String       'Parâmetro SERVER para base de dados: SYASEMIS.
Public gUidSYASEMIS     As String       'Parâmetro UID para base de dados: SYASEMIS.
Public gEndSYASEMIS     As String       'Endereço da base de dados: SYASEMIS.

