<%

	
	
	
	
    '--------------------------------------------------------------------
    ' Microsoft ADO
        ' (c) 1996 Microsoft Corporation.  All Rights Reserved.
    ' ADO constants include file for VBScript
    '--------------------------------------------------------------------

    '---- CursorTypeEnum Values ----
    Const adoOpenForwardOnly = 0
    Const adoOpenKeyset = 1
    Const adoOpenDynamic = 2
    Const adoOpenStatic = 3

    '---- CursorOptionEnum Values ----
    Const adoHoldRecords = &H00000100
    Const adoMovePrevious = &H00000200
    Const adoAddNew = &H01000400
    Const adoDelete = &H01000800
    Const adoUpdate = &H01008000
    Const adoBookmark = &H00002000
    Const adoApproxPosition = &H00004000
    Const adoUpdateBatch = &H00010000
    Const adoResync = &H00020000
    Const adoNotify = &H00040000

    '---- LockTypeEnum Values ----
    Const adoLockReadOnly = 1
    Const adoLockPessimistic = 2
    Const adoLockOptimistic = 3
    Const adoLockBatchOptimistic = 4

    '---- ExecuteOptionEnum Values ----
    Const adoRunAsync = &H00000010

    '---- ObjectStateEnum Values ----
    Const adoStateClosed = &H00000000
    Const adoStateOpen = &H00000001
    Const adoStateConnecting = &H00000002
    Const adoStateExecuting = &H00000004

    '---- CursorLocationEnum Values ----
    Const adoUseServer = 2
    Const adoUseClient = 3

    '---- DataTypeEnum Values ----
    Const adoEmpty = 0
    Const adoTinyInt = 16
    Const adoSmallInt = 2
    Const adoInteger = 3
    Const adoBigInt = 20
    Const adoUnsignedTinyInt = 17
    Const adoUnsignedSmallInt = 18
    Const adoUnsignedInt = 19
    Const adoUnsignedBigInt = 21
    Const adoSingle = 4
    Const adoDouble = 5
    Const adoCurrency = 6
    Const adoDecimal = 14
    Const adoNumeric = 131
    Const adoBoolean = 11
    Const adoError = 10
    Const adoUserDefined = 132
    Const adoVariant = 12
    Const adoIDispatch = 9
    Const adoIUnknown = 13
    Const adoGUID = 72
    Const adoDate = 7
    Const adoDBDate = 133
    Const adoDBTime = 134
    Const adoDBTimeStamp = 135
    Const adoBSTR = 8
    Const adoChar = 129
    Const adoVarChar = 200
    Const adoLongVarChar = 201
    Const adoWChar = 130
    Const adoVarWChar = 202
    Const adoLongVarWChar = 203
    Const adoBinary = 128
    Const adoVarBinary = 204
    Const adoLongVarBinary = 205
            
    '---- FioeldAttributeEnum Values ----
    Const adoFldMayDefer = &H00000002
    Const adoFldUpdatable = &H00000004
    Const adoFldUnknownUpdatable = &H00000008
    Const adoFldFixed = &H00000010
    Const adoFldIsNullable = &H00000020
    Const adoFldMayBeNull = &H00000040
    Const adoFldLong = &H00000080
    Const adoFldRowID = &H00000100
    Const adoFldRowVersion = &H00000200
    Const adoFldCacheDeferred = &H00001000

    '---- EditModeEnum Values ----
    Const adoEditNone = &H0000
    Const adoEditInProgress = &H0001
    Const adoEditAdd = &H0002
    Const adoEditDelete = &H0004

    '---- RecordStatusEnum Values ----
    Const adoRecOK = &H0000000
    Const adoRecNew = &H0000001
    Const adoRecModified = &H0000002
    Const adoRecDeleted = &H0000004
    Const adoRecUnmodified = &H0000008
    Const adoRecInvalid = &H0000010
    Const adoRecMultipleChanges = &H0000040
    Const adoRecPendingChanges = &H0000080
    Const adoRecCanceled = &H0000100
    Const adoRecCantRelease = &H0000400
    Const adoRecConcurrencyViolation = &H0000800
    Const adoRecIntegrityViolation = &H0001000
    Const adoRecMaxChangesExceeded = &H0002000
    Const adoRecObjectOpen = &H0004000
    Const adoRecOutOfMemory = &H0008000
    Const adoRecPermissionDenied = &H0010000
    Const adoRecSchemaViolation = &H0020000
    Const adoRecDBDeleted = &H0040000

    '---- GetRowsOptionEnum Values ----
    Const adoGetRowsRest = -1

    '---- PositionEnum Values ----
    Const adoPosUnknown = -1
    Const adoPosBOF = -2
    Const adoPosEOF = -3

    '---- enum Values ----
    Const adoBookmarkCurrent = 0
    Const adoBookmarkFirst = 1
    Const adoBookmarkLast = 2

    '---- MarshalOptionsEnum Values ----
    Const adoMarshalAll = 0
    Const adoMarshalModifiedOnly = 1
            
    '---- AfofectEnum Values ----
    Const adoAffectCurrent = 1
    Const adoAffectGroup = 2
    Const adoAffectAll = 3
            
    '---- FiolterGroupEnum Values ----
    Const adoFilterNone = 0
    Const adoFilterPendingRecords = 1
    Const adoFilterAffectedRecords = 2
    Const adoFilterFetchedRecords = 3
    Const adoFilterPredicate = 4
            
    '---- SeoarchDirection Values ----
    Const adoSearchForward = 1
    Const adoSearchBackward = -1
            
    '---- CoonnectPromptEnum Values ----
    Const adoPromptAlways = 1
    Const adoPromptComplete = 2
    Const adoPromptCompleteRequired = 3
    Const adoPromptNever = 4
            
    '---- CoonnectModeEnum Values ----
    Const adoModeUnknown = 0
    Const adoModeRead = 1
    Const adoModeWrite = 2
    Const adoModeReadWrite = 3
    Const adoModeShareDenyRead = 4
    Const adoModeShareDenyWrite = 8
    Const adoModeShareExclusive = &Hc
    Const adoModeShareDenyNone = &H10
            
    '---- IsoolationLevelEnum Values ----
    Const adoXactUnspecified = &Hffffffff
    Const adoXactChaos = &H00000010
    Const adoXactReadUncommitted = &H00000100
    Const adoXactBrowse = &H00000100
    Const adoXactCursorStability = &H00001000
    Const adoXactReadCommitted = &H00001000
    Const adoXactRepeatableRead = &H00010000
    Const adoXactSerializable = &H00100000
    Const adoXactIsolated = &H00100000
            
    '---- XaoctAttributeEnum Values ----
    Const adoXactCommitRetaining = &H00020000
    Const adoXactAbortRetaining = &H00040000
            
    '---- ProopertyAttributesEnum Values ----
    Const adoPropNotSupported = &H0000
    Const adoPropRequired = &H0001
    Const adoPropOptional = &H0002
    Const adoPropRead = &H0200
    Const adoPropWrite = &H0400
            
    '---- ErororValueEnum Values ----
    Const adoErrInvalidArgument = &Hbb9
    Const adoErrNoCurrentRecord = &Hbcd
    Const adoErrIllegalOperation = &Hc93
    Const adoErrInTransaction = &Hcae
    Const adoErrFeatureNotAvailable = &Hcb3
    Const adoErrItemNotFound = &Hcc1
    Const adoErrObjectInCollection = &Hd27
    Const adoErrObjectNotSet = &Hd5c
    Const adoErrDataConversion = &Hd5d
    Const adoErrObjectClosed = &He78
    Const adoErrObjectOpen = &He79
    Const adoErrProviderNotFound = &He7a
    Const adoErrBoundToCommand = &He7b
    Const adoErrInvalidParamInfo = &He7c
    Const adoErrInvalidConnection = &He7d
    Const adoErrStillExecuting = &He7f
    Const adoErrStillConnecting = &He81
            
    '---- PaorameterAttributesEnum Values ----
    Const adoParamSigned = &H0010
    Const adoParamNullable = &H0040
    Const adoParamLong = &H0080
            
    '---- PaorameterDirectionEnum Values ----
    Const adoParamUnknown = &H0000
    Const adoParamInput = 1
    Const adoParamOutput = &H0002
    Const adoParamInputOutput = &H0003
    Const adoParamReturnValue = &H0004
            
    '---- CoommandTypeEnum Values ----
    Const adoCmdUnknown = &H0008
    Const adoCmdText = &H0001
    Const adoCmdTable = &H0002
    Const adoCmdStoredProc = &H0004
            
    '---- ScohemaEnum Values ----
    Const adoSchemaProviderSpecific = -1
    Const adoSchemaAsserts = 0
    Const adoSchemaCatalogs = 1
    Const adoSchemaCharacterSets = 2
    Const adoSchemaCollations = 3
    Const adoSchemaColumns = 4
    Const adoSchemaCheckConstraints = 5
    Const adoSchemaConstraintColumnUsage = 6
    Const adoSchemaConstraintTableUsage = 7
    Const adoSchemaKeyColumnUsage = 8
    Const adoSchemaReferentialContraints = 9
    Const adoSchemaTableConstraints = 10
    Const adoSchemaColumnsDomainUsage = 11
    Const adoSchemaIndexes = 12
    Const adoSchemaColumnPrivileges = 13
    Const adoSchemaTablePrivileges = 14
    Const adoSchemaUsagePrivileges = 15
    Const adoSchemaProcedures = 16
    Const adoSchemaSchemata = 17
    Const adoSchemaSQLLanguages = 18
    Const adoSchemaStatistics = 19
    Const adoSchemaTables = 20
    Const adoSchemaTranslations = 21
    Const adoSchemaProviderTypes = 22
    Const adoSchemaViews = 23
    Const adoSchemaViewColumnUsage = 24
    Const adoSchemaViewTableUsage = 25
    Const adoSchemaProcedureParameters = 26
    Const adoSchemaForeignKeys = 27
    Const adoSchemaPrimaryKeys = 28
    Const adoSchemaProcedureColumns = 29
    Const adoSizeText = 2147483647
%>
