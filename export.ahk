class sqllitedb {
	; PRIVATE Properties and Methods
	static Version := ""
	static _SQLiteDLL := A_ScriptDir . "\sqlite3.dll"
	static _RefCount := 0
	static _MinVersion := "3.6"
	; ======================
	; class _table
	; Object returned from method GetTable()
	; _table is an independent object and does not need SQLite after creation at all.
	; ======================
	class _table {
		; ======================
		; CONSTRUCTOR  Create instance variables
		; ======================
		__New() {
			this.ColumnCount := 0          ; Number of columns in the result table         (Integer)
			this.RowCount := 0             ; Number of rows in the result table            (Integer)
			this.ColumnNames := []         ; Names of columns in the result table          (Array)
			this.Rows := []                ; Rows of the result table                      (Array of Arrays)
			this.HasNames := false         ; Does var ColumnNames contain names?           (Bool)
			this.HasRows := false          ; Does var Rows contain rows?                   (Bool)
			this._CurrentRow := 0          ; Row index of last returned row                (Integer)
		}
		; ======================
		; METHOD GetRow      Get row for RowIndex
		; Parameters:        RowIndex    - Index of the row to retrieve, the index of the first row is 1
		;                    ByRef Row   - Variable to pass out the row array
		; return values:     On failure  - false
		;                    On success  - true, Row contains a valid array
		; Remarks:           _CurrentRow is set to RowIndex, so a subsequent call of NextRow() will return the
		;                    following row.
		; ======================
		GetRow(RowIndex, ByRef Row) {
			Row := ""
			if (RowIndex < 1 || RowIndex > this.RowCount)
				return false
			if !this.Rows.HasKey(RowIndex)
				return false
			Row := this.Rows[RowIndex]
			this._CurrentRow := RowIndex
			return true
		}
		; ======================
		; METHOD Next        Get next row depending on _CurrentRow
		; Parameters:        ByRef Row   - Variable to pass out the row array
		; return values:     On failure  - false, -1 for EOR (end of rows)
		;                    On success  - true, Row contains a valid array
		; ======================
		Next(ByRef Row) {
			Row := ""
			if (this._CurrentRow >= this.RowCount)
				return -1
			this._CurrentRow += 1
			if !this.Rows.HasKey(this._CurrentRow)
				return false
			Row := this.Rows[this._CurrentRow]
			return true
		}
		; ======================
		; METHOD Reset       Reset _CurrentRow to zero
		; Parameters:        None
		; return value:      true
		; ======================
		Reset() {
			this._CurrentRow := 0
			return true
		}
	}
	; ======================
	; class _recordset
	; Object returned from method Query()
	; The records (rows) of a recordset can be accessed sequentially per call of Next() starting with the first record.
	; After a call of Reset() calls of Next() will start with the first record again.
	; When the recordset isn't needed any more, call Free() to free the resources.
	; The lifetime of a recordset depends on the lifetime of the related SQLiteDB object.
	; ======================
	class _recordset {
		; ======================
		; CONSTRUCTOR  Create instance variables
		; ======================
		__New() {
			this.ColumnCount := 0         ; Number of columns                             (Integer)
			this.ColumnNames := []        ; Names of columns in the result table          (Array)
			this.HasNames := false        ; Does var ColumnNames contain names?           (Bool)
			this.HasRows := false         ; Does _recordset contain rows?                 (Bool)
			this.CurrentRow := 0          ; Index of current row                          (Integer)
			this.ErrorMsg := ""           ; Last error message                            (String)
			this.ErrorCode := 0           ; Last SQLite error code / ErrorLevel           (Variant)
			this._Handle := 0             ; Query handle                                  (Pointer)
			this._DB := {}                ; SQLiteDB object                               (Object)
		}
		; ======================
		; DESTRUCTOR   Clear instance variables
		; ======================
		__Delete() {
			if (this._Handle) {
				this.Free()
			}
		}
		; ======================
		; METHOD Next        Get next row of query result
		; Parameters:        ByRef Row   - Variable to store the row array
		; return values:     On success  - true, Row contains the row array
		;                    On failure  - false, ErrorMsg / ErrorCode contain additional information
		;                                  -1 for EOR (end of records)
		; ======================
		Next(ByRef Row) {
			static SQLITE_NULL := 5
			static SQLITE_BLOB := 4
			static EOR := -1
			Row := ""
			this.ErrorMsg := ""
			this.ErrorCode := 0
			if !(this._Handle) {
				this.ErrorMsg := "Invalid query handle!"
				return false
			}
			RC := DllCall("SQlite3.dll\sqlite3_step", "Ptr", this._Handle, "Cdecl Int")
			if (ErrorLevel) {
				this.ErrorMsg := "DllCall sqlite3_step failed!"
				this.ErrorCode := ErrorLevel
				return false
			}
			if (RC <> this._DB._returnCode("SQLITE_ROW")) {
				if (RC = this._DB._returnCode("SQLITE_DONE")) {
				this.ErrorMsg := "EOR"
				this.ErrorCode := RC
				return EOR
				}
				this.ErrorMsg := this._DB.ErrMsg()
				this.ErrorCode := RC
				return false
			}
			RC := DllCall("SQlite3.dll\sqlite3_data_count", "Ptr", this._Handle, "Cdecl Int")
			if (ErrorLevel) {
				this.ErrorMsg := "DllCall sqlite3_data_count failed!"
				this.ErrorCode := ErrorLevel
				return false
			}
			if (RC < 1) {
				this.ErrorMsg := "Recordset is empty!"
				this.ErrorCode := this._DB._returnCode("SQLITE_EMPTY")
				return false
			}
			Row := []
			loop, %RC% {
				Column := A_Index - 1
				ColumnType := DllCall("SQlite3.dll\sqlite3_column_type", "Ptr", this._Handle, "Int", Column, "Cdecl Int")
				if (ErrorLevel) {
				this.ErrorMsg := "DllCall sqlite3_column_type failed!"
				this.ErrorCode := ErrorLevel
				return false
				}
				if (ColumnType = SQLITE_NULL) {
				Row[A_Index] := ""
				} else if (ColumnType = SQLITE_BLOB) {
				BlobPtr := DllCall("SQlite3.dll\sqlite3_column_blob", "Ptr", this._Handle, "Int", Column, "Cdecl UPtr")
				BlobSize := DllCall("SQlite3.dll\sqlite3_column_bytes", "Ptr", this._Handle, "Int", Column, "Cdecl Int")
				if (BlobPtr = 0) || (BlobSize = 0) {
					Row[A_Index] := ""
				} else {
					Row[A_Index] := {}
					Row[A_Index].Size := BlobSize
					Row[A_Index].Blob := ""
					Row[A_Index].SetCapacity("Blob", BlobSize)
					Addr := Row[A_Index].GetAddress("Blob")
					DllCall("Kernel32.dll\RtlMoveMemory", "Ptr", Addr, "Ptr", BlobPtr, "Ptr", BlobSize)
				}
				} else {
				StrPtr := DllCall("SQlite3.dll\sqlite3_column_text", "Ptr", this._Handle, "Int", Column, "Cdecl UPtr")
				if (ErrorLevel) {
					this.ErrorMsg := "DllCall sqlite3_column_text failed!"
					this.ErrorCode := ErrorLevel
					return false
				}
				Row[A_Index] := StrGet(StrPtr, "UTF-8")
				}
			}
			this.CurrentRow += 1
			return true
		}
		; ======================
		; METHOD Reset       Reset the result pointer
		; Parameters:        None
		; return values:     On success  - true
		;                    On failure  - false, ErrorMsg / ErrorCode contain additional information
		; Remarks:           After a call of this method you can access the query result via Next() again.
		; ======================
		Reset() {
			this.ErrorMsg := ""
			this.ErrorCode := 0
			if !(this._Handle) {
				this.ErrorMsg := "Invalid query handle!"
				return false
			}
			RC := DllCall("SQlite3.dll\sqlite3_reset", "Ptr", this._Handle, "Cdecl Int")
			if (ErrorLevel) {
				this.ErrorMsg := "DllCall sqlite3_reset failed!"
				this.ErrorCode := ErrorLevel
				return false
			}
			if (RC) {
				this.ErrorMsg := this._DB._ErrMsg()
				this.ErrorCode := RC
				return false
			}
			this.CurrentRow := 0
			return true
		}
		; ======================
		; METHOD Free        Free query result
		; Parameters:        None
		; return values:     On success  - true
		;                    On failure  - false, ErrorMsg / ErrorCode contain additional information
		; Remarks:           After the call of this method further access on the query result is impossible.
		; ======================
		Free() {
			this.ErrorMsg := ""
			this.ErrorCode := 0
			if !(this._Handle)
				return true
			RC := DllCall("SQlite3.dll\sqlite3_finalize", "Ptr", this._Handle, "Cdecl Int")
			if (ErrorLevel) {
				this.ErrorMsg := "DllCall sqlite3_finalize failed!"
				this.ErrorCode := ErrorLevel
				return false
			}
			if (RC) {
				this.ErrorMsg := this._DB._ErrMsg()
				this.ErrorCode := RC
				return false
			}
			this._DB._Queries.Delete(this._Handle)
			this._Handle := 0
			this._DB := 0
			return true
		}
	}
	; ======================
	; class _statement
	; Object returned from method Prepare()
	; The life-cycle of a prepared statement object usually goes like this:
	; 1. Create the prepared statement object (PST) by calling DB.Prepare().
	; 2. Bind values to parameters using the PST.Bind_*() methods of the statement object.
	; 3. Run the SQL by calling PST.Step() one or more times.
	; 4. Reset the prepared statement using PTS.Reset() then go back to step 2. Do this zero or more times.
	; 5. Destroy the object using PST.Finalize().
	; The lifetime of a prepared statement depends on the lifetime of the related SQLiteDB object.
	; ======================
	class _statement {
		; ======================
		; CONSTRUCTOR  Create instance variables
		; ======================
		__New() {
			this.ErrorMsg := ""           ; Last error message                            (String)
			this.ErrorCode := 0           ; Last SQLite error code / ErrorLevel           (Variant)
			this.ParamCount := 0          ; Number of SQL parameters for this statement   (Integer)
			this._Handle := 0             ; Query handle                                  (Pointer)
			this._DB := {}                ; SQLiteDB object                               (Object)
		}
		; ======================
		; DESTRUCTOR   Clear instance variables
		; ======================
		__Delete() {
			if (this._Handle)
				this.Free()
		}
		; ======================
		; METHOD Bind        Bind values to SQL parameters.
		; Parameters:        Index       -  1-based index of the SQL parameter
		;                    Type        -  type of the SQL parameter (currently: Blob/Double/Int/Text)
		;                    Param3      -  type dependent value
		;                    Param4      -  type dependent value
		;                    Param5      -  not used
		; return values:     On success  - true
		;                    On failure  - false, ErrorMsg / ErrorCode contain additional information
		; ======================
		Bind(Index, Type, Param3 := "", Param4 := 0, Param5 := 0) {
			static SQLITE_static := 0
			static SQLITE_TRANSIENT := -1
			static Types := {Blob: 1, Double: 1, Int: 1, Text: 1}
			this.ErrorMsg := ""
			this.ErrorCode := 0
			if !(this._Handle) {
				this.ErrorMsg := "Invalid statement handle!"
				return false
			}
			if (Index < 1) || (Index > this.ParamCount) {
				this.ErrorMsg := "Invalid parameter index!"
				return false
			}
			if (Types[Type] = "") {
				this.ErrorMsg := "Invalid parameter type!"
				return false
			}
			if (Type = "Blob") {
				; Param3 = BLOB pointer, Param4 = BLOB size in bytes
				if Param3 Is Not Integer
				{
				this.ErrorMsg := "Invalid blob pointer!"
				return false
				}
				if Param4 Is Not Integer
				{
				this.ErrorMsg := "Invalid blob size!"
				return false
				}
				; Let SQLite always create a copy of the BLOB
				RC := DllCall("SQlite3.dll\sqlite3_bind_blob", "Ptr", this._Handle, "Int", Index, "Ptr", Param3
							, "Int", Param4, "Ptr", -1, "Cdecl Int")
				if (ErrorLeveL) {
				this.ErrorMsg := "DllCall sqlite3_bind_blob failed!"
				this.ErrorCode := ErrorLevel
				return false
				}
				if (RC) {
				this.ErrorMsg := this._ErrMsg()
				this.ErrorCode := RC
				return false
				}
			}
			else if (Type = "Double") {
				; Param3 = double value
				if Param3 Is Not Float
				{
				this.ErrorMsg := "Invalid value for double!"
				return false
				}
				RC := DllCall("SQlite3.dll\sqlite3_bind_double", "Ptr", this._Handle, "Int", Index, "Double", Param3
							, "Cdecl Int")
				if (ErrorLeveL) {
				this.ErrorMsg := "DllCall sqlite3_bind_double failed!"
				this.ErrorCode := ErrorLevel
				return false
				}
				if (RC) {
				this.ErrorMsg := this._ErrMsg()
				this.ErrorCode := RC
				return false
				}
			}
			else if (Type = "Int") {
				; Param3 = integer value
				if Param3 Is Not Integer
				{
				this.ErrorMsg := "Invalid value for int!"
				return false
				}
				RC := DllCall("SQlite3.dll\sqlite3_bind_int", "Ptr", this._Handle, "Int", Index, "Int", Param3
							, "Cdecl Int")
				if (ErrorLeveL) {
				this.ErrorMsg := "DllCall sqlite3_bind_int failed!"
				this.ErrorCode := ErrorLevel
				return false
				}
				if (RC) {
				this.ErrorMsg := this._ErrMsg()
				this.ErrorCode := RC
				return false
				}
			}
			else if (Type = "Text") {
				; Param3 = zero-terminated string
				this._DB._StrToUTF8(Param3, ByRef UTF8)
				; Let SQLite always create a copy of the text
				RC := DllCall("SQlite3.dll\sqlite3_bind_text", "Ptr", this._Handle, "Int", Index, "Ptr", &UTF8
							, "Int", -1, "Ptr", -1, "Cdecl Int")
				if (ErrorLeveL) {
				this.ErrorMsg := "DllCall sqlite3_bind_text failed!"
				this.ErrorCode := ErrorLevel
				return false
				}
				if (RC) {
				this.ErrorMsg := this._ErrMsg()
				this.ErrorCode := RC
				return false
				}
			}
			return true
		}

		; ======================
		; METHOD Step        Evaluate the prepared statement.
		; Parameters:        None
		; return values:     On success  - true
		;                    On failure  - false, ErrorMsg / ErrorCode contain additional information
		; Remarks:           You must call ST.Reset() before you can call ST.Step() again.
		; ======================
		Step() {
			this.ErrorMsg := ""
			this.ErrorCode := 0
			if !(this._Handle) {
				this.ErrorMsg := "Invalid statement handle!"
				return false
			}
			RC := DllCall("SQlite3.dll\sqlite3_step", "Ptr", this._Handle, "Cdecl Int")
			if (ErrorLevel) {
				this.ErrorMsg := "DllCall sqlite3_step failed!"
				this.ErrorCode := ErrorLevel
				return false
			}
			if (RC <> this._DB._returnCode("SQLITE_DONE"))
			&& (RC <> this._DB._returnCode("SQLITE_ROW")) {
				this.ErrorMsg := this._DB.ErrMsg()
				this.ErrorCode := RC
				return false
			}
			return true
		}
		; ======================
		; METHOD Reset       Reset the prepared statement.
		; Parameters:        ClearBindings  - Clear bound SQL parameter values (true/false)
		; return values:     On success     - true
		;                    On failure     - false, ErrorMsg / ErrorCode contain additional information
		; Remarks:           After a call of this method you can access the query result via Next() again.
		; ======================
		Reset(ClearBindings := true) {
			this.ErrorMsg := ""
			this.ErrorCode := 0
			if !(this._Handle) {
				this.ErrorMsg := "Invalid statement handle!"
				return false
			}
			RC := DllCall("SQlite3.dll\sqlite3_reset", "Ptr", this._Handle, "Cdecl Int")
			if (ErrorLevel) {
				this.ErrorMsg := "DllCall sqlite3_reset failed!"
				this.ErrorCode := ErrorLevel
				return false
			}
			if (RC) {
				this.ErrorMsg := this._DB._ErrMsg()
				this.ErrorCode := RC
				return false
			}
			if (ClearBindings) {
				RC := DllCall("SQlite3.dll\sqlite3_clear_bindings", "Ptr", this._Handle, "Cdecl Int")
				if (ErrorLevel) {
				this.ErrorMsg := "DllCall sqlite3_clear_bindings failed!"
				this.ErrorCode := ErrorLevel
				return false
				}
				if (RC) {
				this.ErrorMsg := this._DB._ErrMsg()
				this.ErrorCode := RC
				return false
				}
			}
			return true
		}
		; ======================
		; METHOD Free        Free the prepared statement object.
		; Parameters:        None
		; return values:     On success  - true
		;                    On failure  - false, ErrorMsg / ErrorCode contain additional information
		; Remarks:           After the call of this method further access on the statement object is impossible.
		; ======================
		Free() {
			this.ErrorMsg := ""
			this.ErrorCode := 0
			if !(this._Handle)
				return true
			RC := DllCall("SQlite3.dll\sqlite3_finalize", "Ptr", this._Handle, "Cdecl Int")
			if (ErrorLevel) {
				this.ErrorMsg := "DllCall sqlite3_finalize failed!"
				this.ErrorCode := ErrorLevel
				return false
			}
			if (RC) {
				this.ErrorMsg := this._DB._ErrMsg()
				this.ErrorCode := RC
				return false
			}
			this._DB._Stmts.Delete(this._Handle)
			this._Handle := 0
			this._DB := 0
			return true
		}
	}
	; ======================
	; CONSTRUCTOR __New
	; ======================
	__New() {
		this._Path := ""                  ; Database path                                 (String)
		this._Handle := 0                 ; Database handle                               (Pointer)
		this._Queries := {}               ; Valid queries                                 (Object)
		this._Stmts := {}                 ; Valid prepared statements                     (Object)
		if (this.Base._RefCount = 0) {
			SQLiteDLL := this.Base._SQLiteDLL
			if !FileExist(SQLiteDLL)
				if FileExist(A_ScriptDir . "\SQLiteDB.ini") {
				IniRead, SQLiteDLL, %A_ScriptDir%\SQLiteDB.ini, Main, DllPath, %SQLiteDLL%
				this.Base._SQLiteDLL := SQLiteDLL
			}
			if !(DLL := DllCall("LoadLibrary", "Str", this.Base._SQLiteDLL, "UPtr")) {
				MsgBox, 16, SQLiteDB Error, % "DLL " . SQLiteDLL . " does not exist!"
				ExitApp
			}
			this.Base.Version := StrGet(DllCall("SQlite3.dll\sqlite3_libversion", "Cdecl UPtr"), "UTF-8")
			SQLVersion := StrSplit(this.Base.Version, ".")
			MinVersion := StrSplit(this.Base._MinVersion, ".")
			if (SQLVersion[1] < MinVersion[1]) || ((SQLVersion[1] = MinVersion[1]) && (SQLVersion[2] < MinVersion[2])){
				DllCall("FreeLibrary", "Ptr", DLL)
				MsgBox, 16, SQLite ERROR, % "Version " . this.Base.Version .  " of SQLite3.dll is not supported!`n`n"
										. "You can download the current version from www.sqlite.org!"
				ExitApp
			}
		}
		this.Base._RefCount += 1
	}
	; ======================
	; DESTRUCTOR __Delete
	; ======================
	__Delete() {
		if (this._Handle)
			this.CloseDB()
		this.Base._RefCount -= 1
		if (this.Base._RefCount = 0) {
			if (DLL := DllCall("GetModuleHandle", "Str", this.Base._SQLiteDLL, "UPtr"))
				DllCall("FreeLibrary", "Ptr", DLL)
		}
	}
	; ======================
	; PRIVATE _StrToUTF8
	; ======================
	_StrToUTF8(Str, ByRef UTF8) {
		VarSetCapacity(UTF8, StrPut(Str, "UTF-8"), 0)
		StrPut(Str, &UTF8, "UTF-8")
		return &UTF8
	}
	; ======================
	; PRIVATE _UTF8ToStr
	; ======================
	_UTF8ToStr(UTF8) {
		return StrGet(UTF8, "UTF-8")
	}
	; ======================
	; PRIVATE _ErrMsg
	; ======================
	_ErrMsg() {
		if (RC := DllCall("SQLite3.dll\sqlite3_errmsg", "Ptr", this._Handle, "Cdecl UPtr"))
			return StrGet(&RC, "UTF-8")
		return ""
	}
	; ======================
	; PRIVATE _ErrCode
	; ======================
	_ErrCode() {
		return DllCall("SQLite3.dll\sqlite3_errcode", "Ptr", this._Handle, "Cdecl Int")
	}
	; ======================
	; PRIVATE _Changes
	; ======================
	_Changes() {
		return DllCall("SQLite3.dll\sqlite3_changes", "Ptr", this._Handle, "Cdecl Int")
	}
	; ======================
	; PRIVATE _returncode
	; ======================
	_returnCode(RC) {
		static RCODE := {SQLITE_OK: 0          ; Successful result
						, SQLITE_ERROR: 1       ; SQL error or missing database
						, SQLITE_INTERNAL: 2    ; NOT USED. Internal logic error in SQLite
						, SQLITE_PERM: 3        ; Access permission denied
						, SQLITE_ABORT: 4       ; Callback routine requested an abort
						, SQLITE_BUSY: 5        ; The database file is locked
						, SQLITE_LOCKED: 6      ; A table in the database is locked
						, SQLITE_NOMEM: 7       ; A malloc() failed
						, SQLITE_READONLY: 8    ; Attempt to write a readonly database
						, SQLITE_INTERRUPT: 9   ; Operation terminated by sqlite3_interrupt()
						, SQLITE_IOERR: 10      ; Some kind of disk I/O error occurred
						, SQLITE_CORRUPT: 11    ; The database disk image is malformed
						, SQLITE_NOTFOUND: 12   ; NOT USED. Table or record not found
						, SQLITE_FULL: 13       ; Insertion failed because database is full
						, SQLITE_CANTOPEN: 14   ; Unable to open the database file
						, SQLITE_PROTOCOL: 15   ; NOT USED. Database lock protocol error
						, SQLITE_EMPTY: 16      ; Database is empty
						, SQLITE_SCHEMA: 17     ; The database schema changed
						, SQLITE_TOOBIG: 18     ; String or BLOB exceeds size limit
						, SQLITE_CONSTRAINT: 19 ; Abort due to constraint violation
						, SQLITE_MISMATCH: 20   ; Data type mismatch
						, SQLITE_MISUSE: 21     ; Library used incorrectly
						, SQLITE_NOLFS: 22      ; Uses OS features not supported on host
						, SQLITE_AUTH: 23       ; Authorization denied
						, SQLITE_FORMAT: 24     ; Auxiliary database format error
						, SQLITE_RANGE: 25      ; 2nd parameter to sqlite3_bind out of range
						, SQLITE_NOTADB: 26     ; File opened that is not a database file
						, SQLITE_ROW: 100       ; sqlite3_step() has another row ready
						, SQLITE_DONE: 101}     ; sqlite3_step() has finished executing
		return RCODE.HasKey(RC) ? RCODE[RC] : ""
	}

	; PUBLIC Interface ------
	; ======================
	; Properties
	; ======================
	ErrorMsg := ""              ; Error message                           (String)
	ErrorCode := 0              ; SQLite error code / ErrorLevel          (Variant)
	Changes := 0                ; Changes made by last call of Exec()     (Integer)
	SQL := ""                   ; Last executed SQL statement             (String)
	; ======================
	; METHOD OpenDB         Open a database
	; Parameters:           DBPath      - Path of the database file
	;                       Access      - Wanted access: "R"ead / "W"rite
	;                       Create      - Create new database in write mode, if it doesn't exist
	; return values:        On success  - true
	;                       On failure  - false, ErrorMsg / ErrorCode contain additional information
	; Remarks:              if DBPath is empty in write mode, a database called ":memory:" is created in memory
	;                       and deletet on call of CloseDB.
	; ======================
	OpenDB(DBPath, Access := "W", Create := true) {
			static SQLITE_OPEN_READONLY  := 0x01 ; Database opened as read-only
			static SQLITE_OPEN_READWRITE := 0x02 ; Database opened as read-write
			static SQLITE_OPEN_CREATE    := 0x04 ; Database will be created if not exists
			static MEMDB := ":memory:"
			this.ErrorMsg := ""
			this.ErrorCode := 0
			HDB := 0
			if (DBPath = "") {
				DBPath := MEMDB
			}
			if (DBPath = this._Path) && (this._Handle) {
				return true
			}
			if (this._Handle) {
				this.ErrorMsg := "You must first close DB " . this._Path . "!"
				return false
			}
			Flags := 0
			Access := SubStr(Access, 1, 1)
			if (Access <> "W") && (Access <> "R") {
				Access := "R"
			}
			Flags := SQLITE_OPEN_READONLY
			if (Access = "W") {
				Flags := SQLITE_OPEN_READWRITE
				if (Create) {
					Flags |= SQLITE_OPEN_CREATE
				}
			}
			this._Path := DBPath
			this._StrToUTF8(DBPath, UTF8)
			RC := DllCall("SQlite3.dll\sqlite3_open_v2", "Ptr", &UTF8, "PtrP", HDB, "Int", Flags, "Ptr", 0, "Cdecl Int")
			if (ErrorLevel) {
				this._Path := ""
				this.ErrorMsg := "DLLCall sqlite3_open_v2 failed!"
				this.ErrorCode := ErrorLevel
				return false
			}
			if (RC) {
				this._Path := ""
				this.ErrorMsg := this._ErrMsg()
				this.ErrorCode := RC
				return false
			}
			this._Handle := HDB
			return true
	}
	; ======================
	; METHOD CloseDB        Close database
	; Parameters:           None
	; return values:        On success  - true
	;                       On failure  - false, ErrorMsg / ErrorCode contain additional information
	; ======================
	CloseDB() {
			this.ErrorMsg := ""
			this.ErrorCode := 0
			this.SQL := ""
			if !(this._Handle) {
				return true
			}
			for Each, Query in this._Queries {
				DllCall("SQlite3.dll\sqlite3_finalize", "Ptr", Query, "Cdecl Int")
			}
			RC := DllCall("SQlite3.dll\sqlite3_close", "Ptr", this._Handle, "Cdecl Int")
			if (ErrorLevel) {
				this.ErrorMsg := "DLLCall sqlite3_close failed!"
				this.ErrorCode := ErrorLevel
				return false
			}
			if (RC) {
				this.ErrorMsg := this._ErrMsg()
				this.ErrorCode := RC
				return false
			}
			this._Path := ""
			this._Handle := ""
			this._Queries := []
			return true
	}
	; ======================
	; METHOD AttachDB       Add another database file to the current database connection
	;                       http://www.sqlite.org/lang_attach.html
	; Parameters:           DBPath      - Path of the database file
	;                       DBAlias     - Database alias name used internally by SQLite
	; return values:        On success  - true
	;                       On failure  - false, ErrorMsg / ErrorCode contain additional information
	; ======================
	AttachDB(DBPath, DBAlias) {
		return this.Exec("ATTACH DATABASE '" . DBPath . "' As " . DBAlias . ";")
	}
	; ======================
	; METHOD DetachDB       Detaches an additional database connection previously attached using AttachDB()
	;                       http://www.sqlite.org/lang_detach.html
	; Parameters:           DBAlias     - Database alias name used with AttachDB()
	; return values:        On success  - true
	;                       On failure  - false, ErrorMsg / ErrorCode contain additional information
	; ======================
	DetachDB(DBAlias) {
		return this.Exec("DETACH DATABASE " . DBAlias . ";")
	}
	; ======================
	; METHOD Exec           Execute SQL statement
	; Parameters:           SQL         - Valid SQL statement
	;                       Callback    - Name of a callback function to invoke for each result row coming out
	;                                     of the evaluated SQL statements.
	;                                     The function must accept 4 parameters:
	;                                     1: SQLiteDB object
	;                                     2: Number of columns
	;                                     3: Pointer to an array of pointers to columns text
	;                                     4: Pointer to an array of pointers to column names
	;                                     The address of the current SQL string is passed in A_EventInfo.
	;                                     if the callback function returns non-zero, DB.Exec() returns SQLITE_ABORT
	;                                     without invoking the callback again and without running any subsequent
	;                                     SQL statements.
	; return values:        On success  - true, the number of changed rows is given in property Changes
	;                       On failure  - false, ErrorMsg / ErrorCode contain additional information
	; ======================
	Exec(SQL, Callback := "") {
		this.ErrorMsg := ""
		this.ErrorCode := 0
		this.SQL := SQL
		if !(this._Handle) {
			this.ErrorMsg := "Invalid database handle!"
			return false
		}
		CBPtr := 0
		Err := 0
		if (FO := Func(Callback)) && (FO.MinParams = 4) {
			CBPtr := RegisterCallback(Callback, "F C", 4, &SQL)
		}
		this._StrToUTF8(SQL, UTF8)
		RC := DllCall("SQlite3.dll\sqlite3_exec", "Ptr", this._Handle, "Ptr", &UTF8, "Int", CBPtr, "Ptr", Object(This)
			, "PtrP", Err, "Cdecl Int")
		CallError := ErrorLevel
		if (CBPtr) {
			DllCall("Kernel32.dll\GlobalFree", "Ptr", CBPtr)
		}
		if (CallError) {
			this.ErrorMsg := "DLLCall sqlite3_exec failed!"
			this.ErrorCode := CallError
			return false
		}
		if (RC) {
			this.ErrorMsg := StrGet(Err, "UTF-8")
			this.ErrorCode := RC
			DllCall("SQLite3.dll\sqlite3_free", "Ptr", Err, "Cdecl")
			return false
		}
		this.Changes := this._Changes()
		return true
	}
	; ======================
	; METHOD GetTable       Get complete result for SELECT query
	; Parameters:           SQL         - SQL SELECT statement
	;                       ByRef TB    - Variable to store the result object (TB _table)
	;                       MaxResult   - Number of rows to return:
	;                          0          Complete result (default)
	;                         -1          return only RowCount and ColumnCount
	;                         -2          return counters and array ColumnNames
	;                          n          return counters and ColumnNames and first n rows
	; return values:        On success  - true, TB contains the result object
	;                       On failure  - false, ErrorMsg / ErrorCode contain additional information
	; ======================
	GetTable(SQL, ByRef TB, MaxResult := 0) {
		TB := ""
		this.ErrorMsg := ""
		this.ErrorCode := 0
		this.SQL := SQL
		if !(this._Handle) {
			this.ErrorMsg := "Invalid database handle!"
			return false
		}
		if !RegExMatch(SQL, "i)^\s*(SELECT|PRAGMA)\s") {
			this.ErrorMsg := A_ThisFunc . " requires a query statement!"
			return false
		}
		Names := ""
		Err := 0, RC := 0, GetRows := 0
		I := 0, Rows := Cols := 0
		Table := 0
		if MaxResult Is Not Integer
			MaxResult := 0
		if (MaxResult < -2)
			MaxResult := 0
		this._StrToUTF8(SQL, UTF8)
		RC := DllCall("SQlite3.dll\sqlite3_get_table", "Ptr", this._Handle, "Ptr", &UTF8, "PtrP", Table
			, "IntP", Rows, "IntP", Cols, "PtrP", Err, "Cdecl Int")
		if (ErrorLevel) {
			this.ErrorMsg := "DLLCall sqlite3_get_table failed!"
			this.ErrorCode := ErrorLevel
			return false
		}
		if (RC) {
			this.ErrorMsg := StrGet(Err, "UTF-8")
			this.ErrorCode := RC
			DllCall("SQLite3.dll\sqlite3_free", "Ptr", Err, "Cdecl")
			return false
		}
		TB := new this._table
		TB.ColumnCount := Cols
		TB.RowCount := Rows
		if (MaxResult = -1) {
			DllCall("SQLite3.dll\sqlite3_free_table", "Ptr", Table, "Cdecl")
			if (ErrorLevel) {
				this.ErrorMsg := "DLLCall sqlite3_free_table failed!"
				this.ErrorCode := ErrorLevel
				return false
			}
			return true
		}
		if (MaxResult = -2)
			GetRows := 0
		else if (MaxResult > 0) && (MaxResult <= Rows)
			GetRows := MaxResult
		else
			GetRows := Rows
		Offset := 0
		Names := Array()
		loop, %Cols% {
			Names[A_Index] := StrGet(NumGet(Table+0, Offset, "UPtr"), "UTF-8")
			Offset += A_PtrSize
		}
		TB.ColumnNames := Names
		TB.HasNames := true
		loop, %GetRows% {
			I := A_Index
			TB.Rows[I] := []
			loop, %Cols% {
				TB.Rows[I][A_Index] := StrGet(NumGet(Table+0, Offset, "UPtr"), "UTF-8")
				Offset += A_PtrSize
			}
		}
		if (GetRows)
			TB.HasRows := true
		DllCall("SQLite3.dll\sqlite3_free_table", "Ptr", Table, "Cdecl")
		if (ErrorLevel) {
			TB := ""
			this.ErrorMsg := "DLLCall sqlite3_free_table failed!"
			this.ErrorCode := ErrorLevel
			return false
		}
		return true
	}
	; ======================
	; Prepared statement 10:54 2019.07.05. by Dixtroy
	;  DB := new SQLiteDB
	;  DB.OpenDB(DBFileName)
	;  DB.Prepare 1 or more, just once
	;  DB.Step 1 or more on prepared one, repeatable
	;  DB.Finalize at the end
	; ======================
	; ======================
	; METHOD Prepare        Prepare database table for further actions.
	; Parameters:           SQL         - SQL statement to be compiled
	;                       ByRef ST    - Variable to store the statement object (class _statement)
	; return values:        On success  - true, ST contains the statement object
	;                       On failure  - false, ErrorMsg / ErrorCode contain additional information
	; Remarks:              You have to pass one ? for each column you want to assign a value later.
	; ======================
	Prepare(SQL, ByRef ST) {
		this.ErrorMsg := ""
		this.ErrorCode := 0
		this.SQL := SQL
		if !(this._Handle) {
			this.ErrorMsg := "Invalid database handle!"
			return false
		}
		if !RegExMatch(SQL, "i)^\s*(INSERT|UPDATE|REPLACE)\s") {
			this.ErrorMsg := A_ThisFunc . " requires an INSERT/UPDATE/REPLACE statement!"
			return false
		}
		Stmt := 0
		this._StrToUTF8(SQL, UTF8)
		RC := DllCall("SQlite3.dll\sqlite3_prepare_v2", "Ptr", this._Handle, "Ptr", &UTF8, "Int", -1
			, "PtrP", Stmt, "Ptr", 0, "Cdecl Int")
		if (ErrorLeveL) {
			this.ErrorMsg := A_ThisFunc . ": DllCall sqlite3_prepare_v2 failed!"
			this.ErrorCode := ErrorLevel
			return false
		}
		if (RC) {
			this.ErrorMsg := A_ThisFunc . ": " . this._ErrMsg()
			this.ErrorCode := RC
			return false
		}
		ST := New this._statement
		ST.ParamCount := DllCall("SQlite3.dll\sqlite3_bind_parameter_count", "Ptr", this._Handle, "Cdecl Int")
		ST._Handle := Stmt
		ST._DB := This
		this._Stmts[Stmt] := Stmt
		return true
	}
	; ======================
	; METHOD Query          Get "recordset" object for prepared SELECT query
	; Parameters:           SQL         - SQL SELECT statement
	;                       ByRef RS    - Variable to store the result object (class _recordset)
	; return values:        On success  - true, RS contains the result object
	;                       On failure  - false, ErrorMsg / ErrorCode contain additional information
	; ======================
	Query(SQL, ByRef RS) {
		RS := ""
		this.ErrorMsg := ""
		this.ErrorCode := 0
		this.SQL := SQL
		ColumnCount := 0
		HasRows := false
		if !(this._Handle) {
			this.ErrorMsg := "Invalid dadabase handle!"
			return false
		}
		if !RegExMatch(SQL, "i)^\s*(SELECT|PRAGMA)\s|") {
			this.ErrorMsg := A_ThisFunc . " requires a query statement!"
			return false
		}
		Query := 0
		this._StrToUTF8(SQL, UTF8)
		RC := DllCall("SQlite3.dll\sqlite3_prepare_v2", "Ptr", this._Handle, "Ptr", &UTF8, "Int", -1
					, "PtrP", Query, "Ptr", 0, "Cdecl Int")
		if (ErrorLeveL) {
			this.ErrorMsg := "DLLCall sqlite3_prepare_v2 failed!"
			this.ErrorCode := ErrorLevel
			return false
		}
		if (RC) {
			this.ErrorMsg := this._ErrMsg()
			this.ErrorCode := RC
			return false
		}
		RC := DllCall("SQlite3.dll\sqlite3_column_count", "Ptr", Query, "Cdecl Int")
		if (ErrorLevel) {
			this.ErrorMsg := "DLLCall sqlite3_column_count failed!"
			this.ErrorCode := ErrorLevel
			return false
		}
		if (RC < 1) {
			this.ErrorMsg := "Query result is empty!"
			this.ErrorCode := this._returnCode("SQLITE_EMPTY")
			return false
		}
		ColumnCount := RC
		Names := []
		loop, %RC% {
			StrPtr := DllCall("SQlite3.dll\sqlite3_column_name", "Ptr", Query, "Int", A_Index - 1, "Cdecl UPtr")
			if (ErrorLevel) {
				this.ErrorMsg := "DLLCall sqlite3_column_name failed!"
				this.ErrorCode := ErrorLevel
				return false
			}
			Names[A_Index] := StrGet(StrPtr, "UTF-8")
		}
		RC := DllCall("SQlite3.dll\sqlite3_step", "Ptr", Query, "Cdecl Int")
		if (ErrorLevel) {
			this.ErrorMsg := "DLLCall sqlite3_step failed!"
			this.ErrorCode := ErrorLevel
			return false
		}
		if (RC = this._returnCode("SQLITE_ROW"))
			HasRows := true
		RC := DllCall("SQlite3.dll\sqlite3_reset", "Ptr", Query, "Cdecl Int")
		if (ErrorLevel) {
			this.ErrorMsg := "DLLCall sqlite3_reset failed!"
			this.ErrorCode := ErrorLevel
			return false
		}
		RS := new this._recordset
		RS.ColumnCount := ColumnCount
		RS.ColumnNames := Names
		RS.HasNames := true
		RS.HasRows := HasRows
		RS._Handle := Query
		RS._DB := This
		this._Queries[Query] := Query
		return true
	}
	; ======================
	; METHOD CreateScalarFunc  Create a scalar application defined function
	; Parameters:              Name  -  the name of the function
	;                          Args  -  the number of arguments that the SQL function takes
	;                          Func  -  a pointer to AHK functions that implement the SQL function
	;                          Enc   -  specifies what text encoding this SQL function prefers for its parameters
	;                          Param -  an arbitrary pointer accessible within the funtion with sqlite3_user_data()
	; return values:           On success  - true
	;                          On failure  - false, ErrorMsg / ErrorCode contain additional information
	; Documentation:           www.sqlite.org/c3ref/create_function.html
	; ======================
	CreateScalarFunc(Name, Args, Func, Enc := 0x0801, Param := 0) {
		; SQLITE_DETERMINISTIC = 0x0800 - the function will always return the same result given the same inputs
		;                                 within a single SQL statement
		; SQLITE_UTF8 = 0x0001
		this.ErrorMsg := ""
		this.ErrorCode := 0
		if !(this._Handle) {
			this.ErrorMsg := "Invalid database handle!"
			return false
		}
		RC := DllCall("SQLite3.dll\sqlite3_create_function", "Ptr", this._Handle, "AStr", Name, "Int", Args, "Int", Enc
			, "Ptr", Param, "Ptr", Func, "Ptr", 0, "Ptr", 0, "Cdecl Int")
		if (ErrorLeveL) {
			this.ErrorMsg := "DllCall sqlite3_create_function failed!"
			this.ErrorCode := ErrorLevel
			return false
		}
		if (RC) {
			this.ErrorMsg := this._ErrMsg()
			this.ErrorCode := RC
			return false
		}
		return true
	}
	; ======================
	; METHOD LastInsertRowID   Get the ROWID of the last inserted row
	; Parameters:              ByRef RowID - Variable to store the ROWID
	; return values:           On success  - true, RowID contains the ROWID
	;                          On failure  - false, ErrorMsg / ErrorCode contain additional information
	; ======================
	LastInsertRowID(ByRef RowID) {
		this.ErrorMsg := ""
		this.ErrorCode := 0
		this.SQL := ""
		if !(this._Handle) {
			this.ErrorMsg := "Invalid database handle!"
			return false
		}
		RowID := 0
		RC := DllCall("SQLite3.dll\sqlite3_last_insert_rowid", "Ptr", this._Handle, "Cdecl Int64")
		if (ErrorLevel) {
			this.ErrorMsg := "DllCall sqlite3_last_insert_rowid failed!"
			this.ErrorCode := ErrorLevel
			return false
		}
		RowID := RC
		return true
	}
	; ======================
	; METHOD TotalChanges   Get the number of changed rows since connecting to the database
	; Parameters:           ByRef Rows  - Variable to store the number of rows
	; return values:        On success  - true, Rows contains the number of rows
	;                       On failure  - false, ErrorMsg / ErrorCode contain additional information
	; ======================
	TotalChanges(ByRef Rows) {
		this.ErrorMsg := ""
		this.ErrorCode := 0
		this.SQL := ""
		if !(this._Handle) {
			this.ErrorMsg := "Invalid database handle!"
			return false
		}
		Rows := 0
		RC := DllCall("SQLite3.dll\sqlite3_total_changes", "Ptr", this._Handle, "Cdecl Int")
		if (ErrorLevel) {
			this.ErrorMsg := "DllCall sqlite3_total_changes failed!"
			this.ErrorCode := ErrorLevel
			return false
		}
		Rows := RC
		return true
	}
	; ======================
	; METHOD SetTimeout     Set the timeout to wait before SQLITE_BUSY or SQLITE_IOERR_BLOCKED is returned,
	;                       when a table is locked.
	; Parameters:           TimeOut     - Time to wait in milliseconds
	; return values:        On success  - true
	;                       On failure  - false, ErrorMsg / ErrorCode contain additional information
	; ======================
	SetTimeout(Timeout := 1000) {
		this.ErrorMsg := ""
		this.ErrorCode := 0
		this.SQL := ""
		if !(this._Handle) {
			this.ErrorMsg := "Invalid database handle!"
			return false
		}
		if Timeout Is Not Integer
			Timeout := 1000
		RC := DllCall("SQLite3.dll\sqlite3_busy_timeout", "Ptr", this._Handle, "Int", Timeout, "Cdecl Int")
		if (ErrorLevel) {
			this.ErrorMsg := "DllCall sqlite3_busy_timeout failed!"
			this.ErrorCode := ErrorLevel
			return false
		}
		if (RC) {
			this.ErrorMsg := this._ErrMsg()
			this.ErrorCode := RC
			return false
		}
		return true
	}
	; ======================
	; METHOD EscapeStr      Escapes special characters in a string to be used as field content
	; Parameters:           Str         - String to be escaped
	;                       Quote       - Add single quotes around the outside of the total string (true / false)
	; return values:        On success  - true
	;                       On failure  - false, ErrorMsg / ErrorCode contain additional information
	; ======================
	EscapeStr(ByRef Str, Quote := true) {
		this.ErrorMsg := ""
		this.ErrorCode := 0
		this.SQL := ""
		if !(this._Handle) {
			this.ErrorMsg := "Invalid database handle!"
			return false
		}
		if Str Is Number
			return true
		VarSetCapacity(OP, 16, 0)
		StrPut(Quote ? "%Q" : "%q", &OP, "UTF-8")
		this._StrToUTF8(Str, UTF8)
		Ptr := DllCall("SQLite3.dll\sqlite3_mprintf", "Ptr", &OP, "Ptr", &UTF8, "Cdecl UPtr")
		if (ErrorLevel) {
			this.ErrorMsg := "DllCall sqlite3_mprintf failed!"
			this.ErrorCode := ErrorLevel
			return false
		}
		Str := this._UTF8ToStr(Ptr)
		DllCall("SQLite3.dll\sqlite3_free", "Ptr", Ptr, "Cdecl")
		return true
	}
	; ======================
	; METHOD StoreBLOB      Use BLOBs as parameters of an INSERT/UPDATE/REPLACE statement.
	; Parameters:           SQL         - SQL statement to be compiled
	;                       BlobArray   - Array of objects containing two keys/value pairs:
	;                                     Addr : Address of the (variable containing the) BLOB.
	;                                     Size : Size of the BLOB in bytes.
	; return values:        On success  - true
	;                       On failure  - false, ErrorMsg / ErrorCode contain additional information
	; Remarks:              for each BLOB in the row you have to specify a ? parameter within the statement. The
	;                       parameters are numbered automatically from left to right starting with 1.
	;                       for each parameter you have to pass an object within BlobArray containing the address
	;                       and the size of the BLOB.
	; ======================
	StoreBLOB(SQL, BlobArray) {
		static SQLITE_static := 0
		static SQLITE_TRANSIENT := -1
		this.ErrorMsg := ""
		this.ErrorCode := 0
		if !(this._Handle) {
			this.ErrorMsg := "Invalid database handle!"
			return false
		}
		if !RegExMatch(SQL, "i)^\s*(INSERT|UPDATE|REPLACE)\s") {
			this.ErrorMsg := A_ThisFunc . " requires an INSERT/UPDATE/REPLACE statement!"
			return false
		}
		Query := 0
		this._StrToUTF8(SQL, UTF8)
		RC := DllCall("SQlite3.dll\sqlite3_prepare_v2", "Ptr", this._Handle, "Ptr", &UTF8, "Int", -1
					, "PtrP", Query, "Ptr", 0, "Cdecl Int")
		if (ErrorLeveL) {
			this.ErrorMsg := A_ThisFunc . ": DllCall sqlite3_prepare_v2 failed!"
			this.ErrorCode := ErrorLevel
			return false
		}
		if (RC) {
			this.ErrorMsg := A_ThisFunc . ": " . this._ErrMsg()
			this.ErrorCode := RC
			return false
		}
		for BlobNum, Blob In BlobArray {
			if !(Blob.Addr) || !(Blob.Size) {
				this.ErrorMsg := A_ThisFunc . ": Invalid parameter BlobArray!"
				this.ErrorCode := ErrorLevel
				return false
			}
			RC := DllCall("SQlite3.dll\sqlite3_bind_blob", "Ptr", Query, "Int", BlobNum, "Ptr", Blob.Addr
				, "Int", Blob.Size, "Ptr", SQLITE_static, "Cdecl Int")
			if (ErrorLeveL) {
				this.ErrorMsg := A_ThisFunc . ": DllCall sqlite3_bind_blob failed!"
				this.ErrorCode := ErrorLevel
				return false
			}
			if (RC) {
				this.ErrorMsg := A_ThisFunc . ": " . this._ErrMsg()
				this.ErrorCode := RC
				return false
			}
		}
		RC := DllCall("SQlite3.dll\sqlite3_step", "Ptr", Query, "Cdecl Int")
		if (ErrorLevel) {
			this.ErrorMsg := A_ThisFunc . ": DllCall sqlite3_step failed!"
			this.ErrorCode := ErrorLevel
			return false
		}
		if (RC) && (RC <> this._returnCode("SQLITE_DONE")) {
			this.ErrorMsg := A_ThisFunc . ": " . this._ErrMsg()
			this.ErrorCode := RC
			return false
		}
		RC := DllCall("SQlite3.dll\sqlite3_finalize", "Ptr", Query, "Cdecl Int")
		if (ErrorLevel) {
			this.ErrorMsg := A_ThisFunc . ": DllCall sqlite3_finalize failed!"
			this.ErrorCode := ErrorLevel
			return false
		}
		if (RC) {
			this.ErrorMsg := A_ThisFunc . ": " . this._ErrMsg()
			this.ErrorCode := RC
			return false
		}
		return true
	}
}
; =========================
; Exemplary custom callback function regexp()
; Parameters:        Context  -  handle to a sqlite3_context object
;                    ArgC     -  number of elements passed in Values (must be 2 for this function)
;                    Values   -  pointer to an array of pointers which can be passed to sqlite3_value_text():
;                                1. Needle
;                                2. Haystack
; return values:     Call sqlite3_result_int() passing 1 (true) for a match, otherwise pass 0 (false).
; =========================
SQLiteDB_RegExp(Context, ArgC, Values) {
	Result := 0
	if (ArgC = 2) {
		AddrN := DllCall("SQLite3.dll\sqlite3_value_text", "Ptr", NumGet(Values + 0, "UPtr"), "Cdecl UPtr")
		AddrH := DllCall("SQLite3.dll\sqlite3_value_text", "Ptr", NumGet(Values + A_PtrSize, "UPtr"), "Cdecl UPtr")
		Result := RegExMatch(StrGet(AddrH, "UTF-8"), StrGet(AddrN, "UTF-8"))
	}
	DllCall("SQLite3.dll\sqlite3_result_int", "Ptr", Context, "Int", !!Result, "Cdecl") ; 0 = false, 1 = trus
}