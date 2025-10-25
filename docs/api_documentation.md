# Classic ASP Framework — API Documentation

_Auto-generated reference for the project's `framework/` includes. Date: October 25, 2025._

---

## Files in `framework/`

- `Database.asp`
- `SearchControl.asp`
- `Table.asp`
- `TableTemplate.asp`
- `Pagination.asp`
- `Upload.asp`
- `Controller.asp`
- `TemplateEngine.asp`
- `Model.asp`
- `RecordsetWrapper.asp`
- `FormWrapper.asp`
- `StoredProcedure.asp`
- `EnvHelper.asp`
- `HTTPHelper.asp`
- `IO.asp`

---

## Overview

This document summarizes the public purpose and API surface of each include in the `framework/` folder so you can quickly discover classes and helper functions when building Classic ASP pages and APIs.

Each file lists the main classes, common public methods, expected inputs/outputs and short usage examples where helpful.

---

## Database.asp

````markdown
# Classic ASP Framework — API Documentation

_Auto-generated reference for the project's `framework/` includes. Date: October 25, 2025._

---

## Files in `framework/`

- `Database.asp`
- `SearchControl.asp`
- `Table.asp`
- `TableTemplate.asp`
- `Pagination.asp`
- `Upload.asp`
- `DebugHelper.asp`
- `Controller.asp`
- `TemplateEngine.asp`
- `Model.asp`
- `RecordsetWrapper.asp`
- `FormWrapper.asp`
- `StoredProcedure.asp`
- `EnvHelper.asp`
- `HTTPHelper.asp`
- `IO.asp`

---

## Overview

This document summarizes the public purpose and API surface of each include in the `framework/` folder so you can quickly discover classes and helper functions when building Classic ASP pages and APIs.

Each file lists the main classes, common public methods, expected inputs/outputs and short usage examples where helpful.

---

## Database.asp

**Purpose:** MSSQL connection wrapper and utility helpers built on ADODB.

**Class:** `MSSQLConnection`

Properties

- `ConnectionString` — connection string used to open ADODB.Connection.

Methods

- `Connect()` — opens and returns an ADODB.Connection object.
- `Close()` — closes the connection.
- `ExecuteQuery(sql)` — executes a SQL SELECT and returns an ADODB.Recordset.
- `ExecuteNonQuery(sql)` — executes INSERT/UPDATE/DELETE; returns affected row count.
- `GetScalar(sql)` — runs a scalar query and returns the first column of the first row.
- `BeginTrans(), CommitTrans(), RollbackTrans()` — transaction helpers.
- `GetLastError()` — returns last recorded error text.

> Notes: Several helpers that build parameterised ADODB.Command objects (in other includes) accept either an ADODB.Connection or a wrapper instance. If you need reuse or transactions across multiple commands, open the connection and reuse it directly via the wrapper's `Connect()` method.

---

## SearchControl.asp

**Purpose:** Build search UIs and translate fields into a parameterized WHERE clause and prepared ADODB.Command objects.

**Class:** `SearchControl`

Public methods

- `AddField(name, label, type, attrs, options)` — register a search field. Supported types: `text`, `select`, `checkbox`, `date`, `textarea`.
- `Render(target, method)` — renders the search form HTML targeting a URL.
- `GetValue(name)` — get the posted/queried value for a field.
- `BuildWhereFromFields()` — constructs a parameterized WHERE clause and an ordered params array. Returns an object/Dictionary with keys: `where` (string) and `params` (array).
- `PrepareCommand(baseSql, connectionOrWrapper)` — convenience to append the WHERE fragment to `baseSql`, create an ADODB.Command, attach parameters in correct order and return the prepared command ready for Execute.

---

## Table.asp

**Purpose:** Lightweight table-level helper for CRUD, counting, custom queries and paged results.

**Class:** `Table`

Public methods

- `Init(tableName, primaryKey, dbWrapper)` — initialize with the table name, primary key column (string) and a database wrapper instance (e.g., `MSSQLConnection`).
- `List(where, paramsArr, orderBy, limit, offset)` — return a Recordset or array of rows matching criteria.
- `FindById(id)` — returns single row by primary key (dictionary/object).
- `Insert(dict)` — inserts a dictionary/object of column->value pairs, returns new id when possible.
- `Update(id, dict)` — updates row identified by id using column values from dict.
- `Delete(id)` — deletes by primary key.
- `Count(where, paramsArr)` — returns total count matching optional where clause.
- `Query(baseSql, paramsArr)` — run a custom SQL with parameter array; returns rows or Recordset depending on usage.
- `PagedList(baseSql, paramsArr, orderBy, pageSize, pageNumber)` — returns a dictionary with keys: `items` (array), `total` (int), `page`, `pageSize`. Uses SQL `OFFSET/FETCH` for efficient paging.

> Implementation notes: The class accepts either an ADODB.Connection or the project's DB wrapper. It internally converts params arrays to command parameters when preparing queries.

---

## TableTemplate.asp

**Purpose:** Server-side HTML table renderer with column formatters, links, sortable headers and expandable/nested rows.

Public methods

- `RenderTable(headers, rows, options)` — renders a complete HTML table. `headers` is an array describing columns (field name, label, sortable boolean, formatter callback name or key, link template, etc.).
- `RenderTableWithPager(tableObj, baseSql, paramsArr, headers, options)` — convenience wrapper that calls `Table.PagedList` and then `RenderTable` and the pagination helper.

> Options supported: column-specific formatter callbacks (date/datetime/currency/custom), per-cell link templates that substitute values like `{id}`, and expandable-row templates which render hidden rows with extra details toggled via small JS.

---

## Pagination.asp

**Purpose:** Render pagers and helper URL builders.

Public methods

- `RenderPagination(total, currentPage, pageSize, options)` — render pager HTML with first/prev/numbered pages/next/last and pageSize choices.
- `MergeQueryString(newParams)` — utility to merge/override querystring parameters into the current URL.
- `BuildPageUrl(page, pageSize)` — returns a URL for a given page and pageSize.
- `GetPagerParams(defaultPageSize)` — parse and normalize `page` and `pageSize` from Request with validation.
- `RenderPaginationFromRequest(total, defaultPageSize, options)` — helper that reads page params from Request and prints a ready pager for the given total.

---

## Upload.asp

**Purpose:** Parse multipart/form-data uploads and save file parts to disk (ADODB.Stream).

**Class:** `Upload`

Public fields & methods

- `Files` (array) — after parsing, contains dictionaries with keys: `FieldName`, `FileName`, `SavedPath`, `Size`, `ContentType`.
- `MaxFileSize`, `AllowedTypes` — configuration for upload size and MIME restrictions.
- `SaveRequestFiles(targetDir)` — parse the current request (multipart), write parts to the `targetDir` and populate `Files`. Returns number of files saved.

Additional properties and behavior (recent updates):

- `AllowedExtensions` — optional array of allowed file extensions (without the leading dot). If set, only files with these extensions are accepted (case-insensitive).
- `DenyExtensions` — optional array of denied extensions (without dot). Used only when `AllowedExtensions` is not set. Files with extensions in this list are rejected.
- `SavedName` — each saved file dictionary now contains a `SavedName` value: the sanitized, unique filename actually written to disk (separate from the original `FileName`).
- `SanitizeFileName(baseName)` — helper function used internally to remove unsafe characters and normalize the filename base before appending the unique suffix.

Behavior notes:

- If `AllowedExtensions` is provided it takes precedence (only those extensions are accepted). If not provided and `DenyExtensions` is provided, files with denied extensions are skipped.
- Extension checks are case-insensitive and operate on the filename extension extracted from the uploaded filename.
- Skipped files are not saved and are not added to the `Files` array; currently no per-file error reasons are returned by default (this can be added if you want rejection reasons recorded).

---

## DebugHelper.asp

**Purpose:** Lightweight debug logger and dump utilities for Classic ASP.

**Class:** `DebugHelper`

Public members & methods

- `LEVEL_FATAL, LEVEL_ERROR, LEVEL_WARN, LEVEL_INFO, LEVEL_DEBUG` — constants defining log levels.
- `LogLevel` — numeric threshold. Messages with level higher (less severe) than this are ignored.
- `LogToScreen` — boolean, when true messages are written to Response.
- `LogToFile(path)` — set an on-disk log file; the helper will append UTF-8 text to this file.
- `Log(msg, level)` — emit a log message.
- `DumpDict(dict, title)` — write all keys/values from a Scripting.Dictionary to the log.
- `DumpArray(arr, title)` — log array contents.
- `DumpRS(rs, title, maxRows)` — log recordset columns and up to `maxRows` rows (default 50).
- `Trace(msg)` — convenience wrapper for a debug-level trace entry.

Usage:

```vbscript
' <!--#include file="framework/DebugHelper.asp" -->
Dim dbg: Set dbg = New DebugHelper
dbg.LogToScreen = True
dbg.LogLevel = dbg.LEVEL_DEBUG
dbg.Log "Starting processing...", dbg.LEVEL_INFO
dbg.DumpDict someDict, "myDict"
```

---

## Controller.asp

**Purpose:** Small page/controller helper for Classic ASP pages and simple JSON endpoints.

**Class:** `Controller`

Public methods

- `Init()` — bootstrap common includes (.env loader, DB factory), and provide `Request` helpers.
- `Param(name, default)` — read GET/POST/request parameters in one place.
- `Render(templatePath, locals)` — render a template via the project's TemplateEngine.
- `Json(obj)` — send JSON response (with proper headers).
- `Redirect(url)` — send a redirect header.
- `CreateDB()` — convenience factory that reads `.env` and returns a configured DB wrapper (`MSSQLConnection`).

---

## TemplateEngine.asp

**Purpose:** Templating helper to load template files and render with variables.

Public methods

- `Load(path)` — read a template file into memory.
- `Set(key, value)` — set a variable available to the template.
- `SetDict(dict)` — set many variables at once.
- `Render()` — return the rendered HTML as a string.
- `RenderToResponse()` — write rendered HTML to Response and stop execution if needed.

---

## Model.asp

**Purpose:** BaseModel to provide per-entity data helpers built on top of the `Table` helper.

**Class:** `BaseModel`

Public methods

- `TableName`, `PrimaryKey` — configuration properties per model.
- `GetDB()` — returns DB wrapper instance (reads env or uses Controller.CreateDB).
- `FindById(id)`, `FindAll()`, `Insert(dict)`, `Update(id,dict)`, `Delete(id)` — convenience CRUD wrappers that call into the `Table` class.
- `RecordsetToDict(rs)` — helper to convert ADODB.Recordset into a dictionary/object or array of dictionaries.

---

## RecordsetWrapper.asp

**Purpose:** Small convenience wrapper around ADODB.Recordset to simplify reading values and converting to native arrays/JSON.

**Class:** `RecordsetWrapper`

Public methods

- `Init(rs)` — attach an ADODB.Recordset.
- `FieldExists(name)` — check field existence on the current recordset.
- `GetValue(name)` — returns a safe value (handles NULLs).
- `ToArrayOfDicts()`, `ToJSON()` — convenience conversion helpers.

---

## FormWrapper.asp

**Purpose:** Helpers to build HTML form controls pre-populated from a recordset or data dictionary.

Public methods

- `Init(rsOrDict)` — pass a recordset row or dictionary to pre-populate values.
- `FieldValue(name)` — return value for field.
- Render helpers: `InputText`, `Hidden`, `TextArea`, `Checkbox`, `SelectFromRS` etc. — each returns an HTML string for the control.

---

## StoredProcedure.asp

**Purpose:** Lightweight wrapper to prepare and execute stored procedures using ADODB.Command.

Public methods

- `Init(connOrWrapper)` — attach a connection or wrapper.
- `AddParam(name, value, type, direction)` — add input/output params.
- `Execute()` — execute the stored procedure and return results / output parameter values.
- `GetOutput(name)` — retrieve an output parameter after execution.

---

## EnvHelper.asp

**Purpose:** Simple .env loader and replacer for configuration values.

Public methods

- `LoadEnv(path)` — parse a .env file and load entries into a dictionary (or into Server variables).
- `GetEnv(key, default)` — return env value or default.
- `ReplaceVars(str)` — substitute environment variables into a string.

---

## HTTPHelper.asp

**Purpose:** Helpers to determine and normalize HTTP methods used by the client (supports method override patterns).

Public methods

- `GetHttpMethod()` — returns normalized method name (GET, POST, PUT, DELETE, etc.).
- `IsGet(), IsPost(), IsPut(), IsDelete()` — boolean helpers.

---

## IO.asp

**Purpose:** Small filesystem helpers using the FileSystemObject and ADODB.Stream when needed.

Public methods

- `MkDirIfNotExists(path)` — create directory recursively.
- `ReadFile(path)`, `WriteFile(path, contents)` — basic file IO convenience wrappers.
- `FileExists(path)`, `DeleteFile(path)` — other convenience helpers.

---

## Examples & usage

### 1) Basic SELECT with `MSSQLConnection`

```vbscript
' include the wrapper
' <!--#include file="framework/Database.asp" -->
Dim db
Set db = Server.CreateObject("MSSQLConnection")
db.ConnectionString = GetEnv("DATABASE_URL")
Set conn = db.Connect()
Set rs = db.ExecuteQuery("SELECT TOP 10 * FROM Users ORDER BY id DESC")
' process rs or use RecordsetWrapper
```

### 2) Using `Table` for PagedList

```vbscript
' include framework/Table.asp and pagination helpers
' <!--#include file="framework/Table.asp" -->
Dim usersTable
Set usersTable = New Table
usersTable.Init "Users", "id", db
Dim pageResult
Set pageResult = usersTable.PagedList("SELECT * FROM Users WHERE active = 1", Array(), "id DESC", 20, 1)
' pageResult has keys: items (array), total, page, pageSize
```

### 3) Building a searchable list with `SearchControl`

```vbscript
' create a search control, add fields, render form
' <!--#include file="framework/SearchControl.asp" -->
Dim sc
Set sc = New SearchControl
sc.AddField "q", "Query", "text"
sc.AddField "status", "Status", "select", , Array(Array("1","Active"),Array("0","Inactive"))
sc.Render "", "GET"
Dim cmd
Set cmd = sc.PrepareCommand("SELECT * FROM Users", db)
Set rs = cmd.Execute()
```

### 4) Upload usage

```vbscript
' <!--#include file="framework/Upload.asp" -->
Dim uploader
Set uploader = New Upload
uploader.MaxFileSize = 10 * 1024 * 1024 ' 10MB
Dim count
count = uploader.SaveRequestFiles(Server.MapPath("/uploads"))
' uploader.Files contains saved metadata
```

Example: allow-list extensions (recommended)

```vbscript
' Include the Upload helper
' <!--#include file="framework/Upload.asp" -->
Dim uploader
Set uploader = New Upload
uploader.AllowedExtensions = Array("jpg","jpeg","png","pdf")
uploader.MaxFileSize = 5 * 1024 * 1024
Dim saved
saved = uploader.SaveRequestFiles(Server.MapPath("/uploads"))
' Each entry in uploader.Files now includes:
'  - FileName (original name)
'  - SavedName (sanitized unique name on disk)
'  - SavedPath, Size, ContentType, FieldName
```

Example: deny-list extensions

```vbscript
' <!--#include file="framework/Upload.asp" -->
Dim uploader
Set uploader = New Upload
uploader.DenyExtensions = Array("exe","asp","aspx","php")
uploader.SaveRequestFiles Server.MapPath("/uploads")
```

---

## Notes & recommendations

- For real production APIs consider exposing and re-using a single ADODB.Connection when executing multiple commands or when using transactions. Some helpers open a connection internally; prefer passing your own connection when you need transaction scope.
- Harden `Upload.asp`: sanitize filenames, generate unique names, stream large uploads and validate file contents (MIME sniffing) before storing on disk.
- REST endpoints (if present) currently may accept form-encoded bodies. Add a small JSON body parser to support `application/json` payloads.
- All SQL helpers use parameter arrays. Always use parameterized queries to avoid SQL injection.

---

*Generated documentation for the `framework` directory.*
