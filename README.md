# Classic ASP Micro-Framework

A compact set of Classic ASP (VBScript) helpers and small-framework pieces to speed up building data-driven pages and simple JSON APIs on Windows/IIS with ADODB and Microsoft SQL Server.

This workspace contains lightweight includes for database access, models, templating, table rendering, pagination, file upload handling, and debug utilities. The goal is practicality: small, readable code you can drop into existing Classic ASP projects and extend.

## Key features

- MSSQL wrapper (`framework/Database.asp`) — connection helpers, query execution and transaction helpers.
- Table-level helper (`framework/Table.asp`) — CRUD shortcuts, custom queries and OFFSET/FETCH paged results.
- Search form builder (`framework/SearchControl.asp`) — render search UIs and prepare parameterized ADODB.Command objects.
- Server-side table renderer (`framework/TableTemplate.asp`) — column formatters, links, sortable headers and expandable rows.
- Pagination helpers (`framework/Pagination.asp`) — URL helpers and pager rendering.
- Upload helper (`framework/Upload.asp`) — multipart parsing, sanitized file names, allow/deny extension lists and saved-file metadata.
- Debug helper (`framework/DebugHelper.asp`) — log levels, file logging, screen output and utilities to dump dictionaries, arrays and recordsets.
- Template engine, controller, model base and recordset wrappers for a lightweight MVC-like flow.
- Example pages and documentation in `docs/` including `docs/api_documentation.md`.

## Quick start

1. Copy this project into your IIS application folder (or include the `framework/` folder into your existing ASP site).
2. Include the helpers at the top of your ASP pages, for example:

```asp
<!--#include file="framework/Database.asp" -->
<!--#include file="framework/Table.asp" -->
<!--#include file="framework/TableTemplate.asp" -->
```

3. Use `MSSQLConnection` or `Controller.CreateDB()` to get a DB connection/wrapper. Use `Table` for CRUD and `Table.PagedList` for paged results.

4. For uploads, include `framework/Upload.asp` and configure `AllowedExtensions` or `DenyExtensions` as needed, then call `SaveRequestFiles` with a safe target directory.

5. For debugging during development, include `framework/DebugHelper.asp` and enable `LogToScreen`.

## Files (high-level)

- `framework/Database.asp` — DB wrapper
- `framework/Table.asp` — table/CRUD helper
- `framework/SearchControl.asp` — search form + WHERE builder
- `framework/TableTemplate.asp` — HTML table renderer
- `framework/Pagination.asp` — pager helpers
- `framework/Upload.asp` — upload parsing and saving (sanitize + allow/deny lists)
- `framework/DebugHelper.asp` — logging and dumps
- `framework/TemplateEngine.asp`, `framework/Controller.asp`, `framework/Model.asp` — templating and MVC scaffolding
- `framework/RecordsetWrapper.asp`, `framework/FormWrapper.asp` — convenience helpers
- `docs/api_documentation.md` — generated reference for the framework includes

## Security & production notes

- Always use parameterized queries or prepared ADODB.Command objects to avoid SQL injection (the helpers provide parameter support).
- Configure `Upload` with a restrictive `AllowedExtensions` list for production and sanitize user-supplied input.
- Prefer reusing a single ADODB.Connection for transaction scopes. Some helpers open connections internally; pass your own ADODB.Connection when you need transaction scope.

## Next steps & suggestions

- Add JSON body parsing for API endpoints (handle application/json for POST/PUT requests).
- Add more explicit error reporting for rejected uploads (reasons and counts).
- Add automated unit tests (limited options for Classic ASP — consider integration tests or smoke scripts).
- Optionally convert docs to HTML/PDF for sharing.
