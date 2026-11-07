# Word MCP for Windows — Design Spec

- **Ngày:** 2026-11-07
- **Trạng thái:** Đã duyệt thiết kế (chat), chờ duyệt spec
- **Phạm vi:** Sub-project 1 của "Microsoft Office MCP for Windows" — chỉ Word. Excel/PowerPoint là sub-project sau, dùng lại `core`.

## 1. Mục tiêu

Một MCP server (stdio) cho phép Claude **đọc và điều khiển Microsoft Word đang chạy** trên Windows qua COM automation, phủ toàn bộ object model của Word 16.0 (Word 2016–2024 / Microsoft 365).

### Tiêu chí thành công
1. Claude Code/Claude Desktop đăng ký được server `word-mcp` và thấy danh sách tool.
2. Một luồng end-to-end chạy được trên Word thật: mở tài liệu → đọc có đánh số đoạn → sửa đoạn → định dạng → bật track changes → chèn và cập nhật TOC → xuất PDF.
3. Bất kỳ thuộc tính/phương thức nào trong object model đều truy cập được qua gateway (`com_get/com_set/com_call`), không cần tool chuyên biệt.
4. Không treo khi Word bận hoặc có hộp thoại mở: trả lỗi rõ ràng trong thời gian giới hạn.

### Ngoài phạm vi
- Xử lý file .docx offline không cần Word (đã có skill docx của Anthropic).
- Tính năng cloud/AI không có trong object model: Copilot, Designer, Editor AI, Dictate, đồng soạn thảo thời gian thực.
- Word for Mac, Word Online.
- Thay đổi cài đặt bảo mật Office hộ người dùng (ví dụ Trust Center).

## 2. Môi trường mục tiêu (đã kiểm tra trên máy dev)

| Thành phần | Giá trị |
|---|---|
| OS | Windows 11 Pro |
| Office | Microsoft 365 (O365ProPlusRetail), 16.0.20527, x64 |
| ProgID | `Word.Application.16` |
| Python | 3.11 (`C:\Users\FPT5\AppData\Local\Programs\Python\Python311`), quản lý bằng `uv` |

Giả định (tin cậy Cao): code viết cho object model 16.0 chạy trên Word 2024 bản perpetual. API chỉ có trên M365 phải được bắt lỗi khi gọi, không được làm sập server.

## 3. Kiến trúc

```
E:\Microsoft office mcp for window\
  pyproject.toml                 # hatchling, deps: mcp>=2.3,<3, pywin32>=306
  README.md
  src/office_mcp/
    __init__.py
    core/
      com_thread.py              # ComWorker: 1 luồng STA + queue
      connection.py              # get_app(progid): attach hoặc launch
      errors.py                  # HRESULT → OfficeMcpError
      gateway.py                 # parser + resolver đường dẫn COM, list_members
      safety.py                  # blocklist, cờ ALLOW_VBA
      serialize.py               # COM value → JSON-safe
    word/
      server.py                  # FastMCP app, main() → entry `word-mcp`
      context.py                 # resolve document + range locator
      tools/
        documents.py  reading.py  content.py  formatting.py
        tables.py     objects.py  layout.py   review.py
        references.py export.py   gateway_tools.py  vba.py
  tests/
    unit/                        # không cần Word
    integration/                 # cần Word, marker `word`
```

### 3.1 ComWorker (`core/com_thread.py`)
- Một luồng daemon gọi `pythoncom.CoInitialize()` (STA) một lần, chạy vòng lặp lấy job từ `queue.Queue`.
- API: `async def run(fn, *args, timeout=60) -> Any`. Tool async gọi `await worker.run(...)`. Bên trong là `concurrent.futures.Future` + `asyncio.wrap_future`.
- Mọi object COM chỉ được tạo, dùng và giải phóng trên luồng này. Không trả object COM ra ngoài worker: chỉ trả giá trị đã serialize.
- Đăng ký `IMessageFilter` (qua `pythoncom.CoRegisterMessageFilter`) để tự retry khi Word trả `RPC_E_CALL_REJECTED` (0x80010001) hoặc `RPC_E_SERVERCALL_RETRYLATER` (0x8001010A). Retry trong tối đa 10 giây rồi báo lỗi `WORD_BUSY`.
  - Nếu message filter không dùng được với pywin32 thì dùng phương án dự phòng: vòng retry thủ công quanh lệnh gọi, bắt `pywintypes.com_error` theo HRESULT.

### 3.2 Connection (`core/connection.py`)
- `get_app()`: thử `win32com.client.GetActiveObject("Word.Application")`. Nếu không có Word đang chạy thì `win32com.client.DispatchEx("Word.Application")` rồi đặt `Visible = True`.
- Dùng late binding (`dynamic.Dispatch`) làm mặc định, để không phụ thuộc cache gen_py (cache này dễ hỏng). Hằng số Word (`wdFormatPDF`...) được khai báo trong `word/constants.py`.
- Cache đối tượng app. Trước mỗi job, kiểm tra app còn sống bằng cách đọc `app.Name`. Nếu lỗi (Word đã bị đóng) thì kết nối lại một lần.
- `DisplayAlerts`: lưu giá trị cũ, đặt `0` (wdAlertsNone) trong job, khôi phục trong `finally`.

### 3.3 Lỗi (`core/errors.py`)
- `OfficeMcpError(code, message, hint=None)`. Tool trả lỗi MCP (`isError`) với text `"[CODE] message — hint"`.
- Mã lỗi: `WORD_NOT_AVAILABLE`, `WORD_BUSY`, `DOC_NOT_FOUND`, `RANGE_INVALID`, `COM_ERROR` (kèm HRESULT hex, mô tả từ `excepinfo`, đường dẫn COM), `BLOCKED`, `VBA_DISABLED`, `VBA_NOT_TRUSTED`, `UNSAVED_CHANGES`, `FILE_EXISTS`, `TIMEOUT`.

### 3.4 Serialize (`core/serialize.py`)
- Chuyển giá trị COM: None/bool/int/float/str giữ nguyên; `pywintypes.datetime` → chuỗi ISO; tuple → list (đệ quy); object COM → `{"_com_object": "<TypeName>"}` (không trả object COM ra ngoài).
- Chuỗi dài: tool đọc có tham số `max_chars` (mặc định 50 000). Khi bị cắt thì trả cờ `truncated: true` và `next_offset`.

## 4. Định vị tài liệu và vùng (`word/context.py`)

### Tài liệu
Mọi tool thao tác tài liệu có tham số tùy chọn `document`:
- bỏ trống → `ActiveDocument`
- tên (`Report.docx`) → khớp `Document.Name`, không phân biệt hoa thường
- đường dẫn đầy đủ → khớp `Document.FullName`
- không có tài liệu khớp → lỗi `DOC_NOT_FOUND` kèm danh sách tài liệu đang mở

### Range locator
Tham số `target` là object, có đúng **một** trong các khóa sau:
| Khóa | Ý nghĩa |
|---|---|
| `{"paragraph": n}` | đoạn thứ n (1-based), bỏ dấu ¶ cuối khi thay thế nội dung |
| `{"paragraphs": [a, b]}` | từ đoạn a đến đoạn b (bao gồm cả hai) |
| `{"start": s, "end": e}` | vị trí ký tự (`Document.Range(s, e)`) |
| `{"bookmark": "name"}` | range của bookmark |
| `{"selection": true}` | vùng đang chọn |
| `{"document": true}` | toàn bộ nội dung chính |
| `{"table": t, "cell": [r, c]}` | ô (1-based) |

Locator sai (index ngoài phạm vi, bookmark không tồn tại) → lỗi `RANGE_INVALID`, nói rõ giới hạn hợp lệ, ví dụ "tài liệu có 42 đoạn".

## 5. Danh mục tool

Tên tool có tiền tố `word_`. Tham số `document` (mục 4) được hiểu là có sẵn ở mọi tool thao tác tài liệu, không ghi lại.

### 5.1 documents
- `word_list_documents()` → `[{name, full_name, saved, active, paragraphs, pages}]`
- `word_open(path, read_only=false, visible=true)`
- `word_create(template=None)`
- `word_save()`; `word_save_as(path, format="docx"|"doc"|"rtf"|"txt"|"html"|"odt"|"pdf"|"xps", overwrite=false)`
- `word_close(save=false, force=false)`: nếu có thay đổi chưa lưu, `save=false` và `force=false` thì báo lỗi `UNSAVED_CHANGES`
- `word_activate()`

### 5.2 reading
- `word_read_text(start_paragraph=1, end_paragraph=None, max_chars=50000, include_style=true)` → mỗi đoạn có `{index, style, text, start, end}`
- `word_get_outline()` → các heading kèm level và số đoạn
- `word_get_structure()` → số section, page, bảng, inline shape, shape, field, comment, revision, bookmark, hyperlink
- `word_get_selection()` → text, start, end, style
- `word_get_document_info()` → built-in properties, số từ/ký tự/trang, track changes đang bật hay tắt, protection
- `word_get_range_formatting(target)` → font, đoạn văn, style

### 5.3 content
- `word_insert_text(text, target, position="after"|"before"|"replace", style=None)`
- `word_insert_paragraph(text, after_paragraph, style=None)`
- `word_delete(target)`
- `word_find(text, match_case, whole_word, wildcards, max_results=100)` → danh sách `{start, end, paragraph, context}`
- `word_find_replace(find, replace, match_case, whole_word, wildcards, replace_all=true, target=None)` → số chỗ đã thay
- `word_select(target)`
- `word_undo(times=1)`, `word_redo(times=1)`

### 5.4 formatting
- `word_format_font(target, name, size, bold, italic, underline, color, highlight, strikethrough, superscript, subscript, all_caps, small_caps)` (mọi tham số đều tùy chọn)
- `word_format_paragraph(target, alignment, left_indent, right_indent, first_line_indent, space_before, space_after, line_spacing, line_spacing_rule, keep_with_next, page_break_before)`
- `word_apply_style(target, style)`
- `word_list_styles(type=None, in_use_only=false)`
- `word_create_or_modify_style(name, type, base_style, font..., paragraph...)`
- `word_apply_list(target, kind="bullet"|"number"|"outline"|"none", level=None)`
- `word_clear_formatting(target)`

### 5.5 tables
- `word_insert_table(rows, cols, target, data=None, style=None, header_row=true)`
- `word_read_table(table)` → mảng 2 chiều chứa text
- `word_set_cells(table, cells=[{row, col, text}])`
- `word_table_add_row(table, after=None)`, `word_table_add_column(table, after=None)`
- `word_table_delete_rows(table, rows)`, `word_table_delete_columns(table, cols)`
- `word_table_merge_cells(table, from_cell, to_cell)`
- `word_table_format(table, style, autofit, borders, shading)`
- `word_delete_table(table)`

### 5.6 objects
- `word_insert_picture(path, target, width=None, height=None, inline=true)`
- `word_insert_shape(kind, left, top, width, height, text=None)` (bao gồm textbox)
- `word_list_shapes()`: liệt kê cả inline shape lẫn shape trôi nổi
- `word_insert_caption(target, label="Figure"|"Table"|"Equation", text)`

### 5.7 layout
- `word_page_setup(section=None, orientation, paper_size, margins..., columns)`
- `word_insert_break(target, kind="page"|"section_next"|"section_continuous"|"column"|"line")`
- `word_set_header_footer(section, which="header"|"footer", kind="primary"|"first"|"even", text)`
- `word_insert_page_numbers(position, alignment, start_at=None)`
- `word_list_sections()`

### 5.8 review
- `word_track_changes(enabled)`
- `word_list_revisions()` → `{index, type, author, date, text, start, end}`
- `word_accept_revisions(indices=None)`, `word_reject_revisions(indices=None)` (`None` = tất cả)
- `word_add_comment(target, text)`, `word_list_comments()`, `word_reply_comment(index, text)`, `word_resolve_comment(index, resolved=true)`, `word_delete_comment(index)`
- `word_protect(type, password=None)`, `word_unprotect(password=None)`: chỉ nhận password do người dùng đưa trong chat

### 5.9 references
- `word_insert_toc(target, levels="1-3", hyperlinks=true)`, `word_update_toc()`
- `word_update_fields(target=None)`
- `word_insert_field(target, code)`
- `word_add_bookmark(name, target)`, `word_list_bookmarks()`, `word_delete_bookmark(name)`
- `word_add_hyperlink(target, address, text=None)`, `word_list_hyperlinks()`
- `word_insert_footnote(target, text, kind="footnote"|"endnote")`
- `word_insert_cross_reference(target, ref_type, ref_item, kind)`

### 5.10 export
- `word_export_pdf(path, range="all"|"current_page"|{from,to}, overwrite=false, open_after=false)` (dùng `ExportAsFixedFormat`)
- `word_print(copies=1, pages=None, printer=None)`: in ra máy in thật, mô tả tool phải ghi rõ điều này
- `word_mail_merge(data_source, destination="new_document"|"printer"|"email", ...)`: đích `email` gửi thư thật, mô tả tool phải cảnh báo

### 5.11 gateway (`core/gateway.py` + `gateway_tools.py`)
- Cú pháp đường dẫn: `Root(.Member)*`, với `Member := Name | Name(args)`. Root là `Application` (mặc định), `ActiveDocument`, `Selection`, hoặc `Documents("name")`. Arg là JSON literal: số, chuỗi trong nháy kép, true/false/null.
  - Ví dụ: `ActiveDocument.Sections(1).PageSetup.Orientation`, `ActiveDocument.Paragraphs(3).Range.Font.Bold`.
- `word_com_get(path)` → giá trị đã serialize
- `word_com_set(path, value)`: gán giá trị cho thuộc tính cuối của đường dẫn
- `word_com_call(path, args=[], kwargs={})`: gọi phương thức cuối, trả kết quả
- `word_list_members(path)` → `[{name, kind: "property"|"method", args, doc?}]` lấy qua `_oleobj_.GetTypeInfo()`. Nếu đối tượng không có type info thì báo lỗi rõ ràng (đã biết trước là giới hạn).
- `word_execute_mso(id_mso)`: chạy lệnh ribbon qua `CommandBars.ExecuteMso`. Mô tả tool phải cảnh báo lệnh này có thể mở hộp thoại; nếu hộp thoại mở thì sẽ gặp `WORD_BUSY`.

### 5.12 vba
- `word_run_vba(code, entry="Main")`: thêm module tạm vào `NormalTemplate.VBProject`, chạy `Application.Run`, rồi xóa module trong `finally`. Trả giá trị trả về của hàm entry (nếu có).
- Cần bật `WORD_MCP_ALLOW_VBA=1`; nếu không thì báo `VBA_DISABLED`. Nếu Trust Center chưa bật "Trust access to the VBA project object model" thì báo `VBA_NOT_TRUSTED` kèm hướng dẫn để người dùng tự bật.

## 6. An toàn (`core/safety.py`)

- Gateway bật mặc định. Có blocklist tên member, không phân biệt hoa thường, áp dụng cho mọi đoạn trong đường dẫn: `Quit`, `Kill`, `Shell`, `SendKeys`, `VBProject`, `VBE`, `MacroContainer`, `OrganizerDelete`. (`word_run_vba` dùng `VBProject` nội bộ, không đi qua gateway nên không bị chặn.)
  - Vi phạm → lỗi `BLOCKED`.
  - Có thể tắt bằng `WORD_MCP_UNSAFE_GATEWAY=1`.
- Blocklist **chỉ để tránh lỗi vô tình**, không phải ranh giới bảo mật. README phải ghi rõ: ai điều khiển được Claude thì cũng điều khiển được Word với quyền của user hiện tại.
- `word_run_vba` tắt mặc định (xem 5.12).
- Mặc định không làm mất dữ liệu: `close` không lưu, `save_as`/`export_pdf` không ghi đè file đã có.
- Server không tự đổi registry hay cài đặt Trust Center.

## 7. Cấu hình

| Biến môi trường | Mặc định | Ý nghĩa |
|---|---|---|
| `WORD_MCP_ALLOW_VBA` | `0` | Cho phép `word_run_vba` |
| `WORD_MCP_UNSAFE_GATEWAY` | `0` | Tắt blocklist gateway |
| `WORD_MCP_CALL_TIMEOUT` | `60` | Timeout mỗi job (giây) |
| `WORD_MCP_BUSY_RETRY` | `10` | Thời gian retry khi Word bận (giây) |

Đăng ký với Claude Code:
```
claude mcp add word -- "E:\Microsoft office mcp for window\.venv\Scripts\word-mcp.exe"
```

## 8. Kiểm thử

- **Unit test** (`tests/unit`, không cần Word):
  - parser đường dẫn gateway: cú pháp hợp lệ/không hợp lệ, arg là chuỗi/số
  - blocklist
  - đổi HRESULT thành lỗi
  - serialize
  - kiểm tra định dạng locator
- **Integration test** (`tests/integration`, marker `@pytest.mark.word`):
  - fixture mở một tài liệu mới bằng `DispatchEx` (instance Word riêng, không đụng tài liệu của người dùng) và đóng trong teardown
  - mỗi nhóm tool có ít nhất một test đọc và một test ghi
  - bắt buộc có: `find_replace`, track changes và accept, TOC và cập nhật, `export_pdf` ra thư mục tạm, gateway get/set, `list_members`
- **End-to-end:** đăng ký server vào Claude Code và chạy luồng ở tiêu chí thành công #2.
- Integration test chỉ chạy được trên máy có Word, không chạy được trên CI Linux.

## 9. Rủi ro đã biết

| Rủi ro | Mức độ | Cách giảm |
|---|---|---|
| Message filter của pywin32 không hoạt động ổn định | Trung bình | Phương án dự phòng: vòng retry thủ công (3.1) |
| Instance Word bị treo do hộp thoại modal (`ExecuteMso`, `print`) | Trung bình | Timeout job; thông báo `WORD_BUSY`; cảnh báo trong mô tả tool |
| `list_members` không có type info cho một số đối tượng | Thấp–Trung bình | Báo lỗi rõ; Claude chuyển sang dùng tài liệu VBA |
| Mô tả ~65 tool chiếm nhiều context | Thấp | Mô tả ngắn gọn; đo thực tế sau khi build |
| Integration test đụng vào tài liệu người dùng | Thấp | Dùng `DispatchEx` để có instance riêng |
