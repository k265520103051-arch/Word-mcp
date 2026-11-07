# word-mcp

MCP server (Python) điều khiển Microsoft Word qua COM trên Windows. 78 công cụ, tên đều bắt đầu bằng `word_`.

## Yêu cầu

- Windows, Microsoft Word 2016–2024 hoặc Microsoft 365 (đã cài, bản desktop)
- Python 3.11
- [uv](https://docs.astral.sh/uv/)

## Cài đặt

```
uv sync
```

## Đăng ký

Claude Code:

```
claude mcp add word -- "E:\Microsoft office mcp for window\.venv\Scripts\word-mcp.exe"
```

Claude Desktop (`claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "word": {
      "command": "E:\Microsoft office mcp for window\.venv\Scripts\word-mcp.exe"
    }
  }
}
```

Sau khi đăng ký, khởi động lại phiên Claude.

## Biến môi trường

| Biến | Mặc định | Ý nghĩa |
|---|---|---|
| `WORD_MCP_ALLOW_VBA` | `0` | `1` để bật `word_run_vba` (xem phần VBA) |
| `WORD_MCP_UNSAFE_GATEWAY` | `0` | `1` để tắt danh sách chặn của gateway (không khuyến nghị) |
| `WORD_MCP_CALL_TIMEOUT` | `60` | Số giây tối đa cho một lệnh, quá hạn trả `TIMEOUT` |
| `WORD_MCP_BUSY_RETRY` | `10` | Số giây thử lại khi Word bận/đang có hộp thoại, quá hạn trả `WORD_BUSY` |

## Định vị vùng (locator) và đơn vị

Hầu hết công cụ nhận `document` (tên hoặc đường dẫn đầy đủ; bỏ trống = tài liệu đang hoạt động) và `target` với đúng một cách định vị:

`{paragraph:n}`, `{paragraphs:[a,b]}`, `{start,end}`, `{bookmark}`, `{selection:true}`, `{document:true}`, `{table, cell:[r,c]}`.

- Độ dài tính bằng point; màu dạng `"#RRGGBB"`.
- Đường dẫn tệp phải là đường dẫn tuyệt đối, nếu không trả `[RANGE_INVALID] Path must be absolute`.
- Không ghi đè tệp có sẵn theo mặc định (`word_save_as`, `word_export_pdf`); `word_close` từ chối khi có thay đổi chưa lưu trừ khi dùng `save` hoặc `force`.
- Lỗi có dạng `[MÃ] thông điệp — gợi ý`; mã: `WORD_NOT_AVAILABLE, WORD_BUSY, DOC_NOT_FOUND, RANGE_INVALID, COM_ERROR, BLOCKED, VBA_DISABLED, VBA_NOT_TRUSTED, UNSAVED_CHANGES, FILE_EXISTS, TIMEOUT`.

## Danh mục công cụ

- Tài liệu: `word_list_documents, word_open, word_create, word_save, word_save_as, word_close, word_activate, word_export_pdf, word_print, word_mail_merge`
- Đọc: `word_read_text, word_get_outline, word_get_structure, word_get_selection, word_get_document_info, word_get_range_formatting`
- Nội dung: `word_insert_text, word_insert_paragraph, word_delete, word_find, word_find_replace, word_select, word_undo, word_redo`
- Định dạng: `word_format_font, word_format_paragraph, word_apply_style, word_list_styles, word_create_or_modify_style, word_apply_list, word_clear_formatting`
- Bảng: `word_insert_table, word_read_table, word_set_cells, word_table_add_row, word_table_add_column, word_table_delete_rows, word_table_delete_columns, word_table_merge_cells, word_table_format, word_delete_table`
- Đối tượng: `word_insert_picture, word_insert_shape, word_list_shapes, word_insert_caption`
- Bố cục: `word_page_setup, word_insert_break, word_set_header_footer, word_insert_page_numbers, word_list_sections`
- Tham chiếu: `word_insert_toc, word_update_toc, word_update_fields, word_insert_field, word_add_bookmark, word_list_bookmarks, word_delete_bookmark, word_add_hyperlink, word_list_hyperlinks, word_insert_footnote, word_insert_cross_reference`
- Rà soát: `word_track_changes, word_list_revisions, word_accept_revisions, word_reject_revisions, word_add_comment, word_list_comments, word_reply_comment, word_resolve_comment, word_delete_comment, word_protect, word_unprotect`
- Gateway: `word_com_get, word_com_set, word_com_call, word_list_members, word_execute_mso`
- VBA: `word_run_vba`

### Ví dụ gateway

- Đọc: `word_com_get` với `path = "ActiveDocument.Paragraphs.Count"`
- Ghi: `word_com_set` với `path = "ActiveDocument.Paragraphs(1).Range.Font.Bold"`, `value = true`
- Gọi hàm: `word_com_call` với `path = "ActiveDocument.Paragraphs(2).Range.InsertAfter"`, `args = ["Xin chào"]`
- Khám phá: `word_list_members` với `path = "ActiveDocument.PageSetup"`

Danh sách chặn (không phân biệt hoa thường, bất kỳ đoạn nào của đường dẫn): `Quit, Kill, Shell, SendKeys, VBProject, VBE, MacroContainer, OrganizerDelete`.

## Kích thước mô tả công cụ

Đo bằng `uv run pytest tests/unit/test_server_tools.py -s`: 78 công cụ, tổng mô tả + input schema ≈ 33.763 ký tự ≈ 8.440 token.

## Mail merge

`word_mail_merge` chuyển nguồn dữ liệu `.csv`/`.txt` thành một tệp `.docx` tạm chứa bảng rồi dùng bảng đó làm nguồn dữ liệu. Lý do: nếu Word mở trực tiếp `.csv`/`.txt` thì sẽ hiện hộp thoại chọn mã hóa (encoding) làm treo lệnh. Tệp tạm do server tạo và dọn.

## In ấn

`word_print` gửi tài liệu tới **máy in thật** đang được chọn trong Windows. Chỉ gọi khi bạn thực sự muốn in. Nếu chỉ cần bản xem, dùng `word_export_pdf`.

## Bảo mật

- Danh sách chặn của gateway **không phải ranh giới bảo mật**; nó chỉ giảm rủi ro gọi nhầm. Đặt `WORD_MCP_UNSAFE_GATEWAY=1` sẽ tắt hẳn nó.
- **VBA tắt theo mặc định.** `word_run_vba` trả `VBA_DISABLED` cho tới khi đặt `WORD_MCP_ALLOW_VBA=1`. Ngoài ra, bạn phải tự bật trong Word: File > Options > Trust Center > Trust Center Settings > Macro Settings > **Trust access to the VBA project object model**. Server không tự bật tùy chọn này; nếu chưa bật sẽ trả `VBA_NOT_TRUSTED`.
- Gateway và VBA có thể thực thi hành vi tùy ý trong Word; chỉ bật khi tin cậy phiên làm việc.
- Log chỉ ghi ra stderr; stdout dành cho kênh MCP.

## Giới hạn đã biết

- Chỉ Windows và Word desktop (không hỗ trợ Word Online, macOS, LibreOffice); không hỗ trợ các ứng dụng Office khác.
- Word chạy một luồng COM: các lệnh được xếp hàng tuần tự; khi Word hiện hộp thoại modal, lệnh trả `WORD_BUSY` thay vì treo.
- Chỉ dùng late binding (không tạo `gen_py`); một số thành viên COM hiếm có thể cần `word_list_members` để dò.
- Gateway: tham số tùy chọn bỏ qua, đối tượng COM không trả về trực tiếp (chỉ tên kiểu).
- Tài liệu rất lớn: `word_read_text` cắt theo ranh giới đoạn (`max_chars`) và trả `next_paragraph` để đọc tiếp.
- Kích thước mô tả công cụ khoảng 8.400 token, ảnh hưởng ngữ cảnh của mô hình.
