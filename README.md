# Goal Dashboard – tự động refresh mỗi giờ

Đẩy **toàn bộ nội dung thư mục này** lên repo GitHub **private** (thư mục này là gốc của repo). Không commit file Excel hay thư mục `site/`, `inputs/` (đã có trong `.gitignore`).

**Sửa giao diện:** chỉ cần sửa file `index.html` rồi upload lại lên GitHub (commit vào nhánh `main`) là dashboard tự build và deploy lại.

## Luồng chạy
1. Bạn upload `GoalsExport_YYYY-MM-DD.xlsx` vào SharePoint › People Insights Hub › Documents › Goal Dashboard.
2. GitHub Action chạy **mỗi giờ (phút thứ 7)**. Nó lấy file Goals có **ngày mới nhất trong tên file** và file danh sách nhân viên (1 file cố định, cập nhật tại chỗ).
3. Nếu cả hai file không đổi so với lần trước → **dừng, không build/deploy**. Nếu có đổi → build lại, deploy lên Azure Static Web Apps, và ghi báo cáo thay đổi nhân sự vào tab Summary của lần chạy (và artifact `roster-changes`, giữ 30 ngày).
4. Header dashboard hiện "Data updated: <ngày trong tên file>".
Chạy tay ngay sau khi upload: Actions › Refresh Goal Dashboard › Run workflow (tick *force* nếu muốn build lại dù file không đổi).

## Cấu hình một lần
**A. Azure app registration (app ID bạn đã có)**
- API permissions › Microsoft Graph › **Application** › `Sites.Selected` (hoặc `Files.Read.All`) › Grant admin consent.
- Nếu dùng `Sites.Selected`: IT cấp quyền **read** cho app trên **cả hai site** (People Insights Hub và VN People Team).

**B. GitHub › Settings › Secrets and variables › Actions**
- Secrets: `TENANT_ID`, `CLIENT_ID`, `CLIENT_SECRET`, `AZURE_STATIC_WEB_APPS_API_TOKEN`
- Variables:
  - `GOALS_SITE_URL` = `https://<tenant>.sharepoint.com/sites/<PeopleInsightsHub>`
  - `GOALS_FOLDER` = `Goal Dashboard`
  - `EMP_SITE_URL` = `https://<tenant>.sharepoint.com/sites/<VNPeopleTeam>`
  - `EMP_FOLDER` = `General/Employee Information`
  - `EMP_FILE` = `Danh sach nhan vien - updated 09.06.2025.xlsx`
  - `SHAREPOINT_ORIGIN` = `https://<tenant>.sharepoint.com` (để cho phép nhúng iframe)

**C. Azure Static Web Apps** (tạo resource gắn với repo này hoặc tạo trống; lấy *deployment token* cho secret `AZURE_STATIC_WEB_APPS_API_TOKEN`)
- Đăng nhập: provider **Microsoft Entra ID** (đã cấu hình trong `staticwebapp.config.json`: mọi trang yêu cầu đăng nhập, không cần mời user thủ công).
- **Cách phân quyền:** trang HTML chỉ có **một file, không chứa dữ liệu**. Sau khi đăng nhập SSO, trang gọi `/api/data` (hàm trong `api_src`). Hàm lấy email người đăng nhập, tra bảng quyền, và chỉ trả đúng dữ liệu của họ (Dept Head → team họ, master → tất cả). Email không có trong file → "not set up yet", không nhận dữ liệu nào.
- Quyền đọc từ file Excel **`Dashboard Access.xlsx`** đặt cùng folder Goal Dashboard trên SharePoint (file mẫu: `Goal Dashboard/Dashboard Access.xlsx`):
  - Sheet `Dept Heads`: mỗi dòng một Dept Head (tên, title, email, branch root). **Đổi Dept Head = sửa dòng đó**; dữ liệu và quyền cập nhật trong vòng 1 giờ (hoặc chạy tay workflow).
  - Sheet `Master access`: email xem được tất cả.
- Email phải trùng email đăng nhập Microsoft (không phân biệt hoa/thường).

## Nhúng vào SharePoint
- Dùng link/nút mở tab mới là chắc nhất. Nhúng iframe (Embed web part) cần: thêm domain của web app vào **HTML Field Security** của site, và người xem đã đăng nhập sẵn (trang đăng nhập Microsoft không chạy trong iframe).

## Báo cáo nhân sự (mỗi lần dữ liệu đổi)
Nhân viên mới, nhân viên chuyển sang nghỉ/quay lại, đổi leader/Dept Head, và người **không có trong danh sách nhân viên** (cần cập nhật file). Quy tắc nghỉ việc: Active nếu còn dòng Active ở Main sheet **hoặc** có trong sheet `List of employee`. Người nghỉ vẫn nằm trong tổng goals nhưng không vào xếp hạng.

## Chạy thử trên máy
```
pip install -r requirements.txt
python build.py --goals GoalsExport_2026-11-10.xlsx --emps "Danh sach nhan vien - updated 09.06.2025.xlsx" --out site
```
(trang `site/index.html` cần `/api/data` và đăng nhập của Azure nên chỉ chạy đầy đủ sau khi deploy; trên máy chỉ kiểm tra build chạy được).
