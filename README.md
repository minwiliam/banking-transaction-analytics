# Banking Transaction Analytics — Power BI Reporting & Phân tích nghiệp vụ đa chiều

Power BI dashboard 5 trang + phân tích nghiệp vụ đa chiều (multi-dimensional driver analysis) trên 20.000 giao dịch ngân hàng — đi từ làm sạch dữ liệu Excel, xây model, viết DAX, đến khuyến nghị hành động cụ thể được kiểm chứng bằng query trực tiếp trên model, có AI hỗ trợ.

**Xem nhanh:** mở [`banking.pbix`](banking.pbix) bằng Power BI Desktop — dữ liệu đã nhúng sẵn, không cần setup gì thêm. Muốn xem cấu trúc model/report dạng file rời (TMDL/PBIR) để chỉnh sửa, dùng `banking.pbip` — xem Mục 6 (Cách setup).

## 1. Bối cảnh vấn đề

Dữ liệu gốc là **một file Excel export duy nhất** — 20.000 dòng giao dịch ngân hàng (01/2023 – 05/2025, dữ liệu 2025 chưa đầy năm). Không có nhiều nguồn rời rạc cần join như CRM/POS/Ads, nhưng đào sâu vào từng field lại lộ ra nhiều bẫy khiến dashboard tưởng đúng nhưng thực chất sai:

| Vấn đề | Rủi ro nếu bỏ qua |
|---|---|
| `Currency` có 2 giá trị: EUR (84,9%) và USD (15,1%), không có bảng tỷ giá | Cộng thẳng `SUM(Amount)` sẽ ra một con số tiền tệ vô nghĩa (gộp 2 đơn vị khác nhau) |
| `CustomerScore` được ghi **theo từng giao dịch**, không theo từng khách hàng (1 khách có thể mang tới 5 mức điểm khác nhau qua các lần giao dịch) | Dùng field này để "phân khúc rủi ro khách hàng" sẽ cho kết luận sai — vì bản chất nó không ổn định ở cấp khách hàng |
| `RecommendedOffer` được gán cố định 1-1 theo `CustomerSegment` (thiết kế loại trừ lẫn nhau) | "Tỷ lệ khớp offer" trông như một KPI hành vi khách hàng, nhưng thực ra chỉ phản chiếu đúng bảng map có sẵn trong model |
| Dữ liệu 2025 chỉ có 5 tháng đầu năm (đến 20/05/2025) | So sánh trực tiếp theo năm (2023 vs 2024 vs 2025) cho ảo giác "sụt giảm mạnh" nếu không chuẩn hóa theo ngày |
| `LatePaymentAmount` chỉ khác 0 với đúng 1 loại giao dịch (`Loan Payment`) | Nhìn biểu đồ "phí theo loại giao dịch" dễ kết luận nhầm là lỗi định giá hoặc thu phí nhắm sai đối tượng |

Nếu build thẳng dashboard trên dữ liệu thô, các con số như tổng doanh thu, fee rate theo phân khúc hay "kênh nào tốt nhất" vẫn trông hợp lý — nhưng sai bản chất, hoặc tệ hơn là suy diễn nhân quả (causal) từ dữ liệu chỉ mang tính quan sát (correlation).

## 2. Solution này là gì

Solution gồm ba phần:

- **Excel/Power Query cleaning pass** — làm sạch dữ liệu thô trước khi đưa vào model (kiểm lỗi, duplicate, missing value, chuẩn hóa text/case, tạo sẵn cột `Total Fee`).
- **Power BI semantic model tối giản** — star-schema 1 fact + 2 dimension, 23 DAX measure và 3 calculated column, không thêm bảng/cột không cần thiết.
- **5-page Power BI report + phân tích nghiệp vụ đa chiều có AI hỗ trợ** — đi từ "đọc số trên chart" sang "tìm nguyên nhân → đánh giá tác động tích cực/tiêu cực → khuyến nghị hành động cụ thể → KPI theo dõi", kiểm chứng bằng query DAX trực tiếp trên model sống.

```text
Excel export (20.000 dòng, 1 file duy nhất)
             │
             ▼
Excel / Power Query
 kiểm lỗi dữ liệu → duplicate → missing values → chuẩn hóa text/case → tạo Total Fee
             │
             ▼
Power BI Semantic Model (TMDL)
 MapOfferCategory (1) ──< FactTransaction (20.000 dòng) >── (1) DimDate
 23 DAX measures · 3 calculated columns
             │
             ▼
Power BI Report (PBIR) — 5 trang
 Executive Overview · Customer & Segment · Transaction & Channel
 Revenue & Friction · Trend & Offer
             │
             ▼
Phân tích nghiệp vụ đa chiều (AI-assisted, qua MCP)
 Observation → cross-tab đa chiều → nguyên nhân → tác động → khuyến nghị → KPI
```

Metric được tính đúng cấp dữ liệu ngay trong DAX measure (`Fee Rate %`, `Offer Alignment Rate %`, `MoM/YoY Growth %`…), không để Power BI tự cộng lại tỷ lệ hay % tăng trưởng từ visual.

## 3. Kết quả và insights

| Chỉ số (view mặc định: EUR) | Kết quả |
|---|---:|
| Tổng số dòng dữ liệu | **20.000** giao dịch |
| Giao dịch (view EUR) | **17.000** |
| Tổng giá trị giao dịch | **~86 triệu €** |
| Tổng phí | **~541.000 €** |
| Khách hàng hoạt động | **~8.000** |
| Giá trị giao dịch trung bình | **~5.050 €** |

Insight chính:

- **Middle Income dẫn đầu mọi KPI — nhưng vì có nhiều khách hàng hơn, không phải vì giao dịch giá trị lớn hơn.** Giá trị giao dịch trung bình gần như bằng nhau ở mọi phân khúc/kênh/loại giao dịch/sản phẩm (4.840€–5.286€). Thứ thật sự khác: số khách hoạt động (Middle 5.111 so với High 4.135, Low 3.112 — nhiều hơn Low Income tới **64%**) và tần suất giao dịch/khách (1,47 so với 1,37 và 1,23).
- **Không có rủi ro tập trung ở bất kỳ tổ hợp kênh × sản phẩm × loại giao dịch nào** — mọi tổ hợp được test đều chênh lệch dưới 20%, giá trị giao dịch trung bình theo kênh chỉ chênh ~2% (5.022€–5.133€).
- **Phí `Loan Payment` cao vượt trội (327K€, gấp ~7 lần trung bình 5 loại còn lại) — đã xác nhận đây là đặc điểm cấu trúc, không phải lỗi định giá.** `LatePaymentAmount` khác 0 ở đúng 100% giao dịch `Loan Payment` và bằng 0 ở mọi loại khác; phí trung bình mỗi giao dịch loại này gần như không đổi (95€–107€) ở toàn bộ 12 tổ hợp kênh × phân khúc.
- **Fee Rate % đồng đều giữa các phân khúc** (Low 0,641% · High 0,634% · Middle 0,625%) — không có bằng chứng thu phí bất công theo phân khúc thu nhập.
- **Tháng 2 giảm hoạt động thật (-7,8% so với trung bình 12 tháng, dùng năm đầy đủ 2023–2024) — nhưng chỉ tập trung ở kênh tự phục vụ.** Bóc theo kênh: ATM giảm **-16,7%**, Online **-9,8%**, Mobile chỉ **-3,5%**, còn Branch gần như không đổi (**+0,2%**). Đây là phát hiện quan trọng nhất về mặt hành động: một chiến dịch tháng 2 áp dụng đều cho mọi kênh sẽ lãng phí ngân sách vào Branch/Mobile — nơi không có vấn đề.
- **Tỷ lệ khớp offer khác nhau rất lớn theo phân khúc — một khoảng trống targeting có thể định lượng bằng tiền.** Middle Income khớp 75,9% giá trị giao dịch với offer được gán, nhưng High Income chỉ 29,8% và Low Income chỉ 26,7% — tức khoảng **33,6 triệu €** giá trị giao dịch của 2 phân khúc này nằm ngoài nhóm sản phẩm mà offer đang nhắm tới.

### Khuyến nghị nghiệp vụ

| # | Khuyến nghị | Ưu tiên |
|---|---|---|
| 1 | Chạy chiến dịch kích hoạt lại tháng 2 **chỉ cho khách dùng ATM và Online** (không áp dụng cho Branch/Mobile) | Cao |
| 2 | Test offer phụ cho High Income và Low Income theo đúng nhóm sản phẩm họ thực sự giao dịch nhiều (Mortgage, Loan) | Cao |
| 3 | Đưa Fee Rate % + cơ cấu loại phí theo phân khúc thành KPI giám sát công bằng phí định kỳ | Cao |
| 4 | Xác nhận lại việc công bố cấu trúc phí trễ hạn (`LatePaymentAmount`) của `Loan Payment` với khách hàng | Trung bình |
| 5 | Đo lường chiến dịch giữ chân theo **số lượng khách hàng** của Middle Income, không theo 1 sản phẩm/kênh cụ thể | Trung bình |
| 6 | Sửa `CustomerScore` về đúng cấp khách hàng (1 khách = 1 điểm ổn định) ở tầng dữ liệu nguồn | Trung bình |
| 7 | Theo dõi định kỳ (chưa cần hành động ngay) độ lệch kênh × sản phẩm — hiện đang rất cân bằng | Monitor |

Mỗi khuyến nghị đều theo cấu trúc WHO/WHAT/WHERE/HOW/KPI và chỉ dùng ngôn ngữ có điều kiện ("có thể", "cân nhắc").

### Giới hạn cần nhớ

- `CustomerScore` ghi theo giao dịch chứ không theo khách hàng (1 khách trung bình mang 1,83 trong 5 mức điểm) → không dùng để phân tích rủi ro khách hàng ổn định cho tới khi sửa ở nguồn.
- Không có bảng tỷ giá EUR/USD → không gộp 2 loại tiền trong cùng 1 phép tính; dùng slicer single-select, mặc định EUR.
- 2025 chỉ có 5 tháng đầu năm — mọi so sánh theo năm với 2025 đều ghi rõ là partial-year.
- Không có field chi phí phục vụ (cost-to-serve) hay lợi nhuận theo khách hàng → không tính được ROI thực tế cho các khuyến nghị ở trên, chỉ định hướng được **chiều** tác động.
- Dữ liệu chỉ mang tính quan sát (observational) — mọi kết luận nguyên nhân dùng ngôn ngữ có điều kiện ("có liên hệ với", "gợi ý rằng"), không khẳng định kết quả

## 4. Luồng xử lý dữ liệu

### Làm sạch (Excel/Power Query)

| Bước | Mục đích |
|---|---|
| Kiểm tra lỗi dữ liệu | Phát hiện giá trị âm bất thường, ngày sai định dạng, điểm tín dụng ngoài khoảng hợp lệ |
| Kiểm tra Duplicate | Đảm bảo không có giao dịch trùng lặp |
| Kiểm tra Missing Values | Quét toàn bộ cột trước khi load vào model |
| Chuẩn hóa Text | Loại khoảng trắng thừa, ký tự rác ở các field phân loại (`TransactionType`, `Channel`, `BranchCity`) |
| Chuẩn hóa chữ hoa/thường | Tránh 1 giá trị bị tách thành 2 do khác cách viết hoa/thường (`"mobile"` vs `"Mobile"`) |
| Tạo cột `Total Fee` | `CreditCardFees + InsuranceFees + LatePaymentAmount`, verify lại phía Power BI (diff = 0) nên không tính lại bằng DAX |

### Data Model (Power BI)

`MapOfferCategory (1) ──< FactTransaction (20.000 dòng) >── (1) DimDate`. `FactTransaction` giữ nguyên grain 1 dòng/giao dịch; `DimDate` là bảng lịch calculated đảm bảo liên tục không thiếu ngày; 3 calculated column (`CreditScoreGroup`, `OfferExpectedCategory`, `OfferAligned`). Model từng có thêm `IncomeGroup` (nhóm theo `MonthlyIncome`) nhưng đã bị xoá vì trùng lặp hoàn toàn với `CustomerSegment` — giữ lại chỉ gây nhiễu khi phân tích.

### DAX Measures (23)

Nhóm theo mục đích: Core KPI (8) · Fee Breakdown (3) · Time Intelligence (6) · Analytical % (4) · Convenience (2). Chi tiết công thức xem [`banking.SemanticModel/definition/tables/FactTransaction.tmdl`](banking.SemanticModel/definition/tables/FactTransaction.tmdl).

### Report (PBIR)

Canvas ở độ phân giải 1600x900.  Mỗi trang gồm: sidebar filter (Date, Segment, Product, Channel, Branch, Currency single-select) + KPI row + lưới 6 chart/table được **đo kích thước riêng theo số category** (không dùng grid đều) để không chart nào bị ẩn dữ liệu sau thanh cuộn và có tooltip mở rộng ở các chart quan trọng.

### Phân tích nghiệp vụ (AI-assisted)

Mọi insight ở Mục 3 đều bắt nguồn từ một query DAX cross-tab đa chiều (Segment × Type × Channel × Product × Time) chạy trực tiếp trên model sống qua MCP, không suy diễn từ một chart đơn lẻ. Observation luôn được tách rõ khỏi diễn giải nguyên nhân.

## 5. Schema dữ liệu

**Input** (file Excel gốc, [`data/Banking_Transactional_Dataset_Cleaned.xlsx`](data/Banking_Transactional_Dataset_Cleaned.xlsx)): `TransactionID, CustomerID, TransactionDate, Amount, TransactionType, ProductCategory, ProductSubcategory, BranchCity, BranchLat, BranchLong, Channel, Currency, CreditCardFees, InsuranceFees, LatePaymentAmount, CustomerScore, MonthlyIncome, CustomerSegment, RecommendedOffer, Total Fee` (các cột lịch gốc bị loại khi load vì đã có `DimDate` thay thế).

**Output** (Power BI model):

| Bảng | Grain | Vai trò |
|---|---|---|
| `FactTransaction` | 1 dòng/giao dịch (20.000 dòng) | Bảng fact trung tâm |
| `DimDate` | 1 dòng/ngày | Trục thời gian, phục vụ drill-down và time intelligence |
| `MapOfferCategory` | 1 dòng/offer (7 dòng) | Map `RecommendedOffer` → `ExpectedCategory` + `MatchShare` |

## 6. Cách setup

Yêu cầu **Power BI Desktop**. Repo có 2 cách mở, tùy mục đích:

### Cách A — Xem nhanh (khuyến nghị cho recruiter)

Mở thẳng [`banking.pbix`](banking.pbix) bằng Power BI Desktop. Dữ liệu đã nhúng sẵn trong file (`DataModel` đóng gói kèm), **không cần refresh hay trỏ lại nguồn dữ liệu** — mở lên là xem được dashboard đầy đủ ngay. Ở mỗi trang, chọn Currency = EUR (mặc định) trước khi đọc số — không chọn cả EUR lẫn USD cùng lúc.

### Cách B — Mở dạng PBIP để chỉnh sửa model/report

Dùng khi cần sửa DAX, thêm visual, hoặc xem cấu trúc file rời (TMDL/PBIR) thay vì file nhị phân.

> **Lưu ý:** file Excel nguồn đã làm sạch (**[`data/Banking_Transactional_Dataset_Cleaned.xlsx`](data/Banking_Transactional_Dataset_Cleaned.xlsx)**) được đóng gói kèm repo — đây là dữ liệu tổng hợp/giả lập (synthetic), không phải dữ liệu khách hàng thật, nên an toàn để public. Power Query trong model hiện vẫn đang trỏ tới đường dẫn cục bộ trên máy dùng để build project (`C:\Users\...\Downloads\...`), nên sau khi clone, bước 2 dưới đây **sẽ báo lỗi refresh cho tới khi bạn trỏ lại đường dẫn** — trỏ về đúng file đã có sẵn trong `data/`.

```text
1. Mở banking.pbip bằng Power BI Desktop
2. Trỏ lại nguồn dữ liệu về file Excel trong repo:
   Transform Data → Data source settings → chọn nguồn Excel → Change Source…
   → trỏ tới data/Banking_Transactional_Dataset_Cleaned.xlsx (nằm ngay trong repo đã clone)
3. Refresh model (Home → Refresh)
4. Ở mỗi trang, chọn Currency = EUR (mặc định) trước khi đọc số — không chọn cả EUR lẫn USD cùng lúc
```

Nếu dùng dữ liệu khách hàng thật cho một project khác (không phải bản demo này), không commit file Excel nguồn hoặc dữ liệu khách hàng thật lên git nếu repo public.

> **Lưu ý khi cập nhật:** `banking.pbix` là bản export tĩnh tại một thời điểm — nếu sau này chỉnh sửa model/report qua `banking.pbip`, nhớ export lại `.pbix` (File → Save As → Power BI files) để 2 bản không bị lệch nhau.

## 7. Guardrails phát triển

Các nguyên tắc này bảo vệ khỏi những sự cố từng xảy ra thật trong lúc build:
- Không `SUM(Amount)` gộp cả `EUR` và `USD` trong cùng 1 phép tính — luôn lọc theo `Currency` trước.
- Currency slicer luôn mặc định EUR và single-select — không đổi sang multi-select, vì sẽ phá vỡ nguyên tắc "không gộp tiền tệ" ở trên.
- `CustomerScore`/`CreditScoreGroup` không được dùng làm cơ sở cho bất kỳ insight nào về hành vi khách hàng ổn định — field này ở cấp giao dịch, không phải cấp khách hàng.
- `Total Fee` đã được tính sẵn ở Excel; không tính lại bằng DAX (đã verify diff = 0, tính lại chỉ tạo thêm rủi ro sai lệch).
- Sau khi refresh model với dữ liệu mới, đối chiếu lại vài mốc tổng (Total Transactions, Total Transaction Value, Total Fees) trước khi công bố số liệu — phát hiện sớm nếu refresh bị lỗi hoặc dữ liệu nguồn có vấn đề.
- Không đổi thiết kế `RecommendedOffer` → `CustomerSegment` (segment-exclusive) mà không cập nhật lại phân tích Offer Alignment liên quan — toàn bộ Insight về offer alignment trong Mục 3 dựa trên giả định 1-offer-1-segment.

## 8. Định hướng phát triển và scale-up

### Giai đoạn 1 — Hoàn thiện chất lượng dữ liệu

- Sửa `CustomerScore` về đúng cấp khách hàng ở nguồn.
- Bổ sung bảng tỷ giá EUR/USD nếu nghiệp vụ thực sự cần gộp currency.
- Xử lý minh bạch phần dữ liệu 2025 partial-year khi dataset được refresh đầy đủ năm.

### Giai đoạn 2 — Kiểm chứng targeting/offer

- Test offer phụ ngoài phân khúc gán sẵn (Khuyến nghị #2), đo response thực tế thay vì chỉ đo alignment theo thiết kế.
- Thu thập thêm dữ liệu conversion nếu muốn đánh giá hiệu quả offer thay vì chỉ alignment.

### Giai đoạn 3 — Tự động hóa báo cáo

- Power BI refresh định kỳ → AI tóm tắt biến động KPI theo kỳ → gửi báo cáo tự động.
- Review mọi khuyến nghị nghiệp vụ hoặc hành động liên quan phí/giá/tiếp cận khách hàng trước khi triển khai.

Mọi mở rộng phải giữ nguyên nguyên tắc: metric tính đúng ở DAX measure, insight luôn tách rõ observation khỏi diễn giải nguyên nhân, và không để BI tool tự suy diễn lại logic nghiệp vụ.
