B1 đọc file .env để đọc các giá trị sau
CLIENT_LIST_FILE_NAME=Client List GFR.xlsx
CLIENT_LIST_SHEET_NAME=Client List GFR
DOCUMENT_LIST_SHEET_NAME=Download Document Log
NUMBER_ITEMS_PER_PAGE=50

B2: Lấy danh sách client cần thực hiện down load:
đọc dữ liệu sheet CLIENT_LIST_SHEET_NAME, có cột Status là Pending

B3: Thực hiện download và ghi log
3.1 setup driver
method gọi ý setup_driver

3.2 thực hiện tìm client theo client name và client number
nhập client name vào input search, gửi enter
phải đảm bảo text thu được có dang client_name | client number, lấy number_documents (hay là total_document từ bước này)
method gợi ý: search_client, check_client_exists

3.3 nếu không tìm được client, hoặc có bất kỳ lỗi nào liên quan tới việc tìm client thì hãy ghi log vào cột Status là Error, cột Description là nội dung lỗi (có thể là nội dung lỗi đã được handle, hoặc chưa handle thì là message của exception, để sau này chúng ta biết đường để quay lại debug)

3.3 nếu tìm được client nhưng number_document 
nếu tìm được client,
lấy number_documents đã tìm được ở trên, lưu lại, để sau này sẽ cập nhật vào cột Total Documents của sheet CLIENT_LIST_SHEET_NAME (ở row đang xử lý)

nhưng nếu number_documents = 0 thì log Status là Warning, Description là Client has no document
nếu number_documents > 0 thì thực hiện export ở cá bước bên dưới

3.4 nếu tìm được client và có number_documents > 0 thì thực hiện export document
Tạo folder client ở trong thư mục donwload dir
tên folder là: client name + "_" + client number

3.4.1 thực hiện tải csv file chứa danh sách document
code gợi ý
EXPORT_LIST_BTN_LOCATOR = (By.XPATH, "//button[contains(text(), 'Export List')]")
        btn_export_list = self.wait.until(
            EC.element_to_be_clickable(EXPORT_LIST_BTN_LOCATOR)
        )
        btn_export_list.click()
        time.sleep(10)

lưu ý khi xử lý csv, vì code là code selenium nên chỉ có thể cấu hình đường dẫn thưu mục download file về ở trình duyệt đang kết nối với selenium, và file tải về sẽ nằm ở đó, trong bối cảnh hiện tại chính là thư mục self.download_dir. nên hiện tại code chỉ có thể xác nhận file vừa tải về bằng cách kiểm tra file gần nhất được tải về

nên trước khi xử lý cho một client nào đó (ngay từ bước tìm kiếm) thì phải kiểm tra trong thư mục self.download_dir xem có tồn tại file nào không (kể cả file tạm, file đang chưa download xong, đến file bất kỳ) thì sẽ xóa hết, các folder thì giữ nguyên, làm việc này để tránh ở bước downlaod zip bên dưới, có những zip download quá lâu, nên khi chuyển sang client kế tiếp thì zip mới tải về được, làm cho code xác định sai file gần nhât được tải về. 

nên file csv được xác định là file csv gần nhất được tải về trong thư mục self.download_dir (vì ngay khi bắt đầu xử lý client hiện tại thì đã xóa các file từ file tạm, file đang download dở, file ở bất kỳ extension nào, chỉ giữ lại folder nên sẽ chỉ còn có đúng 1 file csv sau khi được tải về nằm ở đây)

vì nếu chạy nhiều client nên việc chờ tải file có thể lâu, hãy dùng cơ chế chờ tải file
nếu phát hiện trong thư mục download dir có nhiều hơn 1 file (có thể là file tmp, file đang download, file dưới các dạng extension nào) thì log STatus là Warning, ghi log description phù hợp, và chuyển sang client khác (dĩ nhiên mỗi khi sang client mới thì phải làm sạch các file, đừng xóa folder nhé)

3.4.2 sau khi tìm được file csv, move file vào thư mục 0_csv_ (đây là thư mục chứa các file csv đã tải về, sau khi tải về phải move vào cho gọn thư mục download dir)
đọc file để lấy danh sách các file

file có cấu trúc dữ liệu dạng này
Client Name,Client Number,File Section,Document Type,Description,Year,Document Date,File Size,Document ID,File Type,
"SPEECH THERAPY SOLUTIONS, PLLC","5431","PERMANENT","FINANCIAL PLANNING & RETIREMENT","11.08.20 RETIREMENT CHANGE-TALIA CUSANELLI","","2020-11-09","17.6 KB","0000001X2S","pdf",
"SPEECH THERAPY SOLUTIONS, PLLC","5431","PERMANENT","OTHER","02.15.21 NEW BUSINESS ADDRESS","","2021-02-09","12 KB","0000001X2T","pdf",
"SPEECH THERAPY SOLUTIONS, PLLC","5431","PERMANENT","OTHER","FORM 8655 FOR E-FILED PAYROLL.PDF","","2015-07-06","682.8 KB","0000001X2V","pdf",
"SPEECH THERAPY SOLUTIONS, PLLC","5431","PERMANENT","OTHER","EMPLOYEE INFORMATION","","2019-06-18","706.9 KB","0000001X31","pdf",
"SPEECH THERAPY SOLUTIONS, PLLC","5431","PERMANENT","OTHER","12-14-15 CHANGE OF ADDRESS FOR IRS","","2015-12-14","518.4 KB","0000001XJP","pdf",
"SPEECH THERAPY SOLUTIONS, PLLC","5431","PERMANENT","OTHER","NCDOR SIGNED FORM GEN-53","","2016-06-22","236.5 KB","0000001XW6","pdf",
"SPEECH THERAPY SOLUTIONS, PLLC","5431","PERMANENT","OTHER","VIEWMYPAYCHECK.PDF","","2016-02-29","2.1 MB","0000001ZGD","pdf",

từ file csv sẽ lấy được danh sách file document của client đó

3.4.3 Log danh sách document vào sheet DOCUMENT_LIST_SHEET_NAME
sheet ấy có các cột như sau
Upload status	Upload description	Download Status	Download Description	Client Name	Client Number	File Name	File Path	File Section	Document Type	Description	Year	Document Date	File Size	Document ID	File Type	Download time
thì những cột cần log sau khi đọc được danh sách file document từ csv đã tải về là:
Client Name,Client Number,File Section,Document Type,Description,Year,Document Date,File Size,Document ID,File Type
Trước khi log thì bạn cần kiểm tra dòng với các thông tin Client Name,Client Number,Document ID,Year,File Type như vậy đã có chưa, nếu chưa có thì mới thêm mới, và lưu giữ lại index (để sau đó khi tải file dù thành công hay thất bại thì sẽ log vào cột Download Status và Download Descrition)
nếu được tìm thấy => chỉ lưu index (để sau update log), không cần ghi mới

mục tiêu làm việc này:
- Sheet DOCUMENT_LIST_SHEET_NAME sẽ lưu danh sách document của từng client, và tránh lưu trùng
- Một client có thể chạy lại lần 2 hoặc 3 khi lần đầu tiên việc export docuemnt bị lỗi ở bước xử lý toàn bộ file hay xử lý từng file lẻ
- Việc kiểm tra log đã tồn tại chưa thế này giúp sheet này không chỉ là sheet lưu việc log tải file, mà còn log thông tin các file document, thuận tiện thống kê về sau

3.4.4 thực hiện export
Nếu number documents (hay total docuements) là 1 => phải tải sigle (tải lẻ)
Nếu number documents (hay total docuements) > 1 => phải tải zip (tải multiple)
đặt try catch để bất kỳ lỗi nào xảy ra trong quá trình này thì phải log lại

3.4.4.1 Nếu tải single:
code gợi ý
                 document_data_cells = row.find_elements(config.DOCUMENT_DATA_CELL_LOCATOR[0], config.DOCUMENT_DATA_CELL_LOCATOR[1])
                        document_id = document_data_cells[9].text

                        csv_file_info = csv_documents_dict[document_id]
                        expected_download_file_name_items = []
                        if csv_file_info["Client Name"]:
                            expected_download_file_name_items.append(str(csv_file_info["Client Name"]))
                        if csv_file_info["Year"]:
                            expected_download_file_name_items.append(str(csv_file_info["Year"]))
                        if csv_file_info["Document Type"]:
                            expected_download_file_name_items.append(str(csv_file_info["Document Type"]))
                        if csv_file_info["Description"]:
                            expected_download_file_name_items.append(str(csv_file_info["Description"]))
                        expected_download_file_name = "_".join(expected_download_file_name_items)
                        expected_download_file_name = re.sub(r'[\\/:*?"<>|]', '', expected_download_file_name)
                        expected_download_file_name = ".".join([expected_download_file_name, csv_file_info["File Type"]])
                        doc_year = str(csv_file_info["Year"]) if csv_file_info["Year"] else ""
                        # Kiểm tra nếu file đã tồn tại thì bỏ qua
                        file_exists = False

                        client_dir = "_".join([str(client_name), str(client_number)]) # Thư mục đích cuối cùng
                        print(f"expected_download_file_name: {expected_download_file_name}")
                        client_target_dir_status, client_target_dir = self._get_safe_client_dir(client_dir, self.download_dir)
                        # print(f"client_target_dir: {client_target_dir}")
                        if client_target_dir_status:
                            if doc_year and os.path.exists(os.path.join(client_target_dir, doc_year, expected_download_file_name)):
                                file_exists = True
                                self.downloaded_documents += 1
                                logger.info(f"File '{expected_download_file_name}' đã tồn tại. Bỏ qua phần download.")
                                continue
                            elif not doc_year and os.path.exists(os.path.join(client_target_dir, expected_download_file_name)):
                                file_exists = True
                                self.downloaded_documents += 1
                                logger.info(f"File '{expected_download_file_name}' đã tồn tại. Bỏ qua phần download.")
                                continue

                        if not file_exists:
                            # 3. TRÍCH XUẤT ID VÀ CLICK EXPORT
                            
                            doument_first_cell = row.find_elements(config.DOCUMENT_ROW_FIRST_CELL_LOCATOR[0], config.DOCUMENT_ROW_FIRST_CELL_LOCATOR[1])
                            btns = doument_first_cell[0].find_elements(By.TAG_NAME, "button")
                            
                            btn_export = btns[2]
                            btn_export.click()
                            time.sleep(5)

Vì file tải về chỉ có một file, nhưng chạy bằng selenium nên chúng ta sẽ chỉ biết file tên gì khi chúng ta chờ file được tải về trong thư mục download dir
vì trước đó file csv đã được move vào thư mục 0_csv_ nên expect là thư mục download dir hiện tại ngoài các folder client ra thì sẽ không có file nào
lúc này chờ file được tải về
sau khi file tải về, lấy tên file
rồi kiểm tra xem tên file có trùng với expected_download_file_name không

(lưu ý, khi tải file ở chế độ single, tức tải từng file, file tả về có tên khong có document id phía sau, còn file tải về ở chế độ multiple tức tải zip, thì file có nối thêm document id phía sau)

nếu kiểm tra tên file trung => đây chính là file vừa tải về
dĩ nhiên, nếu trong thư mục donwload dir có nhiều hơn 1 file tại thời điểm này hoặc có bất kỳ lỗi nào trong quá trình tải này, thì báo lỗi, log giá trị cột Download Status, Download Descripion ở sheet DOCUMENT_LIST_SHEET_NAME
log cột Status và Description ở sheet CLIENT_LIST_SHEET_NAME

sau khi lấy được file, thì move file vào client folder, đổi tên để thêm document id phía sau tên, vì file được tải chế độ single thì sẽ không có document id
ví dụ:
file tải về có tên là abc.pdf, file này có document id là 00004556 
=> đổi tên thành abc_00004556.pdf rồi mới move vào thư mục client folder

lưu ý thêm, nếu row data của client đó trong file csv có cột Year not blank, ví dụ là 2017
thì khi move vào client folder, phải move vào thư mục year của client foler đó
cụ thể:
client folder
    2017
        abc_00004556.pdf

Log status ở cột Status	Description của sheet CLIENT_LIST_SHEET_NAME
và cột Download Status	Download Description của sheet DOCUMENT_LIST_SHEET_NAME ở row tương ứng

điền giá trị Number Of Files Downloaded tức là số file thực sự đã get về thành cong được vào cột Number Of Files Downloaded của sheet CLIENT_LIST_SHEET_NAME


3.4.4.1 Nếu tải multiple:
tham khảo method export_document_list của file gofileroom_download_multiple.py
điểm khác biệt cần sửa: hiện tại thì code sẽ thực hiện clich select all checkbox 
select_all_checkout = headers_list[-1]
select_all_checkout.click()
sau đó ấn next page rồi tiếp tục ấn selectall checkbox cho đến khi không chuyển trang được nữa
nhằm chọn tất cả các file document để sau đó ấn download một thể  
download_document_btn = download_document_btns[0]
download_document_btn.click()
time.sleep(1)

export_document_btns = self.wait.until(
    EC.presence_of_all_elements_located(config.EXPORT_DOCUMENT_BTNS_LOCALTOR)
)
export_document_btn = export_document_btns[0]
export_document_btn.click()

time.sleep(1)

btn_ok = self.wait.until(
    EC.presence_of_element_located(config.OK_BTN_LOCALTOR)
)
btn_ok.click()
time.sleep(15)

rồi xử lý chờ file zip đã tải về, get zip, move vào thư mục 0_zip_, rồi giải nén, rồi move file vào client folder (nếu file nằm trong thư mục year thì move vào thư mục year của client folder)

nhưng làm như vậy sẽ gặp bất cập là: trường hợp hợp client folder có tới mấy trăm tới hơn 1000 file cần tải, thì việc yêu cầu taỉ một file zip nặng như vậy tại một thời điểm sẽ làm gofileroom quá tải, và sẽ không tải được file zip dù hết thời gian chờ, nên chúng ta sẽ thực hiện tải zip ở từng page một
ví dụ ta có 950 ducment , mỗi page có 100 docuemnt (Nó tương ứng giá trị NUMBER_ITEMS_PER_PAGE trong file env) => sẽ có 10 page cần tải
thi ở từng page sẽ thực hiện click select all checkbox
select_all_checkout = headers_list[-1]
select_all_checkout.click()
click tải multiple file
download_document_btn = download_document_btns[0]
download_document_btn.click()
time.sleep(1)

export_document_btns = self.wait.until(
    EC.presence_of_all_elements_located(config.EXPORT_DOCUMENT_BTNS_LOCALTOR)
)
export_document_btn = export_document_btns[0]
export_document_btn.click()

time.sleep(1)

btn_ok = self.wait.until(
    EC.presence_of_element_located(config.OK_BTN_LOCALTOR)
)
btn_ok.click()
time.sleep(15)
sau đó chờ zip, get zip, move zip vào thư mục 0_zip_
sau đó trong thư mục 0_zip_ tạo sẵn một thư mục client name + "_" + client number + "_" + "zip" (chỉ tạo lần đầu trước khi tải các file zip của client này, nếu có rồi thì thôi)
sau đó giải nén file zip vừa rồi, rồi move file vào trong thư mục đó

sau khi thực hiện tương tự cho các page còn lại, thì sẽ có được thư mục zip có đầy đủ các file
sau đó sẽ move file vào client folder (nhớ nếu Year not blank thì phải move vào thư mục year nhé)

và cập nhật log ở cả 2 sheet
nhớ cập nhật giá trị Number Of Files Downloaded
và ở từng bước move từng file từ thư mục client name + "_" + client number + "_" + "zip" sang client folder, xử lý file nào file phải đối chiếu file đó trong csv list file, đối chiếu theo document id nhé
trong một danh sách file của csv list file, thì document id là unique để phân biệt
mình gọi ý là nên lặp qua danh sách file của csv list file
tìm file trong thư mục client name + "_" + client number + "_" + "zip" , nếu không tìm thấy thì báo lỗi và log lỗi ở row tương từng
nếu tìm thấy thì move file
với file được tải từ zip về thì document id sẽ nằm cuối tên file, phía trước dấu gạch dưới _

trong phần xử lý từng file này, nếu gặp bất kỳ lỗi nào thì log vào và chuyển sang xử lý file khác
vì có thể xảy ra những lỗi như, tên file quá dài nên không giải nến được folder, hoặc tên file có ký tự đặc biệt chẳng hạn

nhwo tính total file thực sự đã được move thành công sang client folder để sau đó cập nhật vào cột Number Of Files Downloaded nhé


sau khi xử lý hết cho moọt clinet, thì update clietn folder path vào cột Client Folder Path của sheet CLIENT_LIST_SHEET_NAME

sheet này có cấu trúc là Status	Description	Client Name	Client Number	Client Email	Total Documents	Number Of Files Downloaded	Client Folder Path


hãy đọc kỹ và tham khảo code của 2 file nếu cần gofileroom_download_multiple.py và gofileroom_download_single.py

code file gofileroom_download.py cho mình

lưu ý: việc cập nhật lưu file excel phải cẩn thận không làm hỏng file, hoặc cẩn thận kẻo lại ghi đè toàn bộ file, file không có vba