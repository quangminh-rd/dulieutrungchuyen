
// Cấu hình API Google Sheets (chung cho cả ba tab)
const SPREADSHEET_ID = '14R9efcJ2hGE3mCgmJqi6TNbqkm4GFe91LEAuCyCa4O0';
const SPREADSHEET_ID_DANH_SACH_NGUOI_DUNG = '1GSakZ33O0JLrD2Mewl-EAHPBviokl8cLtI5pF1VIT6g';
const SPREADSHEET_ID_DANH_SACH_KHACH_HANG = '1sG87qCUvIZtuJbAv1vS2VRDmHG0ADwFBZQcNR-AzV9Q';
const SPREADSHEET_ID_GIAO_HANG = '1upjqkhTozefUFiugBeHh0sX9MgNhOtt-QJIh7By2rmY';
const API_KEY = 'AIzaSyA9g2qFUolpsu3_HVHOebdZb0NXnQgXlFM';

// Phạm vi dữ liệu
const RANGE_DON_HANG = 'don_hang!A:CS';
const RANGE_DON_HANG_CHI_TIET = 'don_hang_chi_tiet!A:GF';
const RANGE_DANH_SACH_NGUOI_DUNG = 'danh_sach_nguoi_dung!A:Z';
const RANGE_DANH_SACH_KHACH_HANG = 'danh_sach_khach_hang!A:Z';
const RANGE_GIAO_HANG = 'giao_hang!A:Z';

// Biến lưu trữ dữ liệu (chung cho cả ba tab)
let donHangData = [];
let donHangChiTietData = [];
let danhSachNguoiDungData = [];
let danhSachKhachHangData = [];
let giaoHangData = [];

// Lookup tables (chung)
let lookupNguoiDungByMaNV = {};
let lookupNguoiDungByTenNV = {};
let lookupKhachHangById = {};
let lookupGiaoHangByMaBoHang = {};

// Biến kiểm tra đã tải dữ liệu chưa
let isDataLoaded = false;

// Khởi tạo ứng dụng
document.addEventListener('DOMContentLoaded', function () {
    initTabNavigation();
    loadGapiAndInitialize();
});

// Khởi tạo tab navigation
function initTabNavigation() {
    const tabButtons = document.querySelectorAll('.tab-button');
    tabButtons.forEach(button => {
        button.addEventListener('click', function () {
            const tabId = this.getAttribute('data-tab');

            // Cập nhật active class cho buttons
            tabButtons.forEach(btn => btn.classList.remove('active'));
            this.classList.add('active');

            // Hiển thị tab tương ứng
            document.querySelectorAll('.tab-content').forEach(content => {
                content.classList.remove('active');
            });
            document.getElementById(`${tabId}-tab`).classList.add('active');
        });
    });
}

// ==================== HÀM CHUNG ====================

// Tải Google API Client
function loadGapiAndInitialize() {
    updateConnectionStatus('Đang kết nối...', '#fff4e6', '#e67e22');

    gapi.load('client', async () => {
        try {
            await gapi.client.init({
                apiKey: API_KEY,
                discoveryDocs: ['https://sheets.googleapis.com/$discovery/rest?version=v4']
            });

            updateConnectionStatus('Đã kết nối', '#e8f7e8', '#0c9c07');
            console.log('Google Sheets API đã sẵn sàng');

            // Tự động tải dữ liệu ban đầu
            await initialLoadData();

        } catch (error) {
            updateConnectionStatus('Lỗi kết nối', '#ffeaea', '#c00');
            console.error('Lỗi khởi tạo Google API:', error);
        }
    });
}

// Tải dữ liệu ban đầu
async function initialLoadData() {
    try {
        updateConnectionStatus('Đang tải dữ liệu ban đầu...', '#fff4e6', '#e67e22');
        await fetchDataFromSheets();
        updateConnectionStatus('Đã tải dữ liệu', '#e8f7e8', '#0c9c07');
        isDataLoaded = true;

        // Khởi tạo event listeners cho cả 5 tab
        initKhachHangEventListeners();
        initNhapKhoEventListeners();
        initXuatKhoEventListeners();
        initLenSanXuatEventListeners();
        initXuatBaoHanhEventListeners();

    } catch (error) {
        console.error('Lỗi khi tải dữ liệu ban đầu:', error);
        updateConnectionStatus('Lỗi tải dữ liệu', '#ffeaea', '#c00');
    }
}

// Lấy dữ liệu từ Google Sheets
async function fetchDataFromSheets() {
    updateConnectionStatus('Đang tải dữ liệu...', '#fff4e6', '#e67e22');

    try {
        const promises = [
            gapi.client.sheets.spreadsheets.values.get({
                spreadsheetId: SPREADSHEET_ID,
                range: RANGE_DON_HANG,
            }),
            gapi.client.sheets.spreadsheets.values.get({
                spreadsheetId: SPREADSHEET_ID,
                range: RANGE_DON_HANG_CHI_TIET,
            }),
            gapi.client.sheets.spreadsheets.values.get({
                spreadsheetId: SPREADSHEET_ID_DANH_SACH_NGUOI_DUNG,
                range: RANGE_DANH_SACH_NGUOI_DUNG,
            }),
            gapi.client.sheets.spreadsheets.values.get({
                spreadsheetId: SPREADSHEET_ID_DANH_SACH_KHACH_HANG,
                range: RANGE_DANH_SACH_KHACH_HANG,
            }),
            gapi.client.sheets.spreadsheets.values.get({
                spreadsheetId: SPREADSHEET_ID_GIAO_HANG,
                range: RANGE_GIAO_HANG,
            })
        ];

        const results = await Promise.all(promises);

        // Lưu trữ dữ liệu
        donHangData = results[0].result.values || [];
        donHangChiTietData = results[1].result.values || [];
        danhSachNguoiDungData = results[2].result.values || [];
        danhSachKhachHangData = results[3].result.values || [];
        giaoHangData = results[4].result.values || [];

        console.log(`Đã tải ${donHangData.length - 1} dòng từ sheet don_hang`);
        console.log(`Đã tải ${donHangChiTietData.length - 1} dòng từ sheet don_hang_chi_tiet`);

        // Xây dựng lookup tables
        buildLookupTables();

        // Cập nhật options cho select đơn vị phụ trách
        updateDonViPhuTrachOptions('nhap');
        updateDonViPhuTrachOptions('xuat');
        updateDonViPhuTrachOptions('lsx');
        updateDonViPhuTrachOptions('xbh');
        updateDonViPhuTrachOptions('kh');

        // Cập nhật options cho select mã hợp đồng
        updateMaHopDongOptions('nhap');
        updateMaHopDongOptions('xuat');
        updateMaHopDongOptions('lsx');
        updateMaHopDongOptions('xbh');


        // Cập nhật options cho select xưởng sản xuất
        updateXuongSanXuatOptions();

        updateConnectionStatus('Đã tải dữ liệu', '#e8f7e8', '#0c9c07');
        isDataLoaded = true;

    } catch (error) {
        console.error('Lỗi khi tải dữ liệu từ Google Sheets:', error);
        updateConnectionStatus('Lỗi tải dữ liệu', '#ffeaea', '#c00');
        throw error;
    }
}

// Xây dựng lookup tables
function buildLookupTables() {
    // Xây dựng lookup từ danh_sach_nguoi_dung
    if (danhSachNguoiDungData.length > 0) {
        const headers = danhSachNguoiDungData[0];
        const maNVIndex = headers.indexOf('ma_nhan_vien');
        const tenNVIndex = headers.indexOf('ten_nhan_vien');
        const donViIndex = headers.indexOf('don_vi');
        const mnvCongTyIndex = headers.indexOf('mnv_cong_ty');

        for (let i = 1; i < danhSachNguoiDungData.length; i++) {
            const row = danhSachNguoiDungData[i];
            const maNV = maNVIndex >= 0 ? row[maNVIndex] : '';
            const tenNV = tenNVIndex >= 0 ? row[tenNVIndex] : '';
            const donVi = donViIndex >= 0 ? row[donViIndex] : '';
            const mnvCongTy = mnvCongTyIndex >= 0 ? row[mnvCongTyIndex] : '';

            if (maNV) {
                lookupNguoiDungByMaNV[maNV] = { donVi, mnvCongTy, tenNV };
            }
            if (tenNV) {
                lookupNguoiDungByTenNV[tenNV] = { mnvCongTy };
            }
        }
    }

    // Xây dựng lookup từ danh_sach_khach_hang
    if (danhSachKhachHangData.length > 0) {
        const headers = danhSachKhachHangData[0];
        const idIndex = headers.indexOf('id');
        const maKhachHangIndex = headers.indexOf('ma_khach_hang');

        for (let i = 1; i < danhSachKhachHangData.length; i++) {
            const row = danhSachKhachHangData[i];
            const id = idIndex >= 0 ? row[idIndex] : '';
            const maKhachHang = maKhachHangIndex >= 0 ? row[maKhachHangIndex] : '';

            if (id) {
                lookupKhachHangById[id] = maKhachHang;
            }
        }
    }

    // Xây dựng lookup từ giao_hang
    if (giaoHangData.length > 0) {
        const headers = giaoHangData[0];
        const maBoHangIndex = headers.indexOf('ma_bo_hang');
        const loaiViecIndex = headers.indexOf('loai_viec');
        const nguoiGiaoHangIndex = headers.indexOf('nguoi_giao_hang');

        for (let i = 1; i < giaoHangData.length; i++) {
            const row = giaoHangData[i];
            const maBoHang = maBoHangIndex >= 0 ? row[maBoHangIndex] : '';
            const loaiViec = loaiViecIndex >= 0 ? row[loaiViecIndex] : '';
            const nguoiGiaoHang = nguoiGiaoHangIndex >= 0 ? row[nguoiGiaoHangIndex] : '';

            if (maBoHang) {
                lookupGiaoHangByMaBoHang[maBoHang] = { loaiViec, nguoiGiaoHang };
            }
        }
    }
}

// Cập nhật options cho select đơn vị phụ trách (dạng multi-select)
function updateDonViPhuTrachOptions(type) {
    if (donHangData.length === 0) return;

    const donHangHeaders = donHangData[0];
    const donViPhuTrachIndex = donHangHeaders.indexOf('don_vi_phu_trach');
    if (donViPhuTrachIndex === -1) return;

    // Lấy danh sách đơn vị phụ trách duy nhất
    const donViValues = new Set();
    for (let i = 1; i < donHangData.length; i++) {
        const value = donHangData[i][donViPhuTrachIndex];
        if (value && value.trim() !== '') {
            donViValues.add(value.trim());
        }
    }

    const sortedValues = Array.from(donViValues).sort();

    // Map đúng suffix theo từng tab
    let suffix = "";
    if (type === "nhap") suffix = "-nhap";
    else if (type === "xuat") suffix = "-xuat";
    else if (type === "lsx") suffix = "-lsx";
    else if (type === "kh") suffix = "-kh";
    else if (type === "xbh") suffix = "-xbh";
    else return;

    // Container element tương ứng tab
    const containerElement = document.getElementById(`don-vi-phu-trach${suffix}-container`);
    if (!containerElement) return;

    // Xóa nội dung cũ
    containerElement.innerHTML = '';

    // Tạo input tìm kiếm
    const searchInput = document.createElement('input');
    searchInput.type = 'text';
    searchInput.className = 'multi-select-search';
    searchInput.placeholder = 'Tìm kiếm đơn vị...';
    searchInput.id = `don-vi-search${suffix}`;

    containerElement.appendChild(searchInput);

    // Tạo checkbox "Tất cả đơn vị" (mặc định được chọn)
    const allOptionDiv = document.createElement('div');
    allOptionDiv.className = 'multi-select-option select-all-option';

    const allCheckbox = document.createElement('input');
    allCheckbox.type = 'checkbox';
    allCheckbox.id = `don-vi-phu-trach${suffix}-all`;
    allCheckbox.checked = true;

    const allLabel = document.createElement('label');
    allLabel.htmlFor = `don-vi-phu-trach${suffix}-all`;
    allLabel.textContent = 'Tất cả đơn vị';

    allOptionDiv.appendChild(allCheckbox);
    allOptionDiv.appendChild(allLabel);
    containerElement.appendChild(allOptionDiv);

    // Thêm sự kiện cho checkbox "Tất cả đơn vị"
    allCheckbox.addEventListener('change', function () {
        const checkboxes = containerElement.querySelectorAll('input[type="checkbox"]:not(#don-vi-phu-trach' + suffix + '-all)');
        checkboxes.forEach(checkbox => {
            checkbox.checked = this.checked;

            // ✅ Update label
            updateDonViPhuTrachLabel(suffix);
        });

        // Kích hoạt sự kiện change cho filter nếu cần
        if (suffix === '-kh') requireRefilterKhachHang();
        else if (suffix === '-nhap') requireRefilterNhap();
        else if (suffix === '-xuat') requireRefilterXuat();
        else if (suffix === '-lsx') requireRefilterLenSanXuat();
        else if (suffix === '-xbh') requireRefilterXuatBaoHanh();
    });

    // Add options mới
    sortedValues.forEach(value => {
        const optionDiv = document.createElement('div');
        optionDiv.className = 'multi-select-option';

        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.id = `don-vi-phu-trach${suffix}-${value.replace(/\s+/g, '-')}`;
        checkbox.value = value;

        const label = document.createElement('label');
        label.htmlFor = checkbox.id;
        label.textContent = value;

        optionDiv.appendChild(checkbox);
        optionDiv.appendChild(label);
        containerElement.appendChild(optionDiv);

        // Thêm sự kiện cho từng checkbox
        checkbox.addEventListener('change', function () {
            const allCheckbox = document.getElementById(`don-vi-phu-trach${suffix}-all`);
            const checkboxes = containerElement.querySelectorAll('input[type="checkbox"]:not(#don-vi-phu-trach' + suffix + '-all)');
            const checkedCount = Array.from(checkboxes).filter(cb => cb.checked).length;

            // Nếu tất cả các checkbox đều được chọn, chọn "Tất cả đơn vị"
            if (checkedCount === checkboxes.length) {
                allCheckbox.checked = true;
            }
            // Nếu không có checkbox nào được chọn, chọn "Tất cả đơn vị"
            else if (checkedCount === 0) {
                allCheckbox.checked = true;
            }
            // Nếu có một số checkbox được chọn, bỏ chọn "Tất cả đơn vị"
            else {
                allCheckbox.checked = false;
            }

            // ✅ Cập nhật label số lượng đã chọn
            updateDonViPhuTrachLabel(suffix);

            // Kích hoạt sự kiện change cho filter nếu cần
            if (suffix === '-kh') requireRefilterKhachHang();
            else if (suffix === '-nhap') requireRefilterNhap();
            else if (suffix === '-xuat') requireRefilterXuat();
            else if (suffix === '-lsx') requireRefilterLenSanXuat();
            else if (suffix === '-xbh') requireRefilterXuatBaoHanh();
        });
    });

    // Thêm sự kiện tìm kiếm
    searchInput.addEventListener("input", function () {
        const keyword = searchInput.value.trim().toLowerCase();

        // Lấy tất cả option
        const options = Array.from(
            containerElement.querySelectorAll(".multi-select-option")
        );

        // Nếu không có text → đưa các lựa chọn đã chọn lên đầu
        if (keyword === "") {

            // Reset ẩn option
            options.forEach(opt => opt.classList.remove("filtered-out"));

            // Sort: selected lên trước
            options.sort((a, b) => {
                const aChecked = a.querySelector("input").checked;
                const bChecked = b.querySelector("input").checked;

                return bChecked - aChecked;
            });

            // Render lại đúng thứ tự
            options.forEach(opt => containerElement.appendChild(opt));

            return;
        }

        // Nếu có text → chia làm 2 nhóm:
        // (1) match keyword
        // (2) selected nhưng không match keyword

        let matched = [];
        let selectedNotMatched = [];
        let others = [];

        options.forEach(opt => {
            const labelText = opt.textContent.toLowerCase();
            const checked = opt.querySelector("input").checked;

            if (labelText.includes(keyword)) {
                matched.push(opt);
            }
            else if (checked) {
                selectedNotMatched.push(opt);
            }
            else {
                others.push(opt);
            }
        });

        // Ẩn tất cả trước
        options.forEach(opt => opt.classList.add("filtered-out"));

        // Hiện nhóm match
        matched.forEach(opt => {
            opt.classList.remove("filtered-out");
            containerElement.appendChild(opt);
        });

        // Hiện nhóm selected nhưng không match (ngay dưới match)
        selectedNotMatched.forEach(opt => {
            opt.classList.remove("filtered-out");
            containerElement.appendChild(opt);
        });

        // Nhóm còn lại vẫn ẩn
    });


    // Giữ container mở rộng khi focus vào ô tìm kiếm
    searchInput.addEventListener('focus', function () {
        containerElement.classList.add('expanded');
    });

    let hideTimeout;

    containerElement.addEventListener("mouseleave", function () {
        hideTimeout = setTimeout(() => {
            containerElement.classList.remove("expanded");
        }, 100);
    });

    containerElement.addEventListener("mouseenter", function () {
        clearTimeout(hideTimeout);
    });


    document.addEventListener("click", function (e) {
        if (!containerElement.contains(e.target)) {
            containerElement.classList.remove("expanded");
        }
    });


    searchInput.addEventListener('blur', function () {
        setTimeout(() => {
            if (!containerElement.matches(':hover')) {
                containerElement.classList.remove('expanded');
            }
        }, 200);
    });

    updateDonViPhuTrachLabel(suffix);
}

// Hàm lấy danh sách đơn vị phụ trách được chọn từ multi-select
function getSelectedDonViPhuTrach(suffix) {
    const container = document.getElementById(`don-vi-phu-trach${suffix}-container`);
    if (!container) return [];

    const allCheckbox = document.getElementById(`don-vi-phu-trach${suffix}-all`);
    if (allCheckbox && allCheckbox.checked) {
        return []; // Trả về mảng rỗng nếu chọn "Tất cả đơn vị"
    }

    const checkboxes = container.querySelectorAll('input[type="checkbox"]:checked:not(#don-vi-phu-trach' + suffix + '-all)');
    return Array.from(checkboxes).map(cb => cb.value);
}

function updateDonViPhuTrachLabel(suffix) {

    // span hiển thị số lượng đã chọn
    const selectedSpan = document.getElementById(
        `don-vi-selected${suffix}`
    );

    if (!selectedSpan) return;

    const allCheckbox = document.getElementById(
        `don-vi-phu-trach${suffix}-all`
    );

    // Nếu đang chọn "Tất cả đơn vị"
    if (allCheckbox && allCheckbox.checked) {
        selectedSpan.textContent = "Đã chọn: Tất cả";
        return;
    }

    // Nếu chọn từng đơn vị cụ thể
    const selected = getSelectedDonViPhuTrach(suffix);

    selectedSpan.textContent = `Đã chọn: ${selected.length}`;
}

// Cập nhật options cho select mã hợp đồng (dạng multi-select)

function updateMaHopDongOptions(type) {
    if (donHangData.length === 0) return;

    const donHangHeaders = donHangData[0];
    const maHopDongIndex = donHangHeaders.indexOf('ma_hop_dong');
    if (maHopDongIndex === -1) return;

    // Lấy danh sách mã hợp đồng duy nhất (loại bỏ giá trị rỗng)
    const maHopDongValues = new Set();
    for (let i = 1; i < donHangData.length; i++) {
        const value = donHangData[i][maHopDongIndex];
        if (value && value.trim() !== '') {
            maHopDongValues.add(value.trim());
        }
    }

    const sortedValues = Array.from(maHopDongValues).sort();

    // Map đúng suffix theo từng tab
    let suffix = "";
    if (type === "nhap") suffix = "-nhap";
    else if (type === "xuat") suffix = "-xuat";
    else if (type === "lsx") suffix = "-lsx";
    else if (type === "xbh") suffix = "-xbh";
    else return;

    // Container element tương ứng tab
    const containerElement = document.getElementById(`ma-hop-dong${suffix}-container`);
    if (!containerElement) return;

    // Xóa nội dung cũ
    containerElement.innerHTML = '';

    // Tạo input tìm kiếm
    const searchInput = document.createElement('input');
    searchInput.type = 'text';
    searchInput.className = 'multi-select-search';
    searchInput.placeholder = 'Tìm kiếm mã hợp đồng...';
    searchInput.id = `ma-hop-dong-search${suffix}`;

    containerElement.appendChild(searchInput);

    // Tạo checkbox "Tất cả đơn hàng" (mặc định được chọn)
    const allOptionDiv = document.createElement('div');
    allOptionDiv.className = 'multi-select-option select-all-option';

    const allCheckbox = document.createElement('input');
    allCheckbox.type = 'checkbox';
    allCheckbox.id = `ma-hop-dong${suffix}-all`;
    allCheckbox.checked = true;

    const allLabel = document.createElement('label');
    allLabel.htmlFor = `ma-hop-dong${suffix}-all`;
    allLabel.textContent = 'Tất cả đơn hàng';

    allOptionDiv.appendChild(allCheckbox);
    allOptionDiv.appendChild(allLabel);
    containerElement.appendChild(allOptionDiv);

    // Thêm sự kiện cho checkbox "Tất cả đơn hàng"
    allCheckbox.addEventListener('change', function () {
        const checkboxes = containerElement.querySelectorAll('input[type="checkbox"]:not(#ma-hop-dong' + suffix + '-all)');
        checkboxes.forEach(checkbox => {
            checkbox.checked = this.checked;
            updateMaHopDongLabel(suffix);
        });

        // Kích hoạt sự kiện change cho filter nếu cần
        if (suffix === '-nhap') requireRefilterNhap();
        else if (suffix === '-xuat') requireRefilterXuat();
        else if (suffix === '-lsx') requireRefilterLenSanXuat();
        else if (suffix === '-xbh') requireRefilterXuatBaoHanh();
    });

    // Add options mới
    sortedValues.forEach(value => {
        const optionDiv = document.createElement('div');
        optionDiv.className = 'multi-select-option';

        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.id = `ma-hop-dong${suffix}-${value.replace(/\s+/g, '-')}`;
        checkbox.value = value;

        const label = document.createElement('label');
        label.htmlFor = checkbox.id;
        label.textContent = value;

        optionDiv.appendChild(checkbox);
        optionDiv.appendChild(label);
        containerElement.appendChild(optionDiv);

        // Thêm sự kiện cho từng checkbox
        checkbox.addEventListener('change', function () {
            const allCheckbox = document.getElementById(`ma-hop-dong${suffix}-all`);
            const checkboxes = containerElement.querySelectorAll('input[type="checkbox"]:not(#ma-hop-dong' + suffix + '-all)');
            const checkedCount = Array.from(checkboxes).filter(cb => cb.checked).length;

            // Nếu tất cả các checkbox đều được chọn, chọn "Tất cả đơn hàng"
            if (checkedCount === checkboxes.length) {
                allCheckbox.checked = true;
            }
            // Nếu không có checkbox nào được chọn, chọn "Tất cả đơn hàng"
            else if (checkedCount === 0) {
                allCheckbox.checked = true;
            }
            // Nếu có một số checkbox được chọn, bỏ chọn "Tất cả đơn hàng"
            else {
                allCheckbox.checked = false;
            }

            // Cập nhật label số lượng đã chọn
            updateMaHopDongLabel(suffix);

            // Kích hoạt sự kiện change cho filter nếu cần
            if (suffix === '-nhap') requireRefilterNhap();
            else if (suffix === '-xuat') requireRefilterXuat();
            else if (suffix === '-lsx') requireRefilterLenSanXuat();
            else if (suffix === '-xbh') requireRefilterXuatBaoHanh();
        });
    });

    // Thêm sự kiện tìm kiếm
    searchInput.addEventListener("input", function () {
        const keyword = searchInput.value.trim().toLowerCase();

        // Lấy tất cả option
        const options = Array.from(
            containerElement.querySelectorAll(".multi-select-option")
        );

        // Nếu không có text → đưa các lựa chọn đã chọn lên đầu
        if (keyword === "") {

            // Reset ẩn option
            options.forEach(opt => opt.classList.remove("filtered-out"));

            // Sort: selected lên trước
            options.sort((a, b) => {
                const aChecked = a.querySelector("input").checked;
                const bChecked = b.querySelector("input").checked;

                return bChecked - aChecked;
            });

            // Render lại đúng thứ tự
            options.forEach(opt => containerElement.appendChild(opt));

            return;
        }

        // Nếu có text → chia làm 2 nhóm:
        // (1) match keyword
        // (2) selected nhưng không match keyword

        let matched = [];
        let selectedNotMatched = [];
        let others = [];

        options.forEach(opt => {
            const labelText = opt.textContent.toLowerCase();
            const checked = opt.querySelector("input").checked;

            if (labelText.includes(keyword)) {
                matched.push(opt);
            }
            else if (checked) {
                selectedNotMatched.push(opt);
            }
            else {
                others.push(opt);
            }
        });

        // Ẩn tất cả trước
        options.forEach(opt => opt.classList.add("filtered-out"));

        // Hiện nhóm match
        matched.forEach(opt => {
            opt.classList.remove("filtered-out");
            containerElement.appendChild(opt);
        });

        // Hiện nhóm selected nhưng không match (ngay dưới match)
        selectedNotMatched.forEach(opt => {
            opt.classList.remove("filtered-out");
            containerElement.appendChild(opt);
        });

        // Nhóm còn lại vẫn ẩn
    });


    // Giữ container mở rộng khi focus vào ô tìm kiếm
    searchInput.addEventListener('focus', function () {
        containerElement.classList.add('expanded');
    });

    // ✅ Khi di chuột ra ngoài container → thu gọn luôn
    let hideTimeout;

    containerElement.addEventListener("mouseleave", function () {
        hideTimeout = setTimeout(() => {
            containerElement.classList.remove("expanded");
        }, 100);
    });

    containerElement.addEventListener("mouseenter", function () {
        clearTimeout(hideTimeout);
    });


    document.addEventListener("click", function (e) {
        if (!containerElement.contains(e.target)) {
            containerElement.classList.remove("expanded");
        }
    });

    searchInput.addEventListener('blur', function () {
        setTimeout(() => {
            if (!containerElement.matches(':hover')) {
                containerElement.classList.remove('expanded');
            }
        }, 200);
    });

    updateMaHopDongLabel(suffix);
}

// Hàm lấy danh sách mã hợp đồng được chọn từ multi-select
function getSelectedMaHopDong(suffix) {
    const container = document.getElementById(`ma-hop-dong${suffix}-container`);
    if (!container) return [];

    const allCheckbox = document.getElementById(`ma-hop-dong${suffix}-all`);
    if (allCheckbox && allCheckbox.checked) {
        return []; // Trả về mảng rỗng nếu chọn "Tất cả đơn hàng"
    }

    const checkboxes = container.querySelectorAll('input[type="checkbox"]:checked:not(#ma-hop-dong' + suffix + '-all)');
    return Array.from(checkboxes).map(cb => cb.value);
}

function updateMaHopDongLabel(suffix) {
    // span hiển thị số lượng đã chọn
    const selectedSpan = document.getElementById(
        `ma-hop-dong-selected${suffix}`
    );

    if (!selectedSpan) return;

    const allCheckbox = document.getElementById(
        `ma-hop-dong${suffix}-all`
    );

    // Nếu đang chọn "Tất cả đơn hàng"
    if (allCheckbox && allCheckbox.checked) {
        selectedSpan.textContent = "Đã chọn: Tất cả";
        return;
    }

    // Nếu chọn từng mã hợp đồng cụ thể
    const selected = getSelectedMaHopDong(suffix);

    selectedSpan.textContent = `Đã chọn: ${selected.length}`;
}

// Cập nhật options cho select xưởng sản xuất
function updateXuongSanXuatOptions() {
    if (donHangData.length === 0) return;

    const donHangHeaders = donHangData[0];
    const xuongSanXuatIndex = donHangHeaders.indexOf('xuong_san_xuat');
    if (xuongSanXuatIndex === -1) return;

    // Lấy danh sách xưởng duy nhất
    const xuongValues = new Set();

    for (let i = 1; i < donHangData.length; i++) {
        const value = donHangData[i][xuongSanXuatIndex];
        if (value && value.trim() !== '') {
            xuongValues.add(value.trim());
        }
    }

    const sortedValues = Array.from(xuongValues).sort();

    // ✅ Update cho cả 2 tab
    const selectElements = [
        document.getElementById("xuong-san-xuat-lsx"),
        document.getElementById("xuong-san-xuat-xbh")
    ];

    selectElements.forEach(selectElement => {
        if (!selectElement) return;

        // Xóa option cũ (giữ option đầu)
        while (selectElement.options.length > 1) {
            selectElement.remove(1);
        }

        // Thêm option mới
        sortedValues.forEach(value => {
            const option = document.createElement("option");
            option.value = value;
            option.textContent = value;
            selectElement.appendChild(option);
        });
    });
}


// Cập nhật trạng thái kết nối
function updateConnectionStatus(text, bgColor, color) {
    const statusElement = document.getElementById('connection-status');
    const textElement = document.getElementById('status-text');
    const lastUpdateBox = document.getElementById("last-update");
    const lastUpdateText = document.getElementById("last-update-text");

    statusElement.style.backgroundColor = bgColor;
    textElement.textContent = text;
    textElement.style.color = color;

    if (lastUpdateBox) {
        lastUpdateBox.style.backgroundColor = bgColor;
        lastUpdateBox.style.color = color;
    }

    if (lastUpdateText) {
        lastUpdateText.style.color = color;
    }

    updateLastUpdateTime();
    const dot = statusElement.querySelector(".dot");
    if (dot) dot.style.backgroundColor = color;
}

// Hàm chuyển đổi chuỗi ngày dd/mm/yyyy sang đối tượng Date
function parseDDMMYYYY(dateStr) {
    if (!dateStr) return null;
    const parts = dateStr.split('/');
    if (parts.length !== 3) return null;
    const day = parseInt(parts[0], 10);
    const month = parseInt(parts[1], 10) - 1;
    const year = parseInt(parts[2], 10);
    const date = new Date(year, month, day);
    return isNaN(date.getTime()) ? null : date;
}

// Hàm format số cho hiển thị
function formatNumberForDisplay(value) {
    if (!value && value !== 0) return '';
    const num = parseFloat(value.toString().replace(/\./g, '').replace(',', '.'));
    if (isNaN(num)) return value;
    return num.toLocaleString('vi-VN', {
        minimumFractionDigits: 0,
        maximumFractionDigits: 2
    });
}

// Hàm format số cho CSV
function formatNumberForCSV(value) {
    if (!value && value !== 0) return '';
    const num = parseFloat(value.toString().replace(/\./g, '').replace(',', '.'));
    if (isNaN(num)) return value;
    return num;
}

// Format ngày từ input date (yyyy-mm-dd) → dd.mm.yyyy
function formatDateForFilename(dateStr) {
    if (!dateStr) return "";
    const parts = dateStr.split("-");
    if (parts.length !== 3) return dateStr;
    return `${parts[2]}.${parts[1]}.${parts[0]}`;
}

// Cập nhật thời gian
function updateLastUpdateTime() {
    const lastUpdateElement = document.getElementById("last-update-text");
    if (!lastUpdateElement) return;
    const now = new Date();
    const formatted = now.toLocaleTimeString("vi-VN");
    lastUpdateElement.textContent = formatted;
}

// ==================== TAB DỮ LIỆU KHÁCH HÀNG ====================

let filteredResultsKhachHang = [];

function initKhachHangEventListeners() {
    document.getElementById('btn-loc-kh').addEventListener('click', applyFilterKhachHang);
    document.getElementById('btn-reset-kh').addEventListener('click', resetFilterKhachHang);
    document.getElementById('btn-export-kh').addEventListener('click', exportToExcelKhachHang);

    document.getElementById("tu-ngay-kh").addEventListener("change", requireRefilterKhachHang);
    document.getElementById("den-ngay-kh").addEventListener("change", requireRefilterKhachHang);
}

function resetFilterKhachHang() {
    // Reset multi-select đơn vị phụ trách
    const allCheckbox = document.getElementById('don-vi-phu-trach-kh-all');
    if (allCheckbox) {
        allCheckbox.checked = true;
    }

    const checkboxes = document.querySelectorAll('#don-vi-phu-trach-kh-container input[type="checkbox"]:not(#don-vi-phu-trach-kh-all)');
    checkboxes.forEach(checkbox => {
        checkbox.checked = false;
    });

    document.getElementById('tu-ngay-kh').value = '';
    document.getElementById('den-ngay-kh').value = '';

    document.getElementById('results-table-kh').style.display = 'none';
    document.getElementById('no-results-kh').style.display = 'block';
    document.getElementById('results-count-kh').textContent = 'Kết quả: Chưa có dòng nào được lọc.';

    filteredResultsKhachHang = [];
}

async function applyFilterKhachHang() {
    showLoadingKhachHang(true);

    try {
        if (!isDataLoaded) {
            await fetchDataFromSheets();
        }

        const filterOptions = {
            donViPhuTrach: getSelectedDonViPhuTrach('-kh'), // Lấy danh sách đơn vị được chọn
            tuNgay: document.getElementById('tu-ngay-kh').value,
            denNgay: document.getElementById('den-ngay-kh').value
        };

        filteredResultsKhachHang = filterDataKhachHang(filterOptions);
        displayResultsKhachHang(filteredResultsKhachHang);
        showMessageKhachHang(`Đã tìm thấy ${filteredResultsKhachHang.length} dòng phù hợp.`, 'success');

    } catch (error) {
        console.error('Lỗi khi áp dụng bộ lọc:', error);
        showMessageKhachHang('Đã xảy ra lỗi khi tải dữ liệu. Vui lòng thử lại.', 'error');
    } finally {
        showLoadingKhachHang(false);
    }
}

function filterDataKhachHang(filterOptions) {
    if (danhSachKhachHangData.length === 0) {
        return [];
    }

    const khachHangHeaders = danhSachKhachHangData[0];
    const khachHangColumnIndex = {};
    khachHangHeaders.forEach((header, index) => {
        khachHangColumnIndex[header] = index;
    });

    const filteredResults = [];
    let resultId = 1;

    for (let i = 1; i < danhSachKhachHangData.length; i++) {
        const row = danhSachKhachHangData[i];

        if (!passesFilterConditionsKhachHang(row, khachHangColumnIndex, filterOptions)) {
            continue;
        }

        filteredResults.push({
            id: resultId++,
            row: row,
            columnIndex: khachHangColumnIndex
        });
    }

    return filteredResults;
}

function passesFilterConditionsKhachHang(row, columnIndex, filterOptions) {
    const maKhachHang = row[columnIndex['ma_khach_hang']] || '';
    const donViPhuTrach = row[columnIndex['don_vi_phu_trach']] || '';
    const ngayPhatSinh = row[columnIndex['ngay_phat_sinh']] || '';

    // Điều kiện 1: Mã khách hàng không được rỗng
    if (!maKhachHang || maKhachHang.trim() === '') {
        return false;
    }

    // Điều kiện 2: Lọc theo đơn vị phụ trách
    if (filterOptions.donViPhuTrach && filterOptions.donViPhuTrach.length > 0) {
        if (!filterOptions.donViPhuTrach.includes(donViPhuTrach)) {
            return false;
        }
    }

    // Điều kiện 3: Lọc theo ngày báo cáo (danh sách mã khách hàng)
    if (filterOptions.ngayBaoCao) {
        const maKhachHangList = filterOptions.ngayBaoCao.split(',').map(item => item.trim());
        if (!maKhachHangList.includes(maKhachHang)) {
            return false;
        }
    }

    // Điều kiện 4: Lọc theo ngày phát sinh
    if (filterOptions.tuNgay && filterOptions.denNgay) {
        const tuNgay = new Date(filterOptions.tuNgay);
        const denNgay = new Date(filterOptions.denNgay);

        tuNgay.setHours(0, 0, 0, 0);
        denNgay.setHours(23, 59, 59, 999);

        const ngayPhatSinhDate = parseDDMMYYYY(ngayPhatSinh);
        if (!ngayPhatSinhDate) {
            return false;
        }

        if (ngayPhatSinhDate < tuNgay || ngayPhatSinhDate > denNgay) {
            return false;
        }
    }

    return true;
}

function displayResultsKhachHang(results) {
    const resultsBody = document.getElementById('results-body-kh');
    const resultsTable = document.getElementById('results-table-kh');
    const noResults = document.getElementById('no-results-kh');
    const resultsCount = document.getElementById('results-count-kh');

    resultsCount.textContent = `Kết quả: ${results.length} dòng`;
    resultsBody.innerHTML = '';

    if (results.length === 0) {
        resultsTable.style.display = 'none';
        noResults.style.display = 'block';
        return;
    }

    noResults.style.display = 'none';
    resultsTable.style.display = 'table';

    results.forEach(result => {
        const row = result.row;
        const columnIndex = result.columnIndex;

        // Lấy giá trị từ các cột
        const id = row[columnIndex['id']] || '';
        const loaiKhachHang = row[columnIndex['loai_khach_hang']] || '';
        const maKhachHang = row[columnIndex['ma_khach_hang']] || '';
        const tenNguoiLienHe = row[columnIndex['ten_nguoi_lien_he']] || '';
        const diaChiChiTiet = row[columnIndex['dia_chi_chi_tiet']] || '';
        const maSoThue = row[columnIndex['ma_so_thue']] || '';
        const sdtCoDinh = row[columnIndex['sdt_co_dinh']] || '';
        const sdtT1 = row[columnIndex['sdt_t1']] || '';
        const sdtT2 = row[columnIndex['sdt_t2']] || '';
        const fax = row[columnIndex['fax']] || '';
        const emailKhachHang = row[columnIndex['email_khach_hang']] || '';
        const donViPhuTrach = row[columnIndex['don_vi_phu_trach']] || '';
        const ngayPhatSinh = row[columnIndex['ngay_phat_sinh']] || '';

        // Các cột bổ sung (nếu có)
        const soCmnd = row[columnIndex['so_cmnd']] || '';
        const ngayCap = row[columnIndex['ngay_cap']] || '';
        const noiCap = row[columnIndex['noi_cap']] || '';
        const xungHo = row[columnIndex['xung_ho']] || '';
        const hoTenNlh = row[columnIndex['ho_ten_nlh']] || '';
        const chucDanhNlh = row[columnIndex['chuc_danh_nlh']] || '';
        const diaChiNlh = row[columnIndex['dia_chi_nlh']] || '';
        const dtDiDongNlh = row[columnIndex['dt_di_dong_nlh']] || '';
        const dtDiDongKhacNlh = row[columnIndex['dt_di_dong_khac_nlh']] || '';
        const dtCoDinhNlh = row[columnIndex['dt_co_dinh_nlh']] || '';
        const emailNlh = row[columnIndex['email_nlh']] || '';
        const soTaiKhoan = row[columnIndex['so_tai_khoan']] || '';
        const tenNganHang = row[columnIndex['ten_ngan_hang']] || '';
        const chiNhanh = row[columnIndex['chi_nhanh']] || '';
        const tinhTpNganHang = row[columnIndex['tinh_tp_ngan_hang']] || '';

        // Xử lý cột "Là tổ chức/cá nhân"
        let laToChucCaNhan = '';
        if (loaiKhachHang === "Tổ chức") {
            laToChucCaNhan = "0";
        } else if (loaiKhachHang === "Cá nhân") {
            laToChucCaNhan = "1";
        }

        // Xử lý cột "Điện thoại di động"
        let dienThoaiDiDong = '';
        if (sdtT1 && sdtT2) {
            dienThoaiDiDong = `${sdtT1}/${sdtT2}`;
        } else if (sdtT1) {
            dienThoaiDiDong = sdtT1;
        } else if (sdtT2) {
            dienThoaiDiDong = sdtT2;
        }

        // Xử lý cột "Nhóm KH/NCC"
        let nhomKhNcc = '';
        if (donViPhuTrach === "BP. BH1" || donViPhuTrach === "BP. BH2") {
            const ma12 = maKhachHang.substring(0, 12);
            const ma9 = maKhachHang.substring(0, 9);
            const ma6 = maKhachHang.substring(0, 6);
            const ma2 = maKhachHang.substring(0, 2);
            const ma1 = maKhachHang.substring(0, 1);
            nhomKhNcc = `${ma12};${ma9};${ma6};${ma2};${ma1}`;
        }

        const tableRow = document.createElement('tr');
        tableRow.innerHTML = `
    <td>${id}</td>
    <td>${laToChucCaNhan}</td>
    <td>0</td>
    <td>${maKhachHang}</td>
    <td>${tenNguoiLienHe}</td>
    <td>${diaChiChiTiet}</td>
    <td>${maSoThue}</td>
    <td>${sdtCoDinh}</td>
    <td>${dienThoaiDiDong}</td>
    <td>${fax}</td>
    <td>${emailKhachHang}</td>
    <td></td>
    <td>${nhomKhNcc}</td>
    <td>${soCmnd}</td>
    <td>${ngayCap}</td>
    <td>${noiCap}</td>
    <td>${xungHo}</td>
    <td>${hoTenNlh}</td>
    <td>${chucDanhNlh}</td>
    <td>${diaChiNlh}</td>
    <td>${dtDiDongNlh}</td>
    <td>${dtDiDongKhacNlh}</td>
    <td>${dtCoDinhNlh}</td>
    <td>${emailNlh}</td>
    <td>${soTaiKhoan}</td>
    <td>${tenNganHang}</td>
    <td>${chiNhanh}</td>
    <td>${tinhTpNganHang}</td>
    <td>${ngayPhatSinh}</td>
    `;
        resultsBody.appendChild(tableRow);
    });
}

function exportToExcelKhachHang() {
    if (filteredResultsKhachHang.length === 0) {
        showMessageKhachHang('Không có dữ liệu để xuất. Vui lòng thực hiện lọc trước.', 'error');
        return;
    }

    let csvContent = "\uFEFF";
    csvContent += "ID,Là tổ chức/cá nhân,Là nhà cung cấp,Mã khách hàng (*),Tên khách hàng (*),Địa chỉ,Mã số thuế,Điện thoại,Điện thoại di động,Fax,Email,Website,Nhóm KH/NCC,Số CMND,Ngày cấp,Nơi cấp,Xưng hô,Họ và tên NLH,Chức danh NLH,Địa chỉ người liên hệ,ĐT di động NLH,ĐT di động khác của NLH,ĐT cố định NLH,Email người liên hệ,Số tài khoản,Tên ngân hàng,Chi nhánh,Tỉnh/TP TK ngân hàng,Ngày phát sinh\n";

    filteredResultsKhachHang.forEach(result => {
        const row = result.row;
        const columnIndex = result.columnIndex;

        const id = row[columnIndex['id']] || '';
        const loaiKhachHang = row[columnIndex['loai_khach_hang']] || '';
        const maKhachHang = row[columnIndex['ma_khach_hang']] || '';
        const tenNguoiLienHe = row[columnIndex['ten_nguoi_lien_he']] || '';
        const diaChiChiTiet = row[columnIndex['dia_chi_chi_tiet']] || '';
        const maSoThue = row[columnIndex['ma_so_thue']] || '';
        const sdtCoDinh = row[columnIndex['sdt_co_dinh']] || '';
        const sdtT1 = row[columnIndex['sdt_t1']] || '';
        const sdtT2 = row[columnIndex['sdt_t2']] || '';
        const fax = row[columnIndex['fax']] || '';
        const emailKhachHang = row[columnIndex['email_khach_hang']] || '';
        const donViPhuTrach = row[columnIndex['don_vi_phu_trach']] || '';
        const ngayPhatSinh = row[columnIndex['ngay_phat_sinh']] || '';

        // Các cột bổ sung
        const soCmnd = row[columnIndex['so_cmnd']] || '';
        const ngayCap = row[columnIndex['ngay_cap']] || '';
        const noiCap = row[columnIndex['noi_cap']] || '';
        const xungHo = row[columnIndex['xung_ho']] || '';
        const hoTenNlh = row[columnIndex['ho_ten_nlh']] || '';
        const chucDanhNlh = row[columnIndex['chuc_danh_nlh']] || '';
        const diaChiNlh = row[columnIndex['dia_chi_nlh']] || '';
        const dtDiDongNlh = row[columnIndex['dt_di_dong_nlh']] || '';
        const dtDiDongKhacNlh = row[columnIndex['dt_di_dong_khac_nlh']] || '';
        const dtCoDinhNlh = row[columnIndex['dt_co_dinh_nlh']] || '';
        const emailNlh = row[columnIndex['email_nlh']] || '';
        const soTaiKhoan = row[columnIndex['so_tai_khoan']] || '';
        const tenNganHang = row[columnIndex['ten_ngan_hang']] || '';
        const chiNhanh = row[columnIndex['chi_nhanh']] || '';
        const tinhTpNganHang = row[columnIndex['tinh_tp_ngan_hang']] || '';

        // Xử lý cột "Là tổ chức/cá nhân"
        let laToChucCaNhan = '';
        if (loaiKhachHang === "Tổ chức") {
            laToChucCaNhan = "0";
        } else if (loaiKhachHang === "Cá nhân") {
            laToChucCaNhan = "1";
        }

        // Xử lý cột "Điện thoại di động"
        let dienThoaiDiDong = '';
        if (sdtT1 && sdtT2) {
            dienThoaiDiDong = `${sdtT1}/${sdtT2}`;
        } else if (sdtT1) {
            dienThoaiDiDong = sdtT1;
        } else if (sdtT2) {
            dienThoaiDiDong = sdtT2;
        }

        // Xử lý cột "Nhóm KH/NCC"
        let nhomKhNcc = '';
        if (donViPhuTrach === "BP. BH1" || donViPhuTrach === "BP. BH2") {
            const ma12 = maKhachHang.substring(0, 12);
            const ma9 = maKhachHang.substring(0, 9);
            const ma6 = maKhachHang.substring(0, 6);
            const ma2 = maKhachHang.substring(0, 2);
            const ma1 = maKhachHang.substring(0, 1);
            nhomKhNcc = `${ma12};${ma9};${ma6};${ma2};${ma1}`;
        }

        const escapeCSV = (str) => {
            if (!str) return '';
            if (str.includes(',') || str.includes('"') || str.includes('\n')) {
                return '"' + str.replace(/"/g, '""') + '"';
            }
            return str;
        };

        csvContent += `${escapeCSV(id)},${laToChucCaNhan},0,${escapeCSV(maKhachHang)},${escapeCSV(tenNguoiLienHe)},${escapeCSV(diaChiChiTiet)},${escapeCSV(maSoThue)},${escapeCSV(sdtCoDinh)},${escapeCSV(dienThoaiDiDong)},${escapeCSV(fax)},${escapeCSV(emailKhachHang)},,${escapeCSV(nhomKhNcc)},${escapeCSV(soCmnd)},${escapeCSV(ngayCap)},${escapeCSV(noiCap)},${escapeCSV(xungHo)},${escapeCSV(hoTenNlh)},${escapeCSV(chucDanhNlh)},${escapeCSV(diaChiNlh)},${escapeCSV(dtDiDongNlh)},${escapeCSV(dtDiDongKhacNlh)},${escapeCSV(dtCoDinhNlh)},${escapeCSV(emailNlh)},${escapeCSV(soTaiKhoan)},${escapeCSV(tenNganHang)},${escapeCSV(chiNhanh)},${escapeCSV(tinhTpNganHang)},${escapeCSV(ngayPhatSinh)}\n`;
    });

    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement("a");
    const url = URL.createObjectURL(blob);

    link.setAttribute("href", url);
    const tuNgayRaw = document.getElementById("tu-ngay-kh").value;
    const denNgayRaw = document.getElementById("den-ngay-kh").value;
    const tuNgay = formatDateForFilename(tuNgayRaw);
    const denNgay = formatDateForFilename(denNgayRaw);
    const fileName = `Danh sách khách hàng - ${tuNgay} - ${denNgay}.csv`;

    link.setAttribute("download", fileName);
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);

    showMessageKhachHang(`Đã xuất ${filteredResultsKhachHang.length} dòng ra file Excel.`, 'success');
}

function showLoadingKhachHang(show) {
    const loadingElement = document.getElementById('loading-kh');
    loadingElement.style.display = show ? 'block' : 'none';
}

function showMessageKhachHang(message, type) {
    const resultsCount = document.getElementById("results-count-kh");
    resultsCount.className = "results-count";

    if (type === "success") {
        resultsCount.style.backgroundColor = "#e8f7e8";
        resultsCount.style.color = "#0c9c07";
        resultsCount.style.borderLeft = "4px solid #0c9c07";
    }

    if (type === "error") {
        resultsCount.style.backgroundColor = "#ffeaea";
        resultsCount.style.color = "#c00";
        resultsCount.style.borderLeft = "4px solid #c00";
    }

    resultsCount.style.padding = "5px";
    resultsCount.style.borderRadius = "6px";
    resultsCount.style.fontSize = "16px";
    resultsCount.style.fontWeight = "600";
    resultsCount.textContent = message;
}

function requireRefilterKhachHang() {
    document.getElementById('results-table-kh').style.display = 'none';
    document.getElementById('no-results-kh').style.display = 'block';
    document.getElementById('results-count-kh').textContent = 'Kết quả: Chưa có dòng nào được lọc.';
    filteredResultsKhachHang = [];
    showMessageKhachHang("Bạn đã thay đổi bộ lọc. Vui lòng nhấn 'Lọc' để cập nhật kết quả mới.", "error");
}

// ==================== TAB NHẬP KHO ====================

let filteredResultsNhap = [];

function initNhapKhoEventListeners() {
    document.getElementById('btn-loc-nhap').addEventListener('click', applyFilterNhap);
    document.getElementById('btn-reset-nhap').addEventListener('click', resetFilterNhap);
    document.getElementById('btn-export-nhap').addEventListener('click', exportToExcelNhap);

    // Xóa event listener cho input text cũ
    // document.getElementById("ma-hop-dong-nhap").addEventListener("input", requireRefilterNhap);
    document.getElementById("tu-ngay-nhap").addEventListener("change", requireRefilterNhap);
    document.getElementById("den-ngay-nhap").addEventListener("change", requireRefilterNhap);

    document.querySelectorAll("input[name='loai-don-hang-nhap']").forEach(radio => {
        radio.addEventListener("change", requireRefilterNhap);
    });
}

function resetFilterNhap() {
    // Reset multi-select đơn vị phụ trách
    const donViAllCheckbox = document.getElementById('don-vi-phu-trach-nhap-all');
    if (donViAllCheckbox) {
        donViAllCheckbox.checked = true;
    }

    const donViCheckboxes = document.querySelectorAll('#don-vi-phu-trach-nhap-container input[type="checkbox"]:not(#don-vi-phu-trach-nhap-all)');
    donViCheckboxes.forEach(checkbox => {
        checkbox.checked = false;
    });

    // Reset multi-select mã hợp đồng
    const maHopDongAllCheckbox = document.getElementById('ma-hop-dong-nhap-all');
    if (maHopDongAllCheckbox) {
        maHopDongAllCheckbox.checked = true;
    }

    const maHopDongCheckboxes = document.querySelectorAll('#ma-hop-dong-nhap-container input[type="checkbox"]:not(#ma-hop-dong-nhap-all)');
    maHopDongCheckboxes.forEach(checkbox => {
        checkbox.checked = false;
    });

    document.getElementById('tu-ngay-nhap').value = '';
    document.getElementById('den-ngay-nhap').value = '';
    document.getElementById('loai-tat-ca-nhap').checked = true;

    document.getElementById('results-table-nhap').style.display = 'none';
    document.getElementById('no-results-nhap').style.display = 'block';
    document.getElementById('results-count-nhap').textContent = 'Kết quả: Chưa có dòng nào được lọc.';

    filteredResultsNhap = [];
}

async function applyFilterNhap() {
    showLoadingNhap(true);

    try {
        if (!isDataLoaded) {
            await fetchDataFromSheets();
        }

        const filterOptions = {
            maHopDong: getSelectedMaHopDong('-nhap'), // Lấy danh sách mã hợp đồng được chọn
            donViPhuTrach: getSelectedDonViPhuTrach('-nhap'), // Lấy danh sách đơn vị được chọn
            tuNgay: document.getElementById('tu-ngay-nhap').value,
            denNgay: document.getElementById('den-ngay-nhap').value,
            loaiDonHang: document.querySelector('input[name="loai-don-hang-nhap"]:checked').value
        };

        filteredResultsNhap = filterDataNhap(filterOptions);
        displayResultsNhap(filteredResultsNhap);
        showMessageNhap(`Đã tìm thấy ${filteredResultsNhap.length} dòng phù hợp.`, 'success');

    } catch (error) {
        console.error('Lỗi khi áp dụng bộ lọc:', error);
        showMessageNhap('Đã xảy ra lỗi khi tải dữ liệu. Vui lòng thử lại.', 'error');
    } finally {
        showLoadingNhap(false);
    }
}

function filterDataNhap(filterOptions) {
    if (donHangData.length === 0 || donHangChiTietData.length === 0) {
        return [];
    }

    const donHangHeaders = donHangData[0];
    const chiTietHeaders = donHangChiTietData[0];

    const donHangColumnIndex = {};
    donHangHeaders.forEach((header, index) => {
        donHangColumnIndex[header] = index;
    });

    const chiTietColumnIndex = {};
    chiTietHeaders.forEach((header, index) => {
        chiTietColumnIndex[header] = index;
    });

    const filteredResults = [];
    let resultId = 1;

    for (let i = 1; i < donHangChiTietData.length; i++) {
        const chiTietRow = donHangChiTietData[i];
        const maDonHangID = chiTietRow[chiTietColumnIndex['ma_don_hang_id']] || '';
        const donHangRow = donHangData.find(row => {
            return row[donHangColumnIndex['ma_don_hang']] === maDonHangID;
        });

        if (!donHangRow) continue;

        if (!passesFilterConditionsNhap(chiTietRow, donHangRow, chiTietColumnIndex, donHangColumnIndex, filterOptions)) {
            continue;
        }

        filteredResults.push({
            id: resultId++,
            chiTietRow: chiTietRow,
            donHangRow: donHangRow,
            chiTietColumnIndex: chiTietColumnIndex,
            donHangColumnIndex: donHangColumnIndex
        });
    }

    return filteredResults;
}

function passesFilterConditionsNhap(chiTietRow, donHangRow, chiTietColumnIndex, donHangColumnIndex, filterOptions) {
    const nhomSanPham = chiTietRow[chiTietColumnIndex['nhom_san_pham']] || '';
    const mauCua = chiTietRow[chiTietColumnIndex['mau_cua']] || '';
    const tenSanPham = chiTietRow[chiTietColumnIndex['ten_san_pham']] || '';
    const trongLuongPhuKien = donHangRow[donHangColumnIndex['trong_luong_phu_kien']] || '';
    const maHopDong = donHangRow[donHangColumnIndex['ma_hop_dong']] || '';
    const donViPhuTrach = donHangRow[donHangColumnIndex['don_vi_phu_trach']] || '';
    const loaiDonHang = donHangRow[donHangColumnIndex['loai_don_hang']] || '';
    const ngayGiaoHang = donHangRow[donHangColumnIndex['ngay_giao_hang']] || '';
    const ngayNhapKho = donHangRow[donHangColumnIndex['ngay_nhap_kho']] || '';

    const excludedGroups = ["Vật tư phát sinh", "Vật tư khác", "Bảo hành", "Hàng hóa"];
    if (excludedGroups.includes(nhomSanPham)) {
        return false;
    }

    if (mauCua === "Nhân công") {
        return false;
    }

    if (tenSanPham === "Di chuyển" || tenSanPham === "Nhân công") {
        return false;
    }

    if (trongLuongPhuKien === "Tiêu chuẩn" && tenSanPham === "Vật tư khác") {
        return false;
    }

    // Lọc theo mã hợp đồng (multi-select)
    if (filterOptions.maHopDong && filterOptions.maHopDong.length > 0) {
        if (!filterOptions.maHopDong.includes(maHopDong)) {
            return false;
        }
    }

    // Lọc theo đơn vị phụ trách (multi-select)
    if (filterOptions.donViPhuTrach && filterOptions.donViPhuTrach.length > 0) {
        if (!filterOptions.donViPhuTrach.includes(donViPhuTrach)) {
            return false;
        }
    }

    if (filterOptions.tuNgay && filterOptions.denNgay) {
        const tuNgay = new Date(filterOptions.tuNgay);
        const denNgay = new Date(filterOptions.denNgay);

        tuNgay.setHours(0, 0, 0, 0);
        denNgay.setHours(23, 59, 59, 999);

        let ngaySoSanhStr;
        if (loaiDonHang === "Yêu cầu kiểm tra") {
            ngaySoSanhStr = ngayGiaoHang;
        } else {
            ngaySoSanhStr = ngayNhapKho;
        }

        const ngaySoSanh = parseDDMMYYYY(ngaySoSanhStr);
        if (!ngaySoSanh) {
            return false;
        }

        if (ngaySoSanh < tuNgay || ngaySoSanh > denNgay) {
            return false;
        }
    }

    if (filterOptions.loaiDonHang === "Tiêu chuẩn") {
        if (tenSanPham === "Vật tư khác" || trongLuongPhuKien !== "Tiêu chuẩn") {
            return false;
        }
    } else if (filterOptions.loaiDonHang === "Khác chuẩn") {
        if (trongLuongPhuKien !== "Khác chuẩn") {
            return false;
        }
    }

    return true;
}


function calculateThongTinCongTrinh(donHangRow, donHangColumnIndex) {
    const phuongThucBan = donHangRow[donHangColumnIndex['phuong_thuc_ban']] || '';
    const maNhanVien = donHangRow[donHangColumnIndex['ma_nhan_vien']] || '';
    const maKhachHang = donHangRow[donHangColumnIndex['ma_khach_hang']] || '';
    const tenNguoiLienHe = donHangRow[donHangColumnIndex['ten_nguoi_lien_he']] || '';
    const diaChiChiTiet = donHangRow[donHangColumnIndex['dia_chi_chi_tiet']] || '';
    const fileBaoGiaNPP = donHangRow[donHangColumnIndex['file_bao_gia_npp']] || '';
    const maHopDong = donHangRow[donHangColumnIndex['ma_hop_dong']] || '';
    const maBoHang = donHangRow[donHangColumnIndex['ma_bo_hang']] || '';
    const tenKhachHangCuoi = donHangRow[donHangColumnIndex['ten_khach_hang_cuoi']] || '';
    const diaChiKhachHangCuoi = donHangRow[donHangColumnIndex['dia_chi_khach_hang_cuoi']] || '';

    const nguoiDung = lookupNguoiDungByMaNV[maNhanVien] || {};
    const donVi = nguoiDung.donVi || '';
    const maKhachHangFromLookup = lookupKhachHangById[maKhachHang] || '';
    const giaoHang = lookupGiaoHangByMaBoHang[maBoHang] || {};
    const loaiViec = giaoHang.loaiViec || '';
    const nguoiGiaoHang = giaoHang.nguoiGiaoHang || '';
    const nguoiGiaoHangInfo = lookupNguoiDungByMaNV[nguoiGiaoHang] || {};
    const tenNguoiGiaoHang = nguoiGiaoHangInfo.tenNV || '';

    let thongTin = '';

    if (phuongThucBan !== "Bán chéo") {
        thongTin = (donVi === "Quang Minh" ? maKhachHangFromLookup + " - " : "") +
            tenNguoiLienHe + " - " +
            diaChiChiTiet + " - " +
            fileBaoGiaNPP + " - " +
            maHopDong +
            (loaiViec === "Lắp đặt" ? " - " + tenNguoiGiaoHang : "");
    } else {
        thongTin = (donVi === "Quang Minh" ? maKhachHangFromLookup + " - " : "") +
            tenKhachHangCuoi + " - " +
            diaChiKhachHangCuoi + " - " +
            fileBaoGiaNPP + " - " +
            maHopDong +
            (loaiViec === "Lắp đặt" ? " - " + tenNguoiGiaoHang : "");
    }

    return thongTin;
}

function calculateMnvCongTy(donHangRow, donHangColumnIndex) {
    const donViPhuTrach = donHangRow[donHangColumnIndex['don_vi_phu_trach']] || '';
    const maNhanVien = donHangRow[donHangColumnIndex['ma_nhan_vien']] || '';
    const tenNhanVien = donHangRow[donHangColumnIndex['ten_nhan_vien']] || '';

    let mnvCongTy = '';

    if (["BP. BH1", "BP. BH2", "BP. Dịch vụ"].includes(donViPhuTrach)) {
        const nguoiDung = lookupNguoiDungByMaNV[maNhanVien] || {};
        mnvCongTy = nguoiDung.mnvCongTy || '';
    } else {
        const nguoiDung = lookupNguoiDungByTenNV[tenNhanVien] || {};
        if (nguoiDung.mnvCongTy) {
            mnvCongTy = nguoiDung.mnvCongTy;
        } else {
            mnvCongTy = "191002.0.1";
        }
    }

    return mnvCongTy;
}

function displayResultsNhap(results) {
    const resultsBody = document.getElementById('results-body-nhap');
    const resultsTable = document.getElementById('results-table-nhap');
    const noResults = document.getElementById('no-results-nhap');
    const resultsCount = document.getElementById('results-count-nhap');

    resultsCount.textContent = `Kết quả: ${results.length} dòng`;
    resultsBody.innerHTML = '';

    if (results.length === 0) {
        resultsTable.style.display = 'none';
        noResults.style.display = 'block';
        return;
    }

    noResults.style.display = 'none';
    resultsTable.style.display = 'table';

    results.forEach(result => {
        const chiTietRow = result.chiTietRow;
        const donHangRow = result.donHangRow;
        const chiTietColumnIndex = result.chiTietColumnIndex;
        const donHangColumnIndex = result.donHangColumnIndex;

        const maDonHangID = chiTietRow[chiTietColumnIndex['ma_don_hang_id']] || '';
        const maSanPhamTheoDoi = chiTietRow[chiTietColumnIndex['ma_san_pham_theo_doi']] || '';
        const dienGiai = chiTietRow[chiTietColumnIndex['dien_giai']] || '';
        const ghiChu = chiTietRow[chiTietColumnIndex['ghi_chu']] || '';
        const dvt = chiTietRow[chiTietColumnIndex['dvt']] || '';
        const khoiLuong = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['khoi_luong']] || '');
        const donGiaNPP = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['don_gia_npp']] || '');
        const giaBanNPP = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['gia_ban_npp']] || '');

        const loaiDonHang = donHangRow[donHangColumnIndex['loai_don_hang']] || '';
        const ngayGiaoHang = donHangRow[donHangColumnIndex['ngay_giao_hang']] || '';
        const ngayNhapKho = donHangRow[donHangColumnIndex['ngay_nhap_kho']] || '';
        const xuongSanXuat = donHangRow[donHangColumnIndex['xuong_san_xuat']] || '';
        const maHopDong = donHangRow[donHangColumnIndex['ma_hop_dong']] || '';

        let ngayHaChToan = '';
        let ngayChungTu = '';

        if (loaiDonHang === "Yêu cầu kiểm tra") {
            ngayHaChToan = ngayGiaoHang;
            ngayChungTu = ngayGiaoHang;
        } else {
            ngayHaChToan = ngayNhapKho;
            ngayChungTu = ngayNhapKho;
        }

        let kho = '';
        if (xuongSanXuat === "Hà Nội") {
            kho = "K09.TP.CUA.HN";
        } else if (xuongSanXuat === "Long An") {
            kho = "K10.TP.CUA.LA";
        }

        let tenHang = '';
        if (dienGiai !== "") {
            tenHang = dienGiai;
        } else {
            tenHang = ghiChu;
        }

        const thongTinCongTrinh = calculateThongTinCongTrinh(donHangRow, donHangColumnIndex);
        const mnvCongTy = calculateMnvCongTy(donHangRow, donHangColumnIndex);

        const row = document.createElement('tr');
        row.innerHTML = `
    <td>${maDonHangID}</td>
    <td></td>
    <td>0</td>
    <td>${ngayHaChToan}</td>
    <td>${ngayChungTu}</td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td>${thongTinCongTrinh}</td>
    <td>${mnvCongTy}</td>
    <td></td>
    <td>VND</td>
    <td></td>
    <td>${maSanPhamTheoDoi}</td>
    <td>${tenHang}</td>
    <td>${kho}</td>
    <td></td>
    <td>155</td>
    <td>154</td>
    <td>${dvt}</td>
    <td>${khoiLuong}</td>
    <td>${donGiaNPP}</td>
    <td>${giaBanNPP}</td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td>${maSanPhamTheoDoi}</td>
    <td></td>
    <td>${maHopDong}</td>
    <td></td>
    <td></td>
    `;
        resultsBody.appendChild(row);
    });
}

function exportToExcelNhap() {
    if (filteredResultsNhap.length === 0) {
        showMessageNhap('Không có dữ liệu để xuất. Vui lòng thực hiện lọc trước.', 'error');
        return;
    }

    let csvContent = "\uFEFF";
    csvContent += "ID,Hiển thị trên sổ,Loại nhập kho,Ngày hạch toán (*),Ngày chứng từ (*),Số chứng từ (*),Mã đối tượng,Tên đối tượng,Người giao hàng,Diễn giải,Nhân viên bán hàng,Kèm theo,Loại tiền,Tỷ giá,Mã hàng (*),Tên hàng,Kho (*),Hàng hóa giữ hộ/bán hộ,TK Nợ (*),TK Có (*),ĐVT,Số lượng,Đơn giá,Thành tiền,Thành tiền quy đổi,Số lô,Hạn sử dụng,Khoản mục CP,Đơn vị,Đối tượng THCP,Công trình,Đơn đặt hàng,Hợp đồng bán,Mã thống kê\n";

    filteredResultsNhap.forEach(result => {
        const chiTietRow = result.chiTietRow;
        const donHangRow = result.donHangRow;
        const chiTietColumnIndex = result.chiTietColumnIndex;
        const donHangColumnIndex = result.donHangColumnIndex;

        const maDonHangID = chiTietRow[chiTietColumnIndex['ma_don_hang_id']] || '';
        const maSanPhamTheoDoi = chiTietRow[chiTietColumnIndex['ma_san_pham_theo_doi']] || '';
        const dienGiai = chiTietRow[chiTietColumnIndex['dien_giai']] || '';
        const ghiChu = chiTietRow[chiTietColumnIndex['ghi_chu']] || '';
        const dvt = chiTietRow[chiTietColumnIndex['dvt']] || '';
        const khoiLuong = formatNumberForCSV(chiTietRow[chiTietColumnIndex['khoi_luong']] || '');
        const donGiaNPP = formatNumberForCSV(chiTietRow[chiTietColumnIndex['don_gia_npp']] || '');
        const giaBanNPP = formatNumberForCSV(chiTietRow[chiTietColumnIndex['gia_ban_npp']] || '');

        const loaiDonHang = donHangRow[donHangColumnIndex['loai_don_hang']] || '';
        const ngayGiaoHang = donHangRow[donHangColumnIndex['ngay_giao_hang']] || '';
        const ngayNhapKho = donHangRow[donHangColumnIndex['ngay_nhap_kho']] || '';
        const xuongSanXuat = donHangRow[donHangColumnIndex['xuong_san_xuat']] || '';
        const maHopDong = donHangRow[donHangColumnIndex['ma_hop_dong']] || '';

        let ngayHaChToan = '';
        let ngayChungTu = '';

        if (loaiDonHang === "Yêu cầu kiểm tra") {
            ngayHaChToan = ngayGiaoHang;
            ngayChungTu = ngayGiaoHang;
        } else {
            ngayHaChToan = ngayNhapKho;
            ngayChungTu = ngayNhapKho;
        }

        let kho = '';
        if (xuongSanXuat === "Hà Nội") {
            kho = "K09.TP.CUA.HN";
        } else if (xuongSanXuat === "Long An") {
            kho = "K10.TP.CUA.LA";
        }

        let tenHang = '';
        if (dienGiai !== "") {
            tenHang = dienGiai;
        } else {
            tenHang = ghiChu;
        }

        const thongTinCongTrinh = calculateThongTinCongTrinh(donHangRow, donHangColumnIndex);
        const mnvCongTy = calculateMnvCongTy(donHangRow, donHangColumnIndex);

        const escapeCSV = (str) => {
            if (!str) return '';
            if (str.includes(',') || str.includes('"') || str.includes('\n')) {
                return '"' + str.replace(/"/g, '""') + '"';
            }
            return str;
        };

        csvContent += `${maDonHangID},,0,${escapeCSV(ngayHaChToan)},${escapeCSV(ngayChungTu)},,,,,${escapeCSV(thongTinCongTrinh)},${escapeCSV(mnvCongTy)},,VND,,${escapeCSV(maSanPhamTheoDoi)},${escapeCSV(tenHang)},${escapeCSV(kho)},,155,154,${escapeCSV(dvt)},${khoiLuong},${donGiaNPP},${giaBanNPP},,,,,,${escapeCSV(maSanPhamTheoDoi)},,${escapeCSV(maHopDong)},,\n`;
    });

    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement("a");
    const url = URL.createObjectURL(blob);

    link.setAttribute("href", url);
    const tuNgayRaw = document.getElementById("tu-ngay-nhap").value;
    const denNgayRaw = document.getElementById("den-ngay-nhap").value;
    const tuNgay = formatDateForFilename(tuNgayRaw);
    const denNgay = formatDateForFilename(denNgayRaw);
    const loaiDonHang = document.querySelector('input[name="loai-don-hang-nhap"]:checked').value;
    const fileName = `Danh sách nhập kho - ${tuNgay} - ${denNgay} - ${loaiDonHang}.csv`;

    link.setAttribute("download", fileName);
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);

    showMessageNhap(`Đã xuất ${filteredResultsNhap.length} dòng ra file Excel.`, 'success');
}

function showLoadingNhap(show) {
    const loadingElement = document.getElementById('loading-nhap');
    loadingElement.style.display = show ? 'block' : 'none';
}

function showMessageNhap(message, type) {
    const resultsCount = document.getElementById("results-count-nhap");
    resultsCount.className = "results-count";

    if (type === "success") {
        resultsCount.style.backgroundColor = "#e8f7e8";
        resultsCount.style.color = "#0c9c07";
        resultsCount.style.borderLeft = "4px solid #0c9c07";
    }

    if (type === "error") {
        resultsCount.style.backgroundColor = "#ffeaea";
        resultsCount.style.color = "#c00";
        resultsCount.style.borderLeft = "4px solid #c00";
    }

    resultsCount.style.padding = "5px";
    resultsCount.style.borderRadius = "6px";
    resultsCount.style.fontSize = "16px";
    resultsCount.style.fontWeight = "600";
    resultsCount.textContent = message;
}

function requireRefilterNhap() {
    document.getElementById('results-table-nhap').style.display = 'none';
    document.getElementById('no-results-nhap').style.display = 'block';
    document.getElementById('results-count-nhap').textContent = 'Kết quả: Chưa có dòng nào được lọc.';
    filteredResultsNhap = [];
    showMessageNhap("Bạn đã thay đổi bộ lọc. Vui lòng nhấn 'Lọc' để cập nhật kết quả mới.", "error");
}

// ==================== TAB XUẤT KHO ====================

let filteredResultsXuat = [];

function initXuatKhoEventListeners() {
    document.getElementById('btn-loc-xuat').addEventListener('click', applyFilterXuat);
    document.getElementById('btn-reset-xuat').addEventListener('click', resetFilterXuat);
    document.getElementById('btn-export-xuat').addEventListener('click', exportToExcelXuat);

    // Xóa event listener cho input text cũ
    // document.getElementById("ma-hop-dong-xuat").addEventListener("input", requireRefilterXuat);
    document.getElementById("tu-ngay-xuat").addEventListener("change", requireRefilterXuat);
    document.getElementById("den-ngay-xuat").addEventListener("change", requireRefilterXuat);

    document.querySelectorAll("input[name='loai-don-hang-xuat']").forEach(radio => {
        radio.addEventListener("change", requireRefilterXuat);
    });
}

function resetFilterXuat() {
    // Reset multi-select đơn vị phụ trách
    const donViAllCheckbox = document.getElementById('don-vi-phu-trach-xuat-all');
    if (donViAllCheckbox) {
        donViAllCheckbox.checked = true;
    }

    const donViCheckboxes = document.querySelectorAll('#don-vi-phu-trach-xuat-container input[type="checkbox"]:not(#don-vi-phu-trach-xuat-all)');
    donViCheckboxes.forEach(checkbox => {
        checkbox.checked = false;
    });

    // Reset multi-select mã hợp đồng
    const maHopDongAllCheckbox = document.getElementById('ma-hop-dong-xuat-all');
    if (maHopDongAllCheckbox) {
        maHopDongAllCheckbox.checked = true;
    }

    const maHopDongCheckboxes = document.querySelectorAll('#ma-hop-dong-xuat-container input[type="checkbox"]:not(#ma-hop-dong-xuat-all)');
    maHopDongCheckboxes.forEach(checkbox => {
        checkbox.checked = false;
    });

    document.getElementById('tu-ngay-xuat').value = '';
    document.getElementById('den-ngay-xuat').value = '';
    document.getElementById('loai-tat-ca-xuat').checked = true;

    document.getElementById('results-table-xuat').style.display = 'none';
    document.getElementById('no-results-xuat').style.display = 'block';
    document.getElementById('results-count-xuat').textContent = 'Kết quả: Chưa có dòng nào được lọc.';

    filteredResultsXuat = [];
}

async function applyFilterXuat() {
    showLoadingXuat(true);

    try {
        if (!isDataLoaded) {
            await fetchDataFromSheets();
        }

        const filterOptions = {
            maHopDong: getSelectedMaHopDong('-xuat'), // Lấy danh sách mã hợp đồng được chọn
            donViPhuTrach: getSelectedDonViPhuTrach('-xuat'), // Lấy danh sách đơn vị được chọn
            tuNgay: document.getElementById('tu-ngay-xuat').value,
            denNgay: document.getElementById('den-ngay-xuat').value,
            loaiDonHang: document.querySelector('input[name="loai-don-hang-xuat"]:checked').value
        };

        filteredResultsXuat = filterDataXuat(filterOptions);
        displayResultsXuat(filteredResultsXuat);
        showMessageXuat(`Đã tìm thấy ${filteredResultsXuat.length} dòng phù hợp.`, 'success');

    } catch (error) {
        console.error('Lỗi khi áp dụng bộ lọc:', error);
        showMessageXuat('Đã xảy ra lỗi khi tải dữ liệu. Vui lòng thử lại.', 'error');
    } finally {
        showLoadingXuat(false);
    }
}

function filterDataXuat(filterOptions) {
    if (donHangData.length === 0 || donHangChiTietData.length === 0) {
        return [];
    }

    const donHangHeaders = donHangData[0];
    const chiTietHeaders = donHangChiTietData[0];

    const donHangColumnIndex = {};
    donHangHeaders.forEach((header, index) => {
        donHangColumnIndex[header] = index;
    });

    const chiTietColumnIndex = {};
    chiTietHeaders.forEach((header, index) => {
        chiTietColumnIndex[header] = index;
    });

    const filteredResults = [];
    let resultId = 1;

    for (let i = 1; i < donHangChiTietData.length; i++) {
        const chiTietRow = donHangChiTietData[i];
        const maDonHangID = chiTietRow[chiTietColumnIndex['ma_don_hang_id']] || '';
        const donHangRow = donHangData.find(row => {
            return row[donHangColumnIndex['ma_don_hang']] === maDonHangID;
        });

        if (!donHangRow) continue;

        if (!passesFilterConditionsXuat(chiTietRow, donHangRow, chiTietColumnIndex, donHangColumnIndex, filterOptions)) {
            continue;
        }

        filteredResults.push({
            id: resultId++,
            chiTietRow: chiTietRow,
            donHangRow: donHangRow,
            chiTietColumnIndex: chiTietColumnIndex,
            donHangColumnIndex: donHangColumnIndex
        });
    }

    return filteredResults;
}

function passesFilterConditionsXuat(chiTietRow, donHangRow, chiTietColumnIndex, donHangColumnIndex, filterOptions) {
    const nhomSanPham = chiTietRow[chiTietColumnIndex['nhom_san_pham']] || '';
    const tenSanPham = chiTietRow[chiTietColumnIndex['ten_san_pham']] || '';
    const trongLuongPhuKien = donHangRow[donHangColumnIndex['trong_luong_phu_kien']] || '';
    const maHopDong = donHangRow[donHangColumnIndex['ma_hop_dong']] || '';
    const donViPhuTrach = donHangRow[donHangColumnIndex['don_vi_phu_trach']] || '';
    const loaiDonHang = donHangRow[donHangColumnIndex['loai_don_hang']] || '';
    const ngayGiaoHang = donHangRow[donHangColumnIndex['ngay_giao_hang']] || '';
    const ngayXuatKho = donHangRow[donHangColumnIndex['ngay_xuat_kho']] || '';

    if (nhomSanPham === "Bảo hành") {
        return false;
    }

    if (trongLuongPhuKien === "Tiêu chuẩn" && tenSanPham === "Vật tư khác") {
        return false;
    }

    // Lọc theo mã hợp đồng (multi-select)
    if (filterOptions.maHopDong && filterOptions.maHopDong.length > 0) {
        if (!filterOptions.maHopDong.includes(maHopDong)) {
            return false;
        }
    }

    // Lọc theo đơn vị phụ trách (multi-select)
    if (filterOptions.donViPhuTrach && filterOptions.donViPhuTrach.length > 0) {
        if (!filterOptions.donViPhuTrach.includes(donViPhuTrach)) {
            return false;
        }
    }

    if (filterOptions.tuNgay && filterOptions.denNgay) {
        const tuNgay = new Date(filterOptions.tuNgay);
        const denNgay = new Date(filterOptions.denNgay);

        tuNgay.setHours(0, 0, 0, 0);
        denNgay.setHours(23, 59, 59, 999);

        let ngaySoSanhStr;
        if (loaiDonHang === "Yêu cầu kiểm tra") {
            ngaySoSanhStr = ngayGiaoHang;
        } else {
            ngaySoSanhStr = ngayXuatKho;
        }

        const ngaySoSanh = parseDDMMYYYY(ngaySoSanhStr);
        if (!ngaySoSanh) {
            return false;
        }

        if (ngaySoSanh < tuNgay || ngaySoSanh > denNgay) {
            return false;
        }
    }

    if (filterOptions.loaiDonHang === "Tiêu chuẩn") {
        if (tenSanPham === "Vật tư khác" || trongLuongPhuKien !== "Tiêu chuẩn") {
            return false;
        }
    } else if (filterOptions.loaiDonHang === "Khác chuẩn") {
        if (trongLuongPhuKien !== "Khác chuẩn") {
            return false;
        }
    }

    return true;
}

function calculateTaiKhoanKho(nhomSanPham, maSanPhamId) {
    const condition1 = (nhomSanPham !== "Vật tư phát sinh" && nhomSanPham !== "Vật tư khác");
    const condition2 = maSanPhamId && maSanPhamId.includes("*");

    if (condition1 || condition2) {
        return "155";
    } else if (nhomSanPham === "Hàng hóa") {
        return "156";
    } else {
        return "1521";
    }
}

function displayResultsXuat(results) {
    const resultsBody = document.getElementById('results-body-xuat');
    const resultsTable = document.getElementById('results-table-xuat');
    const noResults = document.getElementById('no-results-xuat');
    const resultsCount = document.getElementById('results-count-xuat');

    resultsCount.textContent = `Kết quả: ${results.length} dòng`;
    resultsBody.innerHTML = '';

    if (results.length === 0) {
        resultsTable.style.display = 'none';
        noResults.style.display = 'block';
        return;
    }

    noResults.style.display = 'none';
    resultsTable.style.display = 'table';

    results.forEach(result => {
        const chiTietRow = result.chiTietRow;
        const donHangRow = result.donHangRow;
        const chiTietColumnIndex = result.chiTietColumnIndex;
        const donHangColumnIndex = result.donHangColumnIndex;

        const maDonHangID = chiTietRow[chiTietColumnIndex['ma_don_hang_id']] || '';
        const maSanPhamTheoDoi = chiTietRow[chiTietColumnIndex['ma_san_pham_theo_doi']] || '';
        const dienGiai = chiTietRow[chiTietColumnIndex['dien_giai']] || '';
        const ghiChu = chiTietRow[chiTietColumnIndex['ghi_chu']] || '';
        const dvt = chiTietRow[chiTietColumnIndex['dvt']] || '';
        const khoiLuong = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['khoi_luong']] || '');
        const donGiaNPP = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['don_gia_npp']] || '');
        const giaBanNPP = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['gia_ban_npp']] || '');
        const trongLuongNhom = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['trong_luong_nhom']] || '');
        const viTriLapDat = chiTietRow[chiTietColumnIndex['vi_tri_lap_dat']] || '';
        const nhomSanPham = chiTietRow[chiTietColumnIndex['nhom_san_pham']] || '';
        const maSanPhamId = chiTietRow[chiTietColumnIndex['ma_san_pham_id']] || '';
        const taiKhoanKho = calculateTaiKhoanKho(nhomSanPham, maSanPhamId);

        const loaiDonHang = donHangRow[donHangColumnIndex['loai_don_hang']] || '';
        const ngayGiaoHang = donHangRow[donHangColumnIndex['ngay_giao_hang']] || '';
        const ngayXuatKho = donHangRow[donHangColumnIndex['ngay_xuat_kho']] || '';
        const xuongSanXuat = donHangRow[donHangColumnIndex['xuong_san_xuat']] || '';
        const maHopDong = donHangRow[donHangColumnIndex['ma_hop_dong']] || '';
        const donViPhuTrach = donHangRow[donHangColumnIndex['don_vi_phu_trach']] || '';
        const maNhanVien = donHangRow[donHangColumnIndex['ma_nhan_vien']] || '';
        const maKhachHangID = donHangRow[donHangColumnIndex['ma_khach_hang_id']] || '';
        const tenNguoiLienHe = donHangRow[donHangColumnIndex['ten_nguoi_lien_he']] || '';
        const diaChi = donHangRow[donHangColumnIndex['dia_chi']] || '';
        const diaChiChiTiet = donHangRow[donHangColumnIndex['dia_chi_chi_tiet']] || '';
        const mucChietKhauNPP = formatNumberForDisplay(donHangRow[donHangColumnIndex['muc_chiet_khau_npp']] || '');
        const phiVanChuyenLapDatNPP = formatNumberForDisplay(donHangRow[donHangColumnIndex['phi_van_chuyen_lap_dat_npp']] || '');
        const soTienTTL1 = formatNumberForDisplay(donHangRow[donHangColumnIndex['so_tien_tt_l1']] || '');
        const soTienPhaiThanhToanNPP = formatNumberForDisplay(donHangRow[donHangColumnIndex['so_tien_phai_thanh_toan_npp']] || '');
        const fileBaoGiaNPP = donHangRow[donHangColumnIndex['file_bao_gia_npp']] || '';

        let ngayHaChToan = '';
        let ngayChungTu = '';

        if (loaiDonHang === "Yêu cầu kiểm tra") {
            ngayHaChToan = ngayGiaoHang;
            ngayChungTu = ngayGiaoHang;
        } else {
            ngayHaChToan = ngayXuatKho;
            ngayChungTu = ngayXuatKho;
        }

        let maKhachHang = '';
        let tenKhachHang = '';
        let diaChiHienThi = '';

        if (!["BP. BH1", "BP. BH2", "P. Bán hàng", "BP. Dịch vụ"].includes(donViPhuTrach)) {
            maKhachHang = maNhanVien.substring(0, 9);
            tenKhachHang = donViPhuTrach;
            diaChiHienThi = diaChi;
        } else {
            maKhachHang = maKhachHangID;
            tenKhachHang = tenNguoiLienHe;
            diaChiHienThi = diaChiChiTiet;
        }

        let tenHang = '';
        if (dienGiai !== "") {
            tenHang = dienGiai;
        } else {
            tenHang = ghiChu;
        }

        let kho = '';
        if (xuongSanXuat === "Hà Nội") {
            if (nhomSanPham !== "Vật tư phát sinh" && nhomSanPham !== "Vật tư khác") {
                kho = "K09.TP.CUA.HN";
            } else {
                kho = "K03_SX.HN_152";
            }
        } else if (xuongSanXuat === "Long An") {
            if (nhomSanPham !== "Vật tư phát sinh" && nhomSanPham !== "Vật tư khác") {
                kho = "K10.TP.CUA.LA";
            } else {
                kho = "K04_SX.LA_152";
            }
        }

        const thongTinCongTrinh = calculateThongTinCongTrinh(donHangRow, donHangColumnIndex);
        const mnvCongTy = calculateMnvCongTy(donHangRow, donHangColumnIndex);
        const ghiChuFull = `${viTriLapDat} - ${ghiChu}`;

        const row = document.createElement('tr');
        row.innerHTML = `
    <td>${maDonHangID}</td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td>0</td>
    <td></td>
    <td>${ngayHaChToan}</td>
    <td>${ngayChungTu}</td>
    <td>${maHopDong}</td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td>${maKhachHang}</td>
    <td>${tenKhachHang}</td>
    <td>${diaChiHienThi}</td>
    <td></td>
    <td>${thongTinCongTrinh}</td>
    <td></td>
    <td>${mnvCongTy}</td>
    <td>VND</td>
    <td></td>
    <td>${maSanPhamTheoDoi}</td>
    <td>${tenHang}</td>
    <td></td>
    <td>131</td>
    <td>511</td>
    <td>${dvt}</td>
    <td>${khoiLuong}</td>
    <td></td>
    <td>${donGiaNPP}</td>
    <td>${giaBanNPP}</td>
    <td></td>
    <td>${mucChietKhauNPP}</td>
    <td></td>
    <td></td>
    <td>5111Chietkhau</td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td>${trongLuongNhom}</td>
    <td></td>
    <td></td>
    <td></td>
    <td>33311</td>
    <td></td>
    <td>${kho}</td>
    <td>632</td>
    <td>${taiKhoanKho}</td>
    <td></td>
    <td></td>
    <td></td>
    <td>${phiVanChuyenLapDatNPP}</td>
    <td>${soTienTTL1}</td>
    <td>${ghiChuFull}</td>
    <td>${soTienPhaiThanhToanNPP}</td>
    <td>${fileBaoGiaNPP}</td>
    `;
        resultsBody.appendChild(row);
    });
}

function exportToExcelXuat() {
    if (filteredResultsXuat.length === 0) {
        showMessageXuat('Không có dữ liệu để xuất. Vui lòng thực hiện lọc trước.', 'error');
        return;
    }

    let csvContent = "\uFEFF";
    csvContent += "ID,Hiển thị trên sổ,Hình thức bán hàng,Phương thức thanh toán,Kiêm phiếu xuất kho,XK vào khu phi thuế quan và các TH được coi như XK,Lập kèm hóa đơn,Đã lập hóa đơn,Ngày hạch toán (*),Ngày chứng từ (*),Số chứng từ (*),Số phiếu xuất,Lý do xuất,Mẫu số HĐ,Ký hiệu HĐ,Số hóa đơn,Ngày hóa đơn,Mã khách hàng,Tên khách hàng,Địa chỉ,Mã số thuế,Diễn giải,Nộp vào TK,NV bán hàng,Loại tiền,Tỷ giá,Mã hàng (*),Tên hàng,Hàng khuyến mại,TK Tiền/Chi phí/Nợ (*),TK Doanh thu/Có (*),ĐVT,Số lượng,Đơn giá sau thuế,Đơn giá,Thành tiền,Thành tiền quy đổi,Tỷ lệ CK (%),Tiền chiết khấu,Tiền chiết khấu quy đổi,TK chiết khấu,Giá tính thuế XK,% thuế XK,Tiền thuế XK,TK thuế XK,% thuế GTGT,Tỷ lệ tính thuế (Thuế suất KHAC),Tiền thuế GTGT,Tiền thuế GTGT quy đổi,TK thuế GTGT,HH không TH trên tờ khai thuế GTGT,Kho,TK giá vốn,TK Kho,Đơn giá vốn,Tiền vốn,Hàng hóa giữ hộ/bán hộ,Phí vận chuyển lắp đặt,Chi phí dịch vụ sơn yêu cầu,Ghi chú,Tổng số tiền,Kiểu thanh toán\n";

    filteredResultsXuat.forEach(result => {
        const chiTietRow = result.chiTietRow;
        const donHangRow = result.donHangRow;
        const chiTietColumnIndex = result.chiTietColumnIndex;
        const donHangColumnIndex = result.donHangColumnIndex;

        const maDonHangID = chiTietRow[chiTietColumnIndex['ma_don_hang_id']] || '';
        const maSanPhamTheoDoi = chiTietRow[chiTietColumnIndex['ma_san_pham_theo_doi']] || '';
        const dienGiai = chiTietRow[chiTietColumnIndex['dien_giai']] || '';
        const ghiChu = chiTietRow[chiTietColumnIndex['ghi_chu']] || '';
        const dvt = chiTietRow[chiTietColumnIndex['dvt']] || '';
        const khoiLuong = formatNumberForCSV(chiTietRow[chiTietColumnIndex['khoi_luong']] || '');
        const donGiaNPP = formatNumberForCSV(chiTietRow[chiTietColumnIndex['don_gia_npp']] || '');
        const giaBanNPP = formatNumberForCSV(chiTietRow[chiTietColumnIndex['gia_ban_npp']] || '');
        const trongLuongNhom = formatNumberForCSV(chiTietRow[chiTietColumnIndex['trong_luong_nhom']] || '');
        const viTriLapDat = chiTietRow[chiTietColumnIndex['vi_tri_lap_dat']] || '';
        const nhomSanPham = chiTietRow[chiTietColumnIndex['nhom_san_pham']] || '';
        const maSanPhamId = chiTietRow[chiTietColumnIndex['ma_san_pham_id']] || '';
        const taiKhoanKho = calculateTaiKhoanKho(nhomSanPham, maSanPhamId);

        const loaiDonHang = donHangRow[donHangColumnIndex['loai_don_hang']] || '';
        const ngayGiaoHang = donHangRow[donHangColumnIndex['ngay_giao_hang']] || '';
        const ngayXuatKho = donHangRow[donHangColumnIndex['ngay_xuat_kho']] || '';
        const xuongSanXuat = donHangRow[donHangColumnIndex['xuong_san_xuat']] || '';
        const maHopDong = donHangRow[donHangColumnIndex['ma_hop_dong']] || '';
        const donViPhuTrach = donHangRow[donHangColumnIndex['don_vi_phu_trach']] || '';
        const maNhanVien = donHangRow[donHangColumnIndex['ma_nhan_vien']] || '';
        const maKhachHangID = donHangRow[donHangColumnIndex['ma_khach_hang_id']] || '';
        const tenNguoiLienHe = donHangRow[donHangColumnIndex['ten_nguoi_lien_he']] || '';
        const diaChi = donHangRow[donHangColumnIndex['dia_chi']] || '';
        const diaChiChiTiet = donHangRow[donHangColumnIndex['dia_chi_chi_tiet']] || '';
        const mucChietKhauNPP = formatNumberForCSV(donHangRow[donHangColumnIndex['muc_chiet_khau_npp']] || '');
        const phiVanChuyenLapDatNPP = formatNumberForCSV(donHangRow[donHangColumnIndex['phi_van_chuyen_lap_dat_npp']] || '');
        const soTienTTL1 = formatNumberForCSV(donHangRow[donHangColumnIndex['so_tien_tt_l1']] || '');
        const soTienPhaiThanhToanNPP = formatNumberForCSV(donHangRow[donHangColumnIndex['so_tien_phai_thanh_toan_npp']] || '');
        const fileBaoGiaNPP = donHangRow[donHangColumnIndex['file_bao_gia_npp']] || '';

        let ngayHaChToan = '';
        let ngayChungTu = '';

        if (loaiDonHang === "Yêu cầu kiểm tra") {
            ngayHaChToan = ngayGiaoHang;
            ngayChungTu = ngayGiaoHang;
        } else {
            ngayHaChToan = ngayXuatKho;
            ngayChungTu = ngayXuatKho;
        }

        let maKhachHang = '';
        let tenKhachHang = '';
        let diaChiHienThi = '';

        if (!["BP. BH1", "BP. BH2", "P. Bán hàng", "BP. Dịch vụ"].includes(donViPhuTrach)) {
            maKhachHang = maNhanVien.substring(0, 9);
            tenKhachHang = donViPhuTrach;
            diaChiHienThi = diaChi;
        } else {
            maKhachHang = maKhachHangID;
            tenKhachHang = tenNguoiLienHe;
            diaChiHienThi = diaChiChiTiet;
        }

        let tenHang = '';
        if (dienGiai !== "") {
            tenHang = dienGiai;
        } else {
            tenHang = ghiChu;
        }

        let kho = '';
        if (xuongSanXuat === "Hà Nội") {
            if (nhomSanPham !== "Vật tư phát sinh" && nhomSanPham !== "Vật tư khác") {
                kho = "K09.TP.CUA.HN";
            } else {
                kho = "K03_SX.HN_152";
            }
        } else if (xuongSanXuat === "Long An") {
            if (nhomSanPham !== "Vật tư phát sinh" && nhomSanPham !== "Vật tư khác") {
                kho = "K10.TP.CUA.LA";
            } else {
                kho = "K04_SX.LA_152";
            }
        }

        const thongTinCongTrinh = calculateThongTinCongTrinh(donHangRow, donHangColumnIndex);
        const mnvCongTy = calculateMnvCongTy(donHangRow, donHangColumnIndex);
        const ghiChuFull = `${viTriLapDat} - ${ghiChu}`;

        const escapeCSV = (str) => {
            if (!str) return '';
            if (str.includes(',') || str.includes('"') || str.includes('\n')) {
                return '"' + str.replace(/"/g, '""') + '"';
            }
            return str;
        };

        csvContent += `${escapeCSV(maDonHangID)},,,,,,0,,${escapeCSV(ngayHaChToan)},${escapeCSV(ngayChungTu)},${escapeCSV(maHopDong)},,,,,,,${escapeCSV(maKhachHang)},${escapeCSV(tenKhachHang)},${escapeCSV(diaChiHienThi)},,${escapeCSV(thongTinCongTrinh)},,${escapeCSV(mnvCongTy)},VND,,${escapeCSV(maSanPhamTheoDoi)},${escapeCSV(tenHang)},,131,511,${escapeCSV(dvt)},${khoiLuong},,${donGiaNPP},${giaBanNPP},,${mucChietKhauNPP},,,5111Chietkhau,,,,,${trongLuongNhom},,,,33311,,${escapeCSV(kho)},632,${escapeCSV(taiKhoanKho)},,,,${phiVanChuyenLapDatNPP},${soTienTTL1},${escapeCSV(ghiChuFull)},${soTienPhaiThanhToanNPP},${escapeCSV(fileBaoGiaNPP)}\n`;
    });

    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement("a");
    const url = URL.createObjectURL(blob);

    link.setAttribute("href", url);
    const tuNgayRaw = document.getElementById("tu-ngay-xuat").value;
    const denNgayRaw = document.getElementById("den-ngay-xuat").value;
    const tuNgay = formatDateForFilename(tuNgayRaw);
    const denNgay = formatDateForFilename(denNgayRaw);
    const loaiDonHang = document.querySelector('input[name="loai-don-hang-xuat"]:checked').value;
    const fileName = `Danh sách xuất kho - ${tuNgay} - ${denNgay} - ${loaiDonHang}.csv`;

    link.setAttribute("download", fileName);
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);

    showMessageXuat(`Đã xuất ${filteredResultsXuat.length} dòng ra file Excel.`, 'success');
}

function showLoadingXuat(show) {
    const loadingElement = document.getElementById('loading-xuat');
    loadingElement.style.display = show ? 'block' : 'none';
}

function showMessageXuat(message, type) {
    const resultsCount = document.getElementById("results-count-xuat");
    resultsCount.className = "results-count";

    if (type === "success") {
        resultsCount.style.backgroundColor = "#e8f7e8";
        resultsCount.style.color = "#0c9c07";
        resultsCount.style.borderLeft = "4px solid #0c9c07";
    }

    if (type === "error") {
        resultsCount.style.backgroundColor = "#ffeaea";
        resultsCount.style.color = "#c00";
        resultsCount.style.borderLeft = "4px solid #c00";
    }

    resultsCount.style.padding = "5px";
    resultsCount.style.borderRadius = "6px";
    resultsCount.style.fontSize = "16px";
    resultsCount.style.fontWeight = "600";
    resultsCount.textContent = message;
}

function requireRefilterXuat() {
    document.getElementById('results-table-xuat').style.display = 'none';
    document.getElementById('no-results-xuat').style.display = 'block';
    document.getElementById('results-count-xuat').textContent = 'Kết quả: Chưa có dòng nào được lọc.';
    filteredResultsXuat = [];
    showMessageXuat("Bạn đã thay đổi bộ lọc. Vui lòng nhấn 'Lọc' để cập nhật kết quả mới.", "error");
}

// ==================== TAB LỆNH SẢN XUẤT ====================

let filteredResultsLenSanXuat = [];

function initLenSanXuatEventListeners() {
    document.getElementById('btn-loc-lsx').addEventListener('click', applyFilterLenSanXuat);
    document.getElementById('btn-reset-lsx').addEventListener('click', resetFilterLenSanXuat);
    document.getElementById('btn-export-lsx').addEventListener('click', exportToExcelLenSanXuat);

    // Xóa event listener cho input text cũ
    // document.getElementById("ma-hop-dong-lsx").addEventListener("input", requireRefilterLenSanXuat);
    document.getElementById("xuong-san-xuat-lsx").addEventListener("change", requireRefilterLenSanXuat);
    document.getElementById("tu-ngay-lsx").addEventListener("change", requireRefilterLenSanXuat);
    document.getElementById("den-ngay-lsx").addEventListener("change", requireRefilterLenSanXuat);

    document.querySelectorAll("input[name='loai-don-hang-lsx']").forEach(radio => {
        radio.addEventListener("change", requireRefilterLenSanXuat);
    });
}

function resetFilterLenSanXuat() {
    // Reset multi-select đơn vị phụ trách
    const donViAllCheckbox = document.getElementById('don-vi-phu-trach-lsx-all');
    if (donViAllCheckbox) {
        donViAllCheckbox.checked = true;
    }

    const donViCheckboxes = document.querySelectorAll('#don-vi-phu-trach-lsx-container input[type="checkbox"]:not(#don-vi-phu-trach-lsx-all)');
    donViCheckboxes.forEach(checkbox => {
        checkbox.checked = false;
    });

    // Reset multi-select mã hợp đồng
    const maHopDongAllCheckbox = document.getElementById('ma-hop-dong-lsx-all');
    if (maHopDongAllCheckbox) {
        maHopDongAllCheckbox.checked = true;
    }

    const maHopDongCheckboxes = document.querySelectorAll('#ma-hop-dong-lsx-container input[type="checkbox"]:not(#ma-hop-dong-lsx-all)');
    maHopDongCheckboxes.forEach(checkbox => {
        checkbox.checked = false;
    });

    document.getElementById('xuong-san-xuat-lsx').value = '';
    document.getElementById('tu-ngay-lsx').value = '';
    document.getElementById('den-ngay-lsx').value = '';
    document.getElementById('loai-tat-ca-lsx').checked = true;

    document.getElementById('results-table-lsx').style.display = 'none';
    document.getElementById('no-results-lsx').style.display = 'block';
    document.getElementById('results-count-lsx').textContent = 'Kết quả: Chưa có dòng nào được lọc.';

    filteredResultsLenSanXuat = [];
}

async function applyFilterLenSanXuat() {
    showLoadingLenSanXuat(true);

    try {
        if (!isDataLoaded) {
            await fetchDataFromSheets();
        }

        const filterOptions = {
            maHopDong: getSelectedMaHopDong('-lsx'), // Lấy danh sách mã hợp đồng được chọn
            donViPhuTrach: getSelectedDonViPhuTrach('-lsx'), // Lấy danh sách đơn vị được chọn
            xuongSanXuat: document.getElementById('xuong-san-xuat-lsx').value,
            tuNgay: document.getElementById('tu-ngay-lsx').value,
            denNgay: document.getElementById('den-ngay-lsx').value,
            loaiDonHang: document.querySelector('input[name="loai-don-hang-lsx"]:checked').value
        };

        filteredResultsLenSanXuat = filterDataLenSanXuat(filterOptions);
        displayResultsLenSanXuat(filteredResultsLenSanXuat);
        showMessageLenSanXuat(`Đã tìm thấy ${filteredResultsLenSanXuat.length} dòng phù hợp.`, 'success');

    } catch (error) {
        console.error('Lỗi khi áp dụng bộ lọc:', error);
        showMessageLenSanXuat('Đã xảy ra lỗi khi tải dữ liệu. Vui lòng thử lại.', 'error');
    } finally {
        showLoadingLenSanXuat(false);
    }
}

function filterDataLenSanXuat(filterOptions) {
    if (donHangData.length === 0 || donHangChiTietData.length === 0) {
        return [];
    }

    const donHangHeaders = donHangData[0];
    const chiTietHeaders = donHangChiTietData[0];

    const donHangColumnIndex = {};
    donHangHeaders.forEach((header, index) => {
        donHangColumnIndex[header] = index;
    });

    const chiTietColumnIndex = {};
    chiTietHeaders.forEach((header, index) => {
        chiTietColumnIndex[header] = index;
    });

    const filteredResults = [];
    let resultId = 1;

    for (let i = 1; i < donHangChiTietData.length; i++) {
        const chiTietRow = donHangChiTietData[i];
        const maDonHangID = chiTietRow[chiTietColumnIndex['ma_don_hang_id']] || '';
        const donHangRow = donHangData.find(row => {
            return row[donHangColumnIndex['ma_don_hang']] === maDonHangID;
        });

        if (!donHangRow) continue;

        if (!passesFilterConditionsLenSanXuat(chiTietRow, donHangRow, chiTietColumnIndex, donHangColumnIndex, filterOptions)) {
            continue;
        }

        filteredResults.push({
            id: resultId++,
            chiTietRow: chiTietRow,
            donHangRow: donHangRow,
            chiTietColumnIndex: chiTietColumnIndex,
            donHangColumnIndex: donHangColumnIndex
        });
    }

    return filteredResults;
}

function passesFilterConditionsLenSanXuat(chiTietRow, donHangRow, chiTietColumnIndex, donHangColumnIndex, filterOptions) {
    const nhomSanPham = chiTietRow[chiTietColumnIndex['nhom_san_pham']] || '';
    const mauCua = chiTietRow[chiTietColumnIndex['mau_cua']] || '';
    const tenSanPham = chiTietRow[chiTietColumnIndex['ten_san_pham']] || '';
    const trongLuongPhuKien = donHangRow[donHangColumnIndex['trong_luong_phu_kien']] || '';
    const maHopDong = donHangRow[donHangColumnIndex['ma_hop_dong']] || '';
    const donViPhuTrach = donHangRow[donHangColumnIndex['don_vi_phu_trach']] || '';
    const loaiDonHang = donHangRow[donHangColumnIndex['loai_don_hang']] || '';
    const ngayGiaoHang = donHangRow[donHangColumnIndex['ngay_giao_hang']] || '';
    const ngayNhapKho = donHangRow[donHangColumnIndex['ngay_nhap_kho']] || '';
    const xuongSanXuat = donHangRow[donHangColumnIndex['xuong_san_xuat']] || '';

    // Loại bỏ các nhóm sản phẩm không mong muốn
    const excludedGroups = ["Vật tư phát sinh", "Vật tư khác", "Bảo hành", "Hàng hóa"];
    if (excludedGroups.includes(nhomSanPham)) {
        return false;
    }

    if (mauCua === "Nhân công") {
        return false;
    }

    if (tenSanPham === "Di chuyển" || tenSanPham === "Nhân công") {
        return false;
    }

    // Lọc theo mã hợp đồng (multi-select)
    if (filterOptions.maHopDong && filterOptions.maHopDong.length > 0) {
        if (!filterOptions.maHopDong.includes(maHopDong)) {
            return false;
        }
    }

    // Lọc theo đơn vị phụ trách (multi-select)
    if (filterOptions.donViPhuTrach && filterOptions.donViPhuTrach.length > 0) {
        if (!filterOptions.donViPhuTrach.includes(donViPhuTrach)) {
            return false;
        }
    }

    // Lọc theo ngày
    if (filterOptions.tuNgay && filterOptions.denNgay) {
        const tuNgay = new Date(filterOptions.tuNgay);
        const denNgay = new Date(filterOptions.denNgay);

        tuNgay.setHours(0, 0, 0, 0);
        denNgay.setHours(23, 59, 59, 999);

        let ngaySoSanhStr;
        if (loaiDonHang === "Yêu cầu kiểm tra") {
            ngaySoSanhStr = ngayGiaoHang;
        } else {
            ngaySoSanhStr = ngayNhapKho;
        }

        const ngaySoSanh = parseDDMMYYYY(ngaySoSanhStr);
        if (!ngaySoSanh) {
            return false;
        }

        if (ngaySoSanh < tuNgay || ngaySoSanh > denNgay) {
            return false;
        }
    }

    // Lọc theo loại đơn hàng
    if (filterOptions.loaiDonHang === "Tiêu chuẩn") {
        if (tenSanPham === "Vật tư khác" || trongLuongPhuKien !== "Tiêu chuẩn") {
            return false;
        }
    } else if (filterOptions.loaiDonHang === "Khác chuẩn") {
        if (trongLuongPhuKien !== "Khác chuẩn") {
            return false;
        }
    }

    // Lọc theo xưởng sản xuất
    if (filterOptions.xuongSanXuat && xuongSanXuat !== filterOptions.xuongSanXuat) {
        return false;
    }

    return true;
}

function displayResultsLenSanXuat(results) {
    const resultsBody = document.getElementById('results-body-lsx');
    const resultsTable = document.getElementById('results-table-lsx');
    const noResults = document.getElementById('no-results-lsx');
    const resultsCount = document.getElementById('results-count-lsx');

    resultsCount.textContent = `Kết quả: ${results.length} dòng`;
    resultsBody.innerHTML = '';

    if (results.length === 0) {
        resultsTable.style.display = 'none';
        noResults.style.display = 'block';
        return;
    }

    noResults.style.display = 'none';
    resultsTable.style.display = 'table';

    results.forEach(result => {
        const chiTietRow = result.chiTietRow;
        const donHangRow = result.donHangRow;
        const chiTietColumnIndex = result.chiTietColumnIndex;
        const donHangColumnIndex = result.donHangColumnIndex;

        const maDonHangID = chiTietRow[chiTietColumnIndex['ma_don_hang_id']] || '';
        const maSanPhamTheoDoi = chiTietRow[chiTietColumnIndex['ma_san_pham_theo_doi']] || '';
        const dienGiai = chiTietRow[chiTietColumnIndex['dien_giai']] || '';
        const ghiChu = chiTietRow[chiTietColumnIndex['ghi_chu']] || '';
        const dvt = chiTietRow[chiTietColumnIndex['dvt']] || '';
        const soLuong = chiTietRow[chiTietColumnIndex['so_luong']] || '';
        const khoiLuong = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['khoi_luong']] || '');
        const sttTrongDon = chiTietRow[chiTietColumnIndex['stt_trong_don']] || '';

        const loaiDonHang = donHangRow[donHangColumnIndex['loai_don_hang']] || '';
        const ngayGiaoHang = donHangRow[donHangColumnIndex['ngay_giao_hang']] || '';
        const ngayNhapKho = donHangRow[donHangColumnIndex['ngay_nhap_kho']] || '';
        const xuongSanXuat = donHangRow[donHangColumnIndex['xuong_san_xuat']] || '';
        const maHopDong = donHangRow[donHangColumnIndex['ma_hop_dong']] || '';

        let ngayHaChToan = '';
        let ngayChungTu = '';

        if (loaiDonHang === "Yêu cầu kiểm tra") {
            ngayHaChToan = ngayGiaoHang;
            ngayChungTu = ngayGiaoHang;
        } else {
            ngayHaChToan = ngayNhapKho;
            ngayChungTu = ngayNhapKho;
        }

        let tenHang = '';
        if (dienGiai !== "") {
            tenHang = dienGiai;
        } else {
            tenHang = ghiChu;
        }

        const mnvCongTy = calculateMnvCongTy(donHangRow, donHangColumnIndex);

        // Xác định kho dựa trên xưởng sản xuất
        let kho = '';
        if (xuongSanXuat === "Hà Nội") {
            kho = "K03_SX.HN_152";
        } else if (xuongSanXuat === "Long An") {
            kho = "K04_SX.LA_152";
        }

        // Xác định đối tượng THCP dựa trên xưởng sản xuất
        let doiTuongTHCP = '';
        if (xuongSanXuat === "Hà Nội") {
            doiTuongTHCP = "25.SXT.X1";
        } else if (xuongSanXuat === "Long An") {
            doiTuongTHCP = "25.SXT.X2";
        }

        // Lấy các trường m_vt, kt_m_vt, sl_m_vt (từ 1 đến 20)
        const m_vt_1 = chiTietRow[chiTietColumnIndex['m_vt_1']] || '';
        const kt_m_vt_1 = chiTietRow[chiTietColumnIndex['kt_m_vt_1']] || '';
        const sl_m_vt_1 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_m_vt_1']] || '');
        const m_vt_2 = chiTietRow[chiTietColumnIndex['m_vt_2']] || '';
        const kt_m_vt_2 = chiTietRow[chiTietColumnIndex['kt_m_vt_2']] || '';
        const sl_m_vt_2 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_m_vt_2']] || '');
        const m_vt_3 = chiTietRow[chiTietColumnIndex['m_vt_3']] || '';
        const kt_m_vt_3 = chiTietRow[chiTietColumnIndex['kt_m_vt_3']] || '';
        const sl_m_vt_3 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_m_vt_3']] || '');
        const m_vt_4 = chiTietRow[chiTietColumnIndex['m_vt_4']] || '';
        const kt_m_vt_4 = chiTietRow[chiTietColumnIndex['kt_m_vt_4']] || '';
        const sl_m_vt_4 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_m_vt_4']] || '');
        const m_vt_5 = chiTietRow[chiTietColumnIndex['m_vt_5']] || '';
        const kt_m_vt_5 = chiTietRow[chiTietColumnIndex['kt_m_vt_5']] || '';
        const sl_m_vt_5 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_m_vt_5']] || '');
        const m_vt_6 = chiTietRow[chiTietColumnIndex['m_vt_6']] || '';
        const kt_m_vt_6 = chiTietRow[chiTietColumnIndex['kt_m_vt_6']] || '';
        const sl_m_vt_6 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_m_vt_6']] || '');
        const m_vt_7 = chiTietRow[chiTietColumnIndex['m_vt_7']] || '';
        const kt_m_vt_7 = chiTietRow[chiTietColumnIndex['kt_m_vt_7']] || '';
        const sl_m_vt_7 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_m_vt_7']] || '');
        const m_vt_8 = chiTietRow[chiTietColumnIndex['m_vt_8']] || '';
        const kt_m_vt_8 = chiTietRow[chiTietColumnIndex['kt_m_vt_8']] || '';
        const sl_m_vt_8 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_m_vt_8']] || '');
        const m_vt_9 = chiTietRow[chiTietColumnIndex['m_vt_9']] || '';
        const kt_m_vt_9 = chiTietRow[chiTietColumnIndex['kt_m_vt_9']] || '';
        const sl_m_vt_9 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_m_vt_9']] || '');
        const m_vt_10 = chiTietRow[chiTietColumnIndex['m_vt_10']] || '';
        const kt_m_vt_10 = chiTietRow[chiTietColumnIndex['kt_m_vt_10']] || '';
        const sl_m_vt_10 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_m_vt_10']] || '');
        const m_vt_11 = chiTietRow[chiTietColumnIndex['m_vt_11']] || '';
        const kt_m_vt_11 = chiTietRow[chiTietColumnIndex['kt_m_vt_11']] || '';
        const sl_m_vt_11 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_m_vt_11']] || '');
        const m_vt_12 = chiTietRow[chiTietColumnIndex['m_vt_12']] || '';
        const kt_m_vt_12 = chiTietRow[chiTietColumnIndex['kt_m_vt_12']] || '';
        const sl_m_vt_12 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_m_vt_12']] || '');
        const m_vt_13 = chiTietRow[chiTietColumnIndex['m_vt_13']] || '';
        const kt_m_vt_13 = chiTietRow[chiTietColumnIndex['kt_m_vt_13']] || '';
        const sl_m_vt_13 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_m_vt_13']] || '');
        const m_vt_14 = chiTietRow[chiTietColumnIndex['m_vt_14']] || '';
        const kt_m_vt_14 = chiTietRow[chiTietColumnIndex['kt_m_vt_14']] || '';
        const sl_m_vt_14 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_m_vt_14']] || '');
        const m_vt_15 = chiTietRow[chiTietColumnIndex['m_vt_15']] || '';
        const kt_m_vt_15 = chiTietRow[chiTietColumnIndex['kt_m_vt_15']] || '';
        const sl_m_vt_15 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_m_vt_15']] || '');
        const m_vt_16 = chiTietRow[chiTietColumnIndex['m_vt_16']] || '';
        const kt_m_vt_16 = chiTietRow[chiTietColumnIndex['kt_m_vt_16']] || '';
        const sl_m_vt_16 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_m_vt_16']] || '');
        const m_vt_17 = chiTietRow[chiTietColumnIndex['m_vt_17']] || '';
        const kt_m_vt_17 = chiTietRow[chiTietColumnIndex['kt_m_vt_17']] || '';
        const sl_m_vt_17 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_m_vt_17']] || '');
        const m_vt_18 = chiTietRow[chiTietColumnIndex['m_vt_18']] || '';
        const kt_m_vt_18 = chiTietRow[chiTietColumnIndex['kt_m_vt_18']] || '';
        const sl_m_vt_18 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_m_vt_18']] || '');
        const m_vt_19 = chiTietRow[chiTietColumnIndex['m_vt_19']] || '';
        const kt_m_vt_19 = chiTietRow[chiTietColumnIndex['kt_m_vt_19']] || '';
        const sl_m_vt_19 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_m_vt_19']] || '');
        const m_vt_20 = chiTietRow[chiTietColumnIndex['m_vt_20']] || '';
        const kt_m_vt_20 = chiTietRow[chiTietColumnIndex['kt_m_vt_20']] || '';
        const sl_m_vt_20 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_m_vt_20']] || '');

        // Lấy các trường m2_vt, kt1_m2_vt, kt2_m2_vt, sl_m2_vt (1 đến 2)
        const m2_vt_1 = chiTietRow[chiTietColumnIndex['m2_vt_1']] || '';
        const kt1_m2_vt_1 = chiTietRow[chiTietColumnIndex['kt1_m2_vt_1']] || '';
        const kt2_m2_vt_1 = chiTietRow[chiTietColumnIndex['kt2_m2_vt_1']] || '';
        const sl_m2_vt_1 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_m2_vt_1']] || '');
        const m2_vt_2 = chiTietRow[chiTietColumnIndex['m2_vt_2']] || '';
        const kt1_m2_vt_2 = chiTietRow[chiTietColumnIndex['kt1_m2_vt_2']] || '';
        const kt2_m2_vt_2 = chiTietRow[chiTietColumnIndex['kt2_m2_vt_2']] || '';
        const sl_m2_vt_2 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_m2_vt_2']] || '');

        // Lấy các trường c_vt, sl_c_vt (1 đến 30)
        const c_vt_1 = chiTietRow[chiTietColumnIndex['c_vt_1']] || '';
        const sl_c_vt_1 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_1']] || '');
        const c_vt_2 = chiTietRow[chiTietColumnIndex['c_vt_2']] || '';
        const sl_c_vt_2 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_2']] || '');
        const c_vt_3 = chiTietRow[chiTietColumnIndex['c_vt_3']] || '';
        const sl_c_vt_3 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_3']] || '');
        const c_vt_4 = chiTietRow[chiTietColumnIndex['c_vt_4']] || '';
        const sl_c_vt_4 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_4']] || '');
        const c_vt_5 = chiTietRow[chiTietColumnIndex['c_vt_5']] || '';
        const sl_c_vt_5 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_5']] || '');
        const c_vt_6 = chiTietRow[chiTietColumnIndex['c_vt_6']] || '';
        const sl_c_vt_6 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_6']] || '');
        const c_vt_7 = chiTietRow[chiTietColumnIndex['c_vt_7']] || '';
        const sl_c_vt_7 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_7']] || '');
        const c_vt_8 = chiTietRow[chiTietColumnIndex['c_vt_8']] || '';
        const sl_c_vt_8 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_8']] || '');
        const c_vt_9 = chiTietRow[chiTietColumnIndex['c_vt_9']] || '';
        const sl_c_vt_9 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_9']] || '');
        const c_vt_10 = chiTietRow[chiTietColumnIndex['c_vt_10']] || '';
        const sl_c_vt_10 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_10']] || '');
        const c_vt_11 = chiTietRow[chiTietColumnIndex['c_vt_11']] || '';
        const sl_c_vt_11 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_11']] || '');
        const c_vt_12 = chiTietRow[chiTietColumnIndex['c_vt_12']] || '';
        const sl_c_vt_12 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_12']] || '');
        const c_vt_13 = chiTietRow[chiTietColumnIndex['c_vt_13']] || '';
        const sl_c_vt_13 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_13']] || '');
        const c_vt_14 = chiTietRow[chiTietColumnIndex['c_vt_14']] || '';
        const sl_c_vt_14 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_14']] || '');
        const c_vt_15 = chiTietRow[chiTietColumnIndex['c_vt_15']] || '';
        const sl_c_vt_15 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_15']] || '');
        const c_vt_16 = chiTietRow[chiTietColumnIndex['c_vt_16']] || '';
        const sl_c_vt_16 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_16']] || '');
        const c_vt_17 = chiTietRow[chiTietColumnIndex['c_vt_17']] || '';
        const sl_c_vt_17 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_17']] || '');
        const c_vt_18 = chiTietRow[chiTietColumnIndex['c_vt_18']] || '';
        const sl_c_vt_18 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_18']] || '');
        const c_vt_19 = chiTietRow[chiTietColumnIndex['c_vt_19']] || '';
        const sl_c_vt_19 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_19']] || '');
        const c_vt_20 = chiTietRow[chiTietColumnIndex['c_vt_20']] || '';
        const sl_c_vt_20 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_20']] || '');
        const c_vt_21 = chiTietRow[chiTietColumnIndex['c_vt_21']] || '';
        const sl_c_vt_21 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_21']] || '');
        const c_vt_22 = chiTietRow[chiTietColumnIndex['c_vt_22']] || '';
        const sl_c_vt_22 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_22']] || '');
        const c_vt_23 = chiTietRow[chiTietColumnIndex['c_vt_23']] || '';
        const sl_c_vt_23 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_23']] || '');
        const c_vt_24 = chiTietRow[chiTietColumnIndex['c_vt_24']] || '';
        const sl_c_vt_24 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_24']] || '');
        const c_vt_25 = chiTietRow[chiTietColumnIndex['c_vt_25']] || '';
        const sl_c_vt_25 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_25']] || '');
        const c_vt_26 = chiTietRow[chiTietColumnIndex['c_vt_26']] || '';
        const sl_c_vt_26 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_26']] || '');
        const c_vt_27 = chiTietRow[chiTietColumnIndex['c_vt_27']] || '';
        const sl_c_vt_27 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_27']] || '');
        const c_vt_28 = chiTietRow[chiTietColumnIndex['c_vt_28']] || '';
        const sl_c_vt_28 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_28']] || '');
        const c_vt_29 = chiTietRow[chiTietColumnIndex['c_vt_29']] || '';
        const sl_c_vt_29 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_29']] || '');
        const c_vt_30 = chiTietRow[chiTietColumnIndex['c_vt_30']] || '';
        const sl_c_vt_30 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_30']] || '');

        // Tạo hàng với các cột theo yêu cầu
        const row = document.createElement('tr');
        row.innerHTML = `
    <td>${maDonHangID}</td>
    <td>${sttTrongDon}</td>
    <td>1</td>
    <td>${ngayHaChToan}</td>
    <td>${ngayChungTu}</td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td>Xuất kho NVL sản xuất cho Lệnh sản xuất <${maHopDong}></td>
    <td>${mnvCongTy}</td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td>${ghiChu}</td>
    <td>${kho}</td>
    <td></td>
    <td>154</td>
    <td>1521</td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td>${doiTuongTHCP}</td>
    <td>${maSanPhamTheoDoi}</td>
    <td></td>
    <td>${maHopDong}</td>
    <td></td>
    <td></td>
    <td></td>
    <td>${m_vt_1}</td>
    <td>${kt_m_vt_1}</td>
    <td>${sl_m_vt_1}</td>
    <td>${m_vt_2}</td>
    <td>${kt_m_vt_2}</td>
    <td>${sl_m_vt_2}</td>
    <td>${m_vt_3}</td>
    <td>${kt_m_vt_3}</td>
    <td>${sl_m_vt_3}</td>
    <td>${m_vt_4}</td>
    <td>${kt_m_vt_4}</td>
    <td>${sl_m_vt_4}</td>
    <td>${m_vt_5}</td>
    <td>${kt_m_vt_5}</td>
    <td>${sl_m_vt_5}</td>
    <td>${m_vt_6}</td>
    <td>${kt_m_vt_6}</td>
    <td>${sl_m_vt_6}</td>
    <td>${m_vt_7}</td>
    <td>${kt_m_vt_7}</td>
    <td>${sl_m_vt_7}</td>
    <td>${m_vt_8}</td>
    <td>${kt_m_vt_8}</td>
    <td>${sl_m_vt_8}</td>
    <td>${m_vt_9}</td>
    <td>${kt_m_vt_9}</td>
    <td>${sl_m_vt_9}</td>
    <td>${m_vt_10}</td>
    <td>${kt_m_vt_10}</td>
    <td>${sl_m_vt_10}</td>
    <td>${m_vt_11}</td>
    <td>${kt_m_vt_11}</td>
    <td>${sl_m_vt_11}</td>
    <td>${m_vt_12}</td>
    <td>${kt_m_vt_12}</td>
    <td>${sl_m_vt_12}</td>
    <td>${m_vt_13}</td>
    <td>${kt_m_vt_13}</td>
    <td>${sl_m_vt_13}</td>
    <td>${m_vt_14}</td>
    <td>${kt_m_vt_14}</td>
    <td>${sl_m_vt_14}</td>
    <td>${m_vt_15}</td>
    <td>${kt_m_vt_15}</td>
    <td>${sl_m_vt_15}</td>
    <td>${m_vt_16}</td>
    <td>${kt_m_vt_16}</td>
    <td>${sl_m_vt_16}</td>
    <td>${m_vt_17}</td>
    <td>${kt_m_vt_17}</td>
    <td>${sl_m_vt_17}</td>
    <td>${m_vt_18}</td>
    <td>${kt_m_vt_18}</td>
    <td>${sl_m_vt_18}</td>
    <td>${m_vt_19}</td>
    <td>${kt_m_vt_19}</td>
    <td>${sl_m_vt_19}</td>
    <td>${m_vt_20}</td>
    <td>${kt_m_vt_20}</td>
    <td>${sl_m_vt_20}</td>
    <td>${m2_vt_1}</td>
    <td>${kt1_m2_vt_1}</td>
    <td>${kt2_m2_vt_1}</td>
    <td>${sl_m2_vt_1}</td>
    <td>${m2_vt_2}</td>
    <td>${kt1_m2_vt_2}</td>
    <td>${kt2_m2_vt_2}</td>
    <td>${sl_m2_vt_2}</td>
    <td>${c_vt_1}</td>
    <td>${sl_c_vt_1}</td>
    <td>${c_vt_2}</td>
    <td>${sl_c_vt_2}</td>
    <td>${c_vt_3}</td>
    <td>${sl_c_vt_3}</td>
    <td>${c_vt_4}</td>
    <td>${sl_c_vt_4}</td>
    <td>${c_vt_5}</td>
    <td>${sl_c_vt_5}</td>
    <td>${c_vt_6}</td>
    <td>${sl_c_vt_6}</td>
    <td>${c_vt_7}</td>
    <td>${sl_c_vt_7}</td>
    <td>${c_vt_8}</td>
    <td>${sl_c_vt_8}</td>
    <td>${c_vt_9}</td>
    <td>${sl_c_vt_9}</td>
    <td>${c_vt_10}</td>
    <td>${sl_c_vt_10}</td>
    <td>${c_vt_11}</td>
    <td>${sl_c_vt_11}</td>
    <td>${c_vt_12}</td>
    <td>${sl_c_vt_12}</td>
    <td>${c_vt_13}</td>
    <td>${sl_c_vt_13}</td>
    <td>${c_vt_14}</td>
    <td>${sl_c_vt_14}</td>
    <td>${c_vt_15}</td>
    <td>${sl_c_vt_15}</td>
    <td>${c_vt_16}</td>
    <td>${sl_c_vt_16}</td>
    <td>${c_vt_17}</td>
    <td>${sl_c_vt_17}</td>
    <td>${c_vt_18}</td>
    <td>${sl_c_vt_18}</td>
    <td>${c_vt_19}</td>
    <td>${sl_c_vt_19}</td>
    <td>${c_vt_20}</td>
    <td>${sl_c_vt_20}</td>
    <td>${c_vt_21}</td>
    <td>${sl_c_vt_21}</td>
    <td>${c_vt_22}</td>
    <td>${sl_c_vt_22}</td>
    <td>${c_vt_23}</td>
    <td>${sl_c_vt_23}</td>
    <td>${c_vt_24}</td>
    <td>${sl_c_vt_24}</td>
    <td>${c_vt_25}</td>
    <td>${sl_c_vt_25}</td>
    <td>${c_vt_26}</td>
    <td>${sl_c_vt_26}</td>
    <td>${c_vt_27}</td>
    <td>${sl_c_vt_27}</td>
    <td>${c_vt_28}</td>
    <td>${sl_c_vt_28}</td>
    <td>${c_vt_29}</td>
    <td>${sl_c_vt_29}</td>
    <td>${c_vt_30}</td>
    <td>${sl_c_vt_30}</td>
    <td>${soLuong}</td>
    `;
        resultsBody.appendChild(row);
    });
}

function exportToExcelLenSanXuat() {
    if (filteredResultsLenSanXuat.length === 0) {
        showMessageLenSanXuat('Không có dữ liệu để xuất. Vui lòng thực hiện lọc trước.', 'error');
        return;
    }

    let csvContent = "\uFEFF";
    // Header row với tất cả các cột
    csvContent += "ID,Hiển thị trên sổ,Loại xuất kho,Ngày hạch toán (*),Ngày chứng từ (*),Số chứng từ (*),Mẫu số HĐ,Ký hiệu HĐ,Mã đối tượng,Tên đối tượng,Địa chỉ/Bộ phận,Tên người nhận/Của,Lý do xuất/Về việc,Nhân viên bán hàng,Kèm theo,Số lệnh điều động,Ngày lệnh điều động,Người vận chuyển,Tên người vận chuyển,Hợp đồng số,Phương tiện vận chuyển,Xuất tại kho,Địa chỉ kho xuất,Nhập tại chi nhánh,Tên chi nhánh,MST chi nhánh,Nhập tại kho,Địa chỉ kho nhập,Mã hàng (*),Tên hàng,Là hàng khuyến mại,Kho (*),Hàng hóa giữ hộ/bán hộ,TK Nợ (*),TK Có (*),ĐVT,Số lượng,Đơn giá bán,Thành tiền,Đơn giá vốn,Tiền vốn,Số lô,Hạn sử dụng,Đối tượng,Khoản mục CP,Đơn vị,Đối tượng THCP,Công trình,Đơn đặt hàng,Hợp đồng bán,CP không hợp lý,Mã thống kê,m_vt_1,kt_m_vt_1,sl_m_vt_1,m_vt_2,kt_m_vt_2,sl_m_vt_2,m_vt_3,kt_m_vt_3,sl_m_vt_3,m_vt_4,kt_m_vt_4,sl_m_vt_4,m_vt_5,kt_m_vt_5,sl_m_vt_5,m_vt_6,kt_m_vt_6,sl_m_vt_6,m_vt_7,kt_m_vt_7,sl_m_vt_7,m_vt_8,kt_m_vt_8,sl_m_vt_8,m_vt_9,kt_m_vt_9,sl_m_vt_9,m_vt_10,kt_m_vt_10,sl_m_vt_10,m_vt_11,kt_m_vt_11,sl_m_vt_11,m_vt_12,kt_m_vt_12,sl_m_vt_12,m_vt_13,kt_m_vt_13,sl_m_vt_13,m_vt_14,kt_m_vt_14,sl_m_vt_14,m_vt_15,kt_m_vt_15,sl_m_vt_15,m_vt_16,kt_m_vt_16,sl_m_vt_16,m_vt_17,kt_m_vt_17,sl_m_vt_17,m_vt_18,kt_m_vt_18,sl_m_vt_18,m_vt_19,kt_m_vt_19,sl_m_vt_19,m_vt_20,kt_m_vt_20,sl_m_vt_20,m2_vt_1,kt1_m2_vt_1,kt2_m2_vt_1,sl_m2_vt_1,m2_vt_2,kt1_m2_vt_2,kt2_m2_vt_2,sl_m2_vt_2,c_vt_1,sl_c_vt_1,c_vt_2,sl_c_vt_2,c_vt_3,sl_c_vt_3,c_vt_4,sl_c_vt_4,c_vt_5,sl_c_vt_5,c_vt_6,sl_c_vt_6,c_vt_7,sl_c_vt_7,c_vt_8,sl_c_vt_8,c_vt_9,sl_c_vt_9,c_vt_10,sl_c_vt_10,c_vt_11,sl_c_vt_11,c_vt_12,sl_c_vt_12,c_vt_13,sl_c_vt_13,c_vt_14,sl_c_vt_14,c_vt_15,sl_c_vt_15,c_vt_16,sl_c_vt_16,c_vt_17,sl_c_vt_17,c_vt_18,sl_c_vt_18,c_vt_19,sl_c_vt_19,c_vt_20,sl_c_vt_20,c_vt_21,sl_c_vt_21,c_vt_22,sl_c_vt_22,c_vt_23,sl_c_vt_23,c_vt_24,sl_c_vt_24,c_vt_25,sl_c_vt_25,c_vt_26,sl_c_vt_26,c_vt_27,sl_c_vt_27,c_vt_28,sl_c_vt_28,c_vt_29,sl_c_vt_29,c_vt_30,sl_c_vt_30,Số bộ\n";

    filteredResultsLenSanXuat.forEach(result => {
        const chiTietRow = result.chiTietRow;
        const donHangRow = result.donHangRow;
        const chiTietColumnIndex = result.chiTietColumnIndex;
        const donHangColumnIndex = result.donHangColumnIndex;

        const maDonHangID = chiTietRow[chiTietColumnIndex['ma_don_hang_id']] || '';
        const maSanPhamTheoDoi = chiTietRow[chiTietColumnIndex['ma_san_pham_theo_doi']] || '';
        const dienGiai = chiTietRow[chiTietColumnIndex['dien_giai']] || '';
        const ghiChu = chiTietRow[chiTietColumnIndex['ghi_chu']] || '';
        const dvt = chiTietRow[chiTietColumnIndex['dvt']] || '';
        const soLuong = chiTietRow[chiTietColumnIndex['so_luong']] || '';
        const khoiLuong = formatNumberForCSV(chiTietRow[chiTietColumnIndex['khoi_luong']] || '');
        const sttTrongDon = chiTietRow[chiTietColumnIndex['stt_trong_don']] || '';

        const loaiDonHang = donHangRow[donHangColumnIndex['loai_don_hang']] || '';
        const ngayGiaoHang = donHangRow[donHangColumnIndex['ngay_giao_hang']] || '';
        const ngayNhapKho = donHangRow[donHangColumnIndex['ngay_nhap_kho']] || '';
        const xuongSanXuat = donHangRow[donHangColumnIndex['xuong_san_xuat']] || '';
        const maHopDong = donHangRow[donHangColumnIndex['ma_hop_dong']] || '';

        let ngayHaChToan = '';
        let ngayChungTu = '';

        if (loaiDonHang === "Yêu cầu kiểm tra") {
            ngayHaChToan = ngayGiaoHang;
            ngayChungTu = ngayGiaoHang;
        } else {
            ngayHaChToan = ngayNhapKho;
            ngayChungTu = ngayNhapKho;
        }

        let tenHang = '';
        if (dienGiai !== "") {
            tenHang = dienGiai;
        } else {
            tenHang = ghiChu;
        }

        const mnvCongTy = calculateMnvCongTy(donHangRow, donHangColumnIndex);

        // Xác định kho dựa trên xưởng sản xuất
        let kho = '';
        if (xuongSanXuat === "Hà Nội") {
            kho = "K03_SX.HN_152";
        } else if (xuongSanXuat === "Long An") {
            kho = "K04_SX.LA_152";
        }

        // Xác định đối tượng THCP dựa trên xưởng sản xuất
        let doiTuongTHCP = '';
        if (xuongSanXuat === "Hà Nội") {
            doiTuongTHCP = "25.SXT.X1";
        } else if (xuongSanXuat === "Long An") {
            doiTuongTHCP = "25.SXT.X2";
        }

        // Lấy tất cả các trường bổ sung
        const m_vt_1 = chiTietRow[chiTietColumnIndex['m_vt_1']] || '';
        const kt_m_vt_1 = chiTietRow[chiTietColumnIndex['kt_m_vt_1']] || '';
        const sl_m_vt_1 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_m_vt_1']] || '');
        const m_vt_2 = chiTietRow[chiTietColumnIndex['m_vt_2']] || '';
        const kt_m_vt_2 = chiTietRow[chiTietColumnIndex['kt_m_vt_2']] || '';
        const sl_m_vt_2 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_m_vt_2']] || '');
        const m_vt_3 = chiTietRow[chiTietColumnIndex['m_vt_3']] || '';
        const kt_m_vt_3 = chiTietRow[chiTietColumnIndex['kt_m_vt_3']] || '';
        const sl_m_vt_3 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_m_vt_3']] || '');
        const m_vt_4 = chiTietRow[chiTietColumnIndex['m_vt_4']] || '';
        const kt_m_vt_4 = chiTietRow[chiTietColumnIndex['kt_m_vt_4']] || '';
        const sl_m_vt_4 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_m_vt_4']] || '');
        const m_vt_5 = chiTietRow[chiTietColumnIndex['m_vt_5']] || '';
        const kt_m_vt_5 = chiTietRow[chiTietColumnIndex['kt_m_vt_5']] || '';
        const sl_m_vt_5 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_m_vt_5']] || '');
        const m_vt_6 = chiTietRow[chiTietColumnIndex['m_vt_6']] || '';
        const kt_m_vt_6 = chiTietRow[chiTietColumnIndex['kt_m_vt_6']] || '';
        const sl_m_vt_6 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_m_vt_6']] || '');
        const m_vt_7 = chiTietRow[chiTietColumnIndex['m_vt_7']] || '';
        const kt_m_vt_7 = chiTietRow[chiTietColumnIndex['kt_m_vt_7']] || '';
        const sl_m_vt_7 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_m_vt_7']] || '');
        const m_vt_8 = chiTietRow[chiTietColumnIndex['m_vt_8']] || '';
        const kt_m_vt_8 = chiTietRow[chiTietColumnIndex['kt_m_vt_8']] || '';
        const sl_m_vt_8 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_m_vt_8']] || '');
        const m_vt_9 = chiTietRow[chiTietColumnIndex['m_vt_9']] || '';
        const kt_m_vt_9 = chiTietRow[chiTietColumnIndex['kt_m_vt_9']] || '';
        const sl_m_vt_9 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_m_vt_9']] || '');
        const m_vt_10 = chiTietRow[chiTietColumnIndex['m_vt_10']] || '';
        const kt_m_vt_10 = chiTietRow[chiTietColumnIndex['kt_m_vt_10']] || '';
        const sl_m_vt_10 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_m_vt_10']] || '');
        const m_vt_11 = chiTietRow[chiTietColumnIndex['m_vt_11']] || '';
        const kt_m_vt_11 = chiTietRow[chiTietColumnIndex['kt_m_vt_11']] || '';
        const sl_m_vt_11 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_m_vt_11']] || '');
        const m_vt_12 = chiTietRow[chiTietColumnIndex['m_vt_12']] || '';
        const kt_m_vt_12 = chiTietRow[chiTietColumnIndex['kt_m_vt_12']] || '';
        const sl_m_vt_12 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_m_vt_12']] || '');
        const m_vt_13 = chiTietRow[chiTietColumnIndex['m_vt_13']] || '';
        const kt_m_vt_13 = chiTietRow[chiTietColumnIndex['kt_m_vt_13']] || '';
        const sl_m_vt_13 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_m_vt_13']] || '');
        const m_vt_14 = chiTietRow[chiTietColumnIndex['m_vt_14']] || '';
        const kt_m_vt_14 = chiTietRow[chiTietColumnIndex['kt_m_vt_14']] || '';
        const sl_m_vt_14 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_m_vt_14']] || '');
        const m_vt_15 = chiTietRow[chiTietColumnIndex['m_vt_15']] || '';
        const kt_m_vt_15 = chiTietRow[chiTietColumnIndex['kt_m_vt_15']] || '';
        const sl_m_vt_15 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_m_vt_15']] || '');
        const m_vt_16 = chiTietRow[chiTietColumnIndex['m_vt_16']] || '';
        const kt_m_vt_16 = chiTietRow[chiTietColumnIndex['kt_m_vt_16']] || '';
        const sl_m_vt_16 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_m_vt_16']] || '');
        const m_vt_17 = chiTietRow[chiTietColumnIndex['m_vt_17']] || '';
        const kt_m_vt_17 = chiTietRow[chiTietColumnIndex['kt_m_vt_17']] || '';
        const sl_m_vt_17 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_m_vt_17']] || '');
        const m_vt_18 = chiTietRow[chiTietColumnIndex['m_vt_18']] || '';
        const kt_m_vt_18 = chiTietRow[chiTietColumnIndex['kt_m_vt_18']] || '';
        const sl_m_vt_18 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_m_vt_18']] || '');
        const m_vt_19 = chiTietRow[chiTietColumnIndex['m_vt_19']] || '';
        const kt_m_vt_19 = chiTietRow[chiTietColumnIndex['kt_m_vt_19']] || '';
        const sl_m_vt_19 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_m_vt_19']] || '');
        const m_vt_20 = chiTietRow[chiTietColumnIndex['m_vt_20']] || '';
        const kt_m_vt_20 = chiTietRow[chiTietColumnIndex['kt_m_vt_20']] || '';
        const sl_m_vt_20 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_m_vt_20']] || '');

        // Lấy các trường m2_vt, kt1_m2_vt, kt2_m2_vt, sl_m2_vt (1 đến 2)
        const m2_vt_1 = chiTietRow[chiTietColumnIndex['m2_vt_1']] || '';
        const kt1_m2_vt_1 = chiTietRow[chiTietColumnIndex['kt1_m2_vt_1']] || '';
        const kt2_m2_vt_1 = chiTietRow[chiTietColumnIndex['kt2_m2_vt_1']] || '';
        const sl_m2_vt_1 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_m2_vt_1']] || '');
        const m2_vt_2 = chiTietRow[chiTietColumnIndex['m2_vt_2']] || '';
        const kt1_m2_vt_2 = chiTietRow[chiTietColumnIndex['kt1_m2_vt_2']] || '';
        const kt2_m2_vt_2 = chiTietRow[chiTietColumnIndex['kt2_m2_vt_2']] || '';
        const sl_m2_vt_2 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_m2_vt_2']] || '');

        // Lấy các trường c_vt, sl_c_vt (1 đến 30)
        const c_vt_1 = chiTietRow[chiTietColumnIndex['c_vt_1']] || '';
        const sl_c_vt_1 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_1']] || '');
        const c_vt_2 = chiTietRow[chiTietColumnIndex['c_vt_2']] || '';
        const sl_c_vt_2 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_2']] || '');
        const c_vt_3 = chiTietRow[chiTietColumnIndex['c_vt_3']] || '';
        const sl_c_vt_3 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_3']] || '');
        const c_vt_4 = chiTietRow[chiTietColumnIndex['c_vt_4']] || '';
        const sl_c_vt_4 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_4']] || '');
        const c_vt_5 = chiTietRow[chiTietColumnIndex['c_vt_5']] || '';
        const sl_c_vt_5 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_5']] || '');
        const c_vt_6 = chiTietRow[chiTietColumnIndex['c_vt_6']] || '';
        const sl_c_vt_6 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_6']] || '');
        const c_vt_7 = chiTietRow[chiTietColumnIndex['c_vt_7']] || '';
        const sl_c_vt_7 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_7']] || '');
        const c_vt_8 = chiTietRow[chiTietColumnIndex['c_vt_8']] || '';
        const sl_c_vt_8 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_8']] || '');
        const c_vt_9 = chiTietRow[chiTietColumnIndex['c_vt_9']] || '';
        const sl_c_vt_9 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_9']] || '');
        const c_vt_10 = chiTietRow[chiTietColumnIndex['c_vt_10']] || '';
        const sl_c_vt_10 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_10']] || '');
        const c_vt_11 = chiTietRow[chiTietColumnIndex['c_vt_11']] || '';
        const sl_c_vt_11 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_11']] || '');
        const c_vt_12 = chiTietRow[chiTietColumnIndex['c_vt_12']] || '';
        const sl_c_vt_12 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_12']] || '');
        const c_vt_13 = chiTietRow[chiTietColumnIndex['c_vt_13']] || '';
        const sl_c_vt_13 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_13']] || '');
        const c_vt_14 = chiTietRow[chiTietColumnIndex['c_vt_14']] || '';
        const sl_c_vt_14 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_14']] || '');
        const c_vt_15 = chiTietRow[chiTietColumnIndex['c_vt_15']] || '';
        const sl_c_vt_15 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_15']] || '');
        const c_vt_16 = chiTietRow[chiTietColumnIndex['c_vt_16']] || '';
        const sl_c_vt_16 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_16']] || '');
        const c_vt_17 = chiTietRow[chiTietColumnIndex['c_vt_17']] || '';
        const sl_c_vt_17 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_17']] || '');
        const c_vt_18 = chiTietRow[chiTietColumnIndex['c_vt_18']] || '';
        const sl_c_vt_18 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_18']] || '');
        const c_vt_19 = chiTietRow[chiTietColumnIndex['c_vt_19']] || '';
        const sl_c_vt_19 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_19']] || '');
        const c_vt_20 = chiTietRow[chiTietColumnIndex['c_vt_20']] || '';
        const sl_c_vt_20 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_20']] || '');
        const c_vt_21 = chiTietRow[chiTietColumnIndex['c_vt_21']] || '';
        const sl_c_vt_21 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_21']] || '');
        const c_vt_22 = chiTietRow[chiTietColumnIndex['c_vt_22']] || '';
        const sl_c_vt_22 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_22']] || '');
        const c_vt_23 = chiTietRow[chiTietColumnIndex['c_vt_23']] || '';
        const sl_c_vt_23 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_23']] || '');
        const c_vt_24 = chiTietRow[chiTietColumnIndex['c_vt_24']] || '';
        const sl_c_vt_24 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_24']] || '');
        const c_vt_25 = chiTietRow[chiTietColumnIndex['c_vt_25']] || '';
        const sl_c_vt_25 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_25']] || '');
        const c_vt_26 = chiTietRow[chiTietColumnIndex['c_vt_26']] || '';
        const sl_c_vt_26 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_26']] || '');
        const c_vt_27 = chiTietRow[chiTietColumnIndex['c_vt_27']] || '';
        const sl_c_vt_27 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_27']] || '');
        const c_vt_28 = chiTietRow[chiTietColumnIndex['c_vt_28']] || '';
        const sl_c_vt_28 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_28']] || '');
        const c_vt_29 = chiTietRow[chiTietColumnIndex['c_vt_29']] || '';
        const sl_c_vt_29 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_29']] || '');
        const c_vt_30 = chiTietRow[chiTietColumnIndex['c_vt_30']] || '';
        const sl_c_vt_30 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_30']] || '');

        const escapeCSV = (str) => {
            if (!str) return '';
            if (str.includes(',') || str.includes('"') || str.includes('\n')) {
                return '"' + str.replace(/"/g, '""') + '"';
            }
            return str;
        };

        // Tạo dòng CSV với tất cả các cột
        csvContent += `${escapeCSV(maDonHangID)},${escapeCSV(sttTrongDon)},1,${escapeCSV(ngayHaChToan)},${escapeCSV(ngayChungTu)},,,,,,,,Xuất kho NVL sản xuất cho Lệnh sản xuất <${escapeCSV(maHopDong)}>,${escapeCSV(mnvCongTy)},,,,,,,,,,,,,,,,,,${escapeCSV(kho)},,154,1521,,,,,,,,,,,${escapeCSV(doiTuongTHCP)},${escapeCSV(maSanPhamTheoDoi)},,${escapeCSV(maHopDong)},,,,${escapeCSV(m_vt_1)},${escapeCSV(kt_m_vt_1)},${sl_m_vt_1},${escapeCSV(m_vt_2)},${escapeCSV(kt_m_vt_2)},${sl_m_vt_2},${escapeCSV(m_vt_3)},${escapeCSV(kt_m_vt_3)},${sl_m_vt_3},${escapeCSV(m_vt_4)},${escapeCSV(kt_m_vt_4)},${sl_m_vt_4},${escapeCSV(m_vt_5)},${escapeCSV(kt_m_vt_5)},${sl_m_vt_5},${escapeCSV(m_vt_6)},${escapeCSV(kt_m_vt_6)},${sl_m_vt_6},${escapeCSV(m_vt_7)},${escapeCSV(kt_m_vt_7)},${sl_m_vt_7},${escapeCSV(m_vt_8)},${escapeCSV(kt_m_vt_8)},${sl_m_vt_8},${escapeCSV(m_vt_9)},${escapeCSV(kt_m_vt_9)},${sl_m_vt_9},${escapeCSV(m_vt_10)},${escapeCSV(kt_m_vt_10)},${sl_m_vt_10},${escapeCSV(m_vt_11)},${escapeCSV(kt_m_vt_11)},${sl_m_vt_11},${escapeCSV(m_vt_12)},${escapeCSV(kt_m_vt_12)},${sl_m_vt_12},${escapeCSV(m_vt_13)},${escapeCSV(kt_m_vt_13)},${sl_m_vt_13},${escapeCSV(m_vt_14)},${escapeCSV(kt_m_vt_14)},${sl_m_vt_14},${escapeCSV(m_vt_15)},${escapeCSV(kt_m_vt_15)},${sl_m_vt_15},${escapeCSV(m_vt_16)},${escapeCSV(kt_m_vt_16)},${sl_m_vt_16},${escapeCSV(m_vt_17)},${escapeCSV(kt_m_vt_17)},${sl_m_vt_17},${escapeCSV(m_vt_18)},${escapeCSV(kt_m_vt_18)},${sl_m_vt_18},${escapeCSV(m_vt_19)},${escapeCSV(kt_m_vt_19)},${sl_m_vt_19},${escapeCSV(m_vt_20)},${escapeCSV(kt_m_vt_20)},${sl_m_vt_20},${escapeCSV(m2_vt_1)},${escapeCSV(kt1_m2_vt_1)},${escapeCSV(kt2_m2_vt_1)},${sl_m2_vt_1},${escapeCSV(m2_vt_2)},${escapeCSV(kt1_m2_vt_2)},${escapeCSV(kt2_m2_vt_2)},${sl_m2_vt_2},${escapeCSV(c_vt_1)},${sl_c_vt_1},${escapeCSV(c_vt_2)},${sl_c_vt_2},${escapeCSV(c_vt_3)},${sl_c_vt_3},${escapeCSV(c_vt_4)},${sl_c_vt_4},${escapeCSV(c_vt_5)},${sl_c_vt_5},${escapeCSV(c_vt_6)},${sl_c_vt_6},${escapeCSV(c_vt_7)},${sl_c_vt_7},${escapeCSV(c_vt_8)},${sl_c_vt_8},${escapeCSV(c_vt_9)},${sl_c_vt_9},${escapeCSV(c_vt_10)},${sl_c_vt_10},${escapeCSV(c_vt_11)},${sl_c_vt_11},${escapeCSV(c_vt_12)},${sl_c_vt_12},${escapeCSV(c_vt_13)},${sl_c_vt_13},${escapeCSV(c_vt_14)},${sl_c_vt_14},${escapeCSV(c_vt_15)},${sl_c_vt_15},${escapeCSV(c_vt_16)},${sl_c_vt_16},${escapeCSV(c_vt_17)},${sl_c_vt_17},${escapeCSV(c_vt_18)},${sl_c_vt_18},${escapeCSV(c_vt_19)},${sl_c_vt_19},${escapeCSV(c_vt_20)},${sl_c_vt_20},${escapeCSV(c_vt_21)},${sl_c_vt_21},${escapeCSV(c_vt_22)},${sl_c_vt_22},${escapeCSV(c_vt_23)},${sl_c_vt_23},${escapeCSV(c_vt_24)},${sl_c_vt_24},${escapeCSV(c_vt_25)},${sl_c_vt_25},${escapeCSV(c_vt_26)},${sl_c_vt_26},${escapeCSV(c_vt_27)},${sl_c_vt_27},${escapeCSV(c_vt_28)},${sl_c_vt_28},${escapeCSV(c_vt_29)},${sl_c_vt_29},${escapeCSV(c_vt_30)},${sl_c_vt_30},${soLuong}\n`;
    });

    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement("a");
    const url = URL.createObjectURL(blob);

    link.setAttribute("href", url);
    const tuNgayRaw = document.getElementById("tu-ngay-lsx").value;
    const denNgayRaw = document.getElementById("den-ngay-lsx").value;
    const tuNgay = formatDateForFilename(tuNgayRaw);
    const denNgay = formatDateForFilename(denNgayRaw);
    const loaiDonHang = document.querySelector('input[name="loai-don-hang-lsx"]:checked').value;
    const fileName = `Danh sách lệnh sản xuất - ${tuNgay} - ${denNgay} - ${loaiDonHang}.csv`;

    link.setAttribute("download", fileName);
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);

    showMessageLenSanXuat(`Đã xuất ${filteredResultsLenSanXuat.length} dòng ra file Excel.`, 'success');
}

function showLoadingLenSanXuat(show) {
    const loadingElement = document.getElementById('loading-lsx');
    loadingElement.style.display = show ? 'block' : 'none';
}

function showMessageLenSanXuat(message, type) {
    const resultsCount = document.getElementById("results-count-lsx");
    resultsCount.className = "results-count";

    if (type === "success") {
        resultsCount.style.backgroundColor = "#e8f7e8";
        resultsCount.style.color = "#0c9c07";
        resultsCount.style.borderLeft = "4px solid #0c9c07";
    }

    if (type === "error") {
        resultsCount.style.backgroundColor = "#ffeaea";
        resultsCount.style.color = "#c00";
        resultsCount.style.borderLeft = "4px solid #c00";
    }

    resultsCount.style.padding = "5px";
    resultsCount.style.borderRadius = "6px";
    resultsCount.style.fontSize = "16px";
    resultsCount.style.fontWeight = "600";
    resultsCount.textContent = message;
}

function requireRefilterLenSanXuat() {
    document.getElementById('results-table-lsx').style.display = 'none';
    document.getElementById('no-results-lsx').style.display = 'block';
    document.getElementById('results-count-lsx').textContent = 'Kết quả: Chưa có dòng nào được lọc.';
    filteredResultsLenSanXuat = [];
    showMessageLenSanXuat("Bạn đã thay đổi bộ lọc. Vui lòng nhấn 'Lọc' để cập nhật kết quả mới.", "error");
}

// ==================== TAB LỆNH XUẤT VẬT TƯ BẢO HÀNH ====================

let filteredResultsXuatBaoHanh = [];

function initXuatBaoHanhEventListeners() {
    document.getElementById('btn-loc-xbh').addEventListener('click', applyFilterXuatBaoHanh);
    document.getElementById('btn-reset-xbh').addEventListener('click', resetFilterXuatBaoHanh);
    document.getElementById('btn-export-xbh').addEventListener('click', exportToExcelXuatBaoHanh);

    // Xóa event listener cho input text cũ
    // document.getElementById("ma-hop-dong-xbh").addEventListener("input", requireRefilterXuatBaoHanh);
    document.getElementById("xuong-san-xuat-xbh").addEventListener("change", requireRefilterXuatBaoHanh);
    document.getElementById("tu-ngay-xbh").addEventListener("change", requireRefilterXuatBaoHanh);
    document.getElementById("den-ngay-xbh").addEventListener("change", requireRefilterXuatBaoHanh);

    document.querySelectorAll("input[name='loai-don-hang-xbh']").forEach(radio => {
        radio.addEventListener("change", requireRefilterXuatBaoHanh);
    });
}
function resetFilterXuatBaoHanh() {
    // Reset multi-select đơn vị phụ trách
    const donViAllCheckbox = document.getElementById('don-vi-phu-trach-xbh-all');
    if (donViAllCheckbox) {
        donViAllCheckbox.checked = true;
    }

    const donViCheckboxes = document.querySelectorAll('#don-vi-phu-trach-xbh-container input[type="checkbox"]:not(#don-vi-phu-trach-xbh-all)');
    donViCheckboxes.forEach(checkbox => {
        checkbox.checked = false;
    });

    // Reset multi-select mã hợp đồng
    const maHopDongAllCheckbox = document.getElementById('ma-hop-dong-xbh-all');
    if (maHopDongAllCheckbox) {
        maHopDongAllCheckbox.checked = true;
    }

    const maHopDongCheckboxes = document.querySelectorAll('#ma-hop-dong-xbh-container input[type="checkbox"]:not(#ma-hop-dong-xbh-all)');
    maHopDongCheckboxes.forEach(checkbox => {
        checkbox.checked = false;
    });

    document.getElementById('xuong-san-xuat-xbh').value = '';
    document.getElementById('tu-ngay-xbh').value = '';
    document.getElementById('den-ngay-xbh').value = '';
    document.getElementById('loai-tat-ca-xbh').checked = true;

    document.getElementById('results-table-xbh').style.display = 'none';
    document.getElementById('no-results-xbh').style.display = 'block';
    document.getElementById('results-count-xbh').textContent = 'Kết quả: Chưa có dòng nào được lọc.';

    filteredResultsXuatBaoHanh = [];
}

async function applyFilterXuatBaoHanh() {
    showLoadingXuatBaoHanh(true);

    try {
        if (!isDataLoaded) {
            await fetchDataFromSheets();
        }

        const filterOptions = {
            maHopDong: getSelectedMaHopDong('-xbh'), // Lấy danh sách mã hợp đồng được chọn
            donViPhuTrach: getSelectedDonViPhuTrach('-xbh'), // Lấy danh sách đơn vị được chọn
            xuongSanXuat: document.getElementById('xuong-san-xuat-xbh').value,
            tuNgay: document.getElementById('tu-ngay-xbh').value,
            denNgay: document.getElementById('den-ngay-xbh').value,
            loaiDonHang: document.querySelector('input[name="loai-don-hang-xbh"]:checked').value
        };

        filteredResultsXuatBaoHanh = filterDataXuatBaoHanh(filterOptions);
        displayResultsXuatBaoHanh(filteredResultsXuatBaoHanh);
        showMessageXuatBaoHanh(`Đã tìm thấy ${filteredResultsXuatBaoHanh.length} dòng phù hợp.`, 'success');

    } catch (error) {
        console.error('Lỗi khi áp dụng bộ lọc:', error);
        showMessageXuatBaoHanh('Đã xảy ra lỗi khi tải dữ liệu. Vui lòng thử lại.', 'error');
    } finally {
        showLoadingXuatBaoHanh(false);
    }
}

function filterDataXuatBaoHanh(filterOptions) {
    if (donHangData.length === 0 || donHangChiTietData.length === 0) {
        return [];
    }

    const donHangHeaders = donHangData[0];
    const chiTietHeaders = donHangChiTietData[0];

    const donHangColumnIndex = {};
    donHangHeaders.forEach((header, index) => {
        donHangColumnIndex[header] = index;
    });

    const chiTietColumnIndex = {};
    chiTietHeaders.forEach((header, index) => {
        chiTietColumnIndex[header] = index;
    });

    const filteredResults = [];
    let resultId = 1;

    for (let i = 1; i < donHangChiTietData.length; i++) {
        const chiTietRow = donHangChiTietData[i];
        const maDonHangID = chiTietRow[chiTietColumnIndex['ma_don_hang_id']] || '';
        const donHangRow = donHangData.find(row => {
            return row[donHangColumnIndex['ma_don_hang']] === maDonHangID;
        });

        if (!donHangRow) continue;

        if (!passesFilterConditionsXuatBaoHanh(chiTietRow, donHangRow, chiTietColumnIndex, donHangColumnIndex, filterOptions)) {
            continue;
        }

        filteredResults.push({
            id: resultId++,
            chiTietRow: chiTietRow,
            donHangRow: donHangRow,
            chiTietColumnIndex: chiTietColumnIndex,
            donHangColumnIndex: donHangColumnIndex
        });
    }

    return filteredResults;
}

function passesFilterConditionsXuatBaoHanh(chiTietRow, donHangRow, chiTietColumnIndex, donHangColumnIndex, filterOptions) {
    // Tab xuất bảo hành thường chỉ lọc các dòng thuộc nhóm "Bảo hành"
    const nhomSanPham = chiTietRow[chiTietColumnIndex['nhom_san_pham']] || '';
    const mauCua = chiTietRow[chiTietColumnIndex['mau_cua']] || '';
    const tenSanPham = chiTietRow[chiTietColumnIndex['ten_san_pham']] || '';
    const trongLuongPhuKien = donHangRow[donHangColumnIndex['trong_luong_phu_kien']] || '';
    const maHopDong = donHangRow[donHangColumnIndex['ma_hop_dong']] || '';
    const donViPhuTrach = donHangRow[donHangColumnIndex['don_vi_phu_trach']] || '';
    const loaiDonHang = donHangRow[donHangColumnIndex['loai_don_hang']] || '';
    const ngayGiaoHang = donHangRow[donHangColumnIndex['ngay_giao_hang']] || '';
    const ngayNhapKho = donHangRow[donHangColumnIndex['ngay_nhap_kho']] || '';
    const xuongSanXuat = donHangRow[donHangColumnIndex['xuong_san_xuat']] || '';

    // Chỉ lấy các dòng thuộc nhóm "Bảo hành"
    if (nhomSanPham !== "Bảo hành") {
        return false;
    }

    if (mauCua === "Nhân công") {
        return false;
    }

    if (tenSanPham === "Di chuyển" || tenSanPham === "Nhân công") {
        return false;
    }

    // Lọc theo mã hợp đồng
    if (filterOptions.maHopDong && !maHopDong.includes(filterOptions.maHopDong)) {
        return false;
    }

    // Lọc theo đơn vị phụ trách (multi-select)
    if (filterOptions.donViPhuTrach && filterOptions.donViPhuTrach.length > 0) {
        if (!filterOptions.donViPhuTrach.includes(donViPhuTrach)) {
            return false;
        }
    }

    // Lọc theo ngày
    if (filterOptions.tuNgay && filterOptions.denNgay) {
        const tuNgay = new Date(filterOptions.tuNgay);
        const denNgay = new Date(filterOptions.denNgay);

        tuNgay.setHours(0, 0, 0, 0);
        denNgay.setHours(23, 59, 59, 999);

        let ngaySoSanhStr;
        if (loaiDonHang === "Yêu cầu kiểm tra") {
            ngaySoSanhStr = ngayGiaoHang;
        } else {
            ngaySoSanhStr = ngayNhapKho;
        }

        const ngaySoSanh = parseDDMMYYYY(ngaySoSanhStr);
        if (!ngaySoSanh) {
            return false;
        }

        if (ngaySoSanh < tuNgay || ngaySoSanh > denNgay) {
            return false;
        }
    }

    // Lọc theo loại đơn hàng
    if (filterOptions.loaiDonHang === "Tiêu chuẩn") {
        if (tenSanPham === "Vật tư khác" || trongLuongPhuKien !== "Tiêu chuẩn") {
            return false;
        }
    } else if (filterOptions.loaiDonHang === "Khác chuẩn") {
        if (trongLuongPhuKien !== "Khác chuẩn") {
            return false;
        }
    }

    // Lọc theo xưởng sản xuất
    if (filterOptions.xuongSanXuat && xuongSanXuat !== filterOptions.xuongSanXuat) {
        return false;
    }

    return true;
}

function displayResultsXuatBaoHanh(results) {
    const resultsBody = document.getElementById('results-body-xbh');
    const resultsTable = document.getElementById('results-table-xbh');
    const noResults = document.getElementById('no-results-xbh');
    const resultsCount = document.getElementById('results-count-xbh');

    resultsCount.textContent = `Kết quả: ${results.length} dòng`;
    resultsBody.innerHTML = '';

    if (results.length === 0) {
        resultsTable.style.display = 'none';
        noResults.style.display = 'block';
        return;
    }

    noResults.style.display = 'none';
    resultsTable.style.display = 'table';

    results.forEach(result => {
        const chiTietRow = result.chiTietRow;
        const donHangRow = result.donHangRow;
        const chiTietColumnIndex = result.chiTietColumnIndex;
        const donHangColumnIndex = result.donHangColumnIndex;

        const maDonHangID = chiTietRow[chiTietColumnIndex['ma_don_hang_id']] || '';
        const maSanPhamTheoDoi = chiTietRow[chiTietColumnIndex['ma_san_pham_theo_doi']] || '';
        const dienGiai = chiTietRow[chiTietColumnIndex['dien_giai']] || '';
        const ghiChu = chiTietRow[chiTietColumnIndex['ghi_chu']] || '';
        const dvt = chiTietRow[chiTietColumnIndex['dvt']] || '';
        const soLuong = chiTietRow[chiTietColumnIndex['so_luong']] || '';
        const khoiLuong = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['khoi_luong']] || '');
        const sttTrongDon = chiTietRow[chiTietColumnIndex['stt_trong_don']] || '';

        const ngayXuatKho = donHangRow[donHangColumnIndex['ngay_xuat_kho']] || '';
        const xuongSanXuat = donHangRow[donHangColumnIndex['xuong_san_xuat']] || '';
        const maHopDong = donHangRow[donHangColumnIndex['ma_hop_dong']] || '';
        const maKhachHangID = donHangRow[donHangColumnIndex['ma_khach_hang_id']] || '';
        const tenNguoiLienHe = donHangRow[donHangColumnIndex['ten_nguoi_lien_he']] || '';
        const diaChiChiTiet = donHangRow[donHangColumnIndex['dia_chi_chi_tiet']] || '';

        let tenHang = '';
        if (dienGiai !== "") {
            tenHang = dienGiai;
        } else {
            tenHang = ghiChu;
        }

        const mnvCongTy = calculateMnvCongTy(donHangRow, donHangColumnIndex);

        // Xác định kho dựa trên xưởng sản xuất
        let kho = '';
        if (xuongSanXuat === "Hà Nội") {
            kho = "K03_SX.HN_152";
        } else if (xuongSanXuat === "Long An") {
            kho = "K04_SX.LA_152";
        }

        // Xác định đối tượng THCP dựa trên xưởng sản xuất
        let doiTuongTHCP = '';
        if (xuongSanXuat === "Hà Nội") {
            doiTuongTHCP = "25.SXT.X1";
        } else if (xuongSanXuat === "Long An") {
            doiTuongTHCP = "25.SXT.X2";
        }

        // Lấy các trường m_vt, kt_m_vt, sl_m_vt (từ 1 đến 20) - tương tự như tab lệnh sản xuất
        // Để ngắn gọn, tôi sẽ không liệt kê hết 20 trường ở đây. Bạn có thể copy từ hàm displayResultsLenSanXuat.
        // Tuy nhiên, trong ví dụ này, tôi sẽ chỉ lấy một vài trường đầu tiên để minh họa.
        const m_vt_1 = chiTietRow[chiTietColumnIndex['m_vt_1']] || '';
        const kt_m_vt_1 = chiTietRow[chiTietColumnIndex['kt_m_vt_1']] || '';
        const sl_m_vt_1 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_m_vt_1']] || '');
        const m_vt_2 = chiTietRow[chiTietColumnIndex['m_vt_2']] || '';
        const kt_m_vt_2 = chiTietRow[chiTietColumnIndex['kt_m_vt_2']] || '';
        const sl_m_vt_2 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_m_vt_2']] || '');
        const m_vt_3 = chiTietRow[chiTietColumnIndex['m_vt_3']] || '';
        const kt_m_vt_3 = chiTietRow[chiTietColumnIndex['kt_m_vt_3']] || '';
        const sl_m_vt_3 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_m_vt_3']] || '');
        const m_vt_4 = chiTietRow[chiTietColumnIndex['m_vt_4']] || '';
        const kt_m_vt_4 = chiTietRow[chiTietColumnIndex['kt_m_vt_4']] || '';
        const sl_m_vt_4 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_m_vt_4']] || '');
        const m_vt_5 = chiTietRow[chiTietColumnIndex['m_vt_5']] || '';
        const kt_m_vt_5 = chiTietRow[chiTietColumnIndex['kt_m_vt_5']] || '';
        const sl_m_vt_5 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_m_vt_5']] || '');
        const m_vt_6 = chiTietRow[chiTietColumnIndex['m_vt_6']] || '';
        const kt_m_vt_6 = chiTietRow[chiTietColumnIndex['kt_m_vt_6']] || '';
        const sl_m_vt_6 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_m_vt_6']] || '');
        const m_vt_7 = chiTietRow[chiTietColumnIndex['m_vt_7']] || '';
        const kt_m_vt_7 = chiTietRow[chiTietColumnIndex['kt_m_vt_7']] || '';
        const sl_m_vt_7 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_m_vt_7']] || '');
        const m_vt_8 = chiTietRow[chiTietColumnIndex['m_vt_8']] || '';
        const kt_m_vt_8 = chiTietRow[chiTietColumnIndex['kt_m_vt_8']] || '';
        const sl_m_vt_8 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_m_vt_8']] || '');
        const m_vt_9 = chiTietRow[chiTietColumnIndex['m_vt_9']] || '';
        const kt_m_vt_9 = chiTietRow[chiTietColumnIndex['kt_m_vt_9']] || '';
        const sl_m_vt_9 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_m_vt_9']] || '');
        const m_vt_10 = chiTietRow[chiTietColumnIndex['m_vt_10']] || '';
        const kt_m_vt_10 = chiTietRow[chiTietColumnIndex['kt_m_vt_10']] || '';
        const sl_m_vt_10 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_m_vt_10']] || '');
        const m_vt_11 = chiTietRow[chiTietColumnIndex['m_vt_11']] || '';
        const kt_m_vt_11 = chiTietRow[chiTietColumnIndex['kt_m_vt_11']] || '';
        const sl_m_vt_11 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_m_vt_11']] || '');
        const m_vt_12 = chiTietRow[chiTietColumnIndex['m_vt_12']] || '';
        const kt_m_vt_12 = chiTietRow[chiTietColumnIndex['kt_m_vt_12']] || '';
        const sl_m_vt_12 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_m_vt_12']] || '');
        const m_vt_13 = chiTietRow[chiTietColumnIndex['m_vt_13']] || '';
        const kt_m_vt_13 = chiTietRow[chiTietColumnIndex['kt_m_vt_13']] || '';
        const sl_m_vt_13 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_m_vt_13']] || '');
        const m_vt_14 = chiTietRow[chiTietColumnIndex['m_vt_14']] || '';
        const kt_m_vt_14 = chiTietRow[chiTietColumnIndex['kt_m_vt_14']] || '';
        const sl_m_vt_14 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_m_vt_14']] || '');
        const m_vt_15 = chiTietRow[chiTietColumnIndex['m_vt_15']] || '';
        const kt_m_vt_15 = chiTietRow[chiTietColumnIndex['kt_m_vt_15']] || '';
        const sl_m_vt_15 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_m_vt_15']] || '');
        const m_vt_16 = chiTietRow[chiTietColumnIndex['m_vt_16']] || '';
        const kt_m_vt_16 = chiTietRow[chiTietColumnIndex['kt_m_vt_16']] || '';
        const sl_m_vt_16 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_m_vt_16']] || '');
        const m_vt_17 = chiTietRow[chiTietColumnIndex['m_vt_17']] || '';
        const kt_m_vt_17 = chiTietRow[chiTietColumnIndex['kt_m_vt_17']] || '';
        const sl_m_vt_17 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_m_vt_17']] || '');
        const m_vt_18 = chiTietRow[chiTietColumnIndex['m_vt_18']] || '';
        const kt_m_vt_18 = chiTietRow[chiTietColumnIndex['kt_m_vt_18']] || '';
        const sl_m_vt_18 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_m_vt_18']] || '');
        const m_vt_19 = chiTietRow[chiTietColumnIndex['m_vt_19']] || '';
        const kt_m_vt_19 = chiTietRow[chiTietColumnIndex['kt_m_vt_19']] || '';
        const sl_m_vt_19 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_m_vt_19']] || '');
        const m_vt_20 = chiTietRow[chiTietColumnIndex['m_vt_20']] || '';
        const kt_m_vt_20 = chiTietRow[chiTietColumnIndex['kt_m_vt_20']] || '';
        const sl_m_vt_20 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_m_vt_20']] || '');

        // Lấy các trường m2_vt, kt1_m2_vt, kt2_m2_vt, sl_m2_vt (1 đến 2)
        const m2_vt_1 = chiTietRow[chiTietColumnIndex['m2_vt_1']] || '';
        const kt1_m2_vt_1 = chiTietRow[chiTietColumnIndex['kt1_m2_vt_1']] || '';
        const kt2_m2_vt_1 = chiTietRow[chiTietColumnIndex['kt2_m2_vt_1']] || '';
        const sl_m2_vt_1 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_m2_vt_1']] || '');
        const m2_vt_2 = chiTietRow[chiTietColumnIndex['m2_vt_2']] || '';
        const kt1_m2_vt_2 = chiTietRow[chiTietColumnIndex['kt1_m2_vt_2']] || '';
        const kt2_m2_vt_2 = chiTietRow[chiTietColumnIndex['kt2_m2_vt_2']] || '';
        const sl_m2_vt_2 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_m2_vt_2']] || '');

        // Lấy các trường c_vt, sl_c_vt (1 đến 30)
        const c_vt_1 = chiTietRow[chiTietColumnIndex['c_vt_1']] || '';
        const sl_c_vt_1 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_1']] || '');
        const c_vt_2 = chiTietRow[chiTietColumnIndex['c_vt_2']] || '';
        const sl_c_vt_2 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_2']] || '');
        const c_vt_3 = chiTietRow[chiTietColumnIndex['c_vt_3']] || '';
        const sl_c_vt_3 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_3']] || '');
        const c_vt_4 = chiTietRow[chiTietColumnIndex['c_vt_4']] || '';
        const sl_c_vt_4 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_4']] || '');
        const c_vt_5 = chiTietRow[chiTietColumnIndex['c_vt_5']] || '';
        const sl_c_vt_5 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_5']] || '');
        const c_vt_6 = chiTietRow[chiTietColumnIndex['c_vt_6']] || '';
        const sl_c_vt_6 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_6']] || '');
        const c_vt_7 = chiTietRow[chiTietColumnIndex['c_vt_7']] || '';
        const sl_c_vt_7 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_7']] || '');
        const c_vt_8 = chiTietRow[chiTietColumnIndex['c_vt_8']] || '';
        const sl_c_vt_8 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_8']] || '');
        const c_vt_9 = chiTietRow[chiTietColumnIndex['c_vt_9']] || '';
        const sl_c_vt_9 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_9']] || '');
        const c_vt_10 = chiTietRow[chiTietColumnIndex['c_vt_10']] || '';
        const sl_c_vt_10 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_10']] || '');
        const c_vt_11 = chiTietRow[chiTietColumnIndex['c_vt_11']] || '';
        const sl_c_vt_11 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_11']] || '');
        const c_vt_12 = chiTietRow[chiTietColumnIndex['c_vt_12']] || '';
        const sl_c_vt_12 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_12']] || '');
        const c_vt_13 = chiTietRow[chiTietColumnIndex['c_vt_13']] || '';
        const sl_c_vt_13 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_13']] || '');
        const c_vt_14 = chiTietRow[chiTietColumnIndex['c_vt_14']] || '';
        const sl_c_vt_14 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_14']] || '');
        const c_vt_15 = chiTietRow[chiTietColumnIndex['c_vt_15']] || '';
        const sl_c_vt_15 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_15']] || '');
        const c_vt_16 = chiTietRow[chiTietColumnIndex['c_vt_16']] || '';
        const sl_c_vt_16 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_16']] || '');
        const c_vt_17 = chiTietRow[chiTietColumnIndex['c_vt_17']] || '';
        const sl_c_vt_17 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_17']] || '');
        const c_vt_18 = chiTietRow[chiTietColumnIndex['c_vt_18']] || '';
        const sl_c_vt_18 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_18']] || '');
        const c_vt_19 = chiTietRow[chiTietColumnIndex['c_vt_19']] || '';
        const sl_c_vt_19 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_19']] || '');
        const c_vt_20 = chiTietRow[chiTietColumnIndex['c_vt_20']] || '';
        const sl_c_vt_20 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_20']] || '');
        const c_vt_21 = chiTietRow[chiTietColumnIndex['c_vt_21']] || '';
        const sl_c_vt_21 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_21']] || '');
        const c_vt_22 = chiTietRow[chiTietColumnIndex['c_vt_22']] || '';
        const sl_c_vt_22 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_22']] || '');
        const c_vt_23 = chiTietRow[chiTietColumnIndex['c_vt_23']] || '';
        const sl_c_vt_23 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_23']] || '');
        const c_vt_24 = chiTietRow[chiTietColumnIndex['c_vt_24']] || '';
        const sl_c_vt_24 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_24']] || '');
        const c_vt_25 = chiTietRow[chiTietColumnIndex['c_vt_25']] || '';
        const sl_c_vt_25 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_25']] || '');
        const c_vt_26 = chiTietRow[chiTietColumnIndex['c_vt_26']] || '';
        const sl_c_vt_26 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_26']] || '');
        const c_vt_27 = chiTietRow[chiTietColumnIndex['c_vt_27']] || '';
        const sl_c_vt_27 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_27']] || '');
        const c_vt_28 = chiTietRow[chiTietColumnIndex['c_vt_28']] || '';
        const sl_c_vt_28 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_28']] || '');
        const c_vt_29 = chiTietRow[chiTietColumnIndex['c_vt_29']] || '';
        const sl_c_vt_29 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_29']] || '');
        const c_vt_30 = chiTietRow[chiTietColumnIndex['c_vt_30']] || '';
        const sl_c_vt_30 = formatNumberForDisplay(chiTietRow[chiTietColumnIndex['sl_c_vt_30']] || '');

        // Tạo lý do xuất mới
        const lyDoXuat = `Xuất BH - ${maKhachHangID} - ${tenNguoiLienHe} - ${diaChiChiTiet} - ${maHopDong}`;

        // Tạo hàng với các cột theo yêu cầu
        const row = document.createElement('tr');
        row.innerHTML = `
        <td>${maDonHangID}</td>
        <td>${sttTrongDon}</td>
        <td>3</td>
        <td>${ngayXuatKho}</td>
        <td>${ngayXuatKho}</td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td>${lyDoXuat}</td>
        <td>${mnvCongTy}</td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td>${kho}</td>
        <td></td>
        <td>154</td>
        <td>1521</td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td>${doiTuongTHCP}</td>
        <td>${maSanPhamTheoDoi}</td>
        <td></td>
        <td>${maHopDong}</td>
        <td></td>
        <td></td>
        <td></td>
        <!-- Các cột m_vt, kt_m_vt, sl_m_vt ... (từ 1 đến 20) -->
        <td>${m_vt_1}</td>
        <td>${kt_m_vt_1}</td>
        <td>${sl_m_vt_1}</td>
        <td>${m_vt_2}</td>
        <td>${kt_m_vt_2}</td>
        <td>${sl_m_vt_2}</td>
        <td>${m_vt_3}</td>
        <td>${kt_m_vt_3}</td>
        <td>${sl_m_vt_3}</td>
        <td>${m_vt_4}</td>
        <td>${kt_m_vt_4}</td>
        <td>${sl_m_vt_4}</td>
        <td>${m_vt_5}</td>
        <td>${kt_m_vt_5}</td>
        <td>${sl_m_vt_5}</td>
        <td>${m_vt_6}</td>
        <td>${kt_m_vt_6}</td>
        <td>${sl_m_vt_6}</td>
        <td>${m_vt_7}</td>
        <td>${kt_m_vt_7}</td>
        <td>${sl_m_vt_7}</td>
        <td>${m_vt_8}</td>
        <td>${kt_m_vt_8}</td>
        <td>${sl_m_vt_8}</td>
        <td>${m_vt_9}</td>
        <td>${kt_m_vt_9}</td>
        <td>${sl_m_vt_9}</td>
        <td>${m_vt_10}</td>
        <td>${kt_m_vt_10}</td>
        <td>${sl_m_vt_10}</td>
        <td>${m_vt_11}</td>
        <td>${kt_m_vt_11}</td>
        <td>${sl_m_vt_11}</td>
        <td>${m_vt_12}</td>
        <td>${kt_m_vt_12}</td>
        <td>${sl_m_vt_12}</td>
        <td>${m_vt_13}</td>
        <td>${kt_m_vt_13}</td>
        <td>${sl_m_vt_13}</td>
        <td>${m_vt_14}</td>
        <td>${kt_m_vt_14}</td>
        <td>${sl_m_vt_14}</td>
        <td>${m_vt_15}</td>
        <td>${kt_m_vt_15}</td>
        <td>${sl_m_vt_15}</td>
        <td>${m_vt_16}</td>
        <td>${kt_m_vt_16}</td>
        <td>${sl_m_vt_16}</td>
        <td>${m_vt_17}</td>
        <td>${kt_m_vt_17}</td>
        <td>${sl_m_vt_17}</td>
        <td>${m_vt_18}</td>
        <td>${kt_m_vt_18}</td>
        <td>${sl_m_vt_18}</td>
        <td>${m_vt_19}</td>
        <td>${kt_m_vt_19}</td>
        <td>${sl_m_vt_19}</td>
        <td>${m_vt_20}</td>
        <td>${kt_m_vt_20}</td>
        <td>${sl_m_vt_20}</td>
        <td>${m2_vt_1}</td>
        <td>${kt1_m2_vt_1}</td>
        <td>${kt2_m2_vt_1}</td>
        <td>${sl_m2_vt_1}</td>
        <td>${m2_vt_2}</td>
        <td>${kt1_m2_vt_2}</td>
        <td>${kt2_m2_vt_2}</td>
        <td>${sl_m2_vt_2}</td>
        <td>${c_vt_1}</td>
        <td>${sl_c_vt_1}</td>
        <td>${c_vt_2}</td>
        <td>${sl_c_vt_2}</td>
        <td>${c_vt_3}</td>
        <td>${sl_c_vt_3}</td>
        <td>${c_vt_4}</td>
        <td>${sl_c_vt_4}</td>
        <td>${c_vt_5}</td>
        <td>${sl_c_vt_5}</td>
        <td>${c_vt_6}</td>
        <td>${sl_c_vt_6}</td>
        <td>${c_vt_7}</td>
        <td>${sl_c_vt_7}</td>
        <td>${c_vt_8}</td>
        <td>${sl_c_vt_8}</td>
        <td>${c_vt_9}</td>
        <td>${sl_c_vt_9}</td>
        <td>${c_vt_10}</td>
        <td>${sl_c_vt_10}</td>
        <td>${c_vt_11}</td>
        <td>${sl_c_vt_11}</td>
        <td>${c_vt_12}</td>
        <td>${sl_c_vt_12}</td>
        <td>${c_vt_13}</td>
        <td>${sl_c_vt_13}</td>
        <td>${c_vt_14}</td>
        <td>${sl_c_vt_14}</td>
        <td>${c_vt_15}</td>
        <td>${sl_c_vt_15}</td>
        <td>${c_vt_16}</td>
        <td>${sl_c_vt_16}</td>
        <td>${c_vt_17}</td>
        <td>${sl_c_vt_17}</td>
        <td>${c_vt_18}</td>
        <td>${sl_c_vt_18}</td>
        <td>${c_vt_19}</td>
        <td>${sl_c_vt_19}</td>
        <td>${c_vt_20}</td>
        <td>${sl_c_vt_20}</td>
        <td>${c_vt_21}</td>
        <td>${sl_c_vt_21}</td>
        <td>${c_vt_22}</td>
        <td>${sl_c_vt_22}</td>
        <td>${c_vt_23}</td>
        <td>${sl_c_vt_23}</td>
        <td>${c_vt_24}</td>
        <td>${sl_c_vt_24}</td>
        <td>${c_vt_25}</td>
        <td>${sl_c_vt_25}</td>
        <td>${c_vt_26}</td>
        <td>${sl_c_vt_26}</td>
        <td>${c_vt_27}</td>
        <td>${sl_c_vt_27}</td>
        <td>${c_vt_28}</td>
        <td>${sl_c_vt_28}</td>
        <td>${c_vt_29}</td>
        <td>${sl_c_vt_29}</td>
        <td>${c_vt_30}</td>
        <td>${sl_c_vt_30}</td>
        <td>${soLuong}</td>
        `;
        // Lưu ý: Bạn cần thêm đầy đủ các cột còn thiếu (m_vt_2 đến m_vt_20, m2_vt, c_vt) vào row.innerHTML.
        // Để ngắn gọn, tôi chỉ thêm một vài cột. Bạn có thể copy từ phần hiển thị của tab lệnh sản xuất và dán vào đây.

        resultsBody.appendChild(row);
    });
}

function exportToExcelXuatBaoHanh() {
    if (filteredResultsXuatBaoHanh.length === 0) {
        showMessageXuatBaoHanh('Không có dữ liệu để xuất. Vui lòng thực hiện lọc trước.', 'error');
        return;
    }

    let csvContent = "\uFEFF";
    // Header row với tất cả các cột (giống tab lệnh sản xuất)
    csvContent += "ID,Hiển thị trên sổ,Loại xuất kho,Ngày hạch toán (*),Ngày chứng từ (*),Số chứng từ (*),Mẫu số HĐ,Ký hiệu HĐ,Mã đối tượng,Tên đối tượng,Địa chỉ/Bộ phận,Tên người nhận/Của,Lý do xuất/Về việc,Nhân viên bán hàng,Kèm theo,Số lệnh điều động,Ngày lệnh điều động,Người vận chuyển,Tên người vận chuyển,Hợp đồng số,Phương tiện vận chuyển,Xuất tại kho,Địa chỉ kho xuất,Nhập tại chi nhánh,Tên chi nhánh,MST chi nhánh,Nhập tại kho,Địa chỉ kho nhập,Mã hàng (*),Tên hàng,Là hàng khuyến mại,Kho (*),Hàng hóa giữ hộ/bán hộ,TK Nợ (*),TK Có (*),ĐVT,Số lượng,Đơn giá bán,Thành tiền,Đơn giá vốn,Tiền vốn,Số lô,Hạn sử dụng,Đối tượng,Khoản mục CP,Đơn vị,Đối tượng THCP,Công trình,Đơn đặt hàng,Hợp đồng bán,CP không hợp lý,Mã thống kê,m_vt_1,kt_m_vt_1,sl_m_vt_1,m_vt_2,kt_m_vt_2,sl_m_vt_2,m_vt_3,kt_m_vt_3,sl_m_vt_3,m_vt_4,kt_m_vt_4,sl_m_vt_4,m_vt_5,kt_m_vt_5,sl_m_vt_5,m_vt_6,kt_m_vt_6,sl_m_vt_6,m_vt_7,kt_m_vt_7,sl_m_vt_7,m_vt_8,kt_m_vt_8,sl_m_vt_8,m_vt_9,kt_m_vt_9,sl_m_vt_9,m_vt_10,kt_m_vt_10,sl_m_vt_10,m_vt_11,kt_m_vt_11,sl_m_vt_11,m_vt_12,kt_m_vt_12,sl_m_vt_12,m_vt_13,kt_m_vt_13,sl_m_vt_13,m_vt_14,kt_m_vt_14,sl_m_vt_14,m_vt_15,kt_m_vt_15,sl_m_vt_15,m_vt_16,kt_m_vt_16,sl_m_vt_16,m_vt_17,kt_m_vt_17,sl_m_vt_17,m_vt_18,kt_m_vt_18,sl_m_vt_18,m_vt_19,kt_m_vt_19,sl_m_vt_19,m_vt_20,kt_m_vt_20,sl_m_vt_20,m2_vt_1,kt1_m2_vt_1,kt2_m2_vt_1,sl_m2_vt_1,m2_vt_2,kt1_m2_vt_2,kt2_m2_vt_2,sl_m2_vt_2,c_vt_1,sl_c_vt_1,c_vt_2,sl_c_vt_2,c_vt_3,sl_c_vt_3,c_vt_4,sl_c_vt_4,c_vt_5,sl_c_vt_5,c_vt_6,sl_c_vt_6,c_vt_7,sl_c_vt_7,c_vt_8,sl_c_vt_8,c_vt_9,sl_c_vt_9,c_vt_10,sl_c_vt_10,c_vt_11,sl_c_vt_11,c_vt_12,sl_c_vt_12,c_vt_13,sl_c_vt_13,c_vt_14,sl_c_vt_14,c_vt_15,sl_c_vt_15,c_vt_16,sl_c_vt_16,c_vt_17,sl_c_vt_17,c_vt_18,sl_c_vt_18,c_vt_19,sl_c_vt_19,c_vt_20,sl_c_vt_20,c_vt_21,sl_c_vt_21,c_vt_22,sl_c_vt_22,c_vt_23,sl_c_vt_23,c_vt_24,sl_c_vt_24,c_vt_25,sl_c_vt_25,c_vt_26,sl_c_vt_26,c_vt_27,sl_c_vt_27,c_vt_28,sl_c_vt_28,c_vt_29,sl_c_vt_29,c_vt_30,sl_c_vt_30,Số bộ\n";

    filteredResultsXuatBaoHanh.forEach(result => {
        const chiTietRow = result.chiTietRow;
        const donHangRow = result.donHangRow;
        const chiTietColumnIndex = result.chiTietColumnIndex;
        const donHangColumnIndex = result.donHangColumnIndex;

        const maDonHangID = chiTietRow[chiTietColumnIndex['ma_don_hang_id']] || '';
        const maSanPhamTheoDoi = chiTietRow[chiTietColumnIndex['ma_san_pham_theo_doi']] || '';
        const dienGiai = chiTietRow[chiTietColumnIndex['dien_giai']] || '';
        const ghiChu = chiTietRow[chiTietColumnIndex['ghi_chu']] || '';
        const dvt = chiTietRow[chiTietColumnIndex['dvt']] || '';
        const soLuong = chiTietRow[chiTietColumnIndex['so_luong']] || '';
        const khoiLuong = formatNumberForCSV(chiTietRow[chiTietColumnIndex['khoi_luong']] || '');
        const sttTrongDon = chiTietRow[chiTietColumnIndex['stt_trong_don']] || '';

        const ngayXuatKho = donHangRow[donHangColumnIndex['ngay_xuat_kho']] || '';
        const xuongSanXuat = donHangRow[donHangColumnIndex['xuong_san_xuat']] || '';
        const maHopDong = donHangRow[donHangColumnIndex['ma_hop_dong']] || '';
        const maKhachHangID = donHangRow[donHangColumnIndex['ma_khach_hang_id']] || '';
        const tenNguoiLienHe = donHangRow[donHangColumnIndex['ten_nguoi_lien_he']] || '';
        const diaChiChiTiet = donHangRow[donHangColumnIndex['dia_chi_chi_tiet']] || '';

        let tenHang = '';
        if (dienGiai !== "") {
            tenHang = dienGiai;
        } else {
            tenHang = ghiChu;
        }

        const mnvCongTy = calculateMnvCongTy(donHangRow, donHangColumnIndex);

        // Xác định kho dựa trên xưởng sản xuất
        let kho = '';
        if (xuongSanXuat === "Hà Nội") {
            kho = "K03_SX.HN_152";
        } else if (xuongSanXuat === "Long An") {
            kho = "K04_SX.LA_152";
        }

        // Xác định đối tượng THCP dựa trên xưởng sản xuất
        let doiTuongTHCP = '';
        if (xuongSanXuat === "Hà Nội") {
            doiTuongTHCP = "25.SXT.X1";
        } else if (xuongSanXuat === "Long An") {
            doiTuongTHCP = "25.SXT.X2";
        }

        // Lấy các trường m_vt, kt_m_vt, sl_m_vt (từ 1 đến 20) - tương tự như tab lệnh sản xuất
        // Để ngắn gọn, tôi sẽ không liệt kê hết 20 trường ở đây. Bạn có thể copy từ hàm exportToExcelLenSanXuat.
        const m_vt_1 = chiTietRow[chiTietColumnIndex['m_vt_1']] || '';
        const kt_m_vt_1 = chiTietRow[chiTietColumnIndex['kt_m_vt_1']] || '';
        const sl_m_vt_1 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_m_vt_1']] || '');
        const m_vt_2 = chiTietRow[chiTietColumnIndex['m_vt_2']] || '';
        const kt_m_vt_2 = chiTietRow[chiTietColumnIndex['kt_m_vt_2']] || '';
        const sl_m_vt_2 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_m_vt_2']] || '');
        const m_vt_3 = chiTietRow[chiTietColumnIndex['m_vt_3']] || '';
        const kt_m_vt_3 = chiTietRow[chiTietColumnIndex['kt_m_vt_3']] || '';
        const sl_m_vt_3 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_m_vt_3']] || '');
        const m_vt_4 = chiTietRow[chiTietColumnIndex['m_vt_4']] || '';
        const kt_m_vt_4 = chiTietRow[chiTietColumnIndex['kt_m_vt_4']] || '';
        const sl_m_vt_4 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_m_vt_4']] || '');
        const m_vt_5 = chiTietRow[chiTietColumnIndex['m_vt_5']] || '';
        const kt_m_vt_5 = chiTietRow[chiTietColumnIndex['kt_m_vt_5']] || '';
        const sl_m_vt_5 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_m_vt_5']] || '');
        const m_vt_6 = chiTietRow[chiTietColumnIndex['m_vt_6']] || '';
        const kt_m_vt_6 = chiTietRow[chiTietColumnIndex['kt_m_vt_6']] || '';
        const sl_m_vt_6 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_m_vt_6']] || '');
        const m_vt_7 = chiTietRow[chiTietColumnIndex['m_vt_7']] || '';
        const kt_m_vt_7 = chiTietRow[chiTietColumnIndex['kt_m_vt_7']] || '';
        const sl_m_vt_7 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_m_vt_7']] || '');
        const m_vt_8 = chiTietRow[chiTietColumnIndex['m_vt_8']] || '';
        const kt_m_vt_8 = chiTietRow[chiTietColumnIndex['kt_m_vt_8']] || '';
        const sl_m_vt_8 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_m_vt_8']] || '');
        const m_vt_9 = chiTietRow[chiTietColumnIndex['m_vt_9']] || '';
        const kt_m_vt_9 = chiTietRow[chiTietColumnIndex['kt_m_vt_9']] || '';
        const sl_m_vt_9 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_m_vt_9']] || '');
        const m_vt_10 = chiTietRow[chiTietColumnIndex['m_vt_10']] || '';
        const kt_m_vt_10 = chiTietRow[chiTietColumnIndex['kt_m_vt_10']] || '';
        const sl_m_vt_10 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_m_vt_10']] || '');
        const m_vt_11 = chiTietRow[chiTietColumnIndex['m_vt_11']] || '';
        const kt_m_vt_11 = chiTietRow[chiTietColumnIndex['kt_m_vt_11']] || '';
        const sl_m_vt_11 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_m_vt_11']] || '');
        const m_vt_12 = chiTietRow[chiTietColumnIndex['m_vt_12']] || '';
        const kt_m_vt_12 = chiTietRow[chiTietColumnIndex['kt_m_vt_12']] || '';
        const sl_m_vt_12 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_m_vt_12']] || '');
        const m_vt_13 = chiTietRow[chiTietColumnIndex['m_vt_13']] || '';
        const kt_m_vt_13 = chiTietRow[chiTietColumnIndex['kt_m_vt_13']] || '';
        const sl_m_vt_13 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_m_vt_13']] || '');
        const m_vt_14 = chiTietRow[chiTietColumnIndex['m_vt_14']] || '';
        const kt_m_vt_14 = chiTietRow[chiTietColumnIndex['kt_m_vt_14']] || '';
        const sl_m_vt_14 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_m_vt_14']] || '');
        const m_vt_15 = chiTietRow[chiTietColumnIndex['m_vt_15']] || '';
        const kt_m_vt_15 = chiTietRow[chiTietColumnIndex['kt_m_vt_15']] || '';
        const sl_m_vt_15 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_m_vt_15']] || '');
        const m_vt_16 = chiTietRow[chiTietColumnIndex['m_vt_16']] || '';
        const kt_m_vt_16 = chiTietRow[chiTietColumnIndex['kt_m_vt_16']] || '';
        const sl_m_vt_16 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_m_vt_16']] || '');
        const m_vt_17 = chiTietRow[chiTietColumnIndex['m_vt_17']] || '';
        const kt_m_vt_17 = chiTietRow[chiTietColumnIndex['kt_m_vt_17']] || '';
        const sl_m_vt_17 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_m_vt_17']] || '');
        const m_vt_18 = chiTietRow[chiTietColumnIndex['m_vt_18']] || '';
        const kt_m_vt_18 = chiTietRow[chiTietColumnIndex['kt_m_vt_18']] || '';
        const sl_m_vt_18 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_m_vt_18']] || '');
        const m_vt_19 = chiTietRow[chiTietColumnIndex['m_vt_19']] || '';
        const kt_m_vt_19 = chiTietRow[chiTietColumnIndex['kt_m_vt_19']] || '';
        const sl_m_vt_19 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_m_vt_19']] || '');
        const m_vt_20 = chiTietRow[chiTietColumnIndex['m_vt_20']] || '';
        const kt_m_vt_20 = chiTietRow[chiTietColumnIndex['kt_m_vt_20']] || '';
        const sl_m_vt_20 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_m_vt_20']] || '');

        // Lấy các trường m2_vt, kt1_m2_vt, kt2_m2_vt, sl_m2_vt (1 đến 2)
        const m2_vt_1 = chiTietRow[chiTietColumnIndex['m2_vt_1']] || '';
        const kt1_m2_vt_1 = chiTietRow[chiTietColumnIndex['kt1_m2_vt_1']] || '';
        const kt2_m2_vt_1 = chiTietRow[chiTietColumnIndex['kt2_m2_vt_1']] || '';
        const sl_m2_vt_1 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_m2_vt_1']] || '');
        const m2_vt_2 = chiTietRow[chiTietColumnIndex['m2_vt_2']] || '';
        const kt1_m2_vt_2 = chiTietRow[chiTietColumnIndex['kt1_m2_vt_2']] || '';
        const kt2_m2_vt_2 = chiTietRow[chiTietColumnIndex['kt2_m2_vt_2']] || '';
        const sl_m2_vt_2 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_m2_vt_2']] || '');

        // Lấy các trường c_vt, sl_c_vt (1 đến 30)
        const c_vt_1 = chiTietRow[chiTietColumnIndex['c_vt_1']] || '';
        const sl_c_vt_1 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_1']] || '');
        const c_vt_2 = chiTietRow[chiTietColumnIndex['c_vt_2']] || '';
        const sl_c_vt_2 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_2']] || '');
        const c_vt_3 = chiTietRow[chiTietColumnIndex['c_vt_3']] || '';
        const sl_c_vt_3 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_3']] || '');
        const c_vt_4 = chiTietRow[chiTietColumnIndex['c_vt_4']] || '';
        const sl_c_vt_4 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_4']] || '');
        const c_vt_5 = chiTietRow[chiTietColumnIndex['c_vt_5']] || '';
        const sl_c_vt_5 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_5']] || '');
        const c_vt_6 = chiTietRow[chiTietColumnIndex['c_vt_6']] || '';
        const sl_c_vt_6 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_6']] || '');
        const c_vt_7 = chiTietRow[chiTietColumnIndex['c_vt_7']] || '';
        const sl_c_vt_7 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_7']] || '');
        const c_vt_8 = chiTietRow[chiTietColumnIndex['c_vt_8']] || '';
        const sl_c_vt_8 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_8']] || '');
        const c_vt_9 = chiTietRow[chiTietColumnIndex['c_vt_9']] || '';
        const sl_c_vt_9 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_9']] || '');
        const c_vt_10 = chiTietRow[chiTietColumnIndex['c_vt_10']] || '';
        const sl_c_vt_10 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_10']] || '');
        const c_vt_11 = chiTietRow[chiTietColumnIndex['c_vt_11']] || '';
        const sl_c_vt_11 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_11']] || '');
        const c_vt_12 = chiTietRow[chiTietColumnIndex['c_vt_12']] || '';
        const sl_c_vt_12 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_12']] || '');
        const c_vt_13 = chiTietRow[chiTietColumnIndex['c_vt_13']] || '';
        const sl_c_vt_13 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_13']] || '');
        const c_vt_14 = chiTietRow[chiTietColumnIndex['c_vt_14']] || '';
        const sl_c_vt_14 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_14']] || '');
        const c_vt_15 = chiTietRow[chiTietColumnIndex['c_vt_15']] || '';
        const sl_c_vt_15 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_15']] || '');
        const c_vt_16 = chiTietRow[chiTietColumnIndex['c_vt_16']] || '';
        const sl_c_vt_16 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_16']] || '');
        const c_vt_17 = chiTietRow[chiTietColumnIndex['c_vt_17']] || '';
        const sl_c_vt_17 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_17']] || '');
        const c_vt_18 = chiTietRow[chiTietColumnIndex['c_vt_18']] || '';
        const sl_c_vt_18 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_18']] || '');
        const c_vt_19 = chiTietRow[chiTietColumnIndex['c_vt_19']] || '';
        const sl_c_vt_19 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_19']] || '');
        const c_vt_20 = chiTietRow[chiTietColumnIndex['c_vt_20']] || '';
        const sl_c_vt_20 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_20']] || '');
        const c_vt_21 = chiTietRow[chiTietColumnIndex['c_vt_21']] || '';
        const sl_c_vt_21 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_21']] || '');
        const c_vt_22 = chiTietRow[chiTietColumnIndex['c_vt_22']] || '';
        const sl_c_vt_22 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_22']] || '');
        const c_vt_23 = chiTietRow[chiTietColumnIndex['c_vt_23']] || '';
        const sl_c_vt_23 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_23']] || '');
        const c_vt_24 = chiTietRow[chiTietColumnIndex['c_vt_24']] || '';
        const sl_c_vt_24 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_24']] || '');
        const c_vt_25 = chiTietRow[chiTietColumnIndex['c_vt_25']] || '';
        const sl_c_vt_25 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_25']] || '');
        const c_vt_26 = chiTietRow[chiTietColumnIndex['c_vt_26']] || '';
        const sl_c_vt_26 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_26']] || '');
        const c_vt_27 = chiTietRow[chiTietColumnIndex['c_vt_27']] || '';
        const sl_c_vt_27 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_27']] || '');
        const c_vt_28 = chiTietRow[chiTietColumnIndex['c_vt_28']] || '';
        const sl_c_vt_28 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_28']] || '');
        const c_vt_29 = chiTietRow[chiTietColumnIndex['c_vt_29']] || '';
        const sl_c_vt_29 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_29']] || '');
        const c_vt_30 = chiTietRow[chiTietColumnIndex['c_vt_30']] || '';
        const sl_c_vt_30 = formatNumberForCSV(chiTietRow[chiTietColumnIndex['sl_c_vt_30']] || '');


        // Tạo lý do xuất mới
        const lyDoXuat = `Xuất BH - ${maKhachHangID} - ${tenNguoiLienHe} - ${diaChiChiTiet} - ${maHopDong}`;

        const escapeCSV = (str) => {
            if (!str) return '';
            if (str.includes(',') || str.includes('"') || str.includes('\n')) {
                return '"' + str.replace(/"/g, '""') + '"';
            }
            return str;
        };

        // Tạo dòng CSV với tất cả các cột (cần điền đầy đủ các giá trị)
        // Để ngắn gọn, tôi chỉ điền một số cột, bạn cần điền đầy đủ như trong tab lệnh sản xuất.
        csvContent += `${escapeCSV(maDonHangID)},${escapeCSV(sttTrongDon)},3,${escapeCSV(ngayXuatKho)},${escapeCSV(ngayXuatKho)},,,,,,,,${escapeCSV(lyDoXuat)}>,${escapeCSV(mnvCongTy)},,,,,,,,,,,,,,,,,,${escapeCSV(kho)},,154,1521,,,,,,,,,,,${escapeCSV(doiTuongTHCP)},${escapeCSV(maSanPhamTheoDoi)},,${escapeCSV(maHopDong)},,,,${escapeCSV(m_vt_1)},${escapeCSV(kt_m_vt_1)},${sl_m_vt_1},${escapeCSV(m_vt_2)},${escapeCSV(kt_m_vt_2)},${sl_m_vt_2},${escapeCSV(m_vt_3)},${escapeCSV(kt_m_vt_3)},${sl_m_vt_3},${escapeCSV(m_vt_4)},${escapeCSV(kt_m_vt_4)},${sl_m_vt_4},${escapeCSV(m_vt_5)},${escapeCSV(kt_m_vt_5)},${sl_m_vt_5},${escapeCSV(m_vt_6)},${escapeCSV(kt_m_vt_6)},${sl_m_vt_6},${escapeCSV(m_vt_7)},${escapeCSV(kt_m_vt_7)},${sl_m_vt_7},${escapeCSV(m_vt_8)},${escapeCSV(kt_m_vt_8)},${sl_m_vt_8},${escapeCSV(m_vt_9)},${escapeCSV(kt_m_vt_9)},${sl_m_vt_9},${escapeCSV(m_vt_10)},${escapeCSV(kt_m_vt_10)},${sl_m_vt_10},${escapeCSV(m_vt_11)},${escapeCSV(kt_m_vt_11)},${sl_m_vt_11},${escapeCSV(m_vt_12)},${escapeCSV(kt_m_vt_12)},${sl_m_vt_12},${escapeCSV(m_vt_13)},${escapeCSV(kt_m_vt_13)},${sl_m_vt_13},${escapeCSV(m_vt_14)},${escapeCSV(kt_m_vt_14)},${sl_m_vt_14},${escapeCSV(m_vt_15)},${escapeCSV(kt_m_vt_15)},${sl_m_vt_15},${escapeCSV(m_vt_16)},${escapeCSV(kt_m_vt_16)},${sl_m_vt_16},${escapeCSV(m_vt_17)},${escapeCSV(kt_m_vt_17)},${sl_m_vt_17},${escapeCSV(m_vt_18)},${escapeCSV(kt_m_vt_18)},${sl_m_vt_18},${escapeCSV(m_vt_19)},${escapeCSV(kt_m_vt_19)},${sl_m_vt_19},${escapeCSV(m_vt_20)},${escapeCSV(kt_m_vt_20)},${sl_m_vt_20},${escapeCSV(m2_vt_1)},${escapeCSV(kt1_m2_vt_1)},${escapeCSV(kt2_m2_vt_1)},${sl_m2_vt_1},${escapeCSV(m2_vt_2)},${escapeCSV(kt1_m2_vt_2)},${escapeCSV(kt2_m2_vt_2)},${sl_m2_vt_2},${escapeCSV(c_vt_1)},${sl_c_vt_1},${escapeCSV(c_vt_2)},${sl_c_vt_2},${escapeCSV(c_vt_3)},${sl_c_vt_3},${escapeCSV(c_vt_4)},${sl_c_vt_4},${escapeCSV(c_vt_5)},${sl_c_vt_5},${escapeCSV(c_vt_6)},${sl_c_vt_6},${escapeCSV(c_vt_7)},${sl_c_vt_7},${escapeCSV(c_vt_8)},${sl_c_vt_8},${escapeCSV(c_vt_9)},${sl_c_vt_9},${escapeCSV(c_vt_10)},${sl_c_vt_10},${escapeCSV(c_vt_11)},${sl_c_vt_11},${escapeCSV(c_vt_12)},${sl_c_vt_12},${escapeCSV(c_vt_13)},${sl_c_vt_13},${escapeCSV(c_vt_14)},${sl_c_vt_14},${escapeCSV(c_vt_15)},${sl_c_vt_15},${escapeCSV(c_vt_16)},${sl_c_vt_16},${escapeCSV(c_vt_17)},${sl_c_vt_17},${escapeCSV(c_vt_18)},${sl_c_vt_18},${escapeCSV(c_vt_19)},${sl_c_vt_19},${escapeCSV(c_vt_20)},${sl_c_vt_20},${escapeCSV(c_vt_21)},${sl_c_vt_21},${escapeCSV(c_vt_22)},${sl_c_vt_22},${escapeCSV(c_vt_23)},${sl_c_vt_23},${escapeCSV(c_vt_24)},${sl_c_vt_24},${escapeCSV(c_vt_25)},${sl_c_vt_25},${escapeCSV(c_vt_26)},${sl_c_vt_26},${escapeCSV(c_vt_27)},${sl_c_vt_27},${escapeCSV(c_vt_28)},${sl_c_vt_28},${escapeCSV(c_vt_29)},${sl_c_vt_29},${escapeCSV(c_vt_30)},${sl_c_vt_30},${soLuong}\n`;                // Lưu ý: Bạn cần thay thế '... và các cột khác ...' bằng các giá trị thực tế của các cột còn lại.
    });

    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement("a");
    const url = URL.createObjectURL(blob);

    link.setAttribute("href", url);
    const tuNgayRaw = document.getElementById("tu-ngay-xbh").value;
    const denNgayRaw = document.getElementById("den-ngay-xbh").value;
    const tuNgay = formatDateForFilename(tuNgayRaw);
    const denNgay = formatDateForFilename(denNgayRaw);
    const loaiDonHang = document.querySelector('input[name="loai-don-hang-xbh"]:checked').value;
    const fileName = `Danh sách xuất bảo hành - ${tuNgay} - ${denNgay} - ${loaiDonHang}.csv`;

    link.setAttribute("download", fileName);
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);

    showMessageXuatBaoHanh(`Đã xuất ${filteredResultsXuatBaoHanh.length} dòng ra file Excel.`, 'success');
}

function showLoadingXuatBaoHanh(show) {
    const loadingElement = document.getElementById('loading-xbh');
    loadingElement.style.display = show ? 'block' : 'none';
}

function showMessageXuatBaoHanh(message, type) {
    const resultsCount = document.getElementById("results-count-xbh");
    resultsCount.className = "results-count";

    if (type === "success") {
        resultsCount.style.backgroundColor = "#e8f7e8";
        resultsCount.style.color = "#0c9c07";
        resultsCount.style.borderLeft = "4px solid #0c9c07";
    }

    if (type === "error") {
        resultsCount.style.backgroundColor = "#ffeaea";
        resultsCount.style.color = "#c00";
        resultsCount.style.borderLeft = "4px solid #c00";
    }

    resultsCount.style.padding = "5px";
    resultsCount.style.borderRadius = "6px";
    resultsCount.style.fontSize = "16px";
    resultsCount.style.fontWeight = "600";
    resultsCount.textContent = message;
}

function requireRefilterXuatBaoHanh() {
    document.getElementById('results-table-xbh').style.display = 'none';
    document.getElementById('no-results-xbh').style.display = 'block';
    document.getElementById('results-count-xbh').textContent = 'Kết quả: Chưa có dòng nào được lọc.';
    filteredResultsXuatBaoHanh = [];
    showMessageXuatBaoHanh("Bạn đã thay đổi bộ lọc. Vui lòng nhấn 'Lọc' để cập nhật kết quả mới.", "error");
}