// Thay thế toàn bộ file Code.gs

const ss = SpreadsheetApp.getActiveSpreadsheet();
const ccdcSheet = ss.getSheetByName("DanhMuc_CCDC");
const nvSheet = ss.getSheetByName("DanhMuc_NhanVien");
const lsSheet = ss.getSheetByName("LichSu_GiaoDich");

function doGet(e) {
    return HtmlService.createTemplateFromFile('Index')
        .evaluate()
        .setTitle('Hệ thống Quản lý CCDC')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// --- CÁC HÀM LẤY DỮ LIỆU ĐÃ SỬA LỖI ---

function getInitialData() {
    // Hàm này không thay đổi, nhưng để đây cho đầy đủ
    const tools = ccdcSheet.getDataRange().getValues();
    const employees = nvSheet.getDataRange().getValues();
    return {
        tools: tools,
        employees: employees
    };
}

function getTransactions() {
    // FIX: Thêm kiểm tra sheet rỗng
    if (lsSheet.getLastRow() < 2) {
        // Nếu sheet chỉ có 0 hoặc 1 dòng (tiêu đề), trả về mảng rỗng
        // Trả về cả tiêu đề để giao diện không bị lỗi
        return [lsSheet.getRange("A1:J1").getValues()[0]];
    }
    return lsSheet.getDataRange().getValues();
}

function getBorrowedItemsByEmployeeId(employeeId) {
    // FIX: Thêm kiểm tra sheet rỗng
    if (lsSheet.getLastRow() < 2) {
        return []; // Trả về mảng rỗng ngay lập tức
    }

    const allTransactions = lsSheet.getDataRange().getValues();
    // Bỏ qua hàng tiêu đề để tránh lỗi
    const dataRows = allTransactions.slice(1); 
    
    // FIX: Dùng String() và .trim() để so sánh an toàn tuyệt đối
    const borrowedItems = dataRows.filter(row => {
        const transactionEmployeeId = String(row[2]).trim();
        const targetEmployeeId = String(employeeId).trim();
        const status = String(row[8]).trim();
        return transactionEmployeeId == targetEmployeeId && status == 'Đang mượn';
    });
    
    // Trả về kết quả đã lọc
    return borrowedItems;
}

// --- CÁC HÀM XỬ LÝ NGHIỆP VỤ (Giữ nguyên như trước) ---

function issueMultipleTools(data) {
    try {
        const { employeeId, items } = data;
        const toolsData = ccdcSheet.getDataRange().getValues();
        const employeeData = nvSheet.getDataRange().getValues().find(row => String(row[0]).trim() == String(employeeId).trim());

        if (!employeeData) {
             return { success: false, message: "Lỗi: Không tìm thấy nhân viên đã chọn." };
        }

        let messages = [];

        items.forEach(item => {
            let toolRowIndex = toolsData.findIndex(row => String(row[0]).trim() == String(item.toolId).trim());
            if (toolRowIndex === -1) {
                messages.push(`Lỗi: Không tìm thấy CCDC với mã ${item.toolId}.`);
                return;
            }

            const currentStock = parseInt(toolsData[toolRowIndex][3]);
            if (currentStock < item.quantity) {
                messages.push(`Lỗi: Tồn kho của ${toolsData[toolRowIndex][1]} không đủ (Tồn: ${currentStock}, Yêu cầu: ${item.quantity}).`);
                return;
            }

            ccdcSheet.getRange(toolRowIndex + 1, 4).setValue(currentStock - item.quantity);
            
            const transactionId = "GD" + new Date().getTime() + Math.random().toString(36).substr(2, 5);
            lsSheet.appendRow([
                transactionId, new Date(), employeeId, employeeData[1], item.toolId, toolsData[toolRowIndex][1],
                item.quantity, 'Xuất', 'Đang mượn', ''
            ]);
            messages.push(`Thành công: Đã xuất ${item.quantity} ${toolsData[toolRowIndex][1]}.`);
        });

        return { success: true, message: messages.join('\n') };
    } catch (e) {
        return { success: false, message: "Lỗi hệ thống: " + e.message };
    }
}

function returnMultipleTools(data) {
    try {
        const { itemsToReturn } = data;
        const toolsData = ccdcSheet.getDataRange().getValues();
        const transactions = lsSheet.getDataRange().getValues();
        let messages = [];

        itemsToReturn.forEach(item => {
            let toolRowIndex = toolsData.findIndex(row => String(row[0]).trim() == String(item.toolId).trim());
            if(toolRowIndex !== -1){
                const currentStock = parseInt(toolsData[toolRowIndex][3]);
                ccdcSheet.getRange(toolRowIndex + 1, 4).setValue(currentStock + parseInt(item.quantity));
            }

            let transactionRowIndex = transactions.findIndex(row => String(row[0]).trim() == String(item.transactionId).trim());
             if(transactionRowIndex !== -1){
                lsSheet.getRange(transactionRowIndex + 1, 9).setValue('Đã trả');
            }
            messages.push(`Thành công: Đã nhập lại ${item.quantity} ${item.toolName}.`);
        });
        return { success: true, message: messages.join('\n') };
    } catch (e) {
        return { success: false, message: e.message };
    }
}