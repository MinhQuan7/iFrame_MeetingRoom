// Các hàm tiện ích
function formatDate(dateStr) {
    if (!dateStr) return '';

    // Thử các định dạng khác nhau
    const formats = [
        /^\d{1,2}\/\d{1,2}\/\d{4}$/,  // dd/mm/yyyy
        /^\d{1,2}-\d{1,2}-\d{4}$/,    // dd-mm-yyyy
        /^\d{4}-\d{1,2}-\d{1,2}$/     // yyyy-mm-dd
    ];

    for (let format of formats) {
        if (format.test(dateStr)) {
            const parts = dateStr.split(/[\/\-]/);
            const day = parts[0].padStart(2, '0');
            const month = parts[1].padStart(2, '0');
            const year = parts[2];
            return `${day}/${month}/${year}`;
        }
    }

    // Nếu là số (Excel date serial)
    if (!isNaN(Number(dateStr))) {
        const jsDate = new Date(Date.UTC(1900, 0, Number(dateStr) - 1));
        const day = String(jsDate.getUTCDate()).padStart(2, '0');
        const month = String(jsDate.getUTCMonth() + 1).padStart(2, '0');
        const year = jsDate.getUTCFullYear();
        return `${day}/${month}/${year}`;
    }

    return '';
}

function getCurrentTime() {
    const now = new Date();
    return `${String(now.getHours()).padStart(2, "0")}:${String(now.getMinutes()).padStart(2, "0")}`;
}

function getDayOfWeek(dayString) {
    if (!dayString) return 'Không xác định';
    
    const normalizedDay = dayString.toString().toLowerCase().trim();
    const dayMap = {
        '2': 'Thứ Hai',
        '3': 'Thứ Ba',
        '4': 'Thứ Tư',
        '5': 'Thứ Năm',
        '6': 'Thứ Sáu',
        '7': 'Thứ Bảy',
        'cn': 'Chủ Nhật',
        'monday': 'Thứ Hai',
        'tuesday': 'Thứ Ba',
        'wednesday': 'Thứ Tư',
        'thursday': 'Thứ Năm',
        'friday': 'Thứ Sáu',
        'saturday': 'Thứ Bảy',
        'sunday': 'Chủ Nhật'
    };
    
    return dayMap[normalizedDay] || dayString;
}

// Hàm format ngày từ Excel
function formatExcelDate(excelDate) {
    if (!excelDate) return '';

    // Kiểm tra nếu là số (Excel date serial)
    if (typeof excelDate === 'number') {
        // Chuyển đổi từ Excel date serial sang JavaScript Date
        const jsDate = new Date(Date.UTC(1900, 0, excelDate - 1));
        
        // Format lại ngày
        const day = String(jsDate.getUTCDate()).padStart(2, '0');
        const month = String(jsDate.getUTCMonth() + 1).padStart(2, '0');
        const year = jsDate.getUTCFullYear();

        return `${day}/${month}/${year}`;
    }

    // Nếu đã là chuỗi ngày, thử parse
    if (typeof excelDate === 'string') {
        // Thử parse các định dạng khác nhau
        const parsedDate = new Date(excelDate);
        if (!isNaN(parsedDate)) {
            const day = String(parsedDate.getDate()).padStart(2, '0');
            const month = String(parsedDate.getMonth() + 1).padStart(2, '0');
            const year = parsedDate.getFullYear();
            return `${day}/${month}/${year}`;
        }
    }

    return '';
}

function formatTime(timeStr) {
    if (!timeStr) return '';

    // Loại bỏ các giá trị không phải thời gian
    const roomKeywords = [
        'phòng', 'p.', 'lavender', 'lotus', 'watch', 'sk', 
        'meeting', 'phong', 'room'
    ];

    // Chuyển sang chữ thường và loại bỏ khoảng trắng
    const lowerTimeStr = String(timeStr).toLowerCase().trim();

    // Kiểm tra nếu chứa từ khóa phòng thì bỏ qua
    if (roomKeywords.some(keyword => lowerTimeStr.includes(keyword))) {
        return '';
    }

    // Các định dạng thời gian
    const timeFormats = [
        /^\d{1,2}:\d{2}$/,           // HH:MM
        /^\d{1,2}h\d{2}$/,            // Hh:MM
        /^\d{1,2}h\s?\d{2}$/,         // H h MM
        /^\d{1,2}\.\d{2}$/,           // HH.MM
    ];

    for (let format of timeFormats) {
        if (format.test(lowerTimeStr)) {
            // Chuẩn hóa định dạng
            const timeParts = lowerTimeStr.replace('h', ':').replace('.', ':').split(':');
            const hours = timeParts[0].padStart(2, '0');
            const minutes = timeParts[1].padStart(2, '0');
            return `${hours}:${minutes}`;
        }
    }

    // Nếu là số thập phân (Excel time)
    if (!isNaN(parseFloat(timeStr))) {
        const totalMinutes = Math.round(parseFloat(timeStr) * 24 * 60);
        const hours = Math.floor(totalMinutes / 60);
        const minutes = totalMinutes % 60;
        return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}`;
    }

    return '';
}

function normalizeRoomName(room) {
    if (!room) return "Không xác định";

    // Chuyển sang chữ thường và loại bỏ khoảng trắng thừa
    const normalized = String(room).toLowerCase().trim();

    const roomMap = {
        'lotus': "Phòng Lotus",
        'lavender 1': "Phòng Lavender 1", 
        'lavender 2': "Phòng Lavender 2",
        'p.1': "Phòng Lotus",
        'p.2': "Phòng Lavender 1", 
        'p.3': "Phòng Lavender 2",
        'p. lavender 1': "Phòng Lavender 1",
        'p. lavender 2': "Phòng Lavender 2",
    };

    // Thử map trực tiếp
    if (roomMap[normalized]) return roomMap[normalized];

    // Nếu không, kiểm tra từng từ khóa
    for (let [key, value] of Object.entries(roomMap)) {
        if (normalized.includes(key)) return value;
    }

    return room;
}
function isTimeInRange(current, start, end) {
    // So sánh thời gian dạng HH:MM
    const [currentHour, currentMin] = current.split(':').map(Number);
    const [startHour, startMin] = start.split(':').map(Number);
    const [endHour, endMin] = end.split(':').map(Number);

    const currentTime = currentHour * 60 + currentMin;
    const startTime = startHour * 60 + startMin;
    const endTime = endHour * 60 + endMin;

    return currentTime >= startTime && currentTime <= endTime;
}

// Xử lý tải file Excel
// Đảm bảo đã import thư viện XLSX vào HTML
function processExcelFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();

        reader.onload = function (e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: "array" });
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];

                // In ra toàn bộ dữ liệu để debug
                console.log("Raw sheet data:", XLSX.utils.sheet_to_json(firstSheet, { raw: true }));

                // Chuyển sheet sang JSON với nhiều tùy chọn
                const jsonData = XLSX.utils.sheet_to_json(firstSheet, { 
                    raw: false,  // Chuyển đổi giá trị sang dạng string
                    defval: '',  // Giá trị mặc định nếu ô trống
                });

                // In ra dữ liệu JSON để debug
                console.log("Raw JSON data:", jsonData);

                const formattedData = jsonData
                    .filter(row => {
                        // Lọc các dòng có dữ liệu thực
                        return row['Ngày'] || row['NGÀY'] || row['__EMPTY'];
                    })
                    .map((row, index) => {
                        // Chọn cột phù hợp
                        const date = row['Ngày'] || row['NGÀY'] || row['__EMPTY'] || '';
                        const dayOfWeek = row['Thứ'] || row['THỨ'] || row['__EMPTY_1'] || '';
                        const room = row['Phòng'] || row['PHÒNG'] || row['__EMPTY_2'] || '';
                        const startTime = row['Giờ bắt đầu'] || row['GIỜ BẮT ĐẦU'] || row['__EMPTY_3'] || '';
                        const endTime = row['Giờ kết thúc'] || row['GIỜ KẾT THÚC'] || row['__EMPTY_4'] || '';
                        const duration = row['Thời gian'] || row['THỜI GIAN SỬ DỤNG'] || row['__EMPTY_5'] || '';
                        const purpose = row['Mục đích'] || row['MỤC ĐÍCH SỬ DỤNG'] || row['__EMPTY_6'] || '';
                        const content = row['Nội dung'] || row['NỘI DUNG'] || row['__EMPTY_7'] || '';

                        return {
                            id: index + 1,
                            date: formatDate(date),
                            dayOfWeek: getDayOfWeek(dayOfWeek),
                            room: normalizeRoomName(room),
                            startTime: formatTime(startTime),
                            endTime: formatTime(endTime),
                            duration: duration,
                            purpose: purpose,
                            content: content,
                        };
                    });

                resolve(formattedData);
            } catch (error) {
                console.error("Chi tiết lỗi:", error);
                reject(error);
            }
        };

        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}


// Cập nhật bảng lịch
function updateScheduleTable(data) {
    const tableBody = document.querySelector(".schedule-table");
    const headerRow = tableBody.querySelector(".table-header");

    // Xóa các hàng cũ
    Array.from(tableBody.children)
        .filter(child => child !== headerRow)
        .forEach(child => child.remove());

    // Thêm dữ liệu mới
    data.forEach(meeting => {
        const row = document.createElement("div");
        row.className = "table-row";
        row.setAttribute("role", "row");

        row.innerHTML = `
            <div role="cell">${meeting.id}</div>
            <div role="cell">${meeting.date}</div>
            <div role="cell">${meeting.dayOfWeek}</div>
            <div role="cell">${meeting.room}</div>
            <div role="cell">${meeting.startTime}</div>
            <div role="cell">${meeting.endTime}</div>
            <div role="cell">${meeting.duration}</div>
            <div role="cell">${meeting.purpose}</div>
            <div role="cell">${meeting.content}</div>
        `;

        tableBody.appendChild(row);
    });
}

// Cập nhật trạng thái phòng
function updateRoomStatus(data) {
    const currentDate = formatDate(new Date());
    const currentTime = getCurrentTime();

    const roomMapping = {
        "Phòng Lotus": "P.1",
        "Phòng Lavender 1": "P.2",
        "Phòng Lavender 2": "P.3",
    };

    const todayMeetings = data.filter(meeting => 
        meeting.date === currentDate
    );

    Object.entries(roomMapping).forEach(([fullName, shortName]) => {
        const roomMeeting = todayMeetings.find(
            meeting => meeting.room === fullName
        );
        updateSingleRoomStatus(shortName, roomMeeting, currentTime);
    });
}

// Cập nhật trạng thái từng phòng
function updateSingleRoomStatus(roomCode, meeting, currentTime) {
    // Find the room section by iterating through room sections
    const roomSections = document.querySelectorAll('.room-section');
    const roomSection = Array.from(roomSections).find(section => {
        const roomNumberElement = section.querySelector('.room-number');
        return roomNumberElement && roomNumberElement.textContent.trim() === roomCode;
    });
    
    if (!roomSection) return;

    const titleElement = roomSection.querySelector(".meeting-title");
    const startTimeElement = roomSection.querySelector(".start-time");
    const endTimeElement = roomSection.querySelector(".end-time");
    const statusIndicator = roomSection.querySelector(".status-indicator");

    if (meeting && isTimeInRange(currentTime, meeting.startTime, meeting.endTime)) {
        // Phòng đang có cuộc họp
        titleElement.innerHTML = `<span>Thông tin cuộc họp:</span>${meeting.content}`;
        startTimeElement.innerHTML = `<span>Thời gian bắt đầu:</span>${meeting.startTime}`;
        endTimeElement.innerHTML = `<span>Thời gian kết thúc:</span>${meeting.endTime}`;
        statusIndicator.innerHTML = `
            <div class="indicator-dot busy"></div>
            <div class="status-text">Đang họp</div>
        `;
    } else {
        // Phòng trống
        titleElement.innerHTML = `<span>Thông tin cuộc họp:</span>Trống`;
        startTimeElement.innerHTML = `<span>Thời gian bắt đầu:</span>--:--`;
        endTimeElement.innerHTML = `<span>Thời gian kết thúc:</span>--:--`;
        statusIndicator.innerHTML = `
            <div class="indicator-dot available"></div>
            <div class="status-text">Trống</div>
        `;
    }
}
// Xử lý tải file
function handleFileUpload(file) {
    processExcelFile(file)
        .then(data => {
            // Cập nhật bảng lịch
            updateScheduleTable(data);
            
            // Cập nhật trạng thái phòng
            updateRoomStatus(data);
            
            console.log("Xử lý file thành công:", data);
        })
        .catch(error => {
            console.error("Lỗi khi xử lý file:", error);
            alert("Có lỗi xảy ra khi xử lý file. Vui lòng thử lại.");
        });
}
// Tải file lên server
async function uploadToServer(file, processedData) {
    const formData = new FormData();
    formData.append("meetingFile", file);
    formData.append("processedData", JSON.stringify(processedData));

    try {
        const response = await fetch("/api/upload-meeting", {
            method: "POST",
            body: formData,
        });

        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }

        const result = await response.json();
        console.log("Upload thành công:", result);
        return result;
    } catch (error) {
        console.error("Lỗi khi upload:", error);
        throw error;
    }
}

// Sự kiện tải trang
document.addEventListener("DOMContentLoaded", function () {
    const uploadButton = document.querySelector(".upload-button");

    uploadButton.addEventListener("click", function (event) {
        event.preventDefault();

        const fileInput = document.createElement("input");
        fileInput.type = "file";
        fileInput.accept = ".xlsx, .xls";
        fileInput.style.display = "none";

        fileInput.addEventListener("change", function (e) {
            if (e.target.files.length > 0) {
                const file = e.target.files[0];
                console.log("File đã chọn:", file.name);
                handleFileUpload(file);
            }
        });

        fileInput.click();
    });
});