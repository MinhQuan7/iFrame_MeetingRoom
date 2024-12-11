function getCurrentTime() {
  const now = new Date();
  return `${String(now.getHours()).padStart(2, "0")}:${String(
    now.getMinutes()
  ).padStart(2, "0")}`;
}

function formatTime(timeStr) {
  if (!timeStr) return "";

  console.log("Formatting time value:", timeStr, "Type:", typeof timeStr);

  // Handle Excel time values (numbers between 0 and 1)
  if (typeof timeStr === 'number' || !isNaN(timeStr)) {
    const floatTime = parseFloat(timeStr);
    if (floatTime >= 0 && floatTime <= 1) {
      const totalMinutes = Math.round(floatTime * 24 * 60);
      const hours = Math.floor(totalMinutes / 60);
      const minutes = totalMinutes % 60;
      return `${String(hours).padStart(2, "0")}:${String(minutes).padStart(2, "0")}`;
    }
  }

  // Convert to string and normalize
  const normalizedTime = String(timeStr).toLowerCase().trim()
    .replace(/[^0-9h:\.]/g, '') // Remove all characters except numbers, h, : and .
    .replace(/\s+/g, '');       // Remove all whitespace

  // Handle various time formats
  const timeFormats = {
    colon: /^(\d{1,2}):(\d{2})$/,         // 13:30
    hourMinute: /^(\d{1,2})h(\d{2})$/,    // 13h30
    decimal: /^(\d{1,2})\.(\d{2})$/,      // 13.30
    simple: /^(\d{1,2})(\d{2})$/          // 1330
  };

  for (const [format, regex] of Object.entries(timeFormats)) {
    const match = normalizedTime.match(regex);
    if (match) {
      const [_, hours, minutes] = match;
      const hrs = parseInt(hours, 10);
      const mins = parseInt(minutes, 10);
      
      if (hrs >= 0 && hrs < 24 && mins >= 0 && mins < 60) {
        return `${String(hrs).padStart(2, "0")}:${String(mins).padStart(2, "0")}`;
      }
    }
  }

  return "";
}
function isTimeInRange(current, start, end) {
  // So sánh thời gian dạng HH:MM
  const [currentHour, currentMin] = current.split(":").map(Number);
  const [startHour, startMin] = start.split(":").map(Number);
  const [endHour, endMin] = end.split(":").map(Number);

  const currentTime = currentHour * 60 + currentMin;
  const startTime = startHour * 60 + startMin;
  const endTime = endHour * 60 + endMin;

  return currentTime >= startTime && currentTime <= endTime;
}

function formatDayOfWeek(day) {
  if (!day) return "";
  
  const dayMap = {
    "2": "Thứ Hai",
    "3": "Thứ Ba",
    "4": "Thứ Tư",
    "5": "Thứ Năm",
    "6": "Thứ Sáu",
    "7": "Thứ Bảy",
    "CN": "Chủ Nhật",
    "THỨ 2": "Thứ Hai",
    "THỨ 3": "Thứ Ba",
    "THỨ 4": "Thứ Tư",
    "THỨ 5": "Thứ Năm",
    "THỨ 6": "Thứ Sáu",
    "THỨ 7": "Thứ Bảy",
    "CHỦ NHẬT": "Chủ Nhật"
  };

  const normalizedDay = String(day).trim().toUpperCase();
  return dayMap[normalizedDay] || day;
}

// Hàm format tên phòng
function formatRoomName(room) {
  if (!room) return "";

  const roomMap = {
    "PHÒNG LOTUS": "Phòng Lotus",
    "P.LOTUS": "Phòng Lotus",
    "P.LAVENDER 1": "Phòng Lavender 1",
    "PHÒNG LAVENDER 1": "Phòng Lavender 1",
    "P.LAVENDER 2": "Phòng Lavender 2", 
    "PHÒNG LAVENDER 2": "Phòng Lavender 2"
  };

  const normalizedRoom = String(room).trim().toUpperCase();
  return roomMap[normalizedRoom] || room;
}


// Hàm format thời gian sử dụng
function formatDuration(duration) {
  if (!duration) return "";

  console.log("Formatting duration value:", duration, "Type:", typeof duration);

  // Handle string format "HH:MM"
  if (typeof duration === 'string') {
    const match = duration.trim().match(/^(\d{1,2}):(\d{2})$/);
    if (match) {
      const [_, hours, minutes] = match;
      return `${hours.padStart(2, '0')}:${minutes.padStart(2, '0')}`;
    }
  }

  // Handle numeric values (Excel time)
  if (typeof duration === 'number' || !isNaN(duration)) {
    const floatDuration = parseFloat(duration);
    if (floatDuration > 0) {
      const totalMinutes = Math.round(floatDuration * 24 * 60);
      const hours = Math.floor(totalMinutes / 60);
      const minutes = totalMinutes % 60;
      return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}`;
    }
  }

  return "";
}

// Hàm xác định mục đích sử dụng
function determinePurpose(content) {
  if (!content) return "Khác";

  const contentLower = String(content).toLowerCase();
  
  if (contentLower.includes("họp")) return "Họp";
  if (contentLower.includes("đào tạo")) return "Đào tạo";
  if (contentLower.includes("phỏng vấn") || contentLower.includes("pv")) return "Phỏng vấn";
  
  return "Khác";
}

function processExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = function(e) {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { 
          type: "array",
          cellDates: true,
          dateNF: 'dd/mm/yyyy'
        });

        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const rawData = XLSX.utils.sheet_to_json(firstSheet, {
          raw: true,
          defval: "",
          header: 1
        });

        // Debug: Print raw data
        console.log("Raw Excel Data:", rawData);

        // Find header row with more flexible matching
        const headerRowIndex = rawData.findIndex(row => 
          row.some(cell => 
            String(cell).toLowerCase().match(/giờ|thời gian|start|end|duration/i)
          )
        );

        if (headerRowIndex === -1) {
          console.warn("Warning: Header row not found");
          return reject(new Error("Cannot find header row"));
        }

        // Get header row and find column indices
        const headers = rawData[headerRowIndex].map(h => String(h).toLowerCase().trim());
        console.log("Headers found:", headers);

        // More flexible column matching
        const columnIndices = {
          startTime: headers.findIndex(h => 
            h.includes('giờ bắt đầu') || 
            h.includes('start') || 
            h.includes('bắt đầu') ||
            h === 'start time'
          ),
          endTime: headers.findIndex(h => 
            h.includes('giờ kết thúc') || 
            h.includes('end') || 
            h.includes('kết thúc') ||
            h === 'end time'
          ),
          duration: headers.findIndex(h => 
            h.includes('thời gian sử dụng') || 
            h.includes('duration') || 
            h.includes('thời gian') ||
            h === 'duration time'
          )
        };

        console.log("Column indices found:", columnIndices);

        // Validate column indices
        if (columnIndices.startTime === -1 || columnIndices.endTime === -1 || columnIndices.duration === -1) {
          console.warn("Warning: Some columns not found", columnIndices);
        }

        const meetings = [];
        for (let i = headerRowIndex + 1; i < rawData.length; i++) {
          const row = rawData[i];
          if (!row.some(cell => cell)) continue; // Skip empty rows

          // Log raw values for debugging
          console.log(`Processing row ${i}:`, {
            rawStartTime: row[columnIndices.startTime],
            rawEndTime: row[columnIndices.endTime],
            rawDuration: row[columnIndices.duration]
          });

          // Extract time values with fallback to specific columns if needed
          const startTimeValue = row[columnIndices.startTime] || row[3]; // Fallback to column D
          const endTimeValue = row[columnIndices.endTime] || row[4];     // Fallback to column E
          const durationValue = row[columnIndices.duration] || row[5];    // Fallback to column F

          const meeting = {
            id: meetings.length + 1,
            date: formatDate(row[0]),
            dayOfWeek: formatDayOfWeek(row[1]),
            room: formatRoomName(row[2]),
            startTime: formatTime(startTimeValue),
            endTime: formatTime(endTimeValue),
            duration: formatDuration(durationValue),
            content: row[7] || "",
            purpose: determinePurpose(row[7])
          };

          console.log(`Processed meeting data:`, meeting);
          meetings.push(meeting);
        }

        resolve(meetings);
      } catch (error) {
        console.error("Error processing file:", error);
        reject(error);
      }
    };

    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

// Điều chỉnh hàm formatDate để xử lý nhiều kiểu dữ liệu hơn
function formatDate(dateInput) {
  console.log("formatDate input:", dateInput, "type:", typeof dateInput);

  if (!dateInput) return "";

  try {
    // Xử lý Date object
    if (dateInput instanceof Date) {
      // Kiểm tra date hợp lệ
      if (!isNaN(dateInput.getTime())) {
        return `${String(dateInput.getDate()).padStart(2, "0")}/${String(
          dateInput.getMonth() + 1
        ).padStart(2, "0")}/${dateInput.getFullYear()}`;
      }
    }

    // Chuyển đổi string date thành Date object nếu có thể
    if (typeof dateInput === 'string' && dateInput.includes('GMT')) {
      const date = new Date(dateInput);
      if (!isNaN(date.getTime())) {
        return `${String(date.getDate()).padStart(2, "0")}/${String(
          date.getMonth() + 1
        ).padStart(2, "0")}/${date.getFullYear()}`;
      }
    }

    // Xử lý số serial từ Excel
    if (typeof dateInput === 'number' || !isNaN(Number(dateInput))) {
      const numDate = Number(dateInput);
      const excelDate = new Date(Date.UTC(1900, 0, numDate - 1));
      if (!isNaN(excelDate.getTime())) {
        return `${String(excelDate.getUTCDate()).padStart(2, "0")}/${String(
          excelDate.getUTCMonth() + 1
        ).padStart(2, "0")}/${excelDate.getUTCFullYear()}`;
      }
    }

    // Xử lý chuỗi ngày thông thường
    const dateStr = String(dateInput).trim();
    const formats = [
      /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/, // dd/mm/yyyy
      /^(\d{1,2})-(\d{1,2})-(\d{4})$/, // dd-mm-yyyy
      /^(\d{4})-(\d{1,2})-(\d{1,2})$/, // yyyy-mm-dd
    ];

    for (let format of formats) {
      const match = dateStr.match(format);
      if (match) {
        const [_, part1, part2, part3] = match;
        if (format === formats[2]) {
          return `${part3.padStart(2, "0")}/${part2.padStart(2, "0")}/${part1}`;
        }
        return `${part1.padStart(2, "0")}/${part2.padStart(2, "0")}/${part3}`;
      }
    }

    console.log("Could not parse date:", dateInput);
    return "";
  } catch (error) {
    console.error("Error in formatDate:", error);
    return "";
  }
}

function parseMeetingInfo(cellContent) {
  if (!cellContent) return { purpose: "", content: "" };

  const lines = cellContent.split("\r\n");
  const content = lines[0];
  let purpose = "";

  // Extract purpose from common patterns
  if (content.toLowerCase().includes("họp")) {
    purpose = "Họp";
  } else if (content.toLowerCase().includes("đào tạo")) {
    purpose = "Đào tạo";
  } else if (content.toLowerCase().includes("pv")) {
    purpose = "Phỏng vấn";
  } else {
    purpose = "Khác";
  }

  return {
    purpose,
    content,
  };
}

function calculateEndTime(startTime) {
  if (!startTime) return "";

  // Convert time format (e.g., "7H30" to "8:00")
  const time = startTime.replace("H", ":").replace("h", ":");
  const [hours, minutes] = time.split(":").map(Number);

  // Add 30 minutes for default meeting duration
  let endHours = hours;
  let endMinutes = minutes + 30;

  if (endMinutes >= 60) {
    endHours += 1;
    endMinutes -= 60;
  }

  return `${String(endHours).padStart(2, "0")}:${String(endMinutes).padStart(
    2,
    "0"
  )}`;
}

function calculateDuration(startTime, endTime) {
  if (!startTime || !endTime) return "";

  const start = startTime.replace("H", ":").replace("h", ":");
  const [startHours, startMinutes] = start.split(":").map(Number);
  const [endHours, endMinutes] = endTime.split(":").map(Number);

  const durationMinutes =
    endHours * 60 + endMinutes - (startHours * 60 + startMinutes);
  const hours = Math.floor(durationMinutes / 60);
  const minutes = durationMinutes % 60;

  return `${hours}:${String(minutes).padStart(2, "0")}`;
}

// Cập nhật bảng lịch
function updateScheduleTable(data) {
  const tableBody = document.querySelector(".schedule-table");
  const headerRow = tableBody.querySelector(".table-header");

  // Xóa các hàng cũ
  Array.from(tableBody.children)
    .filter((child) => child !== headerRow)
    .forEach((child) => child.remove());

  // Thêm dữ liệu mới
  data.forEach((meeting) => {
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

  const todayMeetings = data.filter((meeting) => meeting.date === currentDate);

  Object.entries(roomMapping).forEach(([fullName, shortName]) => {
    const roomMeeting = todayMeetings.find(
      (meeting) => meeting.room === fullName
    );
    updateSingleRoomStatus(shortName, roomMeeting, currentTime);
  });
}

// Cập nhật trạng thái từng phòng
function updateSingleRoomStatus(roomCode, meeting, currentTime) {
  // Find the room section by iterating through room sections
  const roomSections = document.querySelectorAll(".room-section");
  const roomSection = Array.from(roomSections).find((section) => {
    const roomNumberElement = section.querySelector(".room-number");
    return (
      roomNumberElement && roomNumberElement.textContent.trim() === roomCode
    );
  });

  if (!roomSection) return;

  const titleElement = roomSection.querySelector(".meeting-title");
  const startTimeElement = roomSection.querySelector(".start-time");
  const endTimeElement = roomSection.querySelector(".end-time");
  const statusIndicator = roomSection.querySelector(".status-indicator");

  if (
    meeting &&
    isTimeInRange(currentTime, meeting.startTime, meeting.endTime)
  ) {
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
    .then((data) => {
      // Cập nhật bảng lịch
      updateScheduleTable(data);

      // Cập nhật trạng thái phòng
      updateRoomStatus(data);

      console.log("Xử lý file thành công:", data);
    })
    .catch((error) => {
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
