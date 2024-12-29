function getCurrentDate() {
  const now = new Date();
  const date = String(now.getDate()).padStart(2, "0");
  const month = String(now.getMonth() + 1).padStart(2, "0"); // +1 vì tháng bắt đầu từ 0
  const year = now.getFullYear();
  return `${date}/${month}/${year}`;
}
function formatTime(timeStr) {
  if (!timeStr) return "";

  console.log("Formatting time value:", timeStr, "Type:", typeof timeStr);

  // Handle Date objects from Excel
  if (timeStr instanceof Date) {
    const hours = timeStr.getHours();
    const minutes = timeStr.getMinutes();
    return `${String(hours).padStart(2, "0")}:${String(minutes).padStart(
      2,
      "0"
    )}`;
  }

  // Handle Excel time values (numbers between 0 and 1)
  if (typeof timeStr === "number" || !isNaN(timeStr)) {
    const floatTime = parseFloat(timeStr);
    if (floatTime >= 0 && floatTime <= 1) {
      const totalMinutes = Math.round(floatTime * 24 * 60);
      const hours = Math.floor(totalMinutes / 60);
      const minutes = totalMinutes % 60;
      return `${String(hours).padStart(2, "0")}:${String(minutes).padStart(
        2,
        "0"
      )}`;
    }
  }

  // Handle string format
  if (typeof timeStr === "string") {
    const normalizedTime = timeStr
      .toLowerCase()
      .trim()
      .replace(/[^0-9h:\.]/g, "")
      .replace(/\s+/g, "");

    const timeFormats = {
      colon: /^(\d{1,2}):(\d{2})$/, // 13:30
      hourMinute: /^(\d{1,2})h(\d{2})$/, // 13h30
      decimal: /^(\d{1,2})\.(\d{2})$/, // 13.30
      simple: /^(\d{1,2})(\d{2})$/, // 1330
    };

    for (const [format, regex] of Object.entries(timeFormats)) {
      const match = normalizedTime.match(regex);
      if (match) {
        const [_, hours, minutes] = match;
        const hrs = parseInt(hours, 10);
        const mins = parseInt(minutes, 10);

        if (hrs >= 0 && hrs < 24 && mins >= 0 && mins < 60) {
          return `${String(hrs).padStart(2, "0")}:${String(mins).padStart(
            2,
            "0"
          )}`;
        }
      }
    }
  }

  return "";
}

// Cập nhật hàm isTimeInRange để xử lý giây
function isTimeInRange(currentTime, startTime, endTime) {
  const current = timeToMinutes(currentTime);
  // Thêm :00 cho giây
  const start = timeToMinutes(`${startTime}:00`);
  const end = timeToMinutes(`${endTime}:00`);
  return current >= start && current <= end;
}

function formatDayOfWeek(day) {
  if (!day) return "";

  const dayMap = {
    2: "Thứ Hai",
    3: "Thứ Ba",
    4: "Thứ Tư",
    5: "Thứ Năm",
    6: "Thứ Sáu",
    7: "Thứ Bảy",
    CN: "Chủ Nhật",
    "THỨ 2": "Thứ Hai",
    "THỨ 3": "Thứ Ba",
    "THỨ 4": "Thứ Tư",
    "THỨ 5": "Thứ Năm",
    "THỨ 6": "Thứ Sáu",
    "THỨ 7": "Thứ Bảy",
    "CHỦ NHẬT": "Chủ Nhật",
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
    "PHÒNG LAVENDER 2": "Phòng Lavender 2",
  };

  const normalizedRoom = String(room).trim().toUpperCase();
  return roomMap[normalizedRoom] || room;
}

// Hàm format thời gian sử dụng
function formatDuration(duration) {
  if (!duration) return "";

  console.log("Formatting duration value:", duration, "Type:", typeof duration);

  // Handle Date objects from Excel
  if (duration instanceof Date) {
    const hours = duration.getHours();
    const minutes = duration.getMinutes();
    return `${String(hours).padStart(2, "0")}:${String(minutes).padStart(
      2,
      "0"
    )}`;
  }

  // Handle string format "HH:MM"
  if (typeof duration === "string") {
    const match = duration.trim().match(/^(\d{1,2}):(\d{2})$/);
    if (match) {
      const [_, hours, minutes] = match;
      return `${hours.padStart(2, "0")}:${minutes.padStart(2, "0")}`;
    }
  }

  // Handle numeric values (Excel time)
  if (typeof duration === "number" || !isNaN(duration)) {
    const floatDuration = parseFloat(duration);
    if (floatDuration > 0) {
      const totalMinutes = Math.round(floatDuration * 24 * 60);
      const hours = Math.floor(totalMinutes / 60);
      const minutes = totalMinutes % 60;
      return `${String(hours).padStart(2, "0")}:${String(minutes).padStart(
        2,
        "0"
      )}`;
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
  if (contentLower.includes("phỏng vấn") || contentLower.includes("pv"))
    return "Phỏng vấn";

  return "Khác";
}

function processExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = async function (e) {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, {
          type: "array",
          cellDates: true,
          dateNF: "dd/mm/yyyy",
        });

        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const rawData = XLSX.utils.sheet_to_json(firstSheet, {
          raw: true,
          defval: "",
          header: 1,
        });

        // Tìm và xử lý header
        const headerRowIndex = rawData.findIndex((row) =>
          row.some((cell) =>
            String(cell)
              .toLowerCase()
              .match(/giờ|thời gian|start|end|duration/i)
          )
        );

        if (headerRowIndex === -1) {
          console.warn("Warning: Header row not found");
          return reject(new Error("Cannot find header row"));
        }

        // Get header row and find column indices
        const headers = rawData[headerRowIndex].map((h) =>
          String(h).toLowerCase().trim()
        );
        console.log("Headers found:", headers);

        // More flexible column matching
        const columnIndices = {
          startTime: headers.findIndex(
            (h) =>
              h.includes("GIỜ BẮT ĐẦU") ||
              h.includes("start") ||
              h.includes("bắt đầu") ||
              h === "start time"
          ),
          endTime: headers.findIndex(
            (h) =>
              h.includes("GIỜ KẾT THÚC") ||
              h.includes("end") ||
              h.includes("kết thúc") ||
              h === "end time"
          ),
          duration: headers.findIndex(
            (h) =>
              h.includes("THỜI GIAN SỬ DỤNG") ||
              h.includes("duration") ||
              h.includes("thời gian") ||
              h === "duration time"
          ),
        };

        console.log("Column indices found:", columnIndices);

        // Validate column indices
        if (
          columnIndices.startTime === -1 ||
          columnIndices.endTime === -1 ||
          columnIndices.duration === -1
        ) {
          console.warn("Warning: Some columns not found", columnIndices);
        }

        const meetings = [];
        for (let i = headerRowIndex + 1; i < rawData.length; i++) {
          const row = rawData[i];
          if (!row.some((cell) => cell)) continue; // Skip empty rows

          // Log raw values for debugging
          console.log(`Processing row ${i}:`, {
            rawStartTime: row[columnIndices.startTime],
            rawEndTime: row[columnIndices.endTime],
            rawDuration: row[columnIndices.duration],
          });

          // Extract time values with fallback to specific columns if needed
          const startTimeValue = row[columnIndices.startTime] || row[3]; // Fallback to column D
          const endTimeValue = row[columnIndices.endTime] || row[4]; // Fallback to column E
          const durationValue = row[columnIndices.duration] || row[5]; // Fallback to column F

          const meeting = {
            id: meetings.length + 1,
            date: formatDate(row[0]),
            dayOfWeek: formatDayOfWeek(row[1]),
            room: formatRoomName(row[2]),
            startTime: formatTime(startTimeValue),
            endTime: formatTime(endTimeValue),
            duration: formatDuration(durationValue),
            content: row[7] || "",
            purpose: determinePurpose(row[7]),
          };

          console.log(`Processed meeting data:`, meeting);
          meetings.push(meeting);
        }
        const conflicts = await validateMeetings(meetings);

        if (conflicts.length > 0) {
          let errorMessage = "Phát hiện xung đột trong lịch họp:\n\n";
          conflicts.forEach((conflict) => {
            errorMessage += `${conflict.message}\n\n`;
          });

          showErrorModal(errorMessage);
          reject(new Error("CONFLICT_ERROR"));
          return;
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

function formatDate(dateInput) {
  console.log("formatDate input:", dateInput, "type:", typeof dateInput);

  if (!dateInput) return "";

  try {
    // Xử lý Date object từ Excel (do cellDates: true)
    if (dateInput instanceof Date) {
      if (!isNaN(dateInput.getTime())) {
        const day = dateInput.getDate() + 1;
        const month = dateInput.getMonth() + 1;
        const year = dateInput.getFullYear();
        return `${String(day).padStart(2, "0")}/${String(month).padStart(
          2,
          "0"
        )}/${year}`;
      }
    }

    // Xử lý chuỗi ngày đã được format sẵn dd/mm/yyyy
    if (typeof dateInput === "string") {
      const dateStr = dateInput.trim();
      const match = dateStr.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
      if (match) {
        const [_, day, month, year] = match;
        return `${String(day).padStart(2, "0")}/${String(month).padStart(
          2,
          "0"
        )}/${year}`;
      }
    }

    // Xử lý số serial từ Excel (nếu không dùng cellDates: true)
    if (typeof dateInput === "number" || !isNaN(Number(dateInput))) {
      const numDate = Number(dateInput);
      // Excel bắt đầu từ 1/1/1900, và trừ đi 2 để điều chỉnh lỗi năm nhuận
      const excelEpoch = new Date(1900, 0, -1);
      const offsetDays = numDate - 1;
      const resultDate = new Date(excelEpoch);
      resultDate.setDate(resultDate.getDate() + offsetDays);

      if (!isNaN(resultDate.getTime())) {
        const day = resultDate.getDate();
        const month = resultDate.getMonth() + 1;
        const year = resultDate.getFullYear();
        return `${String(day).padStart(2, "0")}/${String(month).padStart(
          2,
          "0"
        )}/${year}`;
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
//==========Function Update info from Excel file to MeetingInfo Section========
function updateRoomStatus(data) {
  console.log("Updating room status with data at:", getCurrentTime());

  const currentDate = getCurrentDate();
  const currentTime = getCurrentTime();

  console.log("Current date:", currentDate);
  console.log("Current time:", currentTime);

  const todayMeetings = data.filter((meeting) => {
    const isToday = meeting.date === currentDate;
    console.log(`Meeting date: ${meeting.date}, Is today: ${isToday}`);
    return isToday;
  });

  console.log("Today's meetings:", todayMeetings);

  const roomsToUpdate = ["Phòng Lavender 1", "Phòng Lavender 2", "Phòng Lotus"];
  roomsToUpdate.forEach((roomName) => {
    updateSingleRoomStatus(roomName, todayMeetings, currentTime);
  });
}

function normalizeRoomName(roomName) {
  // Loại bỏ "P. " và chuẩn hóa tên phòng
  return roomName
    .replace(/^(P\.|Phòng)\s*/i, "")
    .trim()
    .toLowerCase();
}

//===New version : Update thểm cả giây vì nếu so sánh mỗi phút thì sẽ sau 1 phút thì mới nhảy kết quả
function getCurrentTime() {
  const now = new Date();
  return `${String(now.getHours()).padStart(2, "0")}:${String(
    now.getMinutes()
  ).padStart(2, "0")}:${String(now.getSeconds()).padStart(2, "0")}`;
}

// Sửa hàm isTimeOverdue để có độ chính xác cao hơn
function isTimeOverdue(endTime, currentTime) {
  const endTimeParts = endTime.split(":");
  const endTimeWithSeconds = `${endTimeParts[0]}:${endTimeParts[1]}:00`;
  const isOverdue =
    timeToMinutes(currentTime) > timeToMinutes(endTimeWithSeconds);

  if (isOverdue) {
    console.log(
      `Meeting overdue check at ${currentTime} for end time ${endTime}`
    );
  }

  return isOverdue;
}

// Hàm để tự động cập nhật thời gian và trạng thái
function startAutoUpdate(data) {
  updateRoomStatus(data);

  const intervalId = setInterval(() => {
    const currentTime = getCurrentTime();
    // Chỉ cập nhật khi thay đổi phút
    if (currentTime.endsWith(":00")) {
      console.log("Auto updating at:", currentTime);
      updateRoomStatus(data);
    }
  }, 1000); // Vẫn kiểm tra mỗi giây nhưng chỉ cập nhật khi đổi phút

  window.autoUpdateInterval = intervalId;
  return () => clearInterval(intervalId);
}
// Thêm biến để lưu trữ trạng thái trước đó
let previousStates = {};

function updateSingleRoomStatus(roomCode, meetings, currentTime) {
  console.log("Updating room status for:", roomCode);

  const roomSections = document.querySelectorAll(".room-section");
  const roomSection = Array.from(roomSections).find(
    (section) =>
      normalizeRoomName(section.querySelector(".room-number").textContent) ===
      normalizeRoomName(roomCode)
  );

  if (!roomSection) {
    console.warn(`No room section found for room code: ${roomCode}`);
    return;
  }

  const titleElement = roomSection.querySelector(".meeting-title");
  const startTimeElement = roomSection.querySelector(".start-time");
  const endTimeElement = roomSection.querySelector(".end-time");
  const statusIndicator = roomSection.querySelector(
    ".status-indicator .status-text"
  );
  const indicatorDot = roomSection.querySelector(
    ".status-indicator .indicator-dot"
  );

  // Lọc các cuộc họp cho phòng hiện tại
  const roomMeetings = meetings.filter(
    (meeting) => normalizeRoomName(meeting.room) === normalizeRoomName(roomCode)
  );

  const upcomingMeetings = roomMeetings.filter(
    (meeting) => !isTimeOverdue(meeting.endTime, currentTime)
  );

  const activeMeeting = upcomingMeetings.find((meeting) =>
    isTimeInRange(currentTime, meeting.startTime, meeting.endTime)
  );

  // Tạo key duy nhất cho phòng
  const roomKey = roomCode;

  // Tạo trạng thái mới
  const newState = {
    isActive: !!activeMeeting,
    content: activeMeeting
      ? activeMeeting.content || activeMeeting.purpose
      : upcomingMeetings.length > 0
      ? upcomingMeetings.map((m) => m.content || m.purpose).join(" | ")
      : "Trống",
    startTime: activeMeeting
      ? activeMeeting.startTime
      : upcomingMeetings.length > 0
      ? upcomingMeetings.map((m) => m.startTime).join(" | ")
      : "--:--",
    endTime: activeMeeting
      ? activeMeeting.endTime
      : upcomingMeetings.length > 0
      ? upcomingMeetings.map((m) => m.endTime).join(" | ")
      : "--:--",
  };

  // So sánh với trạng thái trước đó
  if (
    !previousStates[roomKey] ||
    JSON.stringify(previousStates[roomKey]) !== JSON.stringify(newState)
  ) {
    // Chỉ cập nhật DOM nếu có sự thay đổi
    if (activeMeeting) {
      titleElement.innerHTML = `<span>Thông tin cuộc họp:</span> ${newState.content}`;
      startTimeElement.innerHTML = `<span>Thời gian bắt đầu:</span> ${newState.startTime}`;
      endTimeElement.innerHTML = `<span>Thời gian kết thúc:</span> ${newState.endTime}`;
      statusIndicator.textContent = "Đang họp";
      indicatorDot.classList.remove("available");
      indicatorDot.classList.add("busy");
    } else {
      const nextThreeMeetings = upcomingMeetings
        .sort((a, b) => timeToMinutes(a.startTime) - timeToMinutes(b.startTime))
        .slice(0, 3);

      if (nextThreeMeetings.length > 0) {
        titleElement.innerHTML = `<span>Thông tin cuộc họp:</span> ${newState.content}`;
        startTimeElement.innerHTML = `<span>Thời gian bắt đầu:</span> ${newState.startTime}`;
        endTimeElement.innerHTML = `<span>Thời gian kết thúc:</span> ${newState.endTime}`;
      } else {
        titleElement.innerHTML = `<span>Thông tin cuộc họp:</span> Trống`;
        startTimeElement.innerHTML = `<span>Thời gian bắt đầu:</span> --:--`;
        endTimeElement.innerHTML = `<span>Thời gian kết thúc:</span> --:--`;
      }
      statusIndicator.textContent = "Trống";
      indicatorDot.classList.remove("busy");
      indicatorDot.classList.add("available");
    }

    // Lưu trạng thái mới
    previousStates[roomKey] = newState;
  }
}

if (!Element.prototype.contains) {
  Element.prototype.contains = function (text) {
    return this.textContent.trim().includes(text);
  };
}

// Sửa hàm timeToMinutes để xử lý giây
function timeToMinutes(timeStr) {
  if (!timeStr) return 0;
  const parts = timeStr.split(":");
  const hours = parseInt(parts[0]);
  const minutes = parseInt(parts[1]);
  const seconds = parts.length > 2 ? parseInt(parts[2]) : 0;
  return hours * 3600 + minutes * 60 + seconds;
}

let fileHandle = null;
let lastFileData = null;
let fileCache = {
  data: null,
  lastModified: null,
  reader: new FileReader(),
};

// Hàm kiểm tra thay đổi từ input element
async function checkFileChanges() {
  if (!fileHandle) return;

  try {
    const file = await fileHandle.getFile();
    const fileData = await file.text(); // hoặc arrayBuffer() nếu cần

    // Kiểm tra nếu lastFileData chưa được khởi tạo
    if (lastFileData === null) {
      lastFileData = fileData;
      return;
    }

    // So sánh với dữ liệu cũ
    if (fileData !== lastFileData) {
      console.log("File đã thay đổi, đang cập nhật...");
      const data = await processExcelFile(file);
      updateScheduleTable(data);
      startAutoUpdate(data);
      lastFileData = fileData;

      // Cập nhật cache
      fileCache.data = data;
      fileCache.lastModified = new Date().getTime();

      // Lưu vào localStorage
      try {
        localStorage.setItem(
          "fileCache",
          JSON.stringify({
            data: fileCache.data,
            lastModified: fileCache.lastModified,
          })
        );
      } catch (e) {
        console.error("Không thể lưu vào localStorage:", e);
      }
    }
  } catch (error) {
    console.error("Lỗi khi kiểm tra file:", error);
    // Nếu mất quyền truy cập, dừng checking
    if (error.name === "NotAllowedError") {
      clearInterval(window.fileCheckInterval);
      fileHandle = null;
    }
  }
}

async function handleFileUpload(file) {
  try {
    // Xử lý File System Access API
    try {
      const handles = await window.showOpenFilePicker({
        multiple: false,
        types: [
          {
            description: "Excel Files",
            accept: {
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
                [".xlsx"],
              "application/vnd.ms-excel": [".xls"],
            },
          },
        ],
      });
      fileHandle = handles[0];
      const initialFile = await fileHandle.getFile();
      lastFileData = await initialFile.text();
    } catch (error) {
      console.error("Không thể lấy file handle:", error);
    }

    // Xử lý file và kiểm tra xung đột
    const data = await processExcelFile(file);

    // Nếu không có xung đột, tiếp tục xử lý
    updateScheduleTable(data);
    startAutoUpdate(data);

    // Cập nhật cache
    fileCache.data = data;
    fileCache.lastModified = new Date().getTime();

    // Lưu vào localStorage
    try {
      localStorage.setItem(
        "fileCache",
        JSON.stringify({
          data: fileCache.data,
          lastModified: fileCache.lastModified,
        })
      );
    } catch (e) {
      console.error("Không thể lưu vào localStorage:", e);
    }

    // Thiết lập monitoring
    if (fileHandle) {
      if (window.fileCheckInterval) {
        clearInterval(window.fileCheckInterval);
      }
      window.fileCheckInterval = setInterval(checkFileChanges, 5000);
    }
  } catch (error) {
    console.error("Lỗi xử lý file:", error);
    if (error.message === "CONFLICT_ERROR") {
      // Xung đột đã được xử lý và hiển thị trong modal
      return;
    }
    alert("Lỗi khi xử lý file. Vui lòng thử lại.");
  }
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

// Event listener for file upload
document.addEventListener("DOMContentLoaded", function () {
  const uploadButton = document.querySelector(".upload-button");

  uploadButton.addEventListener("click", async function (event) {
    event.preventDefault();

    try {
      // Thử dùng file handle đã có
      if (fileHandle) {
        const file = await fileHandle.getFile();
        await handleFileUpload(file);
        return;
      }
    } catch (error) {
      console.error("Không thể sử dụng file handle cũ:", error);
      fileHandle = null;
    }

    // Nếu không có file handle hoặc có lỗi, tạo input mới
    const fileInput = document.createElement("input");
    fileInput.type = "file";
    fileInput.accept = ".xlsx, .xls";
    fileInput.style.display = "none";

    fileInput.addEventListener("change", function (e) {
      if (e.target.files.length > 0) {
        const file = e.target.files[0];
        handleFileUpload(file);
      }
    });

    fileInput.click();
  });
});

//========================Update Time ====================
function padZero(num) {
  return num < 10 ? `0${num}` : num;
}

function updateClock() {
  const now = new Date();
  const hours = padZero(now.getHours());
  const minutes = padZero(now.getMinutes());
  const seconds = padZero(now.getSeconds());
  const time = `${hours}:${minutes}:${seconds}`;

  const days = [
    "Chủ Nhật",
    "Thứ Hai",
    "Thứ Ba",
    "Thứ Tư",
    "Thứ Năm",
    "Thứ Sáu",
    "Thứ Bảy",
  ];
  const dayName = days[now.getDay()];
  const date = padZero(now.getDate());
  const month = padZero(now.getMonth() + 1);
  const year = now.getFullYear();
  const dateStr = `${dayName}, ${date}/${month}/${year}`;

  const logoElement = document.querySelector(".logo");
  const currentDateElement = document.querySelector(".current-date");

  if (logoElement) {
    logoElement.innerHTML = `
            <div class="clock-container" style="font-size: 100px; color: #ffffff;justify-content: right;">
                ${time}
            </div>
        `;
  }

  if (currentDateElement) {
    currentDateElement.textContent = dateStr;
    currentDateElement.style.fontSize = "17px"; // Thay đổi kích thước font
    currentDateElement.style.color = "#ffffff"; // Thay đổi màu chữ
    currentDateElement.style.fontWeight = "bold"; // Đậm chữ
    currentDateElement.style.paddingRight = "25px";
  }
}

// Khởi tạo đồng hồ và cập nhật mỗi giây
function initClock() {
  updateClock(); // Cập nhật ngay lập tức
  setInterval(updateClock, 1000); // Cập nhật mỗi giây
}

// Gọi hàm khởi tạo khi trang đã load
document.addEventListener("DOMContentLoaded", initClock);

//==========================New update : Selection when user pick any date=====
document.addEventListener("DOMContentLoaded", function () {
  const datePicker = document.getElementById("meetingDate");

  datePicker.addEventListener("change", function () {
    const selectedDate = new Date(this.value);
    filterMeetingsByDate(selectedDate);
  });
});

function filterMeetingsByDate(selectedDate) {
  const rows = document.querySelectorAll(".schedule-table .table-row");

  rows.forEach((row) => {
    const meetingDateText = row.children[1].textContent; // Cột ngày
    const meetingDate = new Date(
      meetingDateText.split("/").reverse().join("-")
    ); // Chuyển đổi định dạng ngày

    if (meetingDate.toDateString() === selectedDate.toDateString()) {
      row.style.display = ""; // Hiển thị nếu trùng khớp
    } else {
      row.style.display = "none"; // Ẩn nếu không trùng khớp
    }
  });
}
//=======New Update : Kiểm tra thông tin nhập vào từ người dùng - Cảnh báo nếu nhập trùng phòng họp=======
// Hàm kiểm tra xung đột thời gian giữa các cuộc họp
function checkTimeConflict(meeting1, meeting2) {
  const start1 = timeToMinutes(meeting1.startTime);
  const end1 = timeToMinutes(meeting1.endTime);
  const start2 = timeToMinutes(meeting2.startTime);
  const end2 = timeToMinutes(meeting2.endTime);
  return start1 < end2 && start2 < end1;
}

// Hàm kiểm tra xung đột lịch họp
async function validateMeetings(meetings) {
  const conflicts = [];
  const processedMeetings = new Set();

  for (let i = 0; i < meetings.length; i++) {
    const currentMeeting = meetings[i];
    const key = `${currentMeeting.date}_${currentMeeting.room}`;

    // Kiểm tra với các cuộc họp khác cùng ngày và cùng phòng
    for (let j = 0; j < meetings.length; j++) {
      if (i === j) continue;
      const otherMeeting = meetings[j];

      if (
        currentMeeting.date === otherMeeting.date &&
        normalizeRoomName(currentMeeting.room) ===
          normalizeRoomName(otherMeeting.room)
      ) {
        if (checkTimeConflict(currentMeeting, otherMeeting)) {
          const conflictKey = [i, j].sort().join("_");
          if (!processedMeetings.has(conflictKey)) {
            conflicts.push({
              meeting1: currentMeeting,
              meeting2: otherMeeting,
              message:
                `Xung đột lịch họp tại phòng ${currentMeeting.room} ngày ${currentMeeting.date}:\n` +
                `- Cuộc họp 1: "${
                  currentMeeting.content || currentMeeting.purpose
                }" (${currentMeeting.startTime} - ${
                  currentMeeting.endTime
                })\n` +
                `- Cuộc họp 2: "${
                  otherMeeting.content || otherMeeting.purpose
                }" (${otherMeeting.startTime} - ${otherMeeting.endTime})`,
            });
            processedMeetings.add(conflictKey);
          }
        }
      }
    }
  }

  return conflicts;
}

function checkTimeConflict(meeting1, meeting2) {
  const start1 = timeToMinutes(meeting1.startTime);
  const end1 = timeToMinutes(meeting1.endTime);
  const start2 = timeToMinutes(meeting2.startTime);
  const end2 = timeToMinutes(meeting2.endTime);

  // Kiểm tra xem hai khoảng thời gian có giao nhau không
  return start1 < end2 && start2 < end1;
}

// Hàm kiểm tra xung đột cho một cuộc họp mới
function validateNewMeeting(newMeeting, existingMeetings) {
  const conflicts = [];

  // Chỉ kiểm tra các cuộc họp cùng ngày và cùng phòng
  const relevantMeetings = existingMeetings.filter(
    (meeting) =>
      meeting.date === newMeeting.date &&
      normalizeRoomName(meeting.room) === normalizeRoomName(newMeeting.room)
  );

  for (const existingMeeting of relevantMeetings) {
    if (checkTimeConflict(newMeeting, existingMeeting)) {
      conflicts.push({
        conflictWith: existingMeeting,
        type: "TIME_OVERLAP",
        message: `Xung đột với cuộc họp "${
          existingMeeting.content || existingMeeting.purpose
        }" 
                 từ ${existingMeeting.startTime} đến ${
          existingMeeting.endTime
        }`,
      });
    }
  }

  return {
    isValid: conflicts.length === 0,
    conflicts,
  };
}

// Hàm hiển thị modal thông báo lỗi
function showErrorModal(message) {
  // Tạo modal container
  const modalContainer = document.createElement("div");
  modalContainer.className = "error-modal-container";
  modalContainer.style.cssText = `
    position: fixed;
    top: 0;
    left: 0;
    width: 100%;
    height: 100%;
    background-color: rgba(0, 0, 0, 0.5);
    display: flex;
    justify-content: center;
    align-items: center;
    z-index: 1000;
  `;

  // Tạo modal content
  const modalContent = document.createElement("div");
  modalContent.className = "error-modal-content";
  modalContent.style.cssText = `
    background-color: white;
    padding: 20px;
    border-radius: 8px;
    max-width: 80%;
    max-height: 80%;
    overflow-y: auto;
    position: relative;
  `;

  // Tạo tiêu đề
  const title = document.createElement("h3");
  title.textContent = "Lỗi Xung Đột Lịch Họp";
  title.style.color = "#dc3545";

  // Tạo nội dung
  const content = document.createElement("pre");
  content.textContent = message;
  content.style.whiteSpace = "pre-wrap";
  content.style.marginTop = "10px";

  // Tạo nút đóng
  const closeButton = document.createElement("button");
  closeButton.textContent = "Đóng";
  closeButton.style.cssText = `
    margin-top: 15px;
    padding: 8px 16px;
    background-color: #dc3545;
    color: white;
    border: none;
    border-radius: 4px;
    cursor: pointer;
  `;
  closeButton.onclick = () => modalContainer.remove();

  // Ghép các phần tử
  modalContent.appendChild(title);
  modalContent.appendChild(content);
  modalContent.appendChild(closeButton);
  modalContainer.appendChild(modalContent);
  document.body.appendChild(modalContainer);
}

//===================E-Ra Services=============================
const eraWidget = new EraWidget();
const currentId = "Current"; // ID chuỗi cho dòng điện
const voltageId = "Voltage"; // ID chuỗi cho điện áp
const pwId = "Power"; // ID số cho công suất tiêu thụ

let configCurrent = null,
  configVol = null,
  configPower = null;

eraWidget.init({
  onConfiguration: (configuration) => {
    // Lưu các cấu hình khi nhận được từ widget
    configCurrent = configuration.realtime_configs[0]; // Lưu cấu hình dòng điện
    configVol = configuration.realtime_configs[1]; // Lưu cấu hình điện áp
    configPower = configuration.realtime_configs[2]; // Lưu cấu hình power
    actions = configuration.actions; // Lưu các hành động điều khiển
  },

  // Hàm lấy giá trị từ các ID và cập nhật giao diện
  onValues: (values) => {
    if (configVol && values[configVol.id]) {
      const voltageVal = values[configVol.id].value;
      document.getElementById("voltageId").textContent = voltageVal;
    }

    if (configCurrent && values[configCurrent.id]) {
      const currentVal = values[configCurrent.id].value;
      document.getElementById("currentId").textContent = currentVal; // Cập nhật giá trị nhiệt độ
    }
    
    if (configPower && values[configPower.id]) {
      const powerVal = values[configPower.id].value;
      document.getElementById("power").textContent = powerVal; // Cập nhật giá trị công suất tiêu thụ
    } else {
      console.error("Không tìm thấy cấu hình hoặc giá trị cho Power.");
    }
  },
});
