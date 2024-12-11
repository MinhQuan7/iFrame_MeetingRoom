// Các hàm tiện ích
function formatDate(dateStr) {
  if (!dateStr) return "";
  
  // Chuyển đổi thành chuỗi và làm sạch dữ liệu
  const cleanStr = String(dateStr).trim();
  
  // Xử lý ngày từ Excel (số serial)
  if (!isNaN(Number(cleanStr))) {
    const serialDate = Number(cleanStr);
    // Excel bắt đầu từ 1/1/1900, trừ 1 vì JS bắt đầu từ 0
    const jsDate = new Date(Date.UTC(1900, 0, serialDate - 1));
    
    // Kiểm tra tính hợp lệ của ngày
    if (!isNaN(jsDate.getTime())) {
      return `${String(jsDate.getUTCDate()).padStart(2, "0")}/${String(
        jsDate.getUTCMonth() + 1
      ).padStart(2, "0")}/${jsDate.getUTCFullYear()}`;
    }
  }

  // Các định dạng ngày thông thường
  const formats = [
    /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/, // dd/mm/yyyy
    /^(\d{1,2})-(\d{1,2})-(\d{4})$/, // dd-mm-yyyy
    /^(\d{4})-(\d{1,2})-(\d{1,2})$/, // yyyy-mm-dd
  ];

  for (let format of formats) {
    const match = cleanStr.match(format);
    if (match) {
      const parts = match.slice(1); // Lấy các nhóm matched
      if (format === formats[2]) {
        // Nếu là yyyy-mm-dd
        const [year, month, day] = parts;
        return `${day.padStart(2, "0")}/${month.padStart(2, "0")}/${year}`;
      } else {
        const [day, month, year] = parts;
        return `${day.padStart(2, "0")}/${month.padStart(2, "0")}/${year}`;
      }
    }
  }

  return "";
}

function getCurrentTime() {
  const now = new Date();
  return `${String(now.getHours()).padStart(2, "0")}:${String(
    now.getMinutes()
  ).padStart(2, "0")}`;
}

function getDayOfWeek(dayString) {
  if (!dayString) return "Không xác định";

  const normalizedDay = dayString.toString().toLowerCase().trim();
  
  // Mở rộng map để bao gồm thêm các định dạng tiếng Việt
  const dayMap = {
    2: "Thứ Hai",
    3: "Thứ Ba", 
    4: "Thứ Tư",
    5: "Thứ Năm",
    6: "Thứ Sáu",
    7: "Thứ Bảy",
    cn: "Chủ Nhật",
    "thứ 2": "Thứ Hai",
    "thứ 3": "Thứ Ba",
    "thứ 4": "Thứ Tư", 
    "thứ 5": "Thứ Năm",
    "thứ 6": "Thứ Sáu",
    "thứ 7": "Thứ Bảy",
    "chủ nhật": "Chủ Nhật",
    "thu 2": "Thứ Hai",
    "thu 3": "Thứ Ba",
    "thu 4": "Thứ Tư",
    "thu 5": "Thứ Năm", 
    "thu 6": "Thứ Sáu",
    "thu 7": "Thứ Bảy",
    monday: "Thứ Hai",
    tuesday: "Thứ Ba", 
    wednesday: "Thứ Tư",
    thursday: "Thứ Năm",
    friday: "Thứ Sáu",
    saturday: "Thứ Bảy",
    sunday: "Chủ Nhật"
  };

  // Kiểm tra xem có chứa từ "thứ" không
  if (normalizedDay.includes("thứ")) {
    // Tách số từ chuỗi "thứ X"
    const num = normalizedDay.match(/\d+/);
    if (num) {
      return dayMap[num[0]] || dayString;
    }
  }

  return dayMap[normalizedDay] || dayString;
}

// Hàm format ngày từ Excel
function formatExcelDate(excelDate) {
  if (!excelDate) return "";

  // Kiểm tra nếu là số (Excel date serial)
  if (typeof excelDate === "number") {
    // Chuyển đổi từ Excel date serial sang JavaScript Date
    const jsDate = new Date(Date.UTC(1900, 0, excelDate - 1));

    // Format lại ngày
    const day = String(jsDate.getUTCDate()).padStart(2, "0");
    const month = String(jsDate.getUTCMonth() + 1).padStart(2, "0");
    const year = jsDate.getUTCFullYear();

    return `${day}/${month}/${year}`;
  }

  // Nếu đã là chuỗi ngày, thử parse
  if (typeof excelDate === "string") {
    // Thử parse các định dạng khác nhau
    const parsedDate = new Date(excelDate);
    if (!isNaN(parsedDate)) {
      const day = String(parsedDate.getDate()).padStart(2, "0");
      const month = String(parsedDate.getMonth() + 1).padStart(2, "0");
      const year = parsedDate.getFullYear();
      return `${day}/${month}/${year}`;
    }
  }

  return "";
}

function formatTime(timeStr) {
  if (!timeStr) return "";

  // Loại bỏ các giá trị không phải thời gian
  const roomKeywords = [
    "phòng",
    "p.",
    "lavender",
    "lotus",
    "watch",
    "sk",
    "meeting",
    "phong",
    "room",
  ];

  // Chuyển sang chữ thường và loại bỏ khoảng trắng
  const lowerTimeStr = String(timeStr).toLowerCase().trim();

  // Kiểm tra nếu chứa từ khóa phòng thì bỏ qua
  if (roomKeywords.some((keyword) => lowerTimeStr.includes(keyword))) {
    return "";
  }

  // Các định dạng thời gian
  const timeFormats = [
    /^\d{1,2}:\d{2}$/, // HH:MM
    /^\d{1,2}h\d{2}$/, // Hh:MM
    /^\d{1,2}h\s?\d{2}$/, // H h MM
    /^\d{1,2}\.\d{2}$/, // HH.MM
  ];

  for (let format of timeFormats) {
    if (format.test(lowerTimeStr)) {
      // Chuẩn hóa định dạng
      const timeParts = lowerTimeStr
        .replace("h", ":")
        .replace(".", ":")
        .split(":");
      const hours = timeParts[0].padStart(2, "0");
      const minutes = timeParts[1].padStart(2, "0");
      return `${hours}:${minutes}`;
    }
  }

  // Nếu là số thập phân (Excel time)
  if (!isNaN(parseFloat(timeStr))) {
    const totalMinutes = Math.round(parseFloat(timeStr) * 24 * 60);
    const hours = Math.floor(totalMinutes / 60);
    const minutes = totalMinutes % 60;
    return `${String(hours).padStart(2, "0")}:${String(minutes).padStart(
      2,
      "0"
    )}`;
  }

  return "";
}

function normalizeRoomName(room) {
  if (!room) return "Không xác định";

  // Chuyển sang chữ thường và loại bỏ khoảng trắng thừa
  const normalized = String(room).toLowerCase().trim();

  const roomMap = {
    lotus: "Phòng Lotus",
    "lavender 1": "Phòng Lavender 1",
    "lavender 2": "Phòng Lavender 2",
    "p.1": "Phòng Lotus",
    "p.2": "Phòng Lavender 1",
    "p.3": "Phòng Lavender 2",
    "p. lavender 1": "Phòng Lavender 1",
    "p. lavender 2": "Phòng Lavender 2",
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
  const [currentHour, currentMin] = current.split(":").map(Number);
  const [startHour, startMin] = start.split(":").map(Number);
  const [endHour, endMin] = end.split(":").map(Number);

  const currentTime = currentHour * 60 + currentMin;
  const startTime = startHour * 60 + startMin;
  const endTime = endHour * 60 + endMin;

  return currentTime >= startTime && currentTime <= endTime;
}

// Xử lý tải file Excel
// Đảm bảo đã import thư viện XLSX vào HTML
// function processExcelFile(file) {
//   return new Promise((resolve, reject) => {
//     const reader = new FileReader();

//     reader.onload = function (e) {
//       try {
//         const data = new Uint8Array(e.target.result);
//         const workbook = XLSX.read(data, { type: "array" });
//         const firstSheet = workbook.Sheets[workbook.SheetNames[0]];

//         // Get raw data
//         const rawData = XLSX.utils.sheet_to_json(firstSheet, {
//           raw: false,
//           defval: "",
//           header: "A",
//         });

//         // Extract room names from header row
//         const roomNames = {
//           "PHÒNG LOTUS": "Phòng Lotus",
//           "P.LAVENDER 1": "Phòng Lavender 1",
//           "PHÒNG LAVENDER 2": "Phòng Lavender 2",
//         };

//         // Process meetings
//         const meetings = [];
//         let currentDate = "";
//         let currentDay = "";

//         rawData.forEach((row, index) => {
//           // Check if row contains date
//           // if (row['A'] && (row['A'].includes('THỨ') || !isNaN(row['A']))) {
//           //     currentDay = row['A'].includes('THỨ') ? row['A'] : '';
//           //     currentDate = !isNaN(row['A']) ? formatDate(row['A']) : '';
//           // }
//           if (row["A"]) {
//             if (row["A"].includes("THỨ")) {
//               currentDay = row["A"];
//             } else if (!isNaN(Number(row["A"]))) {
//               currentDate = formatDate(row["A"]);
//               console.log("Raw date:", row["A"], "Formatted:", currentDate);
//             }
//           }

//           // Get time from column B
//           const timeSlot = row["B"];
//           if (!timeSlot) return;

//           // Check each room column (C, D, E) for meetings
//           ["C", "D", "E"].forEach((col, roomIndex) => {
//             if (
//               row[col] &&
//               typeof row[col] === "string" &&
//               row[col].trim() !== ""
//             ) {
//               const roomName = Object.values(roomNames)[roomIndex];

//               // Parse meeting details
//               const meetingInfo = parseMeetingInfo(row[col]);

//               meetings.push({
//                 id: meetings.length + 1,
//                 date: currentDate,
//                 dayOfWeek: getDayOfWeek(currentDay),
//                 room: roomName,
//                 startTime: formatTime(timeSlot),
//                 endTime: calculateEndTime(timeSlot),
//                 duration: calculateDuration(
//                   timeSlot,
//                   calculateEndTime(timeSlot)
//                 ),
//                 purpose: meetingInfo.purpose,
//                 content: meetingInfo.content,
//               });
//             }
//           });
//         });

//         resolve(meetings);
//       } catch (error) {
//         console.error("Error processing file:", error);
//         reject(error);
//       }
//     };

//     reader.onerror = reject;
//     reader.readAsArrayBuffer(file);
//   });
// }

function processExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = function (e) {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { 
          type: "array",
          cellDates: true  // Thêm option này để đọc đúng định dạng ngày
        });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];

        // Get raw data với options được điều chỉnh
        const rawData = XLSX.utils.sheet_to_json(firstSheet, {
          raw: true,
          defval: "",
          header: "A",
          dateNF: 'dd/mm/yyyy'  // Định dạng ngày mong muốn
        });

        // Log toàn bộ dữ liệu thô để kiểm tra
        console.log("Raw Excel Data:", rawData);

        // Extract room names from header row
        const roomNames = {
          "PHÒNG LOTUS": "Phòng Lotus",
          "P.LAVENDER 1": "Phòng Lavender 1",
          "PHÒNG LAVENDER 2": "Phòng Lavender 2",
        };

        // Process meetings
        const meetings = [];
        let currentDate = "";
        let currentDay = "";

        // Thêm logging để debug từng row
        rawData.forEach((row, index) => {
          console.log(`Processing row ${index}:`, row);
          console.log(`Value in column A:`, row['A']);
          console.log(`Type of A:`, typeof row['A']);

          // Kiểm tra nếu có giá trị ở cột A
          if (row['A'] !== undefined && row['A'] !== '') {
            const cellValue = row['A'];
            console.log(`Found value in A:`, cellValue, `Type:`, typeof cellValue);

            // Nếu là thứ
            if (typeof cellValue === 'string' && cellValue.toUpperCase().includes('THỨ')) {
              currentDay = cellValue;
              console.log(`Set current day to:`, currentDay);
            } 
            // Nếu là số hoặc ngày
            else {
              // Thử chuyển đổi sang số
              const numValue = Number(cellValue);
              if (!isNaN(numValue)) {
                // Xử lý như Excel serial date
                const excelDate = new Date(Date.UTC(1900, 0, numValue - 1));
                currentDate = formatDate(excelDate);
                console.log(`Converted Excel serial date:`, numValue, `to:`, currentDate);
              } else if (cellValue instanceof Date) {
                // Nếu là Date object
                currentDate = formatDate(cellValue);
                console.log(`Converted Date object to:`, currentDate);
              } else {
                // Thử parse như string
                currentDate = formatDate(cellValue);
                console.log(`Attempted to parse string date:`, cellValue, `Result:`, currentDate);
              }
            }
          }

          // Get time from column B
          const timeSlot = row["B"];
          if (!timeSlot) return;

          // Check each room column (C, D, E) for meetings
          ["C", "D", "E"].forEach((col, roomIndex) => {
            if (row[col] && typeof row[col] === "string" && row[col].trim() !== "") {
              const roomName = Object.values(roomNames)[roomIndex];
              const meetingInfo = parseMeetingInfo(row[col]);

              meetings.push({
                id: meetings.length + 1,
                date: currentDate,
                dayOfWeek: getDayOfWeek(currentDay),
                room: roomName,
                startTime: formatTime(timeSlot),
                endTime: calculateEndTime(timeSlot),
                duration: calculateDuration(timeSlot, calculateEndTime(timeSlot)),
                purpose: meetingInfo.purpose,
                content: meetingInfo.content,
              });

              // Log mỗi meeting được tạo
              console.log(`Created meeting:`, meetings[meetings.length - 1]);
            }
          });
        });

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
