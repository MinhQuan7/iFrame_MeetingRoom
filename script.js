document.addEventListener('DOMContentLoaded', function() {
    const uploadButton = document.querySelector('.upload-button');
    
    uploadButton.addEventListener('click', function(event) {
        event.preventDefault();
        
        // Tạo một input element ẩn để chọn file
        const fileInput = document.createElement('input');
        fileInput.type = 'file';
        fileInput.accept = '.xlsx, .xls, .csv'; // Chỉ chấp nhận file Excel và CSV
        fileInput.style.display = 'none';
        
        // Thêm sự kiện khi file được chọn
        fileInput.addEventListener('change', function(e) {
            if (e.target.files.length > 0) {
                const file = e.target.files[0];
                console.log('File đã chọn:', file.name);
                // Thêm logic xử lý file tại đây
                handleFileUpload(file);
            }
        });
        
        // Kích hoạt click để mở dialog chọn file
        fileInput.click();
    });
});

function handleFileUpload(file) {
    // Tạo FormData object để gửi file
    const formData = new FormData();
    formData.append('meetingFile', file);
    
    // Thêm logic gửi file lên server tại đây
    console.log('Đang xử lý file:', file.name);
    
    // Ví dụ về việc gửi file lên server
    /*
    fetch('/api/upload-meeting', {
        method: 'POST',
        body: formData
    })
    .then(response => response.json())
    .then(data => {
        console.log('Upload thành công:', data);
    })
    .catch(error => {
        console.error('Lỗi khi upload:', error);
    });
    */
}