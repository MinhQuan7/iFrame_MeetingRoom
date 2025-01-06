// Trong index1.html, thêm script để xử lý nút Home
document.addEventListener("DOMContentLoaded", function () {
  const homeButton = document.querySelector(".home-button");

  if (homeButton) {
    homeButton.addEventListener("click", function () {
      // Quay trở lại trang chính
      window.location.href = "index.html";
    });
  }
});
