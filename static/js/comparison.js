const loadingScreen = document.getElementById("loadingScreen");
const loadingStatusText = document.getElementById("loadingStatusText");
const paginationContainer = document.querySelector(".pagination-container");
const dropdownMenu = document.querySelector(".dropdown-menu");

function showLoadingAndNavigate(url, statusMessage) {
  if (!url || url === "#") {
    return;
  }

  loadingStatusText.textContent = statusMessage;
  loadingScreen.style.display = "flex";

  setTimeout(() => {
    window.location.href = url;
  }, 120);
}

if (paginationContainer) {
  paginationContainer.addEventListener("click", function (event) {
    const targetLink = event.target.closest("a");
    if (targetLink && targetLink.getAttribute("href")) {
      event.preventDefault();
      showLoadingAndNavigate(targetLink.getAttribute("href"), "Loading selected page...");
    }
  });
}

if (dropdownMenu) {
  dropdownMenu.addEventListener("click", function (event) {
    const targetLink = event.target.closest("a");

    if (
      targetLink &&
      targetLink.getAttribute("href") &&
      targetLink.getAttribute("href") !== "#"
    ) {
      event.preventDefault();
      showLoadingAndNavigate(targetLink.getAttribute("href"), "Loading selected sheet...");
    }
  });
}

window.addEventListener("beforeunload", function () {
  loadingStatusText.textContent = "Almost finished...";
});
