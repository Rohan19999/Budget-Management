window.FINANCE_API_BASE = "http://localhost:5050";

window.applyFinanceSidebarRoleVisibility = function () {
    const auth = JSON.parse(localStorage.getItem("financeAuth") || "null");
    if (!auth || !auth.role) return;

    const showUserManagement = auth.role === "Finance Admin";
    document.querySelectorAll('a.nav-sub-item[href="usermanagement.html"]').forEach((link) => {
        link.classList.toggle("nav-role-hidden", !showUserManagement);
    });
};

if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", window.applyFinanceSidebarRoleVisibility);
} else {
    window.applyFinanceSidebarRoleVisibility();
}
