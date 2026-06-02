// main.js - 主程序入口
// 动态表单架构：所有表单由 form-renderer.js 根据模板 schema 动态生成

window.addEventListener('DOMContentLoaded', async () => {
    // --- Init Modules ---
    if (window.AppVulnManager) AppVulnManager.init();
    if (window.AppToolbox) AppToolbox.init();
    if (window.AppFormRenderer) AppFormRenderer.init();
    if (window.AppTemplateManager) AppTemplateManager.init();

    // --- Global Refs ---
    const statusDot = document.getElementById('api-status-dot');
    const statusText = document.getElementById('api-status-text');
    const versionWarningBanner = document.getElementById('version-warning-banner');

    function normalizeVersion(raw) {
        if (!raw) return '';
        const text = String(raw).trim();
        return text.toLowerCase().startsWith('v') ? text.slice(1) : text;
    }

    function showVersionWarning(message) {
        if (!versionWarningBanner) return;
        versionWarningBanner.textContent = message;
        versionWarningBanner.style.display = 'block';
    }

    function clearVersionWarning() {
        if (!versionWarningBanner) return;
        versionWarningBanner.textContent = '';
        versionWarningBanner.style.display = 'none';
    }

    // 存储最新更新检查结果，供菜单"检查更新"和点击回调使用
    window._latestUpdateResult = null;
    window._updateToastShown = false;

    /**
     * 显示更新提示 Toast
     */
    function showUpdateToast(latestVersion, downloadUrl) {
        if (window._updateToastShown) return;
        window._updateToastShown = true;

        var toast = document.createElement('div');
        toast.className = 'update-toast';
        toast.innerHTML = '<div class="toast-body">' +
            '<span class="toast-msg">发现新版本 <strong>V' + latestVersion + '</strong></span>' +
            '<span class="toast-actions">' +
            '<a class="toast-link">查看详情</a>' +
            '<span class="toast-close">&times;</span>' +
            '</span></div>';
        document.body.appendChild(toast);

        // 动画滑入
        setTimeout(function () { toast.classList.add('show'); }, 10);

        // 查看详情
        toast.querySelector('.toast-link').onclick = function (e) {
            e.preventDefault();
            if (window.electronAPI && window.electronAPI.openExternal) {
                window.electronAPI.openExternal(downloadUrl);
            } else if (downloadUrl) {
                window.open(downloadUrl, '_blank');
            }
        };

        // 关闭按钮
        toast.querySelector('.toast-close').onclick = function () {
            toast.classList.remove('show');
            setTimeout(function () { if (toast.parentNode) toast.remove(); }, 400);
        };

        // 8 秒后自动消失
        setTimeout(function () {
            toast.classList.remove('show');
            setTimeout(function () { if (toast.parentNode) toast.remove(); }, 400);
        }, 8000);
    }

    /**
     * 检查更新：发现新版本时在版本号上显示蓝色箭头指示器
     */
    async function checkForUpdates() {
        try {
            const result = await AppAPI._request('/api/check-update');
            window._latestUpdateResult = result;
            const verEl = document.getElementById('version-info');
            if (!verEl) return;

            if (result.has_update) {
                verEl.classList.add('version-update-available');
                showUpdateToast(result.latest_version, result.download_url);
            } else {
                verEl.classList.remove('version-update-available');
                verEl.title = '';
                verEl.onclick = null;
            }
        } catch (e) {
            // Silently ignore - update check is non-critical
        }
    }

    /**
     * 供菜单"检查更新"调用的全局入口
     * @param {boolean} showToastOnNoUpdate - 无更新时是否显示 toast
     */
    window._triggerUpdateCheck = async function (showToastOnNoUpdate) {
        try {
            const result = await AppAPI._request('/api/check-update');
            window._latestUpdateResult = result;
            const verEl = document.getElementById('version-info');
            if (!verEl) return;

            if (result && result.has_update) {
                verEl.classList.add('version-update-available');
                showUpdateToast(result.latest_version, result.download_url);
            } else if (showToastOnNoUpdate && result && !result.error) {
                verEl.classList.remove('version-update-available');
                verEl.title = '';
                verEl.onclick = null;
                if (window.AppUtils) AppUtils.showToast('当前已是最新版本', 'info');
            }
        } catch (e) {
            if (showToastOnNoUpdate && window.AppUtils) {
                AppUtils.showToast('检查更新失败', 'error');
            }
        }
    };

    async function checkVersionConsistency(configVersion) {
        try {
            const versionInfo = await AppAPI.getVersionInfo();
            const backendVersion = normalizeVersion(versionInfo.backend_version || configVersion);
            const sharedVersion = normalizeVersion(versionInfo.shared_version || '');
            const electronVersion = normalizeVersion(window.electronConfig && window.electronConfig.appVersion);

            if (electronVersion === '0.0.0' && backendVersion && backendVersion !== '0.0.0') {
                showVersionWarning(`版本异常：桌面端 ${electronVersion} / 后端 ${backendVersion}。通常是打包未包含 backend/shared-config.json，请重新打包发布。`);
                return;
            }

            if (electronVersion && backendVersion && electronVersion !== backendVersion) {
                showVersionWarning(`版本不一致：桌面端 ${electronVersion} / 后端 ${backendVersion}。建议重新构建并发布。`);
                return;
            }

            if (sharedVersion && backendVersion && sharedVersion !== backendVersion) {
                showVersionWarning(`版本不一致：shared-config ${sharedVersion} / backend ${backendVersion}。请先执行 npm run sync-version。`);
                return;
            }

            clearVersionWarning();
        } catch (e) {
            console.warn('Version consistency check failed:', e);
            showVersionWarning('版本一致性检查失败，请确认 /api/version 可访问。');
        }
    }

    // --- Init App ---
    async function initApp() {
        try {
            const config = await AppAPI.checkConnection();

            // 更新连接状态
            statusDot.classList.remove('error');
            statusDot.classList.add('connected');
            statusText.innerText = "Connected";
            statusText.style.color = "green";
            
            const connectionInfo = `后端地址: ${AppAPI.BASE_URL}`;
            statusDot.title = connectionInfo;
            statusText.title = connectionInfo;
            
            const versionEl = document.getElementById('version-info');
            if (versionEl) versionEl.innerText = config.version || '';

            await checkVersionConsistency(config.version || '');
            checkForUpdates();  // async, non-blocking

            // 同步漏洞列表到 VulnManager
            if (window.AppVulnManager && config.vulnerabilities_list) {
                window.AppVulnManager.VULN_LIST = config.vulnerabilities_list;
            }
            
            // 加载模板列表并渲染默认模板表单
            if (window.AppFormRenderer && window.AppFormRenderer.loadTemplateList) {
                await AppFormRenderer.loadTemplateList();
            }

        } catch (e) {
            console.error("Init failed", e);
            clearVersionWarning();
            statusDot.classList.remove('connected');
            statusDot.classList.add('error');
            statusText.innerText = "Connecting...";
            statusText.style.color = "#999";
            setTimeout(initApp, 2000);
        }
    }
    
    initApp();

    // --- Global Shortcuts ---
    document.addEventListener('keydown', (e) => {
        // Ctrl+Enter: 生成报告
        if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
            const btnGen = document.getElementById('btn-dynamic-generate');
            if (btnGen && !btnGen.disabled) btnGen.click();
        }

        // Esc: 关闭模态框
        if (e.key === 'Escape') {
            // 模板详情模态框
            const templateDetailModal = document.getElementById('template-detail-modal');
            if (templateDetailModal && templateDetailModal.style.display !== 'none') {
                templateDetailModal.style.display = 'none';
                return;
            }
            // 图片预览模态框
            const imgModal = document.getElementById('form-image-preview-modal');
            if (imgModal && imgModal.style.display !== 'none') {
                if (window.AppFormRenderer) AppFormRenderer.closeImagePreview();
                return;
            }
            // 工具箱模态框
            const toolbox = document.getElementById('toolbox-modal');
            if (toolbox && toolbox.style.display !== 'none') {
                toolbox.style.display = 'none';
                return;
            }
        }
    });
});
