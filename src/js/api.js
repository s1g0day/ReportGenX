// api.js - API 接口交互

window.AppAPI = {
    // 优先使用 Electron 注入的配置，其次使用 AppConfig，最后使用默认值
    get BASE_URL() {
        return (window.electronConfig && window.electronConfig.apiBaseUrl)
            || (window.AppConfig && window.AppConfig.API && window.AppConfig.API.BASE_URL) 
            || "http://127.0.0.1:8000";
    },

    _extractErrorMessage(errorData, status) {
        return errorData.detail || errorData.message || errorData.error || `API Error: ${status}`;
    },

    _appendTokenToUrl(url) {
        const token = window.electronConfig && window.electronConfig.appApiToken;
        if (!token) {
            return url;
        }
        const separator = url.includes('?') ? '&' : '?';
        return `${url}${separator}app_token=${encodeURIComponent(token)}`;
    },

    _buildAuthHeaders() {
        const headers = {};
        if (window.electronConfig && window.electronConfig.appApiToken) {
            headers['X-App-Token'] = window.electronConfig.appApiToken;
        }
        return headers;
    },

    async _verifyTokenBinding() {
        try {
            const res = await fetch(`${this.BASE_URL}/api/health-auth`, {
                method: 'GET',
                headers: this._buildAuthHeaders()
            });
            return res.ok;
        } catch (_e) {
            return false;
        }
    },

    async _request(endpoint, method = 'GET', body = null) {
        const options = {
            method,
            headers: this._buildAuthHeaders()
        };

        if (body) {
            if (body instanceof FormData) {
                options.body = body;
            } else {
                options.headers['Content-Type'] = 'application/json';
                options.body = JSON.stringify(body);
            }
        }
        
        try {
            const res = await fetch(`${this.BASE_URL}${endpoint}`, options);
            if (!res.ok) {
                const errorData = await res.json().catch(() => ({}));
                const message = this._extractErrorMessage(errorData, res.status);
                if (res.status === 403 && /invalid application token/i.test(message)) {
                    const tokenBound = await this._verifyTokenBinding();
                    if (!tokenBound) {
                        throw new Error('检测到后端令牌不一致，请完全退出并重启应用（关闭旧后端进程后再启动）。');
                    }
                }
                throw new Error(message);
            }
            return await res.json();
        } catch (e) {
            if (e instanceof TypeError) {
                throw new Error(`无法连接后端服务（${this.BASE_URL}）`);
            }
            throw e;
        }
    },

    async _requestBlob(endpoint, method = 'GET', body = null) {
        const options = {
            method,
            headers: this._buildAuthHeaders()
        };

        if (body) {
            if (body instanceof FormData) {
                options.body = body;
            } else {
                options.headers['Content-Type'] = 'application/json';
                options.body = JSON.stringify(body);
            }
        }

        try {
            const res = await fetch(`${this.BASE_URL}${endpoint}`, options);
            if (!res.ok) {
                const errorData = await res.json().catch(() => ({}));
                const message = this._extractErrorMessage(errorData, res.status);
                if (res.status === 403 && /invalid application token/i.test(message)) {
                    const tokenBound = await this._verifyTokenBinding();
                    if (!tokenBound) {
                        throw new Error('检测到后端令牌不一致，请完全退出并重启应用（关闭旧后端进程后再启动）。');
                    }
                }
                throw new Error(message);
            }

            return {
                blob: await res.blob(),
                filename: this._extractDownloadFilename(res.headers.get('Content-Disposition'))
            };
        } catch (e) {
            if (e instanceof TypeError) {
                throw new Error(`无法连接后端服务（${this.BASE_URL}）`);
            }
            throw e;
        }
    },

    _extractDownloadFilename(contentDisposition) {
        if (!contentDisposition) {
            return '';
        }

        const utf8Match = contentDisposition.match(/filename\*=UTF-8''([^;\n]+)/i);
        if (utf8Match && utf8Match[1]) {
            try {
                return decodeURIComponent(utf8Match[1].trim());
            } catch (_e) {
                return utf8Match[1].trim();
            }
        }

        const basicMatch = contentDisposition.match(/filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/i);
        if (basicMatch && basicMatch[1]) {
            return basicMatch[1].replace(/['"]/g, '').trim();
        }

        return '';
    },

    // --- Core ---

    // 初始化检查
    async checkConnection() {
        return this._request('/api/config');
    },

    async getConfig() {
        return this._request('/api/config');
    },

    async getVersionInfo() {
        return this._request('/api/version');
    },
    
    async updateConfig(data) {
        return this._request('/api/update-config', 'POST', data);
    },

    async getPluginRuntimeConfig() {
        return this._request('/api/plugin-runtime-config');
    },

    async getFrontendConfig() {
        return this._request('/api/frontend-config');
    },

    async updatePluginRuntimeConfig(data) {
        return this._request('/api/plugin-runtime-config', 'POST', data);
    },

    // URL 处理
    async processUrl(url) {
        return this._request('/api/process-url', 'POST', { url });
    },

    // 打开文件夹
    async openFolder(path) {
        return this._request('/api/open-folder', 'POST', { path });
    },

    async uploadImage(base64Data, filename) {
        return this._request('/api/upload-image', 'POST', { 
            image_base64: base64Data, 
            filename 
        });
    },

    async backupDatabase() {
        const download = await this._requestBlob('/api/backup-db');
        const filename = download.filename || 'backup.db';
        const blob = download.blob;
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        window.URL.revokeObjectURL(url);

        return {
            success: true,
            filename
        };
    },

    // --- Vulnerabilities ---

    async getVulnerabilities() {
        return this._request('/api/vulnerabilities');
    },

    async saveVulnerability(data) {
        // Decide create or update based on logic or pass ID separately
        // But backend uses PUT for update by ID and POST for create
        // Helper to check if it's an update probably needs ID
        // For simplicity let the caller decide or auto-detect
        return this._request('/api/vulnerabilities', 'POST', data);
    },

    async updateVulnerability(id, data) {
         return this._request(`/api/vulnerabilities/${encodeURIComponent(id)}`, 'PUT', data);
    },

    async deleteVulnerability(id) {
         return this._request(`/api/vulnerabilities/${encodeURIComponent(id)}`, 'DELETE');
     },

    // Vulnerabilities CRUD object for CRUDManager
    Vulnerabilities: {
        async save(data) {
            if (data.id) {
                return window.AppAPI.updateVulnerability(data.id, data);
            } else {
                return window.AppAPI.saveVulnerability(data);
            }
        },
        
        async delete(id) {
            return window.AppAPI.deleteVulnerability(id);
        },

        async importFile(file, overwrite = false) {
            const formData = new FormData();
            formData.append('file', file);
            formData.append('overwrite', overwrite);
            return window.AppAPI._request('/api/vulnerabilities/import', 'POST', formData);
        },

        async exportFile() {
            const resp = await fetch(`${window.AppAPI.BASE_URL}/api/vulnerabilities/export`);
            if (!resp.ok) throw new Error('Export failed');
            const blob = await resp.blob();
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'vulnerabilities.xlsx';
            a.click();
            URL.revokeObjectURL(url);
        },

        async downloadTemplate() {
            const resp = await fetch(`${window.AppAPI.BASE_URL}/api/vulnerabilities/template`);
            if (!resp.ok) throw new Error('Download failed');
            const blob = await resp.blob();
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'vulnerability_template.xlsx';
            a.click();
            URL.revokeObjectURL(url);
        }
    },

     // --- Templates ---
    
    Templates: {
        async list(includeDetails = false) {
            return window.AppAPI._request(`/api/templates?include_details=${includeDetails}`);
        },
        
        async getSchema(id) {
            return window.AppAPI._request(`/api/templates/${id}/schema`);
        },
        
        async getDataSources(id) {
            return window.AppAPI._request(`/api/templates/${id}/data-sources`);
        },
        
        async reload() {
            return window.AppAPI._request('/api/templates/reload', 'POST');
        },
        
        async generate(id, data) {
            return window.AppAPI._request(`/api/templates/${id}/generate`, 'POST', data);
        },
        
        async import(file, overwrite = false) {
            const formData = new FormData();
            formData.append('file', file);
            formData.append('overwrite', overwrite);
            return window.AppAPI._request('/api/templates/import', 'POST', formData);
        },
        
        async batchImport(files, overwrite = false) {
            const formData = new FormData();
            for(let i=0; i<files.length; i++) {
                formData.append('files', files[i]);
            }
            formData.append('overwrite', overwrite);
            return window.AppAPI._request('/api/templates/batch-import', 'POST', formData);
        },
        
        // Export is a download, usually via GET
        exportUrl(id) {
            return window.AppAPI._appendTokenToUrl(`${window.AppAPI.BASE_URL}/api/templates/${id}/export`);
        },

        async export(id) {
            return window.AppAPI._requestBlob(`/api/templates/${id}/export`);
        },

        async batchExport(templateIds) {
            return window.AppAPI._requestBlob('/api/templates/batch-export', 'POST', templateIds);
        },
         
        async save(data) {
            // Templates don't have a traditional save operation
            // This is a placeholder for CRUDManager compatibility
            return { message: '模板操作成功' };
        },
        
        async delete(id) {
             return window.AppAPI._request(`/api/templates/${id}`, 'DELETE');
        },
        
        async setDefault(id) {
            return window.AppAPI._request(`/api/templates/${id}/set-default`, 'PUT');
        }
    },

    // --- Reports ---

    Reports: {
        async list() {
            return window.AppAPI._request('/api/list-reports', 'POST'); 
        },
        
        async save(data) {
            // Reports don't have a traditional save operation
            // This is a placeholder for CRUDManager compatibility
            return { message: '报告操作成功' };
        },
        
        async delete(path) {
            return window.AppAPI._request('/api/delete-report', 'POST', { path });
        },
        
        async merge(files, filename) {
             return window.AppAPI._request('/api/merge-reports', 'POST', { files, output_filename: filename });
        }
    },
    
    // --- ICP ---
    
    Icp: {
        async list() {
            return window.AppAPI._request('/api/icp-list');
        },
        
        async save(data) {
            if (data.id) {
                return window.AppAPI.Icp.update(data.id, data);
            } else {
                return window.AppAPI.Icp.add(data);
            }
        },
        
        async add(data) {
             return window.AppAPI._request('/api/icp-entry', 'POST', data);
        },
        
        async update(id, data) {
             return window.AppAPI._request(`/api/icp-entry/${id}`, 'PUT', data);
        },
        
        async delete(id) {
             return window.AppAPI._request(`/api/icp-entry/${id}`, 'DELETE');
        },
        
        async batchDelete(ids) {
             return window.AppAPI._request('/api/icp-batch-delete', 'POST', { ids });
        },

        async importFile(file, overwrite = false) {
            const formData = new FormData();
            formData.append('file', file);
            formData.append('overwrite', overwrite);
            return window.AppAPI._request('/api/icp-import', 'POST', formData);
        },

        async downloadTemplate() {
            const resp = await fetch(`${window.AppAPI.BASE_URL}/api/icp-template`);
            if (!resp.ok) throw new Error('Download failed');
            const blob = await resp.blob();
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'icp_template.xlsx';
            a.click();
            URL.revokeObjectURL(url);
        },

        async exportFile() {
            const resp = await fetch(`${window.AppAPI.BASE_URL}/api/icp-export`);
            if (!resp.ok) throw new Error('Export failed');
            const blob = await resp.blob();
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'icp_data.xlsx';
            a.click();
            URL.revokeObjectURL(url);
        }
    },

    // --- Web UI ---

    WebUIConfig: {
        get: () => window.AppAPI._request('/api/web-ui-config'),
        set: (enabled) => window.AppAPI._request('/api/web-ui-config', 'POST', { enabled })
    }
};
