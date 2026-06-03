// form-renderer.js - 动态表单渲染器
// 支持动态/静态表单切换、数据源缓存、模板热加载

window.AppFormRenderer = {
    currentSchema: null,
    currentTemplateId: null,
    formData: {},
    dataSources: {},
    behaviors: {},
    
    // @service/ endpoint resolution cache — loaded lazily from /api/services
    serviceMap: null,
    
    // 数据源缓存配置
    dataSourceCache: {},
    // 使用全局配置的缓存过期时间
    get cacheExpiry() {
        return (window.AppConfig && window.AppConfig.CACHE && window.AppConfig.CACHE.EXPIRY_MS) 
            || 5 * 60 * 1000;
    },
    
    init() {
        this.dynamicFormContainer = document.getElementById('dynamic-form-container');
        this.templateSelector = document.getElementById('template-selector');
        this.reloadButton = document.getElementById('btn-reload-templates');
        this.importButton = document.getElementById('btn-import-template');
        this.bindEvents();
    },

    bindEvents() {
        // 模板选择器事件
        if (this.templateSelector) {
            this.templateSelector.addEventListener('change', async (e) => {
                const templateId = e.target.value;
                if (templateId) await this.loadTemplate(templateId);
            });
        }
        
        // 模板刷新按钮事件
        if (this.reloadButton) {
            this.reloadButton.addEventListener('click', async () => {
                await this.reloadTemplates();
            });
        }
        
        // 导入按钮事件
        if (this.importButton) {
            this.importButton.addEventListener('click', () => {
                this.openImportDialog();
            });
        }
    },
    
    async loadTemplateList(skipAutoLoad = false) {
        try {
            const data = await AppAPI.Templates.list();
            if (this.templateSelector && data.templates) {
                this.templateSelector.innerHTML = '';
                data.templates.forEach(t => {
                    const opt = document.createElement('option');
                    opt.value = t.id;
                    opt.textContent = (t.icon || '') + ' ' + t.name + ' (v' + t.version + ')';
                    if (t.id === data.default_template) opt.selected = true;
                    this.templateSelector.appendChild(opt);
                });
                // 只在非跳过模式下自动加载默认模板
                if (!skipAutoLoad && data.default_template) {
                    await this.loadTemplate(data.default_template);
                }
            }
            return data;
        } catch (e) {
            console.error('Load template list failed:', e);
            this.showError('加载模板列表失败: ' + e.message);
            return null;
        }
    },
    
    // 模板热加载
    async reloadTemplates() {
        try {
            if (window.AppUtils) AppUtils.showToast('正在刷新模板...', 'info');
            
            // 保存当前选中的模板ID
            const currentTemplateId = this.currentTemplateId;
            
            // 调用后端热加载API
            const result = await AppAPI.Templates.reload();
            
            if (result.success) {
                // 清空缓存
                this.clearCache();
                
                // 重新加载模板列表（跳过自动加载默认模板）
                await this.loadTemplateList(true);
                
                // 如果当前有选中的模板，重新加载并更新选择器
                if (currentTemplateId && this.templateSelector) {
                    this.templateSelector.value = currentTemplateId;
                    await this.loadTemplate(currentTemplateId);
                }
                
                if (window.AppUtils) AppUtils.showToast(`模板刷新成功！已加载 ${result.loaded_count} 个模板`, 'success');
            } else {
                throw new Error(result.message || '刷新失败');
            }
        } catch (e) {
            console.error('Reload templates failed:', e);
            if (window.AppUtils) AppUtils.showToast('刷新模板失败: ' + e.message, 'error');
        }
    },
    
    // 缓存管理
    clearCache() {
        this.dataSourceCache = {};
    },
    
    getCachedData(key) {
        const cached = this.dataSourceCache[key];
        if (cached && (Date.now() - cached.timestamp < this.cacheExpiry)) {
            return cached.data;
        }
        return null;
    },
    
    setCachedData(key, data) {
        this.dataSourceCache[key] = {
            data: data,
            timestamp: Date.now()
        };
    },

    async loadTemplate(templateId) {
        try {
            // 动态表单模板 - 完整加载和渲染
            const schema = await AppAPI.Templates.getSchema(templateId);
            
            // 使用缓存加载数据源
            const cacheKey = `datasource_${templateId}`;
            let dataSources = this.getCachedData(cacheKey);
            if (!dataSources) {
                try {
                    dataSources = await AppAPI.Templates.getDataSources(templateId);
                    this.setCachedData(cacheKey, dataSources);
                } catch(e) { /* ignore datasource error */ }
            }
            this.dataSources = dataSources || {};
            
            this.currentSchema = schema;
            this.currentTemplateId = templateId;
            this.formData = {};
            this.behaviors = {};
            this.dataChangedBehaviors = {};
            this.behaviorById = {};

            this.loadTemplateStyle(templateId);
            this.parseBehaviors(schema);
            this.renderForm(schema);
            await this.populateDataSources();
            
            window.dispatchEvent(new CustomEvent('template-loaded', { 
                detail: { templateId, schema, isStatic: false } 
            }));
            return schema;
        } catch (e) {
            console.error('Load template failed:', e);
            this.showError('加载模板失败: ' + e.message);
            return null;
        }
    },
    
    showError(message) {
        if (window.AppUtils) {
            AppUtils.showToast(message, 'error');
        } else {
            alert(message);
        }
    },

    parseBehaviors(schema) {
        if (!schema.behaviors) return;
        schema.behaviors.forEach(beh => {
            // Index by behavior ID for trigger_behavior support
            if (beh.id) {
                this.behaviorById[beh.id] = beh;
            }

            // Resolve trigger field and event from both old and new formats
            let triggerField, triggerEvent;
            if (beh.trigger && beh.trigger.field) {
                // Old format: trigger: { field: ..., event: ... }
                triggerField = beh.trigger.field;
                triggerEvent = String(beh.trigger.event || beh.trigger_event || 'change').toLowerCase();
            } else if (beh.trigger_field) {
                // New format: trigger_field + trigger_event at top level
                triggerField = beh.trigger_field;
                triggerEvent = String(beh.trigger_event || 'change').toLowerCase();
            }

            if (!triggerField) return;

            // Route to the correct index based on trigger event type
            if (triggerEvent === 'data_changed') {
                if (!this.dataChangedBehaviors[triggerField]) {
                    this.dataChangedBehaviors[triggerField] = [];
                }
                this.dataChangedBehaviors[triggerField].push(beh);
            } else if (triggerEvent === 'manual') {
                // Manual-only behaviors: stored only by ID, never auto-fired
                if (!this.behaviors[triggerField]) {
                    this.behaviors[triggerField] = [];
                }
                // Still register so trigger_behavior can find by field+id combination;
                // handleChange() will skip these because triggerEvent !== 'change'.
                this.behaviors[triggerField].push(beh);
            } else {
                // Default: 'change' event (backward compatible)
                if (!this.behaviors[triggerField]) {
                    this.behaviors[triggerField] = [];
                }
                this.behaviors[triggerField].push(beh);
            }
        });
    },

    renderForm(schema) {
        if (!this.dynamicFormContainer) return;
        this.dynamicFormContainer.innerHTML = '';
        
        const groups = {};
        (schema.field_groups || []).forEach(g => { groups[g.id] = {...g, fields: []}; });
        if (!groups['default']) groups['default'] = {id:'default',name:'其他',order:999,fields:[]};
        (schema.fields || []).forEach(field => {
            const gid = field.group || 'default';
            if (!groups[gid]) groups[gid] = {id:gid,name:gid,order:100,fields:[]};
            groups[gid].fields.push(field);
        });
        Object.values(groups).filter(g=>g.fields.length>0).sort((a,b)=>(a.order||0)-(b.order||0)).forEach(group => {
            this.dynamicFormContainer.appendChild(this.createGroupCard(group));
        });
        
        // 添加底部操作栏
        const submitSection = document.createElement('div');
        submitSection.className = 'card text-right';
        const zIndex = (window.AppConfig && window.AppConfig.Z_INDEX && window.AppConfig.Z_INDEX.SUBMIT_SECTION) || 100;
        submitSection.style.cssText = `position: sticky; bottom: 0; z-index: ${zIndex}; border-top: 2px solid var(--primary-color);`;
        submitSection.innerHTML = `
            <button type="button" class="btn btn-secondary" id="btn-dynamic-open-folder" style="margin-right: 10px;">打开报告目录</button>
            <button type="button" class="btn btn-secondary" id="btn-dynamic-reset">重置</button>
            <button type="button" class="btn btn-secondary" id="btn-dynamic-preview" style="margin-left: 10px;">预览数据</button>
            <button type="button" class="btn btn-primary" id="btn-dynamic-generate" style="margin-left: 10px; font-size: 16px; padding: 10px 30px;">
                生成报告
            </button>
        `;
        this.dynamicFormContainer.appendChild(submitSection);
        
        // 绑定按钮事件
        const generateBtn = document.getElementById('btn-dynamic-generate');
        const previewBtn = document.getElementById('btn-dynamic-preview');
        const resetBtn = document.getElementById('btn-dynamic-reset');
        const openFolderBtn = document.getElementById('btn-dynamic-open-folder');
        
        if (generateBtn) {
            generateBtn.addEventListener('click', async () => {
                await this.submitReport();
            });
        }
        
        if (previewBtn) {
            previewBtn.addEventListener('click', () => {
                this.previewFormData();
            });
        }
        
        if (resetBtn) {
            resetBtn.addEventListener('click', () => {
                this.resetForm();
            });
        }
        
        if (openFolderBtn) {
            openFolderBtn.addEventListener('click', () => {
                this.openReportFolder();
            });
        }
        
        this.setDefaultValues(schema);
    },
    
    previewFormData() {
        const data = this.collectFormData();
        const json = JSON.stringify(data, null, 2);
        
        // 尝试使用模态框显示
        const modal = document.getElementById('preview-modal');
        const content = document.getElementById('preview-content');
        
        if (modal && content) {
            content.textContent = json;
            modal.style.display = 'flex';
        } else {
            // 回退到 alert
            alert('表单数据预览:\n\n' + json);
        }
    },

    createGroupCard(group) {
        const card = document.createElement('div');
        card.className = 'card';
        
        // 标题行（包含折叠按钮）
        const titleRow = document.createElement('div');
        titleRow.style.cssText = 'display: flex; justify-content: space-between; align-items: center;';
        
        const title = document.createElement('h2');
        title.style.cssText = 'margin: 0; cursor: pointer;';
        title.innerHTML = (group.icon||'') + ' ' + group.name;
        titleRow.appendChild(title);
        
        // 如果组支持折叠，添加折叠按钮
        let toggleBtn = null;
        if (group.collapsed !== undefined) {
            toggleBtn = document.createElement('button');
            toggleBtn.type = 'button';
            toggleBtn.className = 'btn-mini';
            toggleBtn.style.cssText = 'background: #f0f0f0; border: none; padding: 5px 10px; cursor: pointer; border-radius: 4px; font-size: 12px;';
            toggleBtn.innerHTML = group.collapsed ? '展开 ▼' : '收起 ▲';
            titleRow.appendChild(toggleBtn);
        }
        
        card.appendChild(titleRow);
        
        const grid = document.createElement('div');
        grid.className = 'grid';
        group.fields.sort((a,b)=>(a.order||0)-(b.order||0)).forEach(field => {
            const el = this.createField(field);
            if (el) grid.appendChild(el);
        });
        
        // 如果默认收起，隐藏内容
        if (group.collapsed) {
            grid.style.display = 'none';
        }
        
        card.appendChild(grid);
        
        // 绑定折叠事件
        if (toggleBtn) {
            const toggleFn = () => {
                if (grid.style.display === 'none') {
                    grid.style.display = '';
                    toggleBtn.innerHTML = '收起 ▲';
                } else {
                    grid.style.display = 'none';
                    toggleBtn.innerHTML = '展开 ▼';
                }
            };
            toggleBtn.onclick = toggleFn;
            title.onclick = toggleFn;
        }
        
        return card;
    },

    createField(field) {
        if (!window.AppFormRendererFieldOps) {
            throw new Error('AppFormRendererFieldOps is not loaded');
        }
        return window.AppFormRendererFieldOps.createField(this, field);
    },

    createInput(field, pasteBtn = null) {
        if (!window.AppFormRendererFieldOps) {
            throw new Error('AppFormRendererFieldOps is not loaded');
        }
        return window.AppFormRendererFieldOps.createInput(this, field, pasteBtn);
    },
    
    // 创建可搜索下拉框
    createSearchableSelect(field) {
        const container = document.createElement('div');
        container.className = 'searchable-select-container';
        container.style.cssText = 'display: flex; flex-direction: column; gap: 8px;';
        
        // 搜索输入框
        const searchInput = document.createElement('input');
        searchInput.type = 'text';
        searchInput.className = 'search-input';
        searchInput.placeholder = field.search_placeholder || '输入关键词搜索...';
        searchInput.style.cssText = 'padding: 8px 12px; border: 1px solid #ddd; border-radius: 4px;';
        
        // 下拉选择框
        const select = document.createElement('select');
        select.id = field.key;
        select.name = field.key;
        const empty = document.createElement('option');
        empty.value = ''; 
        empty.textContent = '-- 请选择 --';
        select.appendChild(empty);
        
        if (field.source) select.dataset.source = field.source;
        
        // 保存所有选项用于过滤
        container._allOptions = [];
        
        // 搜索过滤逻辑
        searchInput.addEventListener('input', (e) => {
            const term = e.target.value.toLowerCase().trim();
            const options = container._allOptions;
            
            // 保存当前选中值
            const currentVal = select.value;
            
            // 清空并重建选项
            select.innerHTML = '';
            const emptyOpt = document.createElement('option');
            emptyOpt.value = ''; 
            emptyOpt.textContent = '-- 请选择 --';
            select.appendChild(emptyOpt);
            
            options.forEach(opt => {
                if (!term || opt.text.toLowerCase().includes(term)) {
                    const option = document.createElement('option');
                    option.value = opt.value;
                    option.textContent = opt.text;
                    option.dataset.name = opt.text;
                    select.appendChild(option);
                }
            });
            
            // 尝试恢复选中值
            if (currentVal) select.value = currentVal;
        });
        
        // 选择变更事件
        select.addEventListener('change', (e) => {
            const selectedOption = e.target.selectedOptions[0];
            this.formData[field.key] = e.target.value;
            
            // 如果有 on_change 处理
            
            this.handleChange(field, e.target.value);
        });
        
        container.appendChild(searchInput);
        container.appendChild(select);
        
        return container;
    },
    
    createImageUploader(field, multiple, pasteBtn = null) {
        if (!window.AppFormRendererImageOps) {
            throw new Error('AppFormRendererImageOps is not loaded');
        }
        return window.AppFormRendererImageOps.createImageUploader(this, field, multiple, pasteBtn);
    },
    
    // 绑定图片上传事件
    bindImageUploadEvents(field, uploadArea, previewContainer, multiple, pasteBtn = null) {
        if (!window.AppFormRendererImageOps) {
            throw new Error('AppFormRendererImageOps is not loaded');
        }
        return window.AppFormRendererImageOps.bindImageUploadEvents(this, field, uploadArea, previewContainer, multiple, pasteBtn);
    },
    
    // 上传图片到服务器
    async uploadImage(file) {
        return new Promise((resolve) => {
            const reader = new FileReader();
            reader.onload = async (e) => {
                try {
                    const data = await AppAPI.uploadImage(e.target.result, file.name || `screenshot_${Date.now()}.png`);
                    resolve(data.file_path ? data : null);
                } catch (err) {
                    if (window.AppUtils) AppUtils.showToast("上传失败: " + err.message, "error");
                    resolve(null);
                }
            };
            reader.readAsDataURL(file);
        });
    },
    
    // 添加图片项到预览 - 使用 CSS 类替代内联样式
    addImageItem(field, imageInfo, container, multiple) {
        if (!window.AppFormRendererImageOps) {
            throw new Error('AppFormRendererImageOps is not loaded');
        }
        return window.AppFormRendererImageOps.addImageItem(this, field, imageInfo, container, multiple);
    },
    
    // 确保预览模态框存在 - 使用 CSS 类替代内联样式
    ensurePreviewModal() {
        if (!window.AppFormRendererImageOps) {
            throw new Error('AppFormRendererImageOps is not loaded');
        }
        return window.AppFormRendererImageOps.ensurePreviewModal(this);
    },
    
    openImagePreview(src, text) {
        if (!window.AppFormRendererImageOps) {
            throw new Error('AppFormRendererImageOps is not loaded');
        }
        return window.AppFormRendererImageOps.openImagePreview(this, src, text);
    },
    
    closeImagePreview() {
        if (!window.AppFormRendererImageOps) {
            throw new Error('AppFormRendererImageOps is not loaded');
        }
        return window.AppFormRendererImageOps.closeImagePreview(this);
    },

    // 创建分组图片列表（统一 server_type_group + db_screenshot_group）
    createGroupedImageList(field) {
        const self = this;
        const container = document.createElement('div');
        container.className = 'grouped-image-list-container';
        container.id = field.key;

        // 初始化数据
        if (!this.formData[field.key]) {
            this.formData[field.key] = [];
        }

        // 分组列表容器
        const groupList = document.createElement('div');
        groupList.className = 'grouped-image-list';
        container.appendChild(groupList);

        // 添加按钮
        const addBtn = document.createElement('button');
        addBtn.type = 'button';
        addBtn.className = 'btn-add-grouped-image';
        addBtn.textContent = field.add_button_text || '添加';
        container.appendChild(addBtn);

        // 读取字段配置
        const groupField = field.group_field || { type: 'text', key: 'title', label: '标题' };
        const itemFieldKey = field.item_field_key || 'description';
        const imagePathKey = field.image_path_key || 'path';
        const itemLabel = field.item_label || '描述';
        const pasteEnabled = field.paste_enabled !== false;

        // 选项列表（select 类型时使用）
        const groupOptions = field.group_field && field.group_field.options 
            ? field.group_field.options 
            : (field.type_options || []);

        const renderAll = () => {
            groupList.innerHTML = '';
            (self.formData[field.key] || []).forEach((group, idx) => {
                const groupEl = self._createGroupedImageGroup(
                    field, idx, group, groupField, itemFieldKey, imagePathKey, itemLabel,
                    groupOptions, pasteEnabled, renderAll
                );
                groupList.appendChild(groupEl);
            });
        };

        const addGroup = () => {
            const defaultGroup = { items: [] };
            if (groupField.type === 'select') {
                const firstOpt = groupOptions.length > 0 ? groupOptions[0] : '';
                const value = typeof firstOpt === 'object' ? (firstOpt.value || firstOpt.label) : firstOpt;
                defaultGroup[groupField.key] = value;
            } else {
                defaultGroup[groupField.key] = '';
            }
            self.formData[field.key].push(defaultGroup);
            renderAll();
            self.emitDataChanged(field.key);
        };

        addBtn.addEventListener('click', addGroup);

        // 默认添加一个空分组
        if (self.formData[field.key].length === 0) {
            addGroup();
        } else {
            renderAll();
        }

        return container;
    },

    // 创建分组图片列表中的单个分组
    _createGroupedImageGroup(field, idx, group, groupField, itemFieldKey, imagePathKey, itemLabel,
                              groupOptions, pasteEnabled, renderAll) {
        const self = this;

        const groupDiv = document.createElement('div');
        groupDiv.className = 'grouped-image-group-item';

        // 分组头部
        const header = document.createElement('div');
        header.className = 'grouped-image-group-header';

        // 分组字段输入（select 或 text）
        let groupInput;
        if (groupField.type === 'select') {
            groupInput = document.createElement('select');
            groupInput.className = 'grouped-image-select';
            groupOptions.forEach(opt => {
                const option = document.createElement('option');
                option.value = typeof opt === 'object' ? (opt.value || opt.label) : opt;
                option.textContent = typeof opt === 'object' ? (opt.label || opt.value) : opt;
                if (option.value === group[groupField.key]) option.selected = true;
                groupInput.appendChild(option);
            });
            groupInput.addEventListener('change', (e) => {
                group[groupField.key] = e.target.value;
            });
        } else {
            groupInput = document.createElement('input');
            groupInput.type = 'text';
            groupInput.className = 'grouped-image-text-input';
            groupInput.placeholder = groupField.label || '标题';
            groupInput.value = group[groupField.key] || '';
            groupInput.addEventListener('input', (e) => {
                group[groupField.key] = e.target.value;
            });
        }

        header.appendChild(groupInput);

        // 粘贴截图按钮
        let pasteBtn = null;
        if (pasteEnabled) {
            pasteBtn = document.createElement('button');
            pasteBtn.type = 'button';
            pasteBtn.className = 'btn-paste-screenshot';
            pasteBtn.innerHTML = '粘贴截图';
            header.appendChild(pasteBtn);
        }

        // 删除分组按钮
        const delGroupBtn = document.createElement('button');
        delGroupBtn.type = 'button';
        delGroupBtn.className = 'btn-delete-group';
        delGroupBtn.innerHTML = '删除';
        delGroupBtn.onclick = () => {
            self.formData[field.key].splice(idx, 1);
            renderAll();
            self.emitDataChanged(field.key);
        };
        header.appendChild(delGroupBtn);

        groupDiv.appendChild(header);

        // 上传区域
        const uploadArea = document.createElement('div');
        uploadArea.className = 'upload-area';
        uploadArea.innerHTML = `
            <span class="upload-icon">📷</span>
            <p>点击上传或粘贴截图</p>
        `;
        groupDiv.appendChild(uploadArea);

        // 图片列表容器
        const previewContainer = document.createElement('div');
        previewContainer.className = 'image-list-container';
        groupDiv.appendChild(previewContainer);

        // 绑定上传和粘贴事件
        this._bindGroupedImageEvents(field, group, uploadArea, previewContainer, pasteBtn,
                                     itemFieldKey, imagePathKey, itemLabel);

        // 渲染已有图片
        if (group.items && group.items.length > 0) {
            group.items.forEach(item => {
                this._addGroupedImageItem(group, item, previewContainer, itemFieldKey, imagePathKey, itemLabel);
            });
        }

        return groupDiv;
    },

    // 绑定分组图片列表的上传和粘贴事件
    _bindGroupedImageEvents(field, group, uploadArea, previewContainer, pasteBtn,
                            itemFieldKey, imagePathKey, itemLabel) {
        const self = this;

        // 点击上传
        uploadArea.addEventListener('click', () => {
            const input = document.createElement('input');
            input.type = 'file';
            input.accept = field.accept || 'image/*';
            input.onchange = async (ev) => {
                const file = ev.target.files[0];
                if (file) {
                    const result = await self.uploadImage(file);
                    if (result) {
                        const newItem = {};
                        newItem[imagePathKey] = result.file_path;
                        newItem[itemFieldKey] = '';
                        group.items.push(newItem);
                        self._addGroupedImageItem(group, newItem, previewContainer,
                                                  itemFieldKey, imagePathKey, itemLabel, result);
                    }
                }
            };
            input.click();
        });

        // 粘贴按钮事件
        if (pasteBtn) {
            pasteBtn.onclick = async (e) => {
                e.preventDefault();
                e.stopPropagation();
                try {
                    const items = await navigator.clipboard.read();
                    let found = false;
                    for (const item of items) {
                        const imgType = item.types.find(t => t.startsWith('image/'));
                        if (imgType) {
                            found = true;
                            const blob = await item.getType(imgType);
                            const result = await self.uploadImage(blob);
                            if (result) {
                                const newItem = {};
                                newItem[imagePathKey] = result.file_path;
                                newItem[itemFieldKey] = '';
                                group.items.push(newItem);
                                self._addGroupedImageItem(group, newItem, previewContainer,
                                                          itemFieldKey, imagePathKey, itemLabel, result);
                            }
                        }
                    }
                    if (!found && window.AppUtils) {
                        AppUtils.showToast("剪贴板中未发现图片", "info");
                    }
                } catch (err) {
                    if (window.AppUtils) AppUtils.showToast("无法读取剪贴板", "error");
                }
            };
        }
    },

    // 添加分组图片列表中的图片项
    _addGroupedImageItem(group, item, container, itemFieldKey, imagePathKey, itemLabel, imageInfo) {
        const fullUrl = imageInfo
            ? `${window.AppAPI.BASE_URL}${imageInfo.url}`
            : `${window.AppAPI.BASE_URL}/temp/${item[imagePathKey].split(/[/\\]/).pop()}`;

        const wrapper = document.createElement('div');
        wrapper.className = 'evidence-item';

        const imgBox = document.createElement('div');
        const img = document.createElement('img');
        img.src = fullUrl;
        img.onclick = () => this.openImagePreview(img.src, item[itemFieldKey] || '截图预览');
        imgBox.appendChild(img);

        const infoBox = document.createElement('div');
        infoBox.className = 'evidence-info-box';

        const label = document.createElement('label');
        label.innerText = itemLabel + ':';
        label.style.marginBottom = '5px';

        const textarea = document.createElement('textarea');
        textarea.rows = 4;
        textarea.className = 'evidence-textarea';
        textarea.placeholder = '请输入' + itemLabel;
        textarea.value = item[itemFieldKey] || '';
        textarea.addEventListener('input', (e) => { item[itemFieldKey] = e.target.value; });

        const delBtn = document.createElement('button');
        delBtn.type = 'button';
        delBtn.innerText = '删除';
        delBtn.className = 'btn-delete';
        delBtn.style.marginTop = '25px';
        delBtn.style.alignSelf = 'flex-start';
        delBtn.onclick = () => {
            const idx = group.items.indexOf(item);
            if (idx > -1) group.items.splice(idx, 1);
            wrapper.remove();
        };

        infoBox.appendChild(label);
        infoBox.appendChild(textarea);

        wrapper.appendChild(imgBox);
        wrapper.appendChild(infoBox);
        wrapper.appendChild(delBtn);
        container.appendChild(wrapper);
    },

    handleChange(field, value, triggerEvent = 'change') {
        if (this.currentSchema) {
            this.currentSchema.fields.forEach(f => {
                if (f.computed && f.compute_from === field.key && f.compute_rule) {
                    this.setFieldValue(f.key, f.compute_rule[value] || '');
                }
            });
        }

        const trigger = String(triggerEvent || 'change').toLowerCase();
        const fieldBehaviors = this.behaviors[field.key];
        if (fieldBehaviors) {
            const behaviorList = Array.isArray(fieldBehaviors) ? fieldBehaviors : [fieldBehaviors];
            behaviorList.forEach((beh) => {
                const expectedEvent = String(
                    (beh && beh.trigger && beh.trigger.event)
                    || beh?.trigger_event
                    || 'change'
                ).toLowerCase();
                if (expectedEvent === trigger && beh.actions) {
                    this.runActions(beh.actions, value);
                }
            });
        }
        
        // 配置驱动的字段联动 - 替代硬编码的模板判断
        this.handleDependentFieldUpdate(field.key, value);

        // Schema-driven behavior trigger: scan behaviors with trigger.fields (multi-field arrays)
        const schemaBehaviors = this.currentSchema?.behaviors || [];
        for (const beh of schemaBehaviors) {
            if (beh.trigger && Array.isArray(beh.trigger.fields) && beh.trigger.fields.includes(field.key)) {
                this.executeBehaviorFromSchema(beh);
            }
        }
    },

    // 发出 data_changed 事件 — 数组/列表字段内容变更时触发关联行为
    emitDataChanged(fieldKey) {
        const behaviors = this.dataChangedBehaviors[fieldKey];
        if (!behaviors || !behaviors.length) return;
        behaviors.forEach((beh) => {
            if (beh.actions) {
                this.runActions(beh.actions, this.formData[fieldKey]);
            }
        });
    },

    // ── Template CSS loading ──────────────────────────────────────────

    // Dynamic CSS injection for template-specific styles
    loadTemplateStyle(templateId) {
        const baseUrl = window.AppAPI?.BASE_URL || '';
        const cssUrl = `${baseUrl}/api/templates/${templateId}/widgets/style.css`;
        // Remove previously injected template CSS
        document.querySelectorAll('link[data-template-css]').forEach(el => el.remove());
        // Inject new template CSS
        const link = document.createElement('link');
        link.rel = 'stylesheet';
        link.href = cssUrl;
        link.setAttribute('data-template-css', templateId);
        link.onerror = () => link.remove();
        document.head.appendChild(link);
    },

    // ── Widget System (PR2.2) ────────────────────────────────────────

    // Load a template widget JS file and instantiate it
    async loadTemplateWidget(field) {
        const widgetFile = field.widget || 'widget.js';
        const templateId = this.currentSchema?.id;
        if (!templateId) return this._createWidgetError('No template loaded');

        try {
            const baseUrl = window.AppAPI?.BASE_URL || '';
            let url = `${baseUrl}/api/templates/${templateId}/widgets/${widgetFile}`;
            let textResponse = await fetch(url);

            // Fallback: if template-local widget not found, try shared widget
            if (!textResponse.ok) {
                const sharedUrl = `${baseUrl}/api/widgets/shared/${widgetFile}`;
                textResponse = await fetch(sharedUrl);
            }

            if (!textResponse.ok) throw new Error(`Failed to load widget: ${textResponse.status}`);
            const code = await textResponse.text();

            if (!window.__widgetFactories) window.__widgetFactories = {};
            new Function('code', code)(code);

            const factoryFn = window.__widgetFactories[field.widget.replace(/\.js$/, '')];
            if (!factoryFn) {
                return this._createWidgetError(`Widget '${field.widget}' did not register a factory`);
            }

            const callbacks = this._createWidgetCallbacks(field);
            const result = factoryFn(field, callbacks);
            return result?.container || result;
        } catch (e) {
            console.error('Widget load error:', e);
            return this._createWidgetError(`Widget error: ${e.message}`);
        }
    },

    _createWidgetCallbacks(field) {
        const self = this;
        return {
            getData: () => self.formData[field.key] || [],
            setData: (data) => {
                self.formData[field.key] = data;
                self.emitDataChanged(field.key);
                self.updateAllSummaryConfigs();
                self.handleDependentFieldUpdate(field.key, data);
            },
            getFormValue: (key) => self.formData[key],
            setFormValue: (key, value) => self.setFieldValue(key, value),
            uploadImage: (file) => self.uploadImage(file),
            apiRequest: async (endpoint, method, body) => {
                if (endpoint && endpoint.startsWith('@service/')) {
                    endpoint = await self.resolveServiceEndpoint(endpoint);
                }
                return AppAPI._request(endpoint, method || 'GET', body || null);
            },
            dataSources: self.dataSources,
            getConfig: (key) => {
                if (key === 'BASE_URL') return window.AppAPI?.BASE_URL || '';
                return (window.AppConfig && window.AppConfig.THEME && window.AppConfig.THEME[key]) || {};
            },
            toast: (msg) => window.AppUtils?.showToast?.(msg),
            openImagePreview: (src, title) => self.openImagePreview(src, title)
        };
    },

    _createWidgetError(msg) {
        const el = document.createElement('div');
        el.className = 'widget-error';
        el.style.cssText = 'color: #ff4d4f; padding: 12px; border: 1px solid #ff4d4f; border-radius: 4px; background: #fff2f0;';
        el.textContent = msg;
        return el;
    },

    resolveBehaviorTemplate(value) {
        if (typeof value === 'string') {
            return value.replace(/\$\{(\w+)\}/g, (_, key) => this.formData[key] || '');
        }
        if (Array.isArray(value)) {
            return value.map((item) => this.resolveBehaviorTemplate(item));
        }
        if (value && typeof value === 'object') {
            const resolved = {};
            Object.entries(value).forEach(([k, v]) => {
                resolved[k] = this.resolveBehaviorTemplate(v);
            });
            return resolved;
        }
        return value;
    },
    
    async resolveServiceEndpoint(endpoint) {
        if (!endpoint || typeof endpoint !== 'string' || !endpoint.startsWith('@service/')) {
            return endpoint;
        }
        
        // Lazy-load service map from backend (cached in memory)
        if (!this.serviceMap) {
            try {
                const data = await AppAPI._request('/api/services', 'GET');
                this.serviceMap = data.services || {};
            } catch (e) {
                console.error('[FormRenderer] Failed to load service map:', e);
                return endpoint; // Fallback: use raw endpoint (may still work if absolute URL)
            }
        }
        
        // Parse "@service/name/path/to/resource?query" → service name + remaining path
        const prefixEnd = endpoint.indexOf('/', '@service/'.length);
        const serviceName = prefixEnd > 0 ? endpoint.substring(0, prefixEnd) : endpoint;
        const rest = prefixEnd > 0 ? endpoint.substring(prefixEnd) : ''; // includes leading /
        
        const mapped = this.serviceMap[serviceName];
        if (!mapped) {
            console.warn('[FormRenderer] Unknown service:', serviceName);
            return endpoint;
        }
        
        return rest ? `${mapped}${rest}` : mapped;
    },
    
    // 配置驱动的字段联动更新
    handleDependentFieldUpdate(triggerKey, value) {
        const dependentFields = this.currentSchema?.dependent_fields;
        if (!dependentFields) return;
        
        // 遍历所有依赖字段配置
        for (const [targetKey, config] of Object.entries(dependentFields)) {
            // 检查是否由当前字段触发
            const triggers = config.trigger_fields || [config.trigger_field];
            if (!triggers.includes(triggerKey)) continue;
            
            // reapply_presets: 重新应用目标字段的 presets（用于含占位符的 preset 文本刷新）
            if (config.reapply_presets) {
                const targetField = this.currentSchema?.fields?.find(f => f.key === targetKey);
                const currentValue = this.formData[targetKey];
                if (targetField && targetField.presets && targetField.presets[currentValue]) {
                    this.applyPresets(targetField.presets[currentValue]);
                }
            }
            
            // 如果是自动生成类型，调用对应的生成方法
            if (config.auto_generate) {
                this.updateAutoGeneratedField(targetKey, config);
            } else if (config.template) {
                // 模板替换类型
                this.updateTemplateField(targetKey, config.template);
            }
        }
    },
    
    // 更新模板字段（替换占位符）
    updateTemplateField(targetKey, template) {
        const el = document.getElementById(targetKey);
        if (!el) return;
        
        const rendered = template.replace(/\$\{(\w+)\}/g, (_, key) => this.formData[key] || '');
        el.value = rendered;
        this.formData[targetKey] = rendered;
    },
    
    // 更新自动生成字段（配置驱动，通过 dependent_fields 的 computed_template）
    updateAutoGeneratedField(targetKey, config) {
        if (config.computed_template) {
            // Pre-compute derived variables (${_effective_high_vulns} etc.)
            // before template substitution — these are not in formData directly
            this._computeTemplateVars(targetKey, config);
            this.updateTemplateField(targetKey, config.computed_template);
        }
    },
    
    // 更新指定字段中的占位符
    updatePlaceholderInField(fieldKey, placeholder, newValue) {
        const el = document.getElementById(fieldKey);
        if (el && el.value && el.value.includes(placeholder)) {
            el.value = el.value.replace(new RegExp(placeholder, 'g'), newValue || '');
            this.formData[fieldKey] = el.value;
        }
    },

    // 自动计算风险评级
    async runActions(actions, value, _chain = new Set()) {
        for (const a of actions) {
            if (a.type === 'compute' && a.target) {
                if (a.function) {
                    // Built-in compute function: count_by, sum, concat, unique
                    const result = this.executeComputeFunction(a.function, a.args, a.template);
                    this.setFieldValue(a.target, result);
                } else if (a.rules) {
                    this.setFieldValue(a.target, a.rules[value] || '');
                } else if (a.expression) {
                    // Resolve ${field_key} and $field tokens from formData
                    let computed = a.expression.replace(/\$\{(\w+)\}/g, (_, k) => this.formData[k] || '');
                    computed = computed.replace(/\$(\w+)/g, (_, k) => this.formData[k] !== undefined ? String(this.formData[k]) : '');

                    // Check if expression is purely arithmetic after resolution
                    if (/^[\d\s\.\+\-\*\/\(\)]+$/.test(computed)) {
                        try {
                            const arithResult = Function('"use strict"; return (' + computed + ')')();
                            this.setFieldValue(a.target, String(arithResult));
                        } catch (_e) {
                            this.setFieldValue(a.target, computed);
                        }
                    } else {
                        // String template — set as-is after field resolution
                        this.setFieldValue(a.target, computed);
                    }
                }
            } else if (a.type === 'trigger_behavior' && a.target) {
                // Chaining: trigger another behavior by ID
                // Guard against circular chains (A→B→A)
                if (_chain.has(a.target)) {
                    console.warn('[FormRenderer] Circular behavior chain detected:', a.target, Array.from(_chain));
                    continue;
                }
                const targetBeh = this.behaviorById[a.target];
                if (targetBeh && targetBeh.actions) {
                    _chain.add(a.target);
                    await this.runActions(targetBeh.actions, null, _chain);
                    _chain.delete(a.target);
                } else {
                    console.warn('[FormRenderer] trigger_behavior target not found:', a.target);
                }
            } else if (a.type === 'api_call' && a.endpoint) {
                try {
                    // Resolve @service/ prefix if present
                    let ep = a.endpoint.startsWith('@service/')
                        ? await this.resolveServiceEndpoint(a.endpoint)
                        : a.endpoint;
                    ep = this.resolveBehaviorTemplate(ep);
                    const hasParams = a.params && typeof a.params === 'object' && Object.keys(a.params).length > 0;
                    const payload = hasParams ? this.resolveBehaviorTemplate(a.params) : null;
                    const method = String(a.method || (payload ? 'POST' : 'GET')).toUpperCase();

                    if (method === 'GET' && payload) {
                        const search = new URLSearchParams();
                        Object.entries(payload).forEach(([key, raw]) => {
                            if (raw === undefined || raw === null) {
                                return;
                            }
                            if (Array.isArray(raw) || (raw && typeof raw === 'object')) {
                                search.append(key, JSON.stringify(raw));
                                return;
                            }
                            search.append(key, String(raw));
                        });
                        const query = search.toString();
                        if (query) {
                            ep += (ep.includes('?') ? '&' : '?') + query;
                        }
                    }

                    const data = await AppAPI._request(ep, method, method === 'GET' ? null : payload);
                    if (a.result_mapping) {
                        Object.entries(a.result_mapping).forEach(([src, tgt]) => {
                            const v = src.split('.').reduce((o, p) => o && o[p], data);
                            if (v != null) this.setFieldValue(tgt, v);
                        });
                    }
                } catch (e) {
                    console.error('[FormRenderer] API call failed:', e);
                }
            }
        }
    },

    // 内置计算函数 — 供 compute action 的 function 字段调用
    executeComputeFunction(funcName, args, template) {
        args = args || {};
        switch (funcName) {
            case 'count_by': {
                // Count array items where item[prop] === value
                const arr = this.formData[args.field] || [];
                const prop = args.prop;
                const value = args.value;
                let count = 0;
                arr.forEach(item => {
                    if (item[prop] === value) count++;
                });
                return String(count);
            }
            case 'sum': {
                // Sum numeric values of prop across array items
                const arr = this.formData[args.field] || [];
                const prop = args.prop;
                const total = arr.reduce((sum, item) => sum + (parseFloat(item[prop]) || 0), 0);
                return String(total);
            }
            case 'concat': {
                // Concatenate field values, optionally with a template
                const fields = args.fields || [];
                if (template) {
                    let result = template;
                    fields.forEach((f, i) => {
                        result = result.replace(new RegExp('\\{' + f + '\\}', 'g'),
                            this.formData[f] || '');
                    });
                    // Also replace bare $field tokens
                    result = result.replace(/\$(\w+)/g, (_, k) => this.formData[k] || '');
                    return result;
                }
                return fields.map(f => this.formData[f] || '').join('');
            }
            case 'unique_count': {
                // Count unique values of prop across array items
                const arr = this.formData[args.field] || [];
                const prop = args.prop;
                const seen = new Set();
                arr.forEach(item => {
                    const val = item[prop];
                    if (val != null && val !== '') seen.add(String(val));
                });
                return String(seen.size);
            }
            default:
                console.warn('[FormRenderer] Unknown compute function:', funcName);
                return '';
        }
    },

    setDefaultValues(schema) {
        // 1. 先设置所有默认值和自动生成值
        (schema.fields || []).forEach(f => {
            let v = f.default;
            if (v === 'today') v = new Date().toISOString().split('T')[0];
            
            // 处理自动生成字段 (兼容布尔值和字符串 "true")
            const shouldAutoGenerate = f.auto_generate === true || f.auto_generate === 'true';
            if (shouldAutoGenerate && f.auto_generate_rule) {
                v = this.generateAutoValue(f.auto_generate_rule);
            }
            
            if (v) { this.formData[f.key] = v; const el = document.getElementById(f.key); if (el) el.value = v; }
        });
        
        // 2. 再处理计算字段（根据已设置的默认值计算）
        (schema.fields || []).forEach(f => {
            if (f.computed && f.compute_from && f.compute_rule) {
                const sourceValue = this.formData[f.compute_from];
                if (sourceValue && f.compute_rule[sourceValue]) {
                    const computedValue = f.compute_rule[sourceValue];
                    this.formData[f.key] = computedValue;
                    const el = document.getElementById(f.key);
                    if (el) el.value = computedValue;
                }
            }
        });
        
        // 3. 处理 presets（根据默认值应用预设）
        (schema.fields || []).forEach(f => {
            if (f.presets && f.default && f.presets[f.default]) {
                this.applyPresets(f.presets[f.default]);
            }
        });
    },
    
    // 根据规则自动生成值
    generateAutoValue(rule) {
        let result = rule;
        
        const now = new Date();
        const year = now.getFullYear();
        const month = String(now.getMonth() + 1).padStart(2, '0');
        const day = String(now.getDate()).padStart(2, '0');
        
        // 替换 {date:YYYYMMDD} 格式
        result = result.replace('{date:YYYYMMDD}', `${year}${month}${day}`);
        
        // 替换 {date:YYYY-MM-DD} 格式
        result = result.replace('{date:YYYY-MM-DD}', `${year}-${month}-${day}`);
        
        // 替换 {date} 为当前日期 YYYYMMDD
        result = result.replace('{date}', `${year}${month}${day}`);
        
        // 替换 {seq:N} 为 N 位序号（使用随机数模拟）
        result = result.replace(/\{seq:(\d+)\}/g, (match, digits) => {
            const n = parseInt(digits);
            const rand = Math.floor(Math.random() * Math.pow(10, n));
            return String(rand).padStart(n, '0');
        });
        
        // 替换 {random:N} 为 N 位随机数字
        result = result.replace(/\{random:(\d+)\}/g, (match, digits) => {
            const n = parseInt(digits);
            let rand = '';
            for (let i = 0; i < n; i++) {
                rand += Math.floor(Math.random() * 10);
            }
            return rand;
        });
        
        // 替换 {timestamp} 为时间戳
        result = result.replace('{timestamp}', Date.now().toString());
        
        // 替换 {uuid} 为简短 UUID
        result = result.replace('{uuid}', this.generateShortUUID());
        
        return result;
    },
    
    // 生成简短 UUID
    generateShortUUID() {
        return 'xxxx-xxxx'.replace(/x/g, () => {
            return Math.floor(Math.random() * 16).toString(16);
        });
    },

    async populateDataSources() {
        if (!this.currentSchema || !this.dynamicFormContainer) return;
        
        // 处理普通 select
        const selects = this.dynamicFormContainer.querySelectorAll('select[data-source]');
        for (const sel of selects) {
            const src = sel.dataset.source;
            
            // 只有当数据源存在且有数据时才覆盖选项
            if (this.dataSources[src] && Array.isArray(this.dataSources[src]) && this.dataSources[src].length > 0) {
                // 处理 risk_levels 格式的数据源 (带 value/label/color)
                let opts = this.dataSources[src].map(i => {
                    if (typeof i === 'object') {
                        // 支持 {value, label, color} 格式 (risk_levels)
                        if (i.value !== undefined) {
                            return { v: i.value, t: i.label || i.value, color: i.color };
                        }
                        // 支持 {id, name} 格式 (漏洞库等)
                        return { v: i.id || i.name, t: i.name || i.id };
                    }
                    return { v: i, t: i };
                });
                
                const ph = sel.querySelector('option[value=""]');
                sel.innerHTML = '';
                if (ph) sel.appendChild(ph);
                else { const e = document.createElement('option'); e.value=''; e.textContent='-- 请选择 --'; sel.appendChild(e); }
                opts.forEach(o => { 
                    const op = document.createElement('option'); 
                    op.value = o.v; 
                    op.textContent = o.t;
                    // 如果有颜色，设置选项样式
                    if (o.color) {
                        op.style.color = o.color;
                        op.dataset.color = o.color;
                    }
                    sel.appendChild(op); 
                });
                
                // 如果是 searchable_select 容器内的 select，保存选项到容器
                const container = sel.closest('.searchable-select-container');
                if (container) {
                    container._allOptions = opts.map(o => ({ value: o.v, text: o.t, color: o.color }));
                }

                // 恢复 setDefaultValues 设置的默认值（options 重建会被清掉）
                const savedVal = this.formData[sel.id];
                if (savedVal !== undefined && savedVal !== null && savedVal !== '') {
                    sel.value = savedVal;
                }
            }
            // 如果数据源不存在或为空，保留 schema 中定义的静态 options
        }
    },

    collectFormData() {
        const data = {...this.formData};
        if (this.currentSchema) {
            this.currentSchema.fields.forEach(f => {
                // 图片类型字段数据已在 formData 中，不需要从 DOM 获取
                if (f.type === 'image' || f.type === 'image_list' || f.type === 'grouped_image_list') {
                    return;
                }
                // widget 类型：数据已在 formData 中（由 widget 通过 setData 回调管理）
                if (f.type === 'widget') {
                    return;
                }
                // array 类型：数据已在 formData 中
                if (f.type === 'array') {
                    return;
                }
                // checkbox_group 类型：将选中的 ID 转换为描述文本
                if (f.type === 'checkbox_group') {
                    // 直接从文本框获取数据
                    const textarea = document.getElementById(f.key);
                    data[f.key] = textarea ? textarea.value : '';
                    return;
                }
                // checkbox 类型：数据已在 formData 中（通过 change 事件更新）
                if (f.type === 'checkbox') {
                    // 不需要从 DOM 获取，formData 中已有正确值
                    return;
                }
                const el = document.getElementById(f.key);
                if (el && el.value !== undefined) {
                    data[f.key] = el.value || '';
                }
            });
        }
        return data;
    },

    // 创建单个复选框（开关）
    createCheckbox(field) {
        const wrapper = document.createElement('div');
        wrapper.className = 'checkbox-single-wrapper';
        wrapper.style.cssText = 'display: flex; align-items: center; gap: 10px;';
        
        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.id = field.key;
        checkbox.name = field.key;
        
        // 设置默认值
        const defaultVal = field.default === true || field.default === 'true';
        checkbox.checked = defaultVal;
        this.formData[field.key] = defaultVal;
        
        checkbox.addEventListener('change', (e) => {
            this.formData[field.key] = e.target.checked;
            this.handleChange(field, e.target.checked);
        });
        
        const label = document.createElement('span');
        label.textContent = field.help_text || '';
        label.style.cssText = 'color: #666; font-size: 13px;';
        
        wrapper.appendChild(checkbox);
        wrapper.appendChild(label);
        
        return wrapper;
    },

    // 创建通用数组字段
    createArrayField(field) {
        const container = document.createElement('div');
        container.className = 'array-field-container';
        container.id = field.key;
        
        // 初始化数据
        if (!this.formData[field.key]) {
            this.formData[field.key] = [];
        }
        
        // 表格容器
        const tableWrapper = document.createElement('div');
        tableWrapper.className = 'array-table-wrapper';
        tableWrapper.style.cssText = 'overflow-x: auto; margin-bottom: 10px;';
        
        // 创建表格
        const table = document.createElement('table');
        table.className = 'array-field-table';
        table.style.cssText = 'width: 100%; border-collapse: collapse; font-size: 14px;';
        
        // 表头
        const thead = document.createElement('thead');
        const headerRow = document.createElement('tr');
        
        // 编号列
        const thNum = document.createElement('th');
        thNum.style.cssText = 'width: 50px; padding: 8px; border: 1px solid #ddd; background: #f5f5f5;';
        thNum.textContent = '编号';
        headerRow.appendChild(thNum);
        
        // 根据 columns 配置生成表头
        if (field.columns && Array.isArray(field.columns)) {
            field.columns.forEach(col => {
                const th = document.createElement('th');
                const width = col.width || 'auto';
                th.style.cssText = `width: ${width}; padding: 8px; border: 1px solid #ddd; background: #f5f5f5;`;
                th.textContent = col.label || col.key;
                headerRow.appendChild(th);
            });
        }
        
        // 操作列
        const thAction = document.createElement('th');
        thAction.style.cssText = 'width: 60px; padding: 8px; border: 1px solid #ddd; background: #f5f5f5;';
        thAction.textContent = '操作';
        headerRow.appendChild(thAction);
        
        thead.appendChild(headerRow);
        table.appendChild(thead);
        
        // 表体
        const tbody = document.createElement('tbody');
        tbody.id = `${field.key}_tbody`;
        table.appendChild(tbody);
        
        tableWrapper.appendChild(table);
        container.appendChild(tableWrapper);
        
        // 按钮容器
        const btnContainer = document.createElement('div');
        btnContainer.style.cssText = 'display: flex; gap: 10px;';
        
        // 添加按钮
        const addBtn = document.createElement('button');
        addBtn.type = 'button';
        addBtn.className = 'btn btn-secondary';
        addBtn.style.cssText = 'padding: 6px 15px; font-size: 13px;';
        addBtn.innerHTML = `+ 添加${field.label || '记录'}`;
        addBtn.onclick = () => this.addArrayRow(field.key, field);
        btnContainer.appendChild(addBtn);
        
        // 批量导入按钮
        const batchBtn = document.createElement('button');
        batchBtn.type = 'button';
        batchBtn.className = 'btn btn-secondary';
        batchBtn.style.cssText = 'padding: 6px 15px; font-size: 13px; background: #52c41a; color: white;';
        batchBtn.innerHTML = '📋 批量导入';
        batchBtn.onclick = () => this.openBatchImportModal(field);
        btnContainer.appendChild(batchBtn);
        
        container.appendChild(btnContainer);
        
        // 默认添加一行
        setTimeout(() => this.addArrayRow(field.key, field), 0);
        
        return container;
    },

    // 添加数组字段行
    addArrayRow(fieldKey, field) {
        const tbody = document.getElementById(`${fieldKey}_tbody`);
        if (!tbody) return;
        
        const rowIndex = this.formData[fieldKey].length;
        const rowData = {};
        
        // 初始化行数据
        if (field.columns && Array.isArray(field.columns)) {
            field.columns.forEach(col => {
                rowData[col.key] = col.default || '';
            });
        }
        
        this.formData[fieldKey].push(rowData);
        
        const tr = document.createElement('tr');
        tr.dataset.index = rowIndex;
        
        // 编号列
        const tdNum = document.createElement('td');
        tdNum.style.cssText = 'padding: 6px; border: 1px solid #ddd; text-align: center;';
        tdNum.textContent = rowIndex + 1;
        tr.appendChild(tdNum);
        
        // 根据 columns 配置生成单元格
        if (field.columns && Array.isArray(field.columns)) {
            field.columns.forEach(col => {
                tr.appendChild(this.createArrayCell(fieldKey, rowIndex, col, rowData));
            });
        }
        
        // 删除按钮
        const tdDel = document.createElement('td');
        tdDel.style.cssText = 'padding: 6px; border: 1px solid #ddd; text-align: center;';
        const delBtn = document.createElement('button');
        delBtn.type = 'button';
        delBtn.className = 'btn-mini';
        delBtn.style.cssText = 'background: #ff4d4f; color: white; border: none; padding: 4px 8px; cursor: pointer; border-radius: 3px;';
        delBtn.textContent = '删除';
        delBtn.onclick = () => {
            const currentIdx = parseInt(tr.dataset.index);
            this.removeArrayRow(fieldKey, tr, currentIdx);
        };
        tdDel.appendChild(delBtn);
        tr.appendChild(tdDel);
        
        tbody.appendChild(tr);

        // 发出 data_changed 事件 — 驱动行为系统
        this.emitDataChanged(fieldKey);

        // 配置驱动的汇总更新
        this.updateAllSummaryConfigs();
        this.handleDependentFieldUpdate(fieldKey, this.formData[fieldKey]);
    },

    // 创建数组字段单元格
    createArrayCell(fieldKey, rowIndex, col, rowData) {
        const td = document.createElement('td');
        td.style.cssText = 'padding: 4px; border: 1px solid #ddd;';
        
        if (col.type === 'select') {
            // 下拉框
            const select = document.createElement('select');
            select.style.cssText = 'width: 100%; padding: 6px; border: 1px solid #ccc; border-radius: 3px; box-sizing: border-box;';
            
            // 添加选项
            if (col.options && Array.isArray(col.options)) {
                col.options.forEach(opt => {
                    const option = document.createElement('option');
                    option.value = typeof opt === 'object' ? opt.value : opt;
                    option.textContent = typeof opt === 'object' ? (opt.label || opt.value) : opt;
                    if (option.value === rowData[col.key]) {
                        option.selected = true;
                    }
                    select.appendChild(option);
                });
            }
            
            select.addEventListener('change', (e) => {
                rowData[col.key] = e.target.value;
                // 配置驱动：检查是否有 summary_config 与此列关联
                this.triggerSummaryUpdateForColumn(col.key);
            });
            
            td.appendChild(select);
        } else if (col.type === 'textarea') {
            // 多行文本框
            const textarea = document.createElement('textarea');
            textarea.style.cssText = 'width: 100%; padding: 6px; border: 1px solid #ccc; border-radius: 3px; box-sizing: border-box; resize: vertical; min-height: 60px;';
            textarea.placeholder = col.placeholder || '';
            textarea.value = rowData[col.key] || '';
            textarea.rows = col.rows || 3;
            
            textarea.addEventListener('input', (e) => {
                rowData[col.key] = e.target.value;
            });
            
            td.appendChild(textarea);
        } else if (col.type === 'image') {
            // 图片上传
            td.appendChild(this.createArrayImageCell(fieldKey, rowIndex, col, rowData));
        } else {
            // 文本输入框（默认）
            const input = document.createElement('input');
            input.type = 'text';
            input.style.cssText = 'width: 100%; padding: 6px; border: 1px solid #ccc; border-radius: 3px; box-sizing: border-box;';
            input.placeholder = col.placeholder || '';
            input.value = rowData[col.key] || '';
            
            input.addEventListener('input', (e) => {
                rowData[col.key] = e.target.value;
                // 配置驱动：检查是否有 summary_config 与此列关联
                this.triggerSummaryUpdateForColumn(col.key);
                this.handleDependentFieldUpdate(fieldKey, this.formData[fieldKey]);
            });
            
            td.appendChild(input);
        }
        
        return td;
    },
    
    // 创建数组字段中的图片上传单元格
    createArrayImageCell(fieldKey, rowIndex, col, rowData) {
        const self = this;
        const container = document.createElement('div');
        container.className = 'array-image-cell';
        container.style.cssText = 'display: flex; flex-direction: column; gap: 5px; min-width: 120px;';
        
        // 预览区域
        const preview = document.createElement('div');
        preview.className = 'array-image-preview';
        preview.style.cssText = 'width: 100%; min-height: 60px; border: 1px dashed #ccc; border-radius: 4px; display: flex; align-items: center; justify-content: center; cursor: pointer; background: #fafafa;';
        
        // 如果已有图片，显示预览
        if (rowData[col.key]) {
            const img = document.createElement('img');
            img.src = `${window.AppAPI.BASE_URL}/temp/${rowData[col.key].split('/').pop()}`;
            img.style.cssText = 'max-width: 100%; max-height: 80px; object-fit: contain;';
            img.onclick = () => this.openImagePreview(img.src, col.label || '图片预览');
            preview.appendChild(img);
        } else {
            preview.innerHTML = '<span style="color: #999; font-size: 12px;">点击上传</span>';
        }
        
        // 按钮区域
        const btnRow = document.createElement('div');
        btnRow.style.cssText = 'display: flex; gap: 5px;';
        
        // 上传按钮
        const uploadBtn = document.createElement('button');
        uploadBtn.type = 'button';
        uploadBtn.className = 'btn-mini';
        uploadBtn.style.cssText = 'flex: 1; padding: 4px 8px; font-size: 11px; background: #1890ff; color: white; border: none; border-radius: 3px; cursor: pointer;';
        uploadBtn.textContent = '上传';
        
        // 粘贴按钮
        const pasteBtn = document.createElement('button');
        pasteBtn.type = 'button';
        pasteBtn.className = 'btn-mini';
        pasteBtn.style.cssText = 'flex: 1; padding: 4px 8px; font-size: 11px; background: #52c41a; color: white; border: none; border-radius: 3px; cursor: pointer;';
        pasteBtn.textContent = '粘贴';
        
        // 上传点击事件
        const handleUpload = () => {
            const input = document.createElement('input');
            input.type = 'file';
            input.accept = col.accept || 'image/*';
            input.onchange = async (e) => {
                const file = e.target.files[0];
                if (file) {
                    const result = await self.uploadImage(file);
                    if (result) {
                        rowData[col.key] = result.file_path;
                        self.updateArrayImagePreview(preview, result, col);
                    }
                }
            };
            input.click();
        };
        
        preview.onclick = handleUpload;
        uploadBtn.onclick = handleUpload;
        
        // 粘贴点击事件
        pasteBtn.onclick = async () => {
            try {
                const items = await navigator.clipboard.read();
                for (const item of items) {
                    const imgType = item.types.find(t => t.startsWith('image/'));
                    if (imgType) {
                        const blob = await item.getType(imgType);
                        const result = await self.uploadImage(blob);
                        if (result) {
                            rowData[col.key] = result.file_path;
                            self.updateArrayImagePreview(preview, result, col);
                        }
                        break;
                    }
                }
            } catch (err) {
                if (window.AppUtils) AppUtils.showToast('无法读取剪贴板', 'error');
            }
        };
        
        btnRow.appendChild(uploadBtn);
        btnRow.appendChild(pasteBtn);
        
        container.appendChild(preview);
        container.appendChild(btnRow);
        
        return container;
    },
    
    // 更新数组字段中的图片预览
    updateArrayImagePreview(preview, imageInfo, col) {
        const fullUrl = `${window.AppAPI.BASE_URL}${imageInfo.url}`;
        preview.innerHTML = '';
        // 移除预览区域的上传点击事件
        preview.onclick = null;
        preview.style.cursor = 'default';
        
        const img = document.createElement('img');
        img.src = fullUrl;
        img.style.cssText = 'max-width: 100%; max-height: 80px; object-fit: contain; cursor: pointer;';
        img.onclick = (e) => {
            e.stopPropagation(); // 阻止冒泡
            this.openImagePreview(img.src, col.label || '图片预览');
        };
        preview.appendChild(img);
    },
    
    // 打开批量导入弹窗
    openBatchImportModal(field) {
        // 确保弹窗存在
        this.ensureBatchImportModal();
        
        const modal = document.getElementById('batch-import-modal');
        const textarea = document.getElementById('batch-import-textarea');
        const confirmBtn = document.getElementById('batch-import-confirm');
        const formatHint = document.getElementById('batch-import-format');
        
        // 生成格式提示
        const colNames = (field.columns || []).map(c => c.label || c.key).join(' | ');
        formatHint.textContent = `格式：${colNames}（每行一条，支持Tab/逗号/分号分隔）`;
        
        // 清空文本框
        textarea.value = '';
        
        // 绑定确认事件
        confirmBtn.onclick = () => {
            this.processBatchImport(field, textarea.value);
            modal.style.display = 'none';
        };
        
        modal.style.display = 'flex';
        textarea.focus();
    },
    
    // 确保批量导入弹窗存在
    ensureBatchImportModal() {
        if (document.getElementById('batch-import-modal')) return;
        
        const modal = document.createElement('div');
        modal.id = 'batch-import-modal';
        modal.style.cssText = 'display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.5); z-index: 10000; align-items: center; justify-content: center;';
        
        modal.innerHTML = `
            <div style="background: white; border-radius: 8px; padding: 20px; width: 600px; max-width: 90%; max-height: 80vh; overflow: auto;">
                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 15px;">
                    <h3 style="margin: 0;">📋 批量导入</h3>
                    <button type="button" id="batch-import-close" style="background: none; border: none; font-size: 20px; cursor: pointer;">&times;</button>
                </div>
                <p id="batch-import-format" style="color: #666; font-size: 13px; margin-bottom: 10px;"></p>
                <textarea id="batch-import-textarea" rows="10" style="width: 100%; padding: 10px; border: 1px solid #ddd; border-radius: 4px; font-family: monospace; font-size: 13px; box-sizing: border-box;" placeholder="粘贴数据，每行一条记录..."></textarea>
                <div style="margin-top: 15px; text-align: right;">
                    <button type="button" id="batch-import-cancel" class="btn btn-secondary" style="margin-right: 10px;">取消</button>
                    <button type="button" id="batch-import-confirm" class="btn btn-primary">导入</button>
                </div>
            </div>
        `;
        
        document.body.appendChild(modal);
        
        // 绑定关闭事件
        document.getElementById('batch-import-close').onclick = () => modal.style.display = 'none';
        document.getElementById('batch-import-cancel').onclick = () => modal.style.display = 'none';
        modal.onclick = (e) => { if (e.target === modal) modal.style.display = 'none'; };
    },
    
    // 处理批量导入数据
    processBatchImport(field, text) {
        if (!text.trim()) return;
        
        const lines = text.trim().split('\n');
        const columns = field.columns || [];
        let importedCount = 0;
        
        lines.forEach(line => {
            line = line.trim();
            if (!line) return;
            
            // 智能分隔：优先Tab，其次逗号，再次分号，最后多空格
            let values;
            if (line.includes('\t')) {
                values = line.split('\t');
            } else if (line.includes(',')) {
                values = line.split(',');
            } else if (line.includes(';')) {
                values = line.split(';');
            } else {
                values = line.split(/\s{2,}|\s+/);
            }
            
            // 构建行数据
            const rowData = {};
            columns.forEach((col, idx) => {
                let val = (values[idx] || '').trim();
                // 对于 select 类型，尝试匹配选项
                if (col.type === 'select' && col.options && val) {
                    const matched = col.options.find(opt => {
                        const optVal = typeof opt === 'object' ? opt.value : opt;
                        const optLabel = typeof opt === 'object' ? (opt.label || opt.value) : opt;
                        return optVal.toLowerCase() === val.toLowerCase() || 
                               optLabel.toLowerCase() === val.toLowerCase();
                    });
                    if (matched) {
                        val = typeof matched === 'object' ? matched.value : matched;
                    }
                }
                rowData[col.key] = val || col.default || '';
            });
            
            // 添加到表格
            this.addArrayRowWithData(field.key, field, rowData);
            importedCount++;
        });
        
        // 批量导入后触发汇总更新
        this.updateAllSummaryConfigs();
        this.handleDependentFieldUpdate(field.key, this.formData[field.key]);
        
        if (window.AppUtils) {
            AppUtils.showToast(`成功导入 ${importedCount} 条记录`, 'success');
        }
    },
    
    // 添加带数据的数组行
    addArrayRowWithData(fieldKey, field, rowData) {
        const tbody = document.getElementById(`${fieldKey}_tbody`);
        if (!tbody) return;
        
        const rowIndex = this.formData[fieldKey].length;
        this.formData[fieldKey].push(rowData);
        
        const tr = document.createElement('tr');
        tr.dataset.index = rowIndex;
        
        // 编号列
        const tdNum = document.createElement('td');
        tdNum.style.cssText = 'padding: 6px; border: 1px solid #ddd; text-align: center;';
        tdNum.textContent = rowIndex + 1;
        tr.appendChild(tdNum);
        
        // 根据 columns 配置生成单元格
        if (field.columns && Array.isArray(field.columns)) {
            field.columns.forEach(col => {
                tr.appendChild(this.createArrayCell(fieldKey, rowIndex, col, rowData));
            });
        }
        
        // 删除按钮
        const tdDel = document.createElement('td');
        tdDel.style.cssText = 'padding: 6px; border: 1px solid #ddd; text-align: center;';
        const delBtn = document.createElement('button');
        delBtn.type = 'button';
        delBtn.className = 'btn-mini';
        delBtn.style.cssText = 'background: #ff4d4f; color: white; border: none; padding: 4px 8px; cursor: pointer; border-radius: 3px;';
        delBtn.textContent = '删除';
        delBtn.onclick = () => {
            const currentIdx = parseInt(tr.dataset.index);
            this.removeArrayRow(fieldKey, tr, currentIdx);
        };
        tdDel.appendChild(delBtn);
        tr.appendChild(tdDel);
        
        tbody.appendChild(tr);

        // 发出 data_changed 事件 — 驱动行为系统
        this.emitDataChanged(fieldKey);

        // 配置驱动的汇总更新
        this.updateAllSummaryConfigs();
        this.handleDependentFieldUpdate(fieldKey, this.formData[fieldKey]);
    },

    // 删除数组字段行
    removeArrayRow(fieldKey, tr, rowIndex) {
        const tbody = tr.parentElement;
        
        // 从数据中移除
        this.formData[fieldKey].splice(rowIndex, 1);
        
        // 从 DOM 中移除
        tr.remove();
        
        // 重新编号
        const rows = tbody.querySelectorAll('tr');
        rows.forEach((row, idx) => {
            row.dataset.index = idx;
            row.cells[0].textContent = idx + 1;
        });
        
        // 发出 data_changed 事件 — 驱动行为系统
        this.emitDataChanged(fieldKey);

        // 配置驱动的汇总更新
        this.updateAllSummaryConfigs();
        this.handleDependentFieldUpdate(fieldKey, this.formData[fieldKey]);
    },
    
    // ========== 通用汇总计算器 ==========
    
    /**
     * 通用计数汇总生成器
     * @param {Array} items - 数据列表
     * @param {string} typeKey - 类型字段名
     * @param {Object} config - 汇总配置 {type_names, template_zero, template_with_data, connector, last_connector}
     * @returns {string} 汇总描述
     */
    generateCountSummary(items, typeKey, config) {
        if (!items || items.length === 0) {
            return config.template_zero;
        }
        
        // 统计各类型数量
        const typeCounts = {};
        items.forEach(item => {
            const t = item[typeKey] || '其他';
            typeCounts[t] = (typeCounts[t] || 0) + 1;
        });
        
        // 生成描述部分
        const parts = [];
        for (const [type, count] of Object.entries(typeCounts)) {
            if (count > 0) {
                const name = config.type_names?.[type] || `${type}实例`;
                parts.push(`${count}个${name}`);
            }
        }
        
        if (parts.length === 0) return config.template_zero;
        
        const connector = config.connector || '、';
        const lastConnector = config.last_connector || '以及';
        
        let detail = '';
        if (parts.length === 1) {
            detail = parts[0];
        } else {
            detail = parts.slice(0, -1).join(connector) + lastConnector + parts[parts.length - 1];
        }
        
        return config.template_with_data
            .replace('{total}', items.length)
            .replace('{detail}', detail);
    },
    
    /**
     * 通用数据量汇总生成器
     * @param {Array} items - 数据列表
     * @param {string} typeKey - 类型字段名
     * @param {string} countKey - 数量字段名
     * @param {Object} config - 汇总配置
     * @returns {{summary: string, total: number}}
     */
    generateDataSummary(items, typeKey, countKey, config) {
        if (!items || items.length === 0) {
            return { summary: config.template_zero, total: 0 };
        }
        
        // 按类型汇总数量
        const typeTotals = {};
        let grandTotal = 0;
        
        items.forEach(item => {
            const dataType = item[typeKey] || '未知类型';
            const countStr = String(item[countKey] || '0').replace(/,/g, '');
            const count = parseInt(countStr, 10) || 0;
            typeTotals[dataType] = (typeTotals[dataType] || 0) + count;
            grandTotal += count;
        });
        
        if (grandTotal === 0) {
            return { summary: config.template_zero, total: 0 };
        }
        
        const parts = Object.entries(typeTotals)
            .map(([type, total]) => `${type}共计${total.toLocaleString()}条`);
        
        const detail = parts.join(config.connector || '，');
        const summary = config.template_with_data
            .replace('{total}', grandTotal.toLocaleString())
            .replace('{detail}', detail);
        
        return { summary, total: grandTotal };
    },

    // Pre-compute derived variables for computed_template substitution.
    // Schema-driven: reads pre_compute config from the dependent_fields entry.
    // Supports: count vulns, count servers, count DB, data summary.
    _computeTemplateVars(targetKey, config) {
        // Find the computed_template entry for this targetKey (config is the dependent_field entry)
        const ctEntry = config;
        if (!ctEntry || !ctEntry.pre_compute) return;

        const pc = ctEntry.pre_compute;
        const result = {};

        // Count vulns (aggregate internet + intranet vuln arrays)
        if (pc.count_vulns) {
            const cv = pc.count_vulns;
            const allVulns = [];
            (cv.source || []).forEach(srcKey => {
                const arr = this.formData[srcKey] || [];
                allVulns.push(...arr);
            });

            let criticalHigh = 0;
            const systems = new Set();

            allVulns.forEach(vuln => {
                const level = vuln[cv.level_field] || '中危';
                const urlLines = (vuln[cv.url_field] || '').split('\n').filter(l => l.trim()).length;
                const count = Math.max(1, urlLines);

                if (level === '超危' || level === '高危') {
                    criticalHigh += count;
                }

                const sys = (vuln[cv.system_field] || '').trim();
                if (sys) systems.add(sys);
            });

            result.count_high_critical_vulns = criticalHigh;
            result.unique_systems = systems.size;
        }

        // Count servers
        if (pc.count_servers) {
            result.server_count = (this.formData[pc.count_servers.source] || []).length;
        }

        // Count DB connections
        if (pc.count_db) {
            result.db_count = (this.formData[pc.count_db.source] || []).length;
        }

        // Data summary
        if (pc.data_summary) {
            const ds = pc.data_summary;
            result.total_data = this.formData[ds.total_field] || 0;
            const stats = this.formData[ds.statistics_field] || [];
            const types = new Set();
            stats.forEach(s => {
                const t = (s[ds.type_field] || '').trim();
                if (t) types.add(t);
            });
            result.data_type_list = types.size > 0 ? Array.from(types).sort().join('、') : '';
        }

        // Write output variables into formData
        if (pc.output) {
            for (const [targetVar, sourceVar] of Object.entries(pc.output)) {
                this.formData[targetVar] = result[sourceVar] !== undefined ? result[sourceVar] : '';
            }
        }
    },

    /**
     * 配置驱动的汇总更新
     * @param {string} summaryKey - 汇总字段名
     */
    updateSummaryFromConfig(summaryKey) {
        const summaryConfigs = this.currentSchema?.summary_configs;
        if (!summaryConfigs?.[summaryKey]) return;
        
        const config = summaryConfigs[summaryKey];
        const items = this.formData[config.source_field] || [];
        const summaryEl = document.getElementById(summaryKey);
        if (!summaryEl) return;
        
        let summary = '';
        
        if (config.mode === 'count') {
            summary = this.generateCountSummary(items, config.type_key, config);
        } else if (config.mode === 'data') {
            const result = this.generateDataSummary(items, config.type_key, config.count_key, config);
            summary = result.summary;
            if (config.total_field) {
                this.formData[config.total_field] = result.total;
            }
        }
        
        summaryEl.value = summary;
        this.formData[summaryKey] = summary;
    },
    
    // 根据列名触发关联的汇总更新（通过 summary_configs 的 trigger_columns 配置驱动）
    triggerSummaryUpdateForColumn(columnKey) {
        const summaryConfigs = this.currentSchema?.summary_configs;
        if (!summaryConfigs) return;
        for (const [summaryKey, config] of Object.entries(summaryConfigs)) {
            const triggers = config.trigger_columns || [];
            if (triggers.includes(columnKey)) {
                this.updateSummaryFromConfig(summaryKey);
            }
        }
    },

    // 更新所有汇总字段（配置驱动）
    updateAllSummaryConfigs() {
        const configs = this.currentSchema?.summary_configs;
        if (!configs) return;
        for (const key of Object.keys(configs)) {
            this.updateSummaryFromConfig(key);
        }
    },

    // Schema-driven behavior executor: dispatches action types from behaviors
    // with trigger.fields (multi-field array triggers)
    executeBehaviorFromSchema(beh) {
        if (!beh || !beh.actions) return;
        for (const action of beh.actions) {
            if (action.type === 'risk_compute') {
                this._executeRiskCompute(action);
            }
        }
    },

    // Execute risk_compute action: sum source fields, evaluate thresholds, set risk level
    _executeRiskCompute(action) {
        if (!action.source_fields || !action.thresholds) return;

        // Sum source fields
        let total = 0;
        action.source_fields.forEach(sf => {
            total += parseInt(this.formData[sf] || '0', 10);
        });
        if (action.sum_target) {
            this.setFieldValue(action.sum_target, String(total));
        }

        // Evaluate thresholds in order (rely on YAML dict ordering for priority)
        let riskLevel = null;
        const levels = Object.entries(action.thresholds);
        for (const [level, condition] of levels) {
            if (this._evalRiskCondition(condition, action.source_fields)) {
                riskLevel = level;
                break;
            }
        }
        // Fallback: if no threshold matched, default to the last defined level
        if (!riskLevel) {
            if (levels.length === 0) return;
            riskLevel = levels[levels.length - 1][0];
        }

        const targetKey = action.target_key || 'overall_risk_level';
        const riskField = this.currentSchema?.fields?.find(f => f.key === targetKey);
        if (riskField) {
            this.setFieldValue(targetKey, riskLevel);
            if (riskField.presets && riskField.presets[riskLevel]) {
                this.applyPresets(riskField.presets[riskLevel]);
            }
        }
    },

    // Evaluate a risk condition string like "critical>=1 || high>=1 || medium>6"
    // against formData. Comma = OR, & = AND within a level.
    _evalRiskCondition(condition, sourceFields) {
        if (!condition) return false;
        // Build a safe evaluation: replace each known field key with its numeric value
        let expr = condition;
        for (const sf of sourceFields) {
            const val = parseInt(this.formData[sf] || '0', 10);
            expr = expr.replace(new RegExp('\\b' + sf + '\\b', 'g'), String(val));
        }
        // Remaining bare identifiers → treat as 0 (shouldn't happen with valid configs)
        expr = expr.replace(/\b[a-zA-Z_]\w*\b/g, '0');
        // Comma → ||, & → &&
        expr = expr.replace(/,/g, ' || ').replace(/&/g, ' && ');
        try {
            return !!(Function('"use strict"; return (' + expr + ')')());
        } catch (_e) {
            console.warn('[FormRenderer] Failed to evaluate risk condition:', condition, _e);
            return false;
        }
    },

    // 自动计算风险评级（schema-driven: reads from compute_risk_level behavior config）
    // Kept for backward compatibility — autoCalculateRiskLevel() is also called directly
    // from widget-level submit, but now delegates to the schema behavior if available.
    autoCalculateRiskLevel() {
        const behavior = (this.currentSchema?.behaviors || []).find(b => b.id === 'compute_risk_level');
        if (!behavior) return;
        this.executeBehaviorFromSchema(behavior);
    },

    // 创建多选复选框组（带文本框）
    createCheckboxGroup(field) {
        const wrapper = document.createElement('div');
        wrapper.className = 'checkbox-group-wrapper';
        
        // 左侧：复选框列表
        const checkboxContainer = document.createElement('div');
        checkboxContainer.className = 'checkbox-group-container';
        checkboxContainer.id = field.key + '_checkboxes';
        
        // 初始化选中值数组
        if (!this.formData[field.key]) {
            this.formData[field.key] = [];
        }
        
        // 创建复选框列表
        if (field.options && Array.isArray(field.options)) {
            field.options.forEach(opt => {
                const item = document.createElement('div');
                item.className = 'checkbox-item';
                
                const checkbox = document.createElement('input');
                checkbox.type = 'checkbox';
                checkbox.id = `${field.key}_${opt.id}`;
                checkbox.value = opt.id;
                checkbox.dataset.description = opt.description || '';
                
                const label = document.createElement('label');
                label.htmlFor = checkbox.id;
                label.textContent = opt.label;
                label.title = opt.description || '';
                
                checkbox.addEventListener('change', () => {
                    this.updateCheckboxGroupValue(field.key, field);
                });
                
                item.appendChild(checkbox);
                item.appendChild(label);
                checkboxContainer.appendChild(item);
            });
        }
        
        // 右侧：文本框
        const textarea = document.createElement('textarea');
        textarea.id = field.key;
        textarea.className = 'checkbox-group-textarea';
        textarea.rows = 10;
        textarea.placeholder = '选中左侧选项后自动填充，也可直接编辑';
        textarea.addEventListener('input', () => {
            this.formData[field.key + '_text'] = textarea.value;
        });
        
        wrapper.appendChild(checkboxContainer);
        wrapper.appendChild(textarea);
        
        return wrapper;
    },

    // 更新复选框组的值和文本框
    updateCheckboxGroupValue(fieldKey, field) {
        const container = document.getElementById(fieldKey + '_checkboxes');
        if (!container) return;
        
        const checkboxes = container.querySelectorAll('input[type="checkbox"]:checked');
        const selectedIds = Array.from(checkboxes).map(cb => cb.value);
        this.formData[fieldKey] = selectedIds;
        
        // 更新文本框内容
        const textarea = document.getElementById(fieldKey);
        if (textarea && field && field.options) {
            const descriptions = selectedIds.map((id, index) => {
                const opt = field.options.find(o => o.id === id);
                return opt ? `${index + 1}、${opt.description}` : '';
            }).filter(d => d);
            textarea.value = descriptions.join('\n');
        }
    },

    // 设置复选框组的选中状态
    setCheckboxGroupValue(fieldKey, selectedIds) {
        const container = document.getElementById(fieldKey + '_checkboxes');
        if (!container || !Array.isArray(selectedIds)) return;
        
        // 先取消所有选中
        container.querySelectorAll('input[type="checkbox"]').forEach(cb => {
            cb.checked = false;
        });
        
        // 选中指定项
        selectedIds.forEach(id => {
            const cb = container.querySelector(`input[value="${id}"]`);
            if (cb) cb.checked = true;
        });
        
        this.formData[fieldKey] = selectedIds;
        
        // 更新文本框内容
        const field = this.currentSchema?.fields?.find(f => f.key === fieldKey);
        const textarea = document.getElementById(fieldKey);
        if (textarea && field && field.options) {
            const descriptions = selectedIds.map((id, index) => {
                const opt = field.options.find(o => o.id === id);
                return opt ? `${index + 1}、${opt.description}` : '';
            }).filter(d => d);
            textarea.value = descriptions.join('\n');
        }
    },

    validateForm() {
        if (!this.currentSchema) return {valid:false, errors:['No schema']};
        const errors = [], data = this.collectFormData();
        this.currentSchema.fields.forEach(f => {
            if (f.required) {
                // 图片字段的验证
                if (f.type === 'image') {
                    if (!data[f.key]) {
                        errors.push(f.label + ' 为必填项');
                    }
                } else if (f.type === 'image_list') {
                    if (!data[f.key] || !Array.isArray(data[f.key]) || data[f.key].length === 0) {
                        errors.push(f.label + ' 为必填项');
                    }
                } else if (f.type === 'array') {
                    // 数组字段验证
                    if (!data[f.key] || !Array.isArray(data[f.key]) || data[f.key].length === 0) {
                        errors.push(f.label + ' 为必填项，请至少添加一条记录');
                    }
                } else {
                    // 普通字段验证
                    if (!data[f.key] || !data[f.key].toString().trim()) {
                        errors.push(f.label + ' 为必填项');
                    }
                }
            }
            
            // widget 类型数据验证：schema-driven, check empty name fields
            if (f.type === 'widget' && f.validate && data[f.key] && Array.isArray(data[f.key]) && data[f.key].length > 0) {
                const vcfg = f.validate;
                const emptyItems = [];
                data[f.key].forEach((item, idx) => {
                    if (!item[vcfg.name_field] || !item[vcfg.name_field].trim()) {
                        emptyItems.push(idx + 1);
                    }
                });
                if (emptyItems.length > 0) {
                    errors.push(vcfg.item_label + ' ' + emptyItems.join('、') + ' 未填写名称，请填写或删除空条目');
                }
            }
        });
        return {valid: errors.length === 0, errors};
    },

    setFieldValue(key, value) {
        this.formData[key] = value;
        const el = document.getElementById(key);
        if (el) {
            // textarea 和 input 都使用 value 属性
            if (el.tagName === 'TEXTAREA' || el.tagName === 'INPUT') {
                el.value = value;
            } else if (el.tagName === 'SELECT') {
                el.value = value;
            }
            el.dispatchEvent(new Event('change', {bubbles:true}));
        }
    },

    // 应用预设值（用于风险等级联动填充）
    applyPresets(presets) {
        if (!presets || typeof presets !== 'object') return;
        for (const [key, value] of Object.entries(presets)) {
            // 数组类型（checkbox_group）
            if (Array.isArray(value)) {
                this.setCheckboxGroupValue(key, value);
                continue;
            }
            // 字符串类型，替换占位符
            let finalValue = value;
            if (typeof value === 'string' && value.includes('#')) {
                finalValue = value.replace(/#(\w+)#/g, (match, fieldKey) => {
                    return this.formData[fieldKey] || match;
                });
            }
            this.setFieldValue(key, finalValue);
        }
    },

    getFieldValue(key) {
        const el = document.getElementById(key);
        return el ? el.value : (this.formData[key] || '');
    },

    getTemplateId() { return this.currentTemplateId || ''; },

    async submitReport() {
        const v = this.validateForm();
        if (!v.valid) { if (window.AppUtils) AppUtils.showToast(v.errors.join('\n'), 'error'); return null; }
        
        const data = this.collectFormData(), tid = this.getTemplateId();
        if (!tid) { if (window.AppUtils) AppUtils.showToast('请先选择模板', 'error'); return null; }
        
        // 更新按钮状态
        const btn = document.getElementById('btn-dynamic-generate');
        const originalText = btn ? btn.innerText : '';
        if (btn) { btn.disabled = true; btn.innerText = '生成中...'; }
        
        const restoreUI = () => {
            if (btn) { btn.disabled = false; btn.innerText = originalText; }
        };
        
        try {
            const result = await AppAPI.Templates.generate(tid, data);
            restoreUI();
            
            if (result.success) {
                window.lastReportPath = result.report_path;
                if (window.AppUtils) AppUtils.showToast(`报告生成成功！\n路径：${result.report_path}`, 'success');

                // Prompt to save custom vulnerability name to the library (template-driven)
                const vulnSave = this.currentSchema?.vuln_save;
                if (vulnSave && vulnSave.name_field) {
                    const vulnName = data[vulnSave.name_field];
                    if (vulnName && String(vulnName).trim()) {
                        const trimmedName = String(vulnName).trim();
                        const vulnList = (window.AppVulnManager && window.AppVulnManager.VULN_LIST) || [];
                        const exists = vulnList.some(v => {
                            const n = window.AppVulnManager.getValue
                                ? window.AppVulnManager.getValue(v, ['Vuln_Name', 'name', 'vuln_name', '漏洞名称'])
                                : (v.Vuln_Name || v.name || v.vuln_name);
                            return n && String(n).trim() === trimmedName;
                        });

                        if (!exists) {
                            const confirmed = await AppUtils.safeConfirm(
                                `漏洞"${trimmedName}"不在漏洞库中，是否保存到漏洞库？`
                            );
                            if (confirmed) {
                                try {
                                    // Build vulnData from schema mapping
                                    const vulnData = {};
                                    if (vulnSave.mapping) {
                                        for (const [dbKey, formKey] of Object.entries(vulnSave.mapping)) {
                                            vulnData[dbKey] = data[formKey] || '';
                                        }
                                    }
                                    // Fallback: ensure required fields
                                    vulnData.name = vulnData.name || trimmedName;
                                    vulnData.level = vulnData.level || '中危';
                                    await AppAPI.saveVulnerability(vulnData);
                                    AppUtils.showToast('漏洞已保存到漏洞库', 'success');
                                    // Refresh the VULN_LIST in VulnManager
                                    if (window.AppVulnManager && window.AppVulnManager.loadVulnerabilities) {
                                        await window.AppVulnManager.loadVulnerabilities();
                                    }
                                } catch (e) {
                                    console.error('Save vulnerability failed:', e);
                                    AppUtils.showToast('保存漏洞失败: ' + e.message, 'error');
                                }
                            }
                        }
                    }
                }
            } else {
                const errMsg = result.message || result.detail || JSON.stringify(result);
                if (window.AppUtils) AppUtils.showToast('生成失败: ' + errMsg, 'error');
                console.error('Generate report failed:', result);
            }
            return result;
        } catch (e) { 
            restoreUI();
            console.error('Generate report failed:', e); 
            if (window.AppUtils) AppUtils.showToast('网络错误: ' + e.message, 'error');
            return null; 
        }
    },
    
    // 重置表单
    resetForm() {
        // 清空 formData
        this.formData = {};
        
        // 重置所有输入框
        if (this.currentSchema) {
            this.currentSchema.fields.forEach(f => {
                const el = document.getElementById(f.key);
                
                // 处理复杂字段类型
                if (f.type === 'array') {
                    // 清空数组数据
                    this.formData[f.key] = [];
                    // 清空表格内容（保留表头）
                    if (el) {
                        const tbody = el.querySelector('tbody');
                        if (tbody) tbody.innerHTML = '';
                    }
                } else if (f.type === 'widget') {
                    this.formData[f.key] = [];
                    if (el) {
                        el.innerHTML = '';
                        this.loadTemplateWidget(f).then(result => {
                            if (result) el.appendChild(result);
                        });
                    }
                } else if (f.type === 'checkbox_group') {
                    // 清空复选框组数据
                    this.formData[f.key] = [];
                    // 取消所有复选框选中状态
                    if (el) {
                        const checkboxes = el.querySelectorAll('input[type="checkbox"]');
                        checkboxes.forEach((checkbox) => {
                            checkbox.checked = false;
                        });
                        // 清空关联的文本框
                        const textarea = el.querySelector('textarea');
                        if (textarea) textarea.value = '';
                    }
                } else if (f.type === 'checkbox') {
                    // 单个复选框
                    if (el && el.type === 'checkbox') {
                        el.checked = false;
                    }
                } else if (f.type === 'searchable_select') {
                    // 可搜索下拉框：清空搜索框和重置下拉框
                    if (el) {
                        const searchInput = el.querySelector('input[type="text"]');
                        const select = el.querySelector('select');
                        if (searchInput) searchInput.value = '';
                        if (select) {
                            select.selectedIndex = 0;
                            // 触发 input 事件以恢复所有选项
                            if (searchInput) {
                                searchInput.dispatchEvent(new Event('input'));
                            }
                        }
                    }
                } else if (f.type === 'image' || f.type === 'image_list') {
                    // 清空图片预览
                    const preview = document.getElementById(`${f.key}-preview`);
                    if (preview) preview.innerHTML = '';
                    // 重置图片数据
                    if (f.type === 'image_list') {
                        this.formData[f.key] = [];
                    } else {
                        this.formData[f.key] = '';
                    }
                } else if (el) {
                    if (el.tagName === 'SELECT') {
                        el.selectedIndex = 0;
                    } else if (el.tagName === 'TEXTAREA' || el.tagName === 'INPUT') {
                        el.value = '';
                    }
                }
            });
        }
        
        // 重新设置默认值
        if (this.currentSchema) {
            this.setDefaultValues(this.currentSchema);
        }
        
        if (window.AppUtils) AppUtils.showToast('表单已重置', 'info');
    },
    
    // 打开报告目录
    openReportFolder() {
        if (window.lastReportPath) {
            AppAPI.openFolder(window.lastReportPath).catch(e => {
                console.error('Open folder failed:', e);
                if (window.AppUtils) AppUtils.showToast('打开目录失败', 'error');
            });
        } else {
            // 打开默认输出目录
            window.AppAPI.openFolder('output/report').catch(e => {
                console.error('Open folder failed:', e);
                if (window.AppUtils) AppUtils.showToast('打开目录失败', 'error');
            });
        }
    },
    
    // ========== 模板导入/导出功能 ==========
    
    async exportTemplate(templateId) {
        try {
            const tid = templateId || this.currentTemplateId;
            if (!tid) {
                this.showError('请先选择要导出的模板');
                return;
            }
            
            if (window.AppUtils) AppUtils.showToast('正在导出模板...', 'info');

            const download = await window.AppAPI.Templates.export(tid);
            const filename = download.filename || (tid + '_template.zip');
            const blob = download.blob;
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            window.URL.revokeObjectURL(url);
            
            if (window.AppUtils) AppUtils.showToast('模板导出成功!', 'success');
        } catch (e) {
            console.error('Export template failed:', e);
            this.showError('导出模板失败: ' + e.message);
        }
    },
    
    async importTemplate(file) {
        try {
            // 1. 基本验证
            if (!file) {
                this.showError('请选择要导入的模板文件');
                return;
            }
            
            // 2. 文件格式验证
            if (!file.name.endsWith('.zip')) {
                this.showError('只支持导入 .zip 格式的模板包');
                return;
            }
            
            // 3. 文件大小验证（限制50MB）
            const maxSize = (window.AppConfig && window.AppConfig.FILE && window.AppConfig.FILE.MAX_SIZE) || 50 * 1024 * 1024;
            const maxSizeMB = (window.AppConfig && window.AppConfig.FILE && window.AppConfig.FILE.MAX_SIZE_MB) || 50;
            if (file.size > maxSize) {
                this.showError(`模板文件过大，最大支持 ${maxSizeMB}MB`);
                return;
            }
            
            // 4. 显示导入进度提示
            if (window.AppUtils) AppUtils.showToast('正在验证模板文件...', 'info');
            
            const formData = new FormData();
            formData.append('file', file);
            
            // 5. 上传并导入
            if (window.AppUtils) AppUtils.showToast('正在上传模板...', 'info');
            
            const result = await AppAPI.Templates.import(file, false);
            
            if (result.success) {
                // 7. 导入成功
                if (window.AppUtils) {
                    AppUtils.showToast(`模板导入成功: ${result.template_id || ''}`, 'success');
                }
                
                // 8. 刷新模板列表
                await this.reloadTemplates();
                
                // 9. 刷新工具箱中的模板列表
                if (window.AppTemplateManager) {
                    await window.AppTemplateManager.loadTemplateListForManagement();
                }
            } else {
                throw new Error(result.message || '导入失败');
            }
        } catch (e) {
            console.error('Import template failed:', e);
            
            // 友好的错误提示
            let errorMsg = '导入模板失败';
            
            if (e.message.includes('Network')) {
                errorMsg = '网络错误，请检查连接后重试';
            } else if (e.message.includes('schema.yaml')) {
                errorMsg = '模板格式错误：缺少 schema.yaml 文件';
            } else if (e.message.includes('template.docx')) {
                errorMsg = '模板格式错误：缺少 template.docx 文件';
            } else if (e.message.includes('Invalid')) {
                errorMsg = '模板文件无效，请检查文件格式';
            } else if (e.message) {
                errorMsg = '导入失败: ' + e.message;
            }
            
            this.showError(errorMsg);
            if (window.AppUtils) AppUtils.showToast(errorMsg, 'error');
        }
    },
    
    // 打开导入对话框
    openImportDialog() {
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = '.zip';
        input.onchange = async (e) => {
            const file = e.target.files[0];
            if (file) await this.importTemplate(file);
        };
        input.click();
    },
    
    async deleteTemplate(templateId) {
        try {
            const tid = templateId || this.currentTemplateId;
            if (!tid) {
                this.showError('请先选择要删除的模板');
                return;
            }
            
            // 确认删除
            const confirmed = await AppUtils.safeConfirm(`确定要删除模板 "${tid}" 吗？此操作不可恢复！`);
            if (!confirmed) return;
            
            const result = await AppAPI.Templates.delete(tid);
            
            if (result.success) {
                if (window.AppUtils) AppUtils.showToast('模板已删除', 'success');
                // 刷新模板列表
                await this.reloadTemplates();
                // 刷新工具箱中的模板列表
                if (window.AppTemplateManager) {
                    await window.AppTemplateManager.loadTemplateListForManagement();
                }
            } else {
                throw new Error(result.message || '删除失败');
            }
        } catch (e) {
            console.error('Delete template failed:', e);
            this.showError('删除模板失败: ' + e.message);
        }
    },

};
