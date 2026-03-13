// 全局状态
var uploadedImages = []; // {image_id, filename, preview, questionRange}
var recognizedQuestions = [];
var classifications = []; // 分类列表

// 页面加载时初始化
window.onload = function() {
    checkAIConfig();
    loadClassifications();
    setupBeforeUnloadWarning();
};

// 设置页面刷新/关闭前的警告
function setupBeforeUnloadWarning() {
    window.addEventListener('beforeunload', function(e) {
        // 如果有上传的图片或识别结果，提示用户
        if (uploadedImages.length > 0 || recognizedQuestions.length > 0) {
            var message = '您有未保存的数据，确定要离开此页面吗？';
            e.returnValue = message;
            return message;
        }
    });
}

// 加载分类列表
async function loadClassifications() {
    try {
        var response = await fetch('/api/classifications');
        var data = await response.json();
        if (data.status === 'ok') {
            classifications = data.classifications;
        }
    } catch (e) {
        console.error('加载分类失败:', e);
        classifications = ['交际用语', '词义辨析', '时态', '非谓语动词', '定语从句', '状语从句', '情态动词', '名词性从句', '代词'];
    }
}

// 检查AI配置
async function checkAIConfig() {
    var aiInfo = document.getElementById('ai-info');
    try {
        var response = await fetch('/api/ai-info');
        var data = await response.json();

        if (data.status === 'ok') {
            aiInfo.className = 'card alert alert-success';
            aiInfo.innerHTML = '✅ 当前使用AI: <strong>' + data.ai_name + '</strong>';
        } else {
            aiInfo.className = 'card alert alert-error';
            aiInfo.innerHTML = '❌ ' + (data.message || '未配置OCR AI') + '，请在软件设置中配置OCR AI';
        }
    } catch (e) {
        aiInfo.className = 'card alert alert-error';
        aiInfo.innerHTML = '❌ 无法连接到服务器，请检查网络';
        console.error('AI配置检查失败:', e);
    }
}

// 文件上传处理 - 上传到服务器
async function handleFileSelect(e) {
    var files = Array.from(e.target.files);
    if (files.length === 0) return;

    var uploadPromises = files.map(function(file) {
        if (!file.type.startsWith('image/')) {
            console.warn('跳过非图片文件:', file.name);
            return Promise.resolve(null);
        }

        return new Promise(function(resolve) {
            var reader = new FileReader();
            reader.onload = function(e) {
                var preview = e.target.result;

                // 上传到服务器
                var formData = new FormData();
                formData.append('image', file);

                fetch('/api/upload-image', {
                    method: 'POST',
                    body: formData
                })
                .then(function(response) { return response.json(); })
                .then(function(data) {
                    if (data.status === 'ok') {
                        uploadedImages.push({
                            image_id: data.image_id,
                            filename: data.filename,
                            preview: preview,
                            questionRange: ''
                        });
                        console.log('图片上传成功:', data.filename, 'ID:', data.image_id);
                    } else {
                        console.error('图片上传失败:', data.error);
                        alert('图片 ' + file.name + ' 上传失败: ' + data.error);
                    }
                    resolve();
                })
                .catch(function(err) {
                    console.error('图片上传请求失败:', err);
                    alert('图片 ' + file.name + ' 上传失败');
                    resolve();
                });
            };
            reader.readAsDataURL(file);
        });
    });

    await Promise.all(uploadPromises);
    updateImagePreview();

    // 重置文件输入框
    e.target.value = '';
}

document.getElementById('imageInput').addEventListener('change', handleFileSelect);

// 更新图片预览
function updateImagePreview() {
    var container = document.getElementById('imagePreviewContainer');
    var placeholder = document.getElementById('uploadPlaceholder');
    var uploadArea = document.getElementById('uploadArea');
    var imageCount = document.getElementById('imageCount');
    var recognizeBtn = document.getElementById('recognizeBtn');

    imageCount.textContent = uploadedImages.length;
    recognizeBtn.disabled = uploadedImages.length === 0;

    if (uploadedImages.length > 0) {
        placeholder.classList.add('hidden');
        container.classList.remove('hidden');
        uploadArea.classList.add('has-images');

        var html = '';
        uploadedImages.forEach(function(img, index) {
            html += '<div class="image-preview-item" style="position: relative; border: 1px solid #ddd; border-radius: 8px; padding: 10px; background: #fff;">';
            html += '<img src="' + img.preview + '" alt="图片' + (index + 1) + '" style="width: 100%; height: 120px; object-fit: cover; border-radius: 4px; cursor: pointer;" onclick="event.stopPropagation(); previewImage(' + index + ');">';
            html += '<button class="remove-btn" onclick="event.stopPropagation(); removeImage(' + index + ');" title="删除" style="position: absolute; top: 5px; right: 5px; background: #ff4444; color: white; border: none; border-radius: 50%; width: 24px; height: 24px; cursor: pointer; font-size: 16px; line-height: 1;">×</button>';
            html += '<div style="margin-top: 8px; padding: 8px; background: #f0f4ff; border-radius: 4px; font-size: 13px; color: #1565c0; word-break: break-all; font-weight: 500; line-height: 1.4;">📄 ' + img.filename + '</div>';
            html += '<div style="margin-top: 8px;">';
            html += '<label style="font-size: 12px; color: #666; display: block; margin-bottom: 4px;">题号范围:</label>';
            html += '<input type="text" class="img-question-range" data-idx="' + index + '" value="' + (img.questionRange || '') + '" placeholder="如: 1-5" style="width: 100%; padding: 6px; border: 1px solid #ddd; border-radius: 4px; font-size: 13px;" onchange="updateImageQuestionRange(' + index + ', this.value)" onclick="event.stopPropagation();">';
            html += '</div>';
            html += '</div>';
        });
        container.innerHTML = html;
    } else {
        placeholder.classList.remove('hidden');
        container.classList.add('hidden');
        uploadArea.classList.remove('has-images');
        container.innerHTML = '';
    }
}

// 更新图片的题号范围
async function updateImageQuestionRange(index, value) {
    if (uploadedImages[index]) {
        uploadedImages[index].questionRange = value;

        // 同步到服务器
        try {
            await fetch('/api/update-image-range', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    image_id: uploadedImages[index].image_id,
                    question_range: value
                })
            });
        } catch (e) {
            console.error('更新题号范围失败:', e);
        }
    }
}

// 删除图片
async function removeImage(index) {
    if (!uploadedImages[index]) return;

    var image_id = uploadedImages[index].image_id;
    var filename = uploadedImages[index].filename;

    // 先通知服务器删除
    try {
        var response = await fetch('/api/remove-image', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ image_id: image_id })
        });
        var data = await response.json();
        if (data.status !== 'ok') {
            console.error('服务器删除图片失败:', data.error);
        }
    } catch (e) {
        console.error('删除图片请求失败:', e);
    }

    // 从本地数组移除
    uploadedImages.splice(index, 1);
    updateImagePreview();
}

// 预览单张图片
function previewImage(index) {
    var img = uploadedImages[index];
    if (!img) return;

    var modal = document.createElement('div');
    modal.style.cssText = 'position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.9); z-index: 10000; display: flex; flex-direction: column; align-items: center; justify-content: center; padding: 20px;';

    var closeBtn = document.createElement('button');
    closeBtn.textContent = '✕ 关闭';
    closeBtn.style.cssText = 'position: absolute; top: 20px; right: 20px; background: #fff; border: none; padding: 10px 20px; border-radius: 4px; cursor: pointer; font-size: 14px;';
    closeBtn.onclick = function() { document.body.removeChild(modal); };

    var imgEl = document.createElement('img');
    imgEl.src = img.preview;
    imgEl.style.cssText = 'max-width: 90%; max-height: 80%; object-fit: contain; border-radius: 8px;';

    var filenameEl = document.createElement('div');
    filenameEl.textContent = img.filename;
    filenameEl.style.cssText = 'color: #fff; margin-top: 20px; font-size: 16px;';

    modal.appendChild(closeBtn);
    modal.appendChild(imgEl);
    modal.appendChild(filenameEl);
    document.body.appendChild(modal);
}

// 拖拽上传
var uploadArea = document.getElementById('uploadArea');
uploadArea.addEventListener('dragover', function(e) {
    e.preventDefault();
    uploadArea.style.borderColor = '#667eea';
});
uploadArea.addEventListener('dragleave', function() {
    uploadArea.style.borderColor = '#ccc';
});
uploadArea.addEventListener('drop', async function(e) {
    e.preventDefault();
    uploadArea.style.borderColor = '#ccc';
    var files = Array.from(e.dataTransfer.files);

    var uploadPromises = files.map(function(file) {
        if (!file.type.startsWith('image/')) return Promise.resolve(null);

        return new Promise(function(resolve) {
            var reader = new FileReader();
            reader.onload = function(e) {
                var preview = e.target.result;
                var formData = new FormData();
                formData.append('image', file);

                fetch('/api/upload-image', {
                    method: 'POST',
                    body: formData
                })
                .then(function(response) { return response.json(); })
                .then(function(data) {
                    if (data.status === 'ok') {
                        uploadedImages.push({
                            image_id: data.image_id,
                            filename: data.filename,
                            preview: preview,
                            questionRange: ''
                        });
                    }
                    resolve();
                })
                .catch(function() { resolve(); });
            };
            reader.readAsDataURL(file);
        });
    });

    await Promise.all(uploadPromises);
    updateImagePreview();
});

// OCR识别 - 逐张识别显示进度
async function recognize() {
    if (uploadedImages.length === 0) {
        alert('请先上传图片');
        return;
    }

    var loading = document.getElementById('loading');
    var recognizeBtn = document.getElementById('recognizeBtn');
    var loadingDetail = document.getElementById('loadingDetail');

    loading.classList.add('active');
    recognizeBtn.disabled = true;

    var generateAnalysis = document.getElementById('generateAnalysis').checked;
    var totalImages = uploadedImages.length;
    var allQuestions = [];
    var failedImages = [];

    // 逐张识别
    for (var i = 0; i < totalImages; i++) {
        var img = uploadedImages[i];
        loadingDetail.textContent = '正在识别第 ' + (i + 1) + '/' + totalImages + ' 张图片: ' + img.filename + '...';

        try {
            var response = await fetch('/api/recognize', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    image_ids: [img.image_id],
                    generate_analysis: generateAnalysis
                })
            });

            var data = await response.json();

            if (data.status === 'ok') {
                allQuestions = allQuestions.concat(data.questions);
                console.log('第 ' + (i + 1) + '/' + totalImages + ' 张图片识别完成: ' + img.filename + ', 识别出 ' + data.questions.length + ' 道题目');
            } else {
                console.error('第 ' + (i + 1) + ' 张图片识别失败: ' + img.filename + ', 错误: ' + data.error);
                failedImages.push(img.filename);
            }
        } catch (e) {
            console.error('第 ' + (i + 1) + ' 张图片识别失败: ' + img.filename, e);
            failedImages.push(img.filename);
        }
    }

    // 将所有识别的题目添加到列表
    if (allQuestions.length > 0) {
        var startId = recognizedQuestions.length;
        allQuestions.forEach(function(q) {
            recognizedQuestions.push({
                question: q.question,
                A: q.A,
                B: q.B,
                C: q.C,
                D: q.D,
                answer: q.answer,
                classification: q.classification,
                analysis: q.analysis,
                id: startId++,
                selected: true,
                source: 'Web OCR'
            });
        });

        displayQuestions();
    }

    // 显示结果
    var msg = '识别完成！共识别 ' + totalImages + ' 张图片，识别出 ' + allQuestions.length + ' 道题目，共计 ' + recognizedQuestions.length + ' 道题目';
    if (failedImages.length > 0) {
        msg += '\n\n以下 ' + failedImages.length + ' 张图片识别失败:\n' + failedImages.join('\n');
    }
    alert(msg);

    loading.classList.remove('active');
    recognizeBtn.disabled = false;
}

// 显示识别结果
function displayQuestions() {
    var container = document.getElementById('questions-list');
    var resultsSection = document.getElementById('results-section');
    var selectedCount = document.getElementById('selectedCount');
    var selectAll = document.getElementById('selectAll');

    if (recognizedQuestions.length === 0) {
        container.innerHTML = '<div class="empty-state">暂无识别结果</div>';
        selectedCount.textContent = '已选择 0/0 题';
        selectAll.checked = false;
    } else {
        var selectedNum = recognizedQuestions.filter(function(q) { return q.selected; }).length;
        selectedCount.textContent = '已选择 ' + selectedNum + '/' + recognizedQuestions.length + ' 题';
        selectAll.checked = selectedNum === recognizedQuestions.length;

        var html = '';
        recognizedQuestions.forEach(function(q, index) {
            html += renderQuestionItem(q, index);
        });
        container.innerHTML = html;

        // 绑定事件
        recognizedQuestions.forEach(function(q, index) {
            bindQuestionEvents(index);
        });
    }

    resultsSection.classList.remove('hidden');
}

// 渲染单个题目
function renderQuestionItem(q, index) {
    var html = '<div class="question-item" data-index="' + index + '" style="border: 1px solid #e0e0e0; border-radius: 8px; margin-bottom: 10px; overflow: hidden;">';

    // 折叠头部
    html += '<div class="question-header" onclick="toggleQuestionCollapse(' + index + ')" style="display: flex; justify-content: space-between; align-items: center; padding: 12px 15px; background: #f5f5f5; cursor: pointer; user-select: none;">';
    html += '<div style="display: flex; align-items: center; gap: 10px; flex: 1; min-width: 0;">';
    html += '<input type="checkbox" class="q-select" ' + (q.selected ? 'checked' : '') + ' onclick="event.stopPropagation();" style="cursor: pointer;">';
    html += '<span class="question-num" style="font-weight: 500; color: #333;">题目 ' + (index + 1) + '</span>';
    html += '<span class="question-preview" style="color: #666; font-size: 14px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; flex: 1;">' + (q.question || '').substring(0, 30) + '</span>';
    html += '</div>';
    html += '<div style="display: flex; align-items: center; gap: 10px;">';
    html += '<button onclick="deleteQuestion(' + index + ', event)" style="padding: 4px 12px; background: #dc3545; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 12px;" title="删除这道题">删除</button>';
    html += '<span id="collapse-icon-' + index + '" style="transition: transform 0.2s; font-size: 12px;">▼</span>';
    html += '</div>';
    html += '</div>';

    // 可折叠内容
    html += '<div id="question-content-' + index + '" class="question-content" style="padding: 15px; background: #fff;">';

    // 题目
    html += '<div class="form-group"><label>题目</label>';
    html += '<textarea class="q-question" data-idx="' + index + '" style="width: 100%; min-height: 60px; padding: 8px; border: 1px solid #ddd; border-radius: 4px; resize: vertical;">' + (q.question || '') + '</textarea></div>';

    // 选项
    html += '<div class="form-row" style="display: grid; grid-template-columns: 1fr 1fr; gap: 10px;">';
    html += '<div class="form-group"><label>选项 A</label><input type="text" class="q-A" data-idx="' + index + '" value="' + (q.A || '') + '" style="width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px;"></div>';
    html += '<div class="form-group"><label>选项 B</label><input type="text" class="q-B" data-idx="' + index + '" value="' + (q.B || '') + '" style="width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px;"></div>';
    html += '<div class="form-group"><label>选项 C</label><input type="text" class="q-C" data-idx="' + index + '" value="' + (q.C || '') + '" style="width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px;"></div>';
    html += '<div class="form-group"><label>选项 D</label><input type="text" class="q-D" data-idx="' + index + '" value="' + (q.D || '') + '" style="width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px;"></div>';
    html += '</div>';

    // 答案和分类
    html += '<div class="form-row" style="display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-top: 10px;">';
    html += '<div class="form-group"><label>正确答案</label>';
    html += '<select class="q-answer" data-idx="' + index + '" style="width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px;">';
    html += '<option value="">请选择</option>';
    html += '<option value="A" ' + (q.answer === 'A' ? 'selected' : '') + '>A</option>';
    html += '<option value="B" ' + (q.answer === 'B' ? 'selected' : '') + '>B</option>';
    html += '<option value="C" ' + (q.answer === 'C' ? 'selected' : '') + '>C</option>';
    html += '<option value="D" ' + (q.answer === 'D' ? 'selected' : '') + '>D</option>';
    html += '</select></div>';

    html += '<div class="form-group"><label>分类</label>';
    html += '<select class="q-classification" data-idx="' + index + '" style="width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px;">';
    html += getClassificationOptions(q.classification);
    html += '</select></div>';
    html += '</div>';

    // 来源
    html += '<div class="form-group" style="margin-top: 10px;"><label>来源</label>';
    html += '<input type="text" class="q-source" data-idx="' + index + '" value="' + (q.source || '') + '" style="width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px;"></div>';

    // 解析
    var analysisText = (q.analysis !== undefined && q.analysis !== null) ? q.analysis : '';
    html += '<div class="form-group" style="margin-top: 10px;"><label>解析</label>';
    html += '<textarea class="q-analysis" data-idx="' + index + '" placeholder="暂无解析" style="width: 100%; min-height: 60px; padding: 8px; border: 1px solid #ddd; border-radius: 4px; resize: vertical;">' + analysisText + '</textarea></div>';

    html += '</div>';
    html += '</div>';

    return html;
}

// 获取分类选项
function getClassificationOptions(selected) {
    var classList = classifications.length > 0 ? classifications : ['交际用语', '词义辨析', '时态', '非谓语动词', '定语从句', '状语从句', '情态动词', '名词性从句', '代词'];
    var isValidSelection = selected && classList.indexOf(selected) !== -1;
    var html = '<option value="" ' + (!isValidSelection ? 'selected' : '') + '>请选择分类</option>';
    classList.forEach(function(c) {
        html += '<option value="' + c + '" ' + (c === selected ? 'selected' : '') + '>' + c + '</option>';
    });
    return html;
}

// 绑定题目事件
function bindQuestionEvents(index) {
    var item = document.querySelector('.question-item[data-index="' + index + '"]');
    if (!item) return;

    item.querySelector('.q-select').onchange = function() {
        recognizedQuestions[index].selected = this.checked;
        updateSelectedCount();
    };

    item.querySelector('.q-question').onchange = function() {
        recognizedQuestions[index].question = this.value;
        updateQuestionPreview(index);
    };

    item.querySelector('.q-A').onchange = function() {
        recognizedQuestions[index].A = this.value;
    };

    item.querySelector('.q-B').onchange = function() {
        recognizedQuestions[index].B = this.value;
    };

    item.querySelector('.q-C').onchange = function() {
        recognizedQuestions[index].C = this.value;
    };

    item.querySelector('.q-D').onchange = function() {
        recognizedQuestions[index].D = this.value;
    };

    item.querySelector('.q-answer').onchange = function() {
        recognizedQuestions[index].answer = this.value;
    };

    item.querySelector('.q-classification').onchange = function() {
        recognizedQuestions[index].classification = this.value;
    };

    item.querySelector('.q-source').onchange = function() {
        recognizedQuestions[index].source = this.value;
    };

    var analysisField = item.querySelector('.q-analysis');
    if (analysisField) {
        analysisField.onchange = function() {
            recognizedQuestions[index].analysis = this.value;
        };
    }
}

// 切换题目折叠
function toggleQuestionCollapse(index) {
    var content = document.getElementById('question-content-' + index);
    var icon = document.getElementById('collapse-icon-' + index);
    if (!content || !icon) return;

    if (content.style.display === 'none') {
        content.style.display = 'block';
        icon.style.transform = 'rotate(0deg)';
    } else {
        content.style.display = 'none';
        icon.style.transform = 'rotate(-90deg)';
    }
}

// 删除单道题目
function deleteQuestion(index, event) {
    if (event) {
        event.stopPropagation();
    }

    if (!confirm('确定要删除题目 ' + (index + 1) + ' 吗？')) {
        return;
    }

    // 从数组中删除该题目
    recognizedQuestions.splice(index, 1);

    // 重新显示所有题目（序号会自动更新）
    displayQuestions();

    console.log('已删除题目，剩余 ' + recognizedQuestions.length + ' 道题');
}

// 更新题目预览
function updateQuestionPreview(index) {
    var previewEl = document.querySelector('.question-item[data-index="' + index + '"] .question-preview');
    if (previewEl && recognizedQuestions[index]) {
        var question = recognizedQuestions[index].question || '';
        previewEl.textContent = question.substring(0, 30) + (question.length > 30 ? '...' : '');
    }
}

// 更新选中数量
function updateSelectedCount() {
    var selectedNum = recognizedQuestions.filter(function(q) { return q.selected; }).length;
    document.getElementById('selectedCount').textContent = '已选择 ' + selectedNum + '/' + recognizedQuestions.length + ' 题';
    document.getElementById('selectAll').checked = selectedNum === recognizedQuestions.length && recognizedQuestions.length > 0;
}

// 全选/取消全选
function toggleSelectAll() {
    var selectAll = document.getElementById('selectAll');
    recognizedQuestions.forEach(function(q) { q.selected = selectAll.checked; });
    displayQuestions();
}

// 折叠/展开所有题目
function toggleAllQuestions() {
    if (recognizedQuestions.length === 0) return;

    var allCollapsed = true;
    for (var i = 0; i < recognizedQuestions.length; i++) {
        var content = document.getElementById('question-content-' + i);
        if (content && content.style.display !== 'none') {
            allCollapsed = false;
            break;
        }
    }

    for (var i = 0; i < recognizedQuestions.length; i++) {
        var content = document.getElementById('question-content-' + i);
        var icon = document.getElementById('collapse-icon-' + i);
        if (!content || !icon) continue;

        if (allCollapsed) {
            content.style.display = 'block';
            icon.style.transform = 'rotate(0deg)';
        } else {
            content.style.display = 'none';
            icon.style.transform = 'rotate(-90deg)';
        }
    }
}

// 清空结果
async function clearResults() {
    recognizedQuestions = [];
    document.getElementById('questions-list').innerHTML = '';
    document.getElementById('results-section').classList.add('hidden');
    document.getElementById('selectedCount').textContent = '已选择 0/0 题';
    document.getElementById('selectAll').checked = false;
}

// 清空所有
async function clearAll() {
    // 清空识别结果
    await clearResults();

    // 通知服务器清空图片
    try {
        await fetch('/api/clear-images', { method: 'POST' });
    } catch (e) {
        console.error('清空服务器图片失败:', e);
    }

    // 清空本地图片列表
    uploadedImages = [];
    updateImagePreview();

    document.getElementById('questionRange').value = '';
}

// 添加选中的题目到题库
async function addSelectedToDatabase() {
    var selectedQuestions = recognizedQuestions.filter(function(q) { return q.selected; });

    if (selectedQuestions.length === 0) {
        alert('请先选择要添加的题目');
        return;
    }

    // 验证数据完整性
    var errors = [];
    for (var i = 0; i < selectedQuestions.length; i++) {
        var q = selectedQuestions[i];
        var questionNum = i + 1;

        if (!q.question || !q.question.trim()) {
            errors.push('题目 ' + questionNum + '：题目内容不能为空');
        }
        if (!q.A || !q.A.trim()) {
            errors.push('题目 ' + questionNum + '：选项A不能为空');
        }
        if (!q.B || !q.B.trim()) {
            errors.push('题目 ' + questionNum + '：选项B不能为空');
        }
        if (!q.C || !q.C.trim()) {
            errors.push('题目 ' + questionNum + '：选项C不能为空');
        }
        if (!q.D || !q.D.trim()) {
            errors.push('题目 ' + questionNum + '：选项D不能为空');
        }
        if (!q.answer) {
            errors.push('题目 ' + questionNum + '：必须选择正确答案');
        }
        if (!q.classification) {
            errors.push('题目 ' + questionNum + '：必须选择分类');
        }
        if (!q.source || !q.source.trim()) {
            errors.push('题目 ' + questionNum + '：来源不能为空');
        }
    }

    if (errors.length > 0) {
        alert('数据检查失败，请完善以下信息：\n\n' + errors.join('\n'));
        return;
    }

    // 发送到服务器
    try {
        var response = await fetch('/api/import', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ questions: selectedQuestions })
        });

        var data = await response.json();

        if (data.status === 'ok') {
            alert('成功导入 ' + data.imported + ' 道题目到题库！');
            // 移除已导入的题目
            recognizedQuestions = recognizedQuestions.filter(function(q) { return !q.selected; });
            displayQuestions();
        } else {
            alert('导入失败: ' + data.error);
        }
    } catch (e) {
        console.error('导入请求失败:', e);
        alert('导入请求失败，请检查网络连接');
    }
}
