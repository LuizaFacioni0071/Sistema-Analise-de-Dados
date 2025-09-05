document.addEventListener('DOMContentLoaded', function() {
    // --- Elementos Globais e Funções Auxiliares ---
    const choiceContainer = document.getElementById('workflow-choice-container');
    const analysisWorkflow = document.getElementById('analysis-workflow-card');
    const updateWorkflow = document.getElementById('update-workflow-card');
    const backButtons = document.querySelectorAll('.btn-back');
    const loader = document.getElementById('loader');
    const loaderText = document.getElementById('loader-text');
    const toastContainer = document.getElementById('toast-container');
    const showLoader = (message) => { loaderText.textContent = message; loader.classList.remove('hidden'); };
    const hideLoader = () => loader.classList.add('hidden');
    const showToast = (message, type = 'info') => {
        const toast = document.createElement('div');
        toast.className = `toast toast-${type}`;
        toast.textContent = message;
        toastContainer.appendChild(toast);
        setTimeout(() => {
            toast.classList.add('show');
            setTimeout(() => { toast.classList.remove('show'); setTimeout(() => toast.remove(), 500); }, 5000);
        }, 100);
    };
    
    // --- Lógica de Navegação Principal ---
    document.querySelectorAll('.choice-card').forEach(card => {
        card.addEventListener('click', () => {
            const workflow = card.dataset.workflow;
            choiceContainer.classList.add('hidden');
            if (workflow === 'analysis') analysisWorkflow.classList.remove('hidden');
            else if (workflow === 'update') updateWorkflow.classList.remove('hidden');
        });
    });
    backButtons.forEach(btn => btn.addEventListener('click', () => window.location.reload()));

    // =======================================================
    // LÓGICA DO FLUXO 1: ANÁLISE DE INCONSISTÊNCIAS
    // =======================================================
    const analysisForm = document.getElementById('analysis-upload-form');
    const analysisFileInput = document.getElementById('analysis-file-input');
    const analysisFileLabel = document.getElementById('analysis-file-label');
    const analysisStep1 = document.getElementById('analysis-step1-card');
    const analysisStep2 = document.getElementById('analysis-step2-card');
    const analysisStep3 = document.getElementById('analysis-step3-card');
    const analysisSummaryCard = document.getElementById('analysis-summary-card');
    const analysisSheetSelect = document.getElementById('analysis-sheet-select');
    const analysisColumnsContainer = document.getElementById('analysis-columns-container');
    
    analysisFileInput.addEventListener('change', () => { analysisFileLabel.textContent = analysisFileInput.files[0]?.name || 'Clique para selecionar um ficheiro...'; });
    
    analysisForm.addEventListener('submit', async (e) => {
        e.preventDefault();
        showLoader('A processar ficheiro...');
        try {
            const response = await fetch('/api/upload_for_analysis', { method: 'POST', body: new FormData(analysisForm) });
            const data = await response.json();
            if (!response.ok) throw new Error(data.error);
            analysisStep1.classList.add('hidden');
            analysisStep2.classList.remove('hidden');
            document.getElementById('analysis-step2-title').textContent = `Passo 2: Analisando "${data.fileName}"`;
            analysisSheetSelect.innerHTML = data.sheets.map(s => `<option value="${s}">${s}</option>`).join('');
            populateAnalysisColumns(data.columns);
        } catch (err) { showToast(err.message, 'error'); } finally { hideLoader(); }
    });

    analysisSheetSelect.addEventListener('change', async (e) => {
        showLoader('A carregar colunas...');
        try {
            const response = await fetch('/api/get_columns_for_analysis', {
                method: 'POST', headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({ sheetName: e.target.value })
            });
            const data = await response.json();
            if (!response.ok) throw new Error(data.error);
            populateAnalysisColumns(data.columns);
            analysisStep3.classList.add('hidden');
        } catch (err) { showToast(err.message, 'error'); } finally { hideLoader(); }
    });

    function populateAnalysisColumns(columns) {
        analysisColumnsContainer.innerHTML = columns.map(col => `
            <div class="checkbox-wrapper">
                <input type="checkbox" id="col-${col.replace(/[^a-zA-Z0-9]/g, '-')}" value="${col}">
                <label for="col-${col.replace(/[^a-zA-Z0-9]/g, '-')}">${col}</label>
            </div>`).join('');
    }
    
    document.getElementById('analyze-btn').addEventListener('click', async () => {
        const selectedColumns = Array.from(analysisColumnsContainer.querySelectorAll('input:checked')).map(cb => cb.value);
        if (selectedColumns.length === 0) return showToast('Selecione pelo menos uma coluna.', 'info');
        showLoader('A analisar dados...');
        try {
            const response = await fetch('/api/analyze', {
                method: 'POST', headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({ sheetTitle: analysisSheetSelect.value, columns: selectedColumns })
            });
            const result = await response.json();
            if (!response.ok) throw new Error(result.error);
            document.getElementById('current-analysis-tab').textContent = analysisSheetSelect.value;
            displayInconsistentRows(result);
            analysisStep3.classList.remove('hidden');
        } catch (err) { showToast(err.message, 'error'); } finally { hideLoader(); }
    });

    function displayInconsistentRows(result) {
        const container = document.getElementById('inconsistencies-table-container');
        if (result.rows.length === 0) {
            container.innerHTML = '<p class="no-inconsistency">Nenhuma inconsistência encontrada para esta aba!</p>';
            document.getElementById('confirm-decisions-btn').disabled = false;
            return;
        }
        document.getElementById('confirm-decisions-btn').disabled = false;
        const headers = `<th>Remover?</th>${result.headers.map(h => `<th>${h}</th>`).join('')}<th>Problema</th>`;
        
        const body = result.rows.map(row => `
            <tr>
                <td><input type="checkbox" class="remove-row-checkbox" value="${row.rowIndex}"></td>
                ${result.headers.map(h => `<td>${row[h] || ''}</td>`).join('')}
                <td class="issue-cell">${row.issue || ''}</td>
            </tr>`).join('');
        container.innerHTML = `<table><thead><tr>${headers}</tr></thead><tbody>${body}</tbody></table>`;
    }

    document.getElementById('confirm-decisions-btn').addEventListener('click', async () => {
        const sheetTitle = analysisSheetSelect.value;
        const checkboxes = document.querySelectorAll('#inconsistencies-table-container .remove-row-checkbox:checked');
        const rowsToRemove = Array.from(checkboxes).map(cb => parseInt(cb.value));

        showLoader(`A guardar decisões para a aba "${sheetTitle}"...`);
        try {
            const response = await fetch('/api/stage_tab_changes', {
                method: 'POST', headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({ sheetTitle, rowsToRemove })
            });
            const data = await response.json();
            if (!response.ok) throw new Error(data.error || 'Erro ao guardar');
            
            showToast(data.message, 'success');
            analysisStep3.classList.add('hidden');
            updateSummary(sheetTitle, rowsToRemove.length);
        } catch(err) {
            showToast(err.message, 'error');
        } finally {
            hideLoader();
        }
    });

    function updateSummary(sheetTitle, removalCount) {
        analysisSummaryCard.classList.remove('hidden');
        const summaryContainer = document.getElementById('analysis-summary-container');
        
        let existingEntry = summaryContainer.querySelector(`[data-sheet="${sheetTitle}"]`);
        if (existingEntry) {
            existingEntry.querySelector('span').textContent = `${removalCount} linha(s) para remover`;
        } else {
            summaryContainer.innerHTML += `<div class="summary-item" data-sheet="${sheetTitle}"><strong>${sheetTitle}:</strong> <span>${removalCount} linha(s) para remover</span></div>`;
        }
    }

    document.getElementById('process-all-btn').addEventListener('click', async () => {
        showLoader('A processar todas as abas e a gerar ficheiro...');
        try {
            const response = await fetch('/api/process_all_staged_changes', { method: 'POST' });
            if (!response.ok) { const err = await response.json(); throw new Error(err.error); }
            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            const disposition = response.headers.get('Content-Disposition');
            a.download = disposition ? /filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/.exec(disposition)[1].replace(/['"]/g, '') : 'processado_final.xlsx';
            a.click();
            a.remove();
            window.URL.revokeObjectURL(url);
            showToast('Ficheiro final gerado com sucesso!', 'success');
        } catch (err) { showToast(err.message, 'error'); } finally { hideLoader(); }
    });


    // =======================================================
    // LÓGICA DO FLUXO 2: ATUALIZAR PLANILHA
    // =======================================================
    let baseFileSheets = [], updateFileSheets = [], mergeQueue = [];
    const baseUploadForm = document.getElementById('base-upload-form');
    const updateUploadForm = document.getElementById('update-upload-form');
    const baseSheetSelect = document.getElementById('base-sheet-select');
    const updateSheetSelect = document.getElementById('update-sheet-select');
    const keyColumnSelect = document.getElementById('key-column-select');
    const addToQueueBtn = document.getElementById('add-to-queue-btn');
    const processAllUpdateBtn = document.getElementById('process-all-update-btn');
    const mergeQueueList = document.getElementById('merge-queue-list');
    
    document.getElementById('base-file-input').addEventListener('change', (e) => { document.getElementById('base-file-label').textContent = e.target.files[0]?.name || '...'; });
    document.getElementById('update-file-input').addEventListener('change', (e) => { document.getElementById('update-file-label').textContent = e.target.files[0]?.name || '...'; });
    
    baseUploadForm.addEventListener('submit', async (e) => {
        e.preventDefault();
        showLoader('A carregar ficheiro base...');
        try {
            const response = await fetch('/api/upload_base', { method: 'POST', body: new FormData(baseUploadForm) });
            const data = await response.json();
            if (!response.ok) throw new Error(data.error);
            baseFileSheets = data.sheets;
            baseSheetSelect.innerHTML = baseFileSheets.map(s => `<option value="${s}">${s}</option>`).join('');
            document.getElementById('update-stepA-card').classList.add('hidden');
            document.getElementById('update-stepB-card').classList.remove('hidden');
        } catch (err) { showToast(err.message, 'error'); } finally { hideLoader(); }
    });

    updateUploadForm.addEventListener('submit', async (e) => {
        e.preventDefault();
        showLoader('A carregar ficheiro de atualização...');
        try {
            const response = await fetch('/api/upload_update', { method: 'POST', body: new FormData(updateUploadForm) });
            const data = await response.json();
            if (!response.ok) throw new Error(data.error);
            updateFileSheets = data.sheets;
            updateSheetSelect.innerHTML = updateFileSheets.map(s => `<option value="${s}">${s}</option>`).join('');
            document.getElementById('update-stepB-card').classList.add('hidden');
            document.getElementById('update-stepC-card').classList.remove('hidden');
            await updateKeyColumnOptions();
        } catch (err) { showToast(err.message, 'error'); } finally { hideLoader(); }
    });

    async function updateKeyColumnOptions() {
        showLoader('A procurar colunas em comum...');
        keyColumnSelect.disabled = true;
        try {
            const response = await fetch('/api/get_common_columns', {
                method: 'POST', headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ baseSheet: baseSheetSelect.value, updateSheet: updateSheetSelect.value })
            });
            const data = await response.json();
            if (!response.ok) throw new Error(data.error);
            if (data.commonColumns.length === 0) { keyColumnSelect.innerHTML = '<option>Nenhuma coluna em comum</option>'; }
            else { keyColumnSelect.innerHTML = data.commonColumns.map(c => `<option value="${c}">${c}</option>`).join(''); keyColumnSelect.disabled = false; }
        } catch (err) { showToast(err.message, 'error'); } finally { hideLoader(); }
    }

    baseSheetSelect.addEventListener('change', updateKeyColumnOptions);
    updateSheetSelect.addEventListener('change', updateKeyColumnOptions);

    addToQueueBtn.addEventListener('click', () => {
        const instruction = { baseTab: baseSheetSelect.value, updateTab: updateSheetSelect.value, keyColumn: keyColumnSelect.value };
        if (!instruction.keyColumn || keyColumnSelect.disabled) return showToast('Selecione uma coluna chave válida.', 'error');
        mergeQueue.push(instruction);
        renderMergeQueue();
    });

    function renderMergeQueue() {
        mergeQueueList.innerHTML = mergeQueue.map((item, index) => `<li><span class="queue-item-text">Atualizar Aba <span>${item.baseTab}</span> com <span>${item.updateTab}</span> usando a chave <span>${item.keyColumn}</span></span><button class="queue-item-remove" data-index="${index}">&times;</button></li>`).join('');
        processAllUpdateBtn.disabled = mergeQueue.length === 0;
    }
    
    mergeQueueList.addEventListener('click', (e) => {
        if (e.target.classList.contains('queue-item-remove')) {
            mergeQueue.splice(e.target.dataset.index, 1);
            renderMergeQueue();
        }
    });

    processAllUpdateBtn.addEventListener('click', async () => {
        showLoader('A processar todas as fusões...');
        try {
            const response = await fetch('/api/process_multi_tab_update', {
                method: 'POST', headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ instructions: mergeQueue })
            });
            if (!response.ok) { const err = await response.json(); throw new Error(err.error); }
            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            const disposition = response.headers.get('Content-Disposition');
            a.download = disposition ? /filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/.exec(disposition)[1].replace(/['"]/g, '') : 'atualizado.xlsx';
            a.click(); a.remove(); window.URL.revokeObjectURL(url);
            showToast('Ficheiro final gerado com sucesso!', 'success');
        } catch (err) { showToast(err.message, 'error'); } finally { hideLoader(); }
    });
});