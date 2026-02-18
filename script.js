class DataViewerApp {
  constructor() {
    this.rawData = [];
    this.visibleData = [];
    this.columns = [];
    this.colSettings = {};
    this.filterSummary = document.getElementById("filterSummaryBar");

    // Pagination params
    this.pageSize = 100;
    this.currentPage = 1;

    // Sort/Filter state
    this.sortCol = null;
    this.sortAsc = true;
    this.searchQuery = "";
    this.activeMenuCol = null;

    // Excel Workbook State
    this.currentWorkbook = null;
    this.tempRawMatrix = [];
    this.tempHeaderIdx = 0;

    this.initElements();
    this.initEvents();
  }

  initElements() {
    this.els = {
      fileInput: document.getElementById("fileInput"),
      emptyState: document.getElementById("emptyState"),
      loadingState: document.getElementById("loadingState"),
      tableWrapper: document.getElementById("tableWrapper"),
      footer: document.getElementById("appFooter"),
      tbody: document.getElementById("tBody"),
      thead: document.getElementById("tHead"),
      tfoot: document.getElementById("tFoot"),
      colMenu: document.getElementById("colMenu"),
      exportMenu: document.getElementById("exportMenu"),
      globalSearch: document.getElementById("globalSearch"),
      dragOverlay: document.getElementById("dragOverlay"),
      ctxMenu: document.getElementById("columnContextMenu"),
      colListContainer: document.getElementById("colListContainer"),
      reportTitle: document.getElementById("reportTitle"),
      reportAuthor: document.getElementById("reportAuthor"),
      sheetModal: document.getElementById("sheetModal"),
      sheetList: document.getElementById("sheetList"),
      exportModal: document.getElementById("exportModal"),
      confirmTitle: document.getElementById("confirmTitle"),
      confirmAuthor: document.getElementById("confirmAuthor"),
      btnConfirmExport: document.getElementById("btnConfirmExport"),
      btnCancelExport: document.getElementById("btnCancelExport"),

      // ELEMENTOS NUEVOS PARA ESTRUCTURA
      structureModal: document.getElementById("structureModal"),
      previewTable: document.getElementById("previewTable"),
      footerSkipCount: document.getElementById("footerSkipCount"),
      selectedHeaderDisplay: document.getElementById(
        "selectedHeaderIndexDisplay"
      ),
      btnConfirmStructure: document.getElementById("btnConfirmStructure"),
      btnCancelStructure: document.getElementById("btnCancelStructure"),
      csvMapModal: document.getElementById("csvMapModal"),
      mapLocalidad: document.getElementById("mapLocalidad"),
      mapScanCode: document.getElementById("mapScanCode"),
      mapProducto: document.getElementById("mapProducto"),
      mapPedido: document.getElementById("mapPedido"),
      mapOrdenCompra: document.getElementById("mapOrdenCompra"),
      chkAutoOC: document.getElementById("chkAutoOC"),
      previewAutoOC: document.getElementById("previewAutoOC"),
      chkManualLocalidad: document.getElementById("chkManualLocalidad"),
      inputManualLocalidad: document.getElementById("inputManualLocalidad")
    };
  }

  initEvents() {
    // File Inputs
    this.els.fileInput.addEventListener("change", (e) =>
      this.handleFile(e.target.files[0])
    );

    // Drag & Drop
    window.addEventListener("dragover", (e) => {
      e.preventDefault();
      this.els.dragOverlay.classList.add("active");
    });
    window.addEventListener("dragleave", (e) => {
      if (e.target === this.els.dragOverlay)
        this.els.dragOverlay.classList.remove("active");
    });
    window.addEventListener("drop", (e) => {
      e.preventDefault();
      this.els.dragOverlay.classList.remove("active");
      if (e.dataTransfer.files.length) this.handleFile(e.dataTransfer.files[0]);
    });

    // Pagination
    document
      .getElementById("btnPrev")
      .addEventListener("click", () => this.changePage(-1));
    document
      .getElementById("btnNext")
      .addEventListener("click", () => this.changePage(1));
    document.getElementById("pageSize").addEventListener("change", (e) => {
      this.pageSize = parseInt(e.target.value);
      this.currentPage = 1;
      this.render();
    });

    // Search
    this.els.globalSearch.addEventListener("input", (e) => {
      this.searchQuery = e.target.value.toLowerCase();
      this.currentPage = 1;
      this.processData();
    });

    // Modal Sheet Selection
    document
      .getElementById("btnCloseSheetModal")
      .addEventListener("click", () => {
        this.els.sheetModal.classList.remove("active");
        this.setLoading(false);
        this.resetState();
      });

    // Export Modal Events
    this.els.btnCancelExport.addEventListener("click", () => {
      this.els.exportModal.classList.remove("active");
      this.pendingExportFormat = null;
    });

    this.els.btnConfirmExport.addEventListener("click", () => {
      this.els.reportTitle.value = this.els.confirmTitle.value;
      this.els.reportAuthor.value = this.els.confirmAuthor.value;
      this.executeExport(this.pendingExportFormat);
      this.els.exportModal.classList.remove("active");
      this.pendingExportFormat = null;
    });

    // EVENTOS NUEVOS PARA ESTRUCTURA
    this.els.btnCancelStructure.addEventListener("click", () => {
      this.els.structureModal.classList.remove("active");
      this.resetState();
    });

    this.els.btnConfirmStructure.addEventListener("click", () => {
      this.applyStructureAndLoad();
    });

    this.els.footerSkipCount.addEventListener("input", () => {
      this.renderPreviewTableRows();
    });

    // Click Outside
    document.addEventListener("click", (e) => {
      if (!e.target.closest("#btnColumns") && !e.target.closest("#colMenu"))
        this.els.colMenu.classList.remove("show");
      if (!e.target.closest("#btnExport") && !e.target.closest("#exportMenu"))
        this.els.exportMenu.classList.remove("show");
      if (
        !e.target.closest("#columnContextMenu") &&
        !e.target.closest(".btn-col-menu")
      )
        this.els.ctxMenu.classList.remove("show");
    });
  }

  // --- CORE FILE HANDLING --- //

  async handleFile(file) {
    if (!file) return;

    this.resetState();
    this.setLoading(true);
    this.els.fileInput.value = "";

    const fname = file.name.replace(/\.[^/.]+$/, "");
    this.els.reportTitle.value = fname;

    try {
      await this.forceRender();
      const data = await this.readFileAsync(file);

      const workbook = XLSX.read(data, { type: "array", cellDates: true });
      this.currentWorkbook = workbook;

      if (workbook.Props && workbook.Props.Author) {
        this.els.reportAuthor.value = workbook.Props.Author;
      } else {
        this.els.reportAuthor.value = "Usuario";
      }

      if (workbook.SheetNames.length === 0)
        throw new Error("Archivo vacío o inválido.");

      if (workbook.SheetNames.length > 1) {
        this.showSheetSelection(workbook.SheetNames);
        this.setLoading(false);
      } else {
        await this.loadSheetData(workbook.SheetNames[0]);
      }
    } catch (err) {
      console.error(err);
      this.showToast(`Error: ${err.message}`, "error");
      this.setLoading(false);
    }
  }

  readFileAsync(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => resolve(new Uint8Array(e.target.result));
      reader.onerror = (e) => reject(new Error("Error de lectura de archivo"));
      reader.readAsArrayBuffer(file);
    });
  }

  forceRender() {
    return new Promise((resolve) =>
      requestAnimationFrame(() => setTimeout(resolve, 0))
    );
  }

  resetState() {
    this.rawData = [];
    this.visibleData = [];
    this.columns = [];
    this.currentWorkbook = null;
    this.sortCol = null;
    this.searchQuery = "";
    this.els.globalSearch.value = "";
    this.els.tableWrapper.classList.add("hidden");
    this.els.footer.classList.add("hidden");
    this.filterSummary.classList.add("hidden");
    this.els.emptyState.classList.remove("hidden");
    this.els.thead.innerHTML = "";
    this.els.tbody.innerHTML = "";
    this.els.tfoot.innerHTML = "";
    this.tempRawMatrix = [];
  }

  showSheetSelection(sheets) {
    const list = this.els.sheetList;
    list.innerHTML = "";
    sheets.forEach((sheet) => {
      const btn = document.createElement("div");
      btn.className = "sheet-btn";
      btn.innerHTML = `<span style="font-weight:600">${sheet}</span> <i class="ph ph-caret-right"></i>`;
      btn.onclick = async () => {
        this.els.sheetModal.classList.remove("active");
        this.setLoading(true);
        await this.forceRender();
        await this.loadSheetData(sheet);
      };
      list.appendChild(btn);
    });
    this.els.sheetModal.classList.add("active");
  }

  // MODIFICADO: Ahora carga raw data y abre el selector
  async loadSheetData(sheetName) {
    try {
      if (!this.currentWorkbook) throw new Error("No hay libro cargado.");

      const sheet = this.currentWorkbook.Sheets[sheetName];
      // Leemos como matriz (header: 1)
      const rawMatrix = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        defval: ""
      });

      if (rawMatrix.length === 0)
        throw new Error(`La hoja "${sheetName}" está vacía.`);

      this.tempRawMatrix = rawMatrix;
      this.tempHeaderIdx = 0;
      this.els.footerSkipCount.value = 0;

      // LLAMADA A LA FUNCIÓN NUEVA
      this.openStructureSelector();
    } catch (err) {
      this.showToast(err.message, "error");
      this.setLoading(false);
    }
  }

  // --- NUEVAS FUNCIONES DE ESTRUCTURA (Dentro de la clase) ---

  openStructureSelector() {
    this.setLoading(false);
    this.els.sheetModal.classList.remove("active");
    this.els.structureModal.classList.add("active");

    // Intentar adivinar el header (primera fila con más de 1 dato)
    const likelyHeader = this.tempRawMatrix.findIndex(
      (row) => row && row.filter((c) => c).length > 1
    );
    this.tempHeaderIdx = likelyHeader >= 0 ? likelyHeader : 0;

    this.renderPreviewTableRows();
  }

  renderPreviewTableRows() {
    const table = this.els.previewTable;
    table.innerHTML = "";
    const footerSkip = parseInt(this.els.footerSkipCount.value) || 0;
    const totalRows = this.tempRawMatrix.length;

    this.els.selectedHeaderDisplay.innerText = `Fila ${this.tempHeaderIdx + 1}`;

    // Renderizar primeras 50 filas
    const limit = Math.min(this.tempRawMatrix.length, 50);

    for (let i = 0; i < limit; i++) {
      this.buildPreviewRow(table, i, totalRows, footerSkip);
    }

    // Si hay más, mostrar salto y el final
    if (totalRows > limit) {
      const tr = document.createElement("tr");
      tr.innerHTML = `<td colspan="100" style="text-align:center; padding:4px; font-style:italic; color:var(--text-muted)">... ${
        totalRows - limit
      } filas más ...</td>`;
      table.appendChild(tr);

      const startEnd = Math.max(limit, totalRows - 5);
      for (let i = startEnd; i < totalRows; i++) {
        this.buildPreviewRow(table, i, totalRows, footerSkip);
      }
    }
  }

  buildPreviewRow(table, index, totalRows, footerSkip) {
    const rowData = this.tempRawMatrix[index];
    if (!rowData) return;

    const tr = document.createElement("tr");

    const isHeader = index === this.tempHeaderIdx;
    const isIgnoredTop = index < this.tempHeaderIdx;
    const isIgnoredBottom = index >= totalRows - footerSkip;

    if (isHeader) tr.className = "preview-header";
    else if (isIgnoredTop || isIgnoredBottom) tr.className = "preview-ignored";

    tr.onclick = () => {
      this.tempHeaderIdx = index;
      this.renderPreviewTableRows();
    };

    // Numero de fila
    const tdNum = document.createElement("td");
    tdNum.className = "preview-row-num";
    tdNum.innerText = index + 1;
    tr.appendChild(tdNum);

    // Datos (limitado a 8 columnas para vista previa)
    const colLimit = Math.min(rowData.length, 8);
    for (let j = 0; j < colLimit; j++) {
      const td = document.createElement("td");
      td.innerText = rowData[j] !== undefined ? rowData[j] : "";
      tr.appendChild(td);
    }

    if (rowData.length > colLimit) {
      const td = document.createElement("td");
      td.innerText = "...";
      tr.appendChild(td);
    }

    table.appendChild(tr);
  }

  applyStructureAndLoad() {
    // 1. Cerrar el modal
    this.els.structureModal.classList.remove("active");

    // 2. Mostrar "Cargando..."
    this.setLoading(true);

    // Usamos setTimeout para dar tiempo al navegador de renderizar el spinner antes de bloquearse procesando
    setTimeout(() => {
      try {
        const headerIdx = this.tempHeaderIdx;
        const footerSkip = parseInt(this.els.footerSkipCount.value) || 0;

        // Validar que exista la fila header
        const headerRow = this.tempRawMatrix[headerIdx];
        if (!headerRow || headerRow.length === 0)
          throw new Error("La fila de encabezado seleccionada está vacía.");

        // 3. Definir nombres de Columnas (Manejo de duplicados y vacíos)
        const columns = [];
        headerRow.forEach((colName, idx) => {
          // Si el nombre es nulo/vacío, le ponemos "Columna_X"
          let safeName =
            colName !== undefined &&
            colName !== null &&
            String(colName).trim() !== ""
              ? String(colName).trim()
              : `Columna_${idx + 1}`;

          // Evitar nombres duplicados (ej: si hay dos columnas "Fecha", la segunda será "Fecha_1")
          if (columns.includes(safeName)) {
            let c = 1;
            while (columns.includes(`${safeName}_${c}`)) c++;
            safeName = `${safeName}_${c}`;
          }
          columns.push(safeName);
        });

        // 4. Construir el JSON final
        const startIndex = headerIdx + 1;
        const endIndex = this.tempRawMatrix.length - footerSkip;
        const jsonData = [];

        for (let i = startIndex; i < endIndex; i++) {
          const rowArr = this.tempRawMatrix[i];
          // Saltamos filas que sean null o vacías
          if (!rowArr || rowArr.length === 0) continue;

          const rowObj = {};
          let hasData = false;

          // Mapeamos array -> objeto usando las columnas definidas
          columns.forEach((colKey, colIdx) => {
            const cellVal = rowArr[colIdx];
            // Guardamos valor, o string vacío si es undefined
            rowObj[colKey] = cellVal !== undefined ? cellVal : "";

            // Verificamos si la fila tiene al menos un dato real
            if (cellVal !== undefined && cellVal !== "" && cellVal !== null)
              hasData = true;
          });

          // Solo agregamos la fila si tiene algún dato
          if (hasData) jsonData.push(rowObj);
        }

        if (jsonData.length === 0)
          throw new Error(
            "No se encontraron datos válidos con esta estructura."
          );

        // 5. Inicializar la tabla con la data procesada
        this.initData(jsonData);

        // =========================================================
        // CORRECCIÓN CRÍTICA: Desactivar el loading aquí
        // =========================================================
        this.setLoading(false);

        this.showToast(
          `Estructura aplicada. ${jsonData.length} filas cargadas.`,
          "success"
        );
      } catch (err) {
        console.error(err);
        this.showToast("Error procesando estructura: " + err.message, "error");
        // Asegurarnos de quitar el loading si hubo error
        this.setLoading(false);
      }
    }, 50); // Pequeño delay para asegurar que la UI responda
  }

  // --- FIN FUNCIONES NUEVAS ---

  initData(data) {
    this.rawData = data;
    this.columns = Object.keys(data[0]);

    this.colSettings = {};
    this.columns.forEach((col) => {
      this.colSettings[col] = {
        hidden: false,
        type: this.inferType(col, data),
        activeFilters: null,
        decimals: 2,
        dateStyle: "short"
      };
    });

    this.els.globalSearch.disabled = false;
    this.buildColumnPicker();
    this.processData();

    this.els.emptyState.classList.add("hidden");
    this.els.tableWrapper.classList.remove("hidden");
    this.els.footer.classList.remove("hidden");
  }

  // --- DATA PROCESSING & FILTER LOGIC --- //

  inferType(colName, data) {
    const lower = colName.toLowerCase();
    if (lower.match(/^(id|cod|sku|isbn|ean|item|ref|dni|ruc)/i)) return "text";
    if (lower.match(/(código|codigo|identificador)/i)) return "text";

    const sample = data.slice(0, 100).find((row) => row[colName] !== "");
    if (!sample) return "text";
    const val = sample[colName];

    if (val instanceof Date) return "date";
    if (typeof val === "number") {
      if (lower.match(/(precio|costo|total|valor|importe|venta|compra)/))
        return "currency";
      if (Number.isInteger(val)) return "integer";
      return "number";
    }
    if (String(val).startsWith("http")) return "link";
    return "text";
  }

  processData() {
    let processed = this.rawData.filter((row) => {
      return this.columns.every((col) => {
        const settings = this.colSettings[col];
        if (!settings.activeFilters) return true;
        return settings.activeFilters.has(String(row[col]));
      });
    });

    if (this.searchQuery) {
      processed = processed.filter((row) => {
        return Object.entries(row).some(([key, val]) => {
          if (this.colSettings[key].hidden) return false;
          return String(val).toLowerCase().includes(this.searchQuery);
        });
      });
    }

    if (this.sortCol) {
      processed.sort((a, b) => {
        let va = a[this.sortCol],
          vb = b[this.sortCol];
        if (typeof va === "string") va = va.toLowerCase();
        if (typeof vb === "string") vb = vb.toLowerCase();

        if (va < vb) return this.sortAsc ? -1 : 1;
        if (va > vb) return this.sortAsc ? 1 : -1;
        return 0;
      });
    }

    this.visibleData = processed;
    this.updatePaginationInfo();
    this.renderHeaders();
    this.render();
    this.renderFooterTotals();
    this.renderFilterSummary();
  }

  renderFilterSummary() {
    const bar = this.filterSummary;
    bar.innerHTML = "";
    const activeCols = this.columns.filter(
      (c) => this.colSettings[c].activeFilters !== null
    );

    if (activeCols.length === 0) {
      bar.classList.add("hidden");
      return;
    }

    bar.classList.remove("hidden");
    bar.innerHTML = `<span style="font-size:12px; font-weight:600; color:var(--text-muted)">Filtros activos:</span>`;

    activeCols.forEach((col) => {
      const chip = document.createElement("div");
      chip.className = "filter-chip";
      chip.innerHTML = `<span>${col}</span> <i class="ph ph-x" onclick="app.clearColFilter('${col}')"></i>`;
      bar.appendChild(chip);
    });

    if (activeCols.length > 1) {
      const clearAll = document.createElement("span");
      clearAll.className = "clear-filters-btn";
      clearAll.innerText = "Limpiar Todo";
      clearAll.onclick = () => {
        activeCols.forEach((c) => (this.colSettings[c].activeFilters = null));
        this.processData();
      };
      bar.appendChild(clearAll);
    }
  }

  // --- RENDERERS --- //

  renderHeaders() {
    this.els.thead.innerHTML = "";
    const tr = document.createElement("tr");

    this.columns.forEach((col) => {
      // 1. Validación de seguridad: Verificar que existan settings para la columna
      if (!this.colSettings[col] || this.colSettings[col].hidden) return;

      const th = document.createElement("th");
      const settings = this.colSettings[col];
      const alignClass = this.getAlignClass(settings.type);
      const isSorted = this.sortCol === col;
      const hasFilter = settings.activeFilters !== null;

      // Lógica del icono de ordenamiento
      const iconClass = isSorted
        ? this.sortAsc
          ? "ph-arrow-up"
          : "ph-arrow-down"
        : "";

      // 2. CORRECCIÓN CRÍTICA: Escapar comillas simples para evitar errores en onclick
      const safeCol = String(col).replace(/'/g, "\\'");

      // Template String limpio
      th.innerHTML = `
        <div class="th-content ${alignClass}">
          <div class="btn-col-menu ${hasFilter ? "active" : ""}" 
               onclick="app.openColumnMenu(event, '${safeCol}')">
             <i class="ph ${
               hasFilter ? "ph-funnel ph-fill" : "ph-dots-three-vertical"
             }"></i>
          </div>
          <div class="th-title" onclick="app.sortBy('${safeCol}')">
            <span>${col}</span>
            ${isSorted ? `<i class="ph ${iconClass} sort-icon"></i>` : ""}
          </div>
        </div>
      `;
      tr.appendChild(th);
    });
    this.els.thead.appendChild(tr);
  }

  render() {
    this.els.tbody.innerHTML = "";
    const start = (this.currentPage - 1) * this.pageSize;
    const end = start + this.pageSize;
    const pageData = this.visibleData.slice(start, end);
    const fragment = document.createDocumentFragment();

    pageData.forEach((row) => {
      const tr = document.createElement("tr");
      this.columns.forEach((col) => {
        if (this.colSettings[col].hidden) return;
        const td = document.createElement("td");
        const config = this.colSettings[col];
        td.className = this.getAlignClass(config.type);
        td.innerHTML = this.formatValue(row[col], config);
        td.addEventListener("dblclick", () => this.enableEditing(td, row, col));
        tr.appendChild(td);
      });
      fragment.appendChild(tr);
    });
    this.els.tbody.appendChild(fragment);
    this.updateFooterUI();
  }

  renderFooterTotals() {
    this.els.tfoot.innerHTML = "";
    let hasTotals = false;
    const tr = document.createElement("tr");

    this.columns.forEach((col, idx) => {
      if (this.colSettings[col].hidden) return;
      const td = document.createElement("td");
      const config = this.colSettings[col];
      const type = config.type;

      if (["number", "currency", "integer", "percent"].includes(type)) {
        const sum = this.visibleData.reduce(
          (acc, r) => acc + (parseFloat(r[col]) || 0),
          0
        );
        if (sum !== 0 && type !== "percent") {
          hasTotals = true;
          td.className = "text-right";
          td.innerHTML = this.formatValue(sum, config);
        }
      }
      if (idx === 0 && !hasTotals) td.innerText = "Totales";
      tr.appendChild(td);
    });
    if (hasTotals) this.els.tfoot.appendChild(tr);
  }

  sortBy(col) {
    if (this.sortCol === col) this.sortAsc = !this.sortAsc;
    else {
      this.sortCol = col;
      this.sortAsc = true;
    }
    this.processData();
  }

  changePage(delta) {
    const maxPages = Math.ceil(this.visibleData.length / this.pageSize);
    const newPage = this.currentPage + delta;
    if (newPage >= 1 && newPage <= maxPages) {
      this.currentPage = newPage;
      this.render();
      this.els.tableWrapper.scrollTop = 0;
    }
  }

  updatePaginationInfo() {
    const total = this.visibleData.length;
    document.getElementById(
      "statusMsg"
    ).innerText = `${total.toLocaleString()} registros encontrados`;
    this.updateFooterUI();
  }

  updateFooterUI() {
    const maxPages = Math.ceil(this.visibleData.length / this.pageSize) || 1;
    document.getElementById("currPage").innerText = this.currentPage;
    document.getElementById("totalPages").innerText = maxPages;
    document.getElementById("btnPrev").disabled = this.currentPage <= 1;
    document.getElementById("btnNext").disabled = this.currentPage >= maxPages;
  }

  // --- COLUMN MENUS & FILTERS --- //

  openColumnMenu(e, col) {
    e.stopPropagation();
    this.activeMenuCol = col;
    const menu = this.els.ctxMenu;

    const rect = e.currentTarget.getBoundingClientRect();
    let top = rect.bottom + 5;
    let left = rect.left;
    if (left + 280 > window.innerWidth) left = window.innerWidth - 290;
    if (left < 0) left = 10;

    menu.style.top = top + "px";
    menu.style.left = left + "px";

    this.renderMenuContent(col, menu);
    menu.classList.add("show");

    const menuRect = menu.getBoundingClientRect();
    if (menuRect.bottom > window.innerHeight) {
      menu.style.top = "auto";
      menu.style.bottom = "10px";
    }
  }

  renderMenuContent(col, container) {
    const settings = this.colSettings[col];

    const relevantRows = this.rawData.filter((row) => {
      return this.columns.every((c) => {
        if (c === col) return true;
        const s = this.colSettings[c];
        if (!s.activeFilters) return true;
        return s.activeFilters.has(String(row[c]));
      });
    });
    const uniqueVals = [
      ...new Set(relevantRows.map((r) => String(r[col])))
    ].sort();

    let extraControls = "";
    if (["number", "currency", "percent"].includes(settings.type)) {
      extraControls += `
             <div style="margin-top:8px; display:flex; align-items:center; justify-content:space-between;">
                <label class="col-menu-label" style="margin:0">Decimales</label>
                <input type="number" min="0" max="6" class="form-input form-input-sm" style="width:60px" value="${settings.decimals}" 
                       onchange="app.changeColDecimal('${col}', this.value)">
             </div>
          `;
    }

    if (settings.type === "currency") {
      extraControls += `
             <div style="margin-top:8px">
                <label class="col-menu-label" style="margin-bottom:2px">Simbolo</label>
                <select class="form-select" onchange="app.changeColCurrency('${col}', this.value)">
                   <option value="PEN" ${
                     settings.currency === "PEN" ? "selected" : ""
                   }>S/ (PEN)</option>
                   <option value="USD" ${
                     settings.currency === "USD" ? "selected" : ""
                   }>$ (USD)</option>
                   <option value="EUR" ${
                     settings.currency === "EUR" ? "selected" : ""
                   }>€ (EUR)</option>
                   <option value="GBP" ${
                     settings.currency === "GBP" ? "selected" : ""
                   }>£ (GBP)</option>
                   <option value="JPY" ${
                     settings.currency === "JPY" ? "selected" : ""
                   }>¥ (JPY)</option>
                </select>
             </div>
          `;
    }

    if (["date", "datetime"].includes(settings.type)) {
      extraControls += `
             <div style="margin-top:8px">
                <label class="col-menu-label" style="margin-bottom:2px">Estilo</label>
                <select class="form-select" onchange="app.changeColDateStyle('${col}', this.value)">
                   <option value="short" ${
                     settings.dateStyle === "short" ? "selected" : ""
                   }>Corto (DD/MM/YYYY)</option>
                   <option value="medium" ${
                     settings.dateStyle === "medium" ? "selected" : ""
                   }>Medio (04 ene 2026)</option>
                   <option value="long" ${
                     settings.dateStyle === "long" ? "selected" : ""
                   }>Largo (4 de enero...)</option>
                   <option value="full" ${
                     settings.dateStyle === "full" ? "selected" : ""
                   }>Texto (Lunes...)</option>
                </select>
             </div>
          `;
    }

    container.innerHTML = `
        <div class="col-menu-section">
          <label class="col-menu-label">Formato</label>
          <select class="form-select" onchange="app.changeColFormat('${col}', this.value)">
            <option value="auto" ${
              settings.type === "auto" ? "selected" : ""
            }>Automático</option>
            <option value="text" ${
              settings.type === "text" ? "selected" : ""
            }>Texto</option>
            <option value="number" ${
              settings.type === "number" ? "selected" : ""
            }>Número</option>
            <option value="integer" ${
              settings.type === "integer" ? "selected" : ""
            }>Entero</option>
            <option value="currency" ${
              settings.type === "currency" ? "selected" : ""
            }>Moneda</option>
            <option value="percent" ${
              settings.type === "percent" ? "selected" : ""
            }>Porcentaje (%)</option>
            <option value="date" ${
              settings.type === "date" ? "selected" : ""
            }>Fecha</option>
            <option value="datetime" ${
              settings.type === "datetime" ? "selected" : ""
            }>Fecha y Hora</option>
            <option value="time" ${
              settings.type === "time" ? "selected" : ""
            }>Hora</option>
            <option value="link" ${
              settings.type === "link" ? "selected" : ""
            }>Link</option>
          </select>
          ${extraControls}
        </div>
        <div class="col-menu-section">
          <label class="col-menu-label">Filtrar (${uniqueVals.length})</label>
          <input type="text" class="form-input form-input-sm" placeholder="Buscar..." oninput="app.filterMenuSearch(this.value)">
          <div class="filter-list" id="filterListContainer"></div>
          <div style="display:flex; justify-content:space-between; margin-top:8px;">
             <button class="btn btn-sm" onclick="app.clearColFilter('${col}')">Limpiar</button>
             <button class="btn btn-sm btn-primary" onclick="app.applyColFilter('${col}')">Aplicar</button>
          </div>
        </div>
      `;

    const filterContainer = container.querySelector("#filterListContainer");
    const allDiv = document.createElement("div");
    allDiv.className = "filter-item";
    allDiv.innerHTML = `<input type="checkbox" id="chkAllFilters" ${
      settings.activeFilters === null ? "checked" : ""
    }> <span>(Seleccionar Todo)</span>`;
    allDiv.onclick = (ev) => {
      if (ev.target.tagName !== "INPUT") {
        const chk = allDiv.querySelector("input");
        chk.checked = !chk.checked;
      }
      const state = allDiv.querySelector("input").checked;
      filterContainer
        .querySelectorAll(".val-chk")
        .forEach((c) => (c.checked = state));
    };
    filterContainer.appendChild(allDiv);

    const maxItems = 2000;
    const isLimited = uniqueVals.length > maxItems;
    const valsToShow = isLimited ? uniqueVals.slice(0, maxItems) : uniqueVals;

    valsToShow.forEach((val) => {
      const div = document.createElement("div");
      div.className = "filter-item val-item";
      div.setAttribute("data-val", val.toLowerCase());
      const displayVal = val === "" ? "(Vacío)" : val;
      const isChecked =
        settings.activeFilters === null
          ? true
          : settings.activeFilters.has(val);
      div.innerHTML = `<input type="checkbox" class="val-chk" value="${val}" ${
        isChecked ? "checked" : ""
      }> <span>${displayVal}</span>`;
      div.onclick = (ev) => {
        if (ev.target.tagName !== "INPUT") {
          const chk = div.querySelector("input");
          chk.checked = !chk.checked;
        }
        if (!div.querySelector("input").checked) {
          const ac = container.querySelector("#chkAllFilters");
          if (ac) ac.checked = false;
        }
      };
      filterContainer.appendChild(div);
    });

    if (isLimited) {
      const w = document.createElement("div");
      w.style.fontSize = "10px";
      w.style.color = "var(--text-muted)";
      w.innerText = `...y ${uniqueVals.length - maxItems} más.`;
      filterContainer.appendChild(w);
    }
  }

  filterMenuSearch(term) {
    term = term.toLowerCase();
    const items = document.querySelectorAll("#filterListContainer .val-item");
    items.forEach((el) => {
      el.style.display = el.getAttribute("data-val").includes(term)
        ? "flex"
        : "none";
    });
  }

  applyColFilter(col) {
    const inputs = document.querySelectorAll("#filterListContainer .val-chk");
    const allChk = document.getElementById("chkAllFilters");

    if (allChk && allChk.checked) {
      this.colSettings[col].activeFilters = null;
    } else {
      const selected = new Set();
      inputs.forEach((inp) => {
        if (inp.checked) selected.add(inp.value);
      });
      this.colSettings[col].activeFilters = selected;
    }
    this.els.ctxMenu.classList.remove("show");
    this.currentPage = 1;
    this.processData();
  }

  clearColFilter(col) {
    this.colSettings[col].activeFilters = null;
    this.els.ctxMenu.classList.remove("show");
    this.processData();
  }

  // --- LIVE SETTINGS UPDATES (No Close) --- //

  changeColFormat(col, type) {
    this.colSettings[col].type = type;
    this.updateAndKeepMenu(col);
  }

  changeColDecimal(col, val) {
    this.colSettings[col].decimals = parseInt(val) || 0;
    this.updateAndKeepMenu(col);
  }

  changeColDateStyle(col, val) {
    this.colSettings[col].dateStyle = val;
    this.updateAndKeepMenu(col);
  }

  changeColCurrency(col, val) {
    this.colSettings[col].currency = val;
    this.updateAndKeepMenu(col);
  }

  updateAndKeepMenu(col) {
    this.render();
    this.renderHeaders();
    this.renderFooterTotals();
    this.renderMenuContent(col, this.els.ctxMenu);
  }

  // --- EDICIÓN DE DATOS --- //

  enableEditing(td, row, col) {
    // Evitar re-entradas si ya se está editando
    if (td.querySelector("input")) return;

    const currentVal = row[col];
    const config = this.colSettings[col];
    const type = config.type;

    // Guardar el HTML original por si cancela (tecla Esc)
    const originalHtml = td.innerHTML;

    td.classList.add("cell-editing");
    td.innerHTML = "";

    // Crear Input
    const input = document.createElement("input");
    input.className = "table-input";

    // Configurar tipo de input
    if (["number", "currency", "integer", "percent"].includes(type)) {
      input.type = "number";
      input.step = "any"; // Permitir decimales
      input.value = currentVal;
    } else if (type === "date") {
      // Para inputs tipo fecha, necesitamos formato yyyy-MM-dd
      input.type = "date";
      try {
        if (currentVal instanceof Date)
          input.value = currentVal.toISOString().split("T")[0];
        else if (currentVal)
          input.value = new Date(currentVal).toISOString().split("T")[0];
      } catch (e) {
        input.value = "";
      }
    } else {
      input.type = "text";
      input.value = currentVal !== undefined ? currentVal : "";
    }

    // Eventos del Input
    input.addEventListener("blur", () =>
      this.saveEdit(td, row, col, input.value)
    );
    input.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        input.blur(); // Dispara el guardado
      } else if (e.key === "Escape") {
        // Cancelar edición
        td.classList.remove("cell-editing");
        td.innerHTML = originalHtml;
      }
    });

    td.appendChild(input);
    input.focus();
  }

  saveEdit(td, row, col, newVal) {
    const config = this.colSettings[col];
    const type = config.type;
    let finalVal = newVal;

    // Convertir tipos de datos
    if (["number", "currency", "percent"].includes(type)) {
      if (newVal === "") finalVal = 0;
      else finalVal = parseFloat(newVal);
    } else if (type === "integer") {
      if (newVal === "") finalVal = 0;
      else finalVal = parseInt(newVal);
    } else if (type === "date" || type === "datetime") {
      // Intentar mantenerlo como objeto Date si el origen era Date
      if (newVal) {
        // El input date devuelve yyyy-mm-dd, forzamos zona horaria local para no perder el día
        const parts = newVal.split("-");
        finalVal = new Date(parts[0], parts[1] - 1, parts[2]);
      } else {
        finalVal = "";
      }
    }

    // Actualizar DATOS (Referencia en memoria)
    row[col] = finalVal;

    // Restaurar vista
    td.classList.remove("cell-editing");
    td.innerHTML = this.formatValue(finalVal, config);

    // Recalcular Totales del Footer
    this.renderFooterTotals();

    // Opcional: Feedback visual
    td.style.backgroundColor = "rgba(16, 185, 129, 0.1)"; // Verde suave flash
    setTimeout(() => (td.style.backgroundColor = ""), 500);
  }

  // --- HELPERS : COLUMN PICKER --- //
  buildColumnPicker() {
    this.renderColumnList(this.columns);
  }
  renderColumnList(cols) {
    this.els.colListContainer.innerHTML = "";
    cols.forEach((col) => {
      const item = document.createElement("div");
      item.className = "dropdown-item";
      item.innerHTML = `<input type="checkbox" ${
        !this.colSettings[col].hidden ? "checked" : ""
      }><span>${col}</span>`;
      item.onclick = (e) => {
        const chk = item.querySelector("input");
        if (e.target !== chk) chk.checked = !chk.checked;
        this.colSettings[col].hidden = !chk.checked;
        this.processData();
      };
      this.els.colListContainer.appendChild(item);
    });
  }
  filterColumnList(term) {
    const lower = term.toLowerCase();
    this.renderColumnList(
      this.columns.filter((c) => c.toLowerCase().includes(lower))
    );
  }
  toggleAllColumns(show) {
    this.columns.forEach((col) => (this.colSettings[col].hidden = !show));
    this.buildColumnPicker();
    this.processData();
  }

  // --- HELPERS : EXPORT --- //
  exportTo(format) {
    this.els.exportMenu.classList.remove("show");
    if (!this.visibleData.length)
      return this.showToast("No hay datos", "error");
    this.pendingExportFormat = format;
    this.els.confirmTitle.value = this.els.reportTitle.value;
    this.els.confirmAuthor.value = this.els.reportAuthor.value;
    this.els.exportModal.classList.add("active");
  }

  executeExport(format) {
    let fname = this.els.reportTitle.value.trim() || "Reporte";
    fname = fname.replace(/[^a-z0-9_\-\sáéíóúñ]/gi, "_");
    const author = this.els.reportAuthor.value.trim();
    const timestamp = new Date().toLocaleString();

    const exportData = this.visibleData.map((row) => {
      const newRow = {};
      this.columns.forEach((col) => {
        if (this.colSettings[col].hidden) return;
        const val = row[col];
        const config = this.colSettings[col];
        const type = config.type;

        if (format === "xlsx") {
          newRow[col] = type === "text" ? String(val) : val;
        } else {
          const d = document.createElement("div");
          d.innerHTML = this.formatValue(val, config);
          newRow[col] = d.textContent.trim();
        }
      });
      return newRow;
    });

    if (format === "xlsx") {
      const ws = XLSX.utils.json_to_sheet(exportData);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Data");
      XLSX.writeFile(wb, `${fname}.xlsx`);
    } else if (format === "html") {
      const esc = (str) => {
        if (str === null || str === undefined) return "";
        return String(str)
          .replace(/&/g, "&amp;")
          .replace(/</g, "&lt;")
          .replace(/>/g, "&gt;")
          .replace(/"/g, "&quot;")
          .replace(/'/g, "&#039;");
      };

      let rowsHtml = "";
      let totalRowHtml = "";
      const totals = {};

      this.columns.forEach((col) => {
        const type = this.colSettings[col].type;
        if (
          ["number", "currency", "integer", "percent"].includes(type) &&
          !this.colSettings[col].hidden
        ) {
          totals[col] = 0;
        }
      });

      this.visibleData.forEach((row) => {
        let tr = "<tr>";
        this.columns.forEach((col) => {
          if (this.colSettings[col].hidden) return;
          const config = this.colSettings[col];
          const type = config.type;
          let val = row[col];
          let cellHtml = "";
          let alignClass = "text-left";
          let cssClass = "col-text";
          let dataAttrs = "";

          const isNum = typeof val === "number";
          if (isNum && !["date", "datetime", "time"].includes(type)) {
            alignClass = "text-right";
            cssClass = "col-num";
            dataAttrs = ` data-val="${val}"`;
            if (totals[col] !== undefined) totals[col] += val;
          }

          if (
            type === "link" ||
            (typeof val === "string" && val.startsWith("http"))
          ) {
            cellHtml = `<a href="${esc(val)}" target="_blank">${esc(val)}</a>`;
          } else if (
            ["date", "datetime"].includes(type) ||
            val instanceof Date
          ) {
            try {
              const d = new Date(val);
              if (!isNaN(d)) {
                const isDateOnly = type === "date";
                const iso = d.toISOString();
                const fmtVal = isDateOnly
                  ? iso.split("T")[0]
                  : iso.slice(0, 16);
                const inputType = isDateOnly ? "date" : "datetime-local";
                cellHtml = `<input type="${inputType}" value="${fmtVal}" readonly onclick="this.showPicker()" onkeydown="return false">`;
              } else {
                cellHtml = esc(val);
              }
            } catch (e) {
              cellHtml = esc(val);
            }
          } else {
            let txt = String(val);
            if (!isNum && txt.length < 20 && txt.length > 2) {
              const up = txt.toUpperCase();
              let badge = "";
              if (up.match(/APROB|OK|ACTIVO|PAGADO|COMPLET|SI|SÍ/))
                badge = "success";
              else if (up.match(/ERROR|RECHAZ|CANCEL|BAJA|NO|FALLO/))
                badge = "danger";
              else if (up.match(/PEND|REV|PROC|ESPERA/)) badge = "warning";
              else if (up.match(/NUEVO|INFO/)) badge = "info";

              if (badge)
                cellHtml = `<span class="badge ${badge}">${esc(txt)}</span>`;
              else cellHtml = esc(txt);
            } else if (isNum) {
              cellHtml = this.formatValue(val, config);
            } else {
              cellHtml = esc(txt);
            }
          }
          tr += `<td class="${cssClass} ${alignClass}"${dataAttrs}>${cellHtml}</td>`;
        });
        tr += "</tr>";
        rowsHtml += tr;
      });

      let hasTotals = Object.keys(totals).length > 0;
      if (hasTotals) {
        totalRowHtml = '<tr class="row-total">';
        this.columns.forEach((col, idx) => {
          if (this.colSettings[col].hidden) return;
          let td = "";
          if (idx === 0) td = "<td>TOTAL</td>";
          else {
            if (totals[col] !== undefined) {
              const config = this.colSettings[col];
              const sum = totals[col];
              const fmtVal = this.formatValue(sum, config);
              td = `<td class="text-right" data-sum="1" data-fmt="${esc(
                config.currency || ""
              )}">${fmtVal}</td>`;
            } else {
              td = "<td></td>";
            }
          }
          totalRowHtml += td;
        });
        totalRowHtml += "</tr>";
      }

      let headersHtml = "";
      let colIndex = 0;
      let filterOptions = '<option value="all">Todas las columnas</option>';
      this.columns.forEach((col) => {
        if (this.colSettings[col].hidden) return;
        const isNum = ["number", "currency", "integer", "percent"].includes(
          this.colSettings[col].type
        );
        const align = isNum ? "text-right" : "text-left";
        headersHtml += `<th class="${align}" onclick="sortGrid(${colIndex})">${esc(
          col
        )}</th>`;
        filterOptions += `<option value="${colIndex}">${esc(col)}</option>`;
        colIndex++;
      });

      const fullHtml = `<!DOCTYPE html>
<html lang='es'>
<head>
  <meta charset='UTF-8'>
  <meta name='viewport' content='width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no'>
  <title>${esc(fname)}</title>
  <style>
    :root { --bg-body: #f8fafc; --bg-gradient: linear-gradient(135deg, #f0f9ff 0%, #e0e7ff 100%); --bg-card: #ffffff; --text-main: #0f172a; --text-muted: #64748b; --primary: #0f172a; --accent: #0ea5e9; --border: #e2e8f0; --table-head: #f8fafc; --table-head-text: #334155; --row-hover: #f1f5f9; --shadow-soft: 0 10px 30px -10px rgba(0,0,0,0.08); --shadow-float: 0 20px 25px -5px rgba(0, 0, 0, 0.1); --link-color: #0284c7; --modal-item-bg: #f8fafc; }
    [data-theme='dark'] { --bg-body: #0f172a; --bg-gradient: linear-gradient(135deg, #0f172a 0%, #1e1b4b 100%); --bg-card: #1e293b; --text-main: #f8fafc; --text-muted: #94a3b8; --primary: #818cf8; --accent: #38bdf8; --border: #334155; --table-head: #1e293b; --table-head-text: #cbd5e1; --row-hover: #334155; --shadow-soft: 0 10px 30px -10px rgba(0,0,0,0.5); --link-color: #7dd3fc; --modal-item-bg: #0f172a; }
    *, *::before, *::after { box-sizing: border-box; -webkit-tap-highlight-color: transparent; transition: all 0.2s ease; }
    html, body { height: 100%; height: 100dvh; margin: 0; padding: 0; overflow: hidden; font-family: 'Inter', system-ui, -apple-system, sans-serif; background: var(--bg-gradient); color: var(--text-main); }
    a { color: var(--link-color); text-decoration: none; font-weight: 500; }
    a:hover { text-decoration: underline; }
    .page { height: 100%; display: flex; flex-direction: column; padding: 24px; max-width: 2000px; margin: 0 auto; gap: 20px; align-items: center; }
    .header-container { display: flex; flex-direction: column; gap: 20px; flex-shrink: 0; width: 100%; max-width: 100%; }
    .title-area { text-align: center; }
    .title-area h1 { font-size: 28px; font-weight: 800; letter-spacing: -0.5px; margin: 0; background: linear-gradient(to right, var(--primary), var(--accent)); -webkit-background-clip: text; -webkit-text-fill-color: transparent; }
    .subtitle-area { font-size: 12px; color: var(--text-muted); font-weight: 400; font-style: italic; }
    .meta-row { display: flex; justify-content: space-between; align-items: center; gap: 12px; margin: 0 auto; }
    .actions-area { display: flex; align-items: center; gap: 12px; }
    .author-pill { display: inline-flex; align-items: center; gap: 6px; padding: 4px 10px; background: rgba(79, 70, 229, 0.1); color: var(--primary); border-radius: 20px; font-size: 11px; font-weight: 600; letter-spacing: 0.5px; }
    .btn-group { display: flex; gap: 8px; }
    .btn { display: inline-flex; align-items: center; justify-content: center; height: 36px; border-radius: 99px; border: 1px solid var(--border); background: var(--bg-card); color: var(--text-muted); cursor: pointer; padding: 0 14px; font-size: 13px; font-weight: 600; gap: 6px; }
    .btn:hover { background: var(--bg-body); color: var(--primary); transform: translateY(-2px); box-shadow: 0 4px 12px rgba(0,0,0,0.1); border-color: var(--primary); }
    .btn-primary { background: var(--bg-card); color: #fff; border: none; }
    .btn-primary:hover { background: var(--accent); color: #fff; box-shadow: 0 4px 12px rgba(14, 165, 233, 0.3); }
    .search-floater { width: 100%; max-width: 600px; height: 40px; background: var(--bg-card); border-radius: 16px; padding: 3px; align-items: center; box-shadow: var(--shadow-float); display: flex; gap: 8px; border: 1px solid var(--border); animation: floatUp 0.6s ease-out; }
    .select-wrapper { position: relative; border-right: 1px solid var(--border); }
    .filter-select { appearance: none; background: transparent; border: none; padding: 10px 30px 10px 16px; font-size: 13px; font-weight: 600; color: var(--text-main); cursor: pointer; outline: none; height: 100%; }
    .filter-select option { background-color: var(--bg-card); color: var(--text-main); }
    .select-arrow { position: absolute; right: 8px; top: 50%; transform: translateY(-50%); pointer-events: none; color: var(--text-muted); width: 12px; }
    .search-input-wrapper { flex-grow: 1; position: relative; }
    .search-input { width: 100%; border: none; background: transparent; padding: 10px 12px; font-size: 14px; color: var(--text-main); outline: none; }
    .table-card { width: fit-content; max-width: 100%; background: rgba(255, 255, 255, 0.7); backdrop-filter: blur(10px); border-radius: 20px; box-shadow: var(--shadow-soft); flex-grow: 1; overflow: hidden; border: 1px solid rgba(255,255,255,0.5); display: flex; flex-direction: column; }
    [data-theme='dark'] .table-card { background: rgba(30, 41, 59, 0.7); border-color: rgba(255,255,255,0.1); }
    .table-container { overflow: auto; flex-grow: 1; position: relative; width: 100%; }
    table { width: auto; border-collapse: separate; border-spacing: 0; font-size: 13px; }
    th, td { white-space: nowrap; }
    thead th { position: sticky; top: 0; background: var(--bg-card); color: var(--text-muted); padding: 10px 16px; font-weight: 700; text-transform: uppercase; font-size: 14px; letter-spacing: 0.8px; border-bottom: 2px solid var(--border); cursor: pointer; z-index: 20; transition: background 0.2s; }
    thead th:hover { background: var(--row-hover); color: var(--primary); }
    thead th::after { content: ''; display: inline-block; margin-left: 8px; vertical-align: middle; border-left: 4px solid transparent; border-right: 4px solid transparent; opacity: 0; transition: opacity 0.2s; }
    thead th:hover::after { opacity: 0.5; border-top: 4px solid currentColor; }
    thead th.asc::after { opacity: 1; border-bottom: 4px solid var(--accent); border-top: none; }
    thead th.desc::after { opacity: 1; border-top: 4px solid var(--accent); border-bottom: none; }
    .text-left { text-align: left; } .text-right { text-align: right; }
    td { padding: 8px 16px; border-bottom: 1px solid var(--border); color: var(--text-main); font-weight: 500; height: 38px; }
    tbody tr:hover { background-color: var(--row-hover); }
    .row-total td { position: sticky; bottom: 0; background-color: var(--bg-card); border-top: 2px solid var(--primary); color: var(--primary); font-weight: 800; z-index: 30; box-shadow: 0 -4px 20px rgba(0,0,0,0.1); padding: 10px 16px; }
    input[type=date], input[type=datetime-local] { border: 0; padding: 0; margin: 0; background: transparent; font-family: inherit; font-size: inherit; color: inherit; cursor: pointer; outline: none; width: auto; }
    .badge { padding: 4px 10px; border-radius: 8px; font-size: 11px; font-weight: 700; }
    .badge.success { background: #dcfce7; color: #166534; } .badge.danger { background: #fee2e2; color: #991b1b; } .badge.warning { background: #fef9c3; color: #854d0e; } .badge.info { background: #e0f2fe; color: #075985; }
    .modal-overlay { position: fixed; top: 0; left: 0; width: 100vw; height: 100vh; background: rgba(15, 23, 42, 0.6); backdrop-filter: blur(4px); z-index: 9999; opacity: 0; visibility: hidden; transition: all 0.25s ease; display: flex; align-items: center; justify-content: center; padding: 20px; }
    .modal-overlay.active { opacity: 1; visibility: visible; }
    .modal-card { background: var(--bg-card); width: 100%; max-width: 450px; max-height: 85vh; border-radius: 12px; box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.25); transform: scale(0.95); transition: all 0.25s; display: flex; flex-direction: column; overflow: hidden; border: 1px solid var(--border); }
    .modal-overlay.active .modal-card { transform: scale(1); }
    .modal-header { padding: 16px 24px; border-bottom: 1px solid var(--border); display: flex; justify-content: space-between; align-items: center; background: var(--bg-card); z-index: 10; }
    .modal-title { font-size: 18px; font-weight: 700; margin:0; color: var(--primary); }
    .modal-body { padding: 0; overflow-y: auto; display: flex; flex-direction: column; }
    .detail-item { padding: 16px 24px; border-bottom: 1px dashed var(--border); display: flex; flex-direction: column; gap: 4px; }
    .detail-item:last-child { border-bottom: none; }
    .detail-label { font-size: 11px; text-transform: uppercase; color: var(--text-muted); font-weight: 600; letter-spacing: 0.5px; }
    .detail-value { font-size: 15px; color: var(--text-main); font-weight: 500; word-break: break-word; line-height: 1.5; }
    @keyframes floatUp { from { opacity: 0; transform: translateY(20px); } to { opacity: 1; transform: translateY(0); } }
    @media (max-width: 768px) { .page { padding: 0; height: 100dvh; gap: 0; align-items: stretch; } .table-card { border-radius: 0; border: none; background: var(--bg-card); width: 100%; } table { width: 100%; } .header-container { padding: 16px; background: var(--bg-card); border-bottom: 1px solid var(--border); gap: 12px; } .meta-row { flex-direction: column; gap: 12px; align-items: flex-start; } .actions-area { width: 100%; justify-content: space-between; } .search-floater { max-width: 100%; box-shadow: none; background: var(--bg-body); } .title-area h1 { font-size: 22px; text-align: left; } }
    @media print {
      html, body, .page, .table-card, .table-container { height: auto !important; min-height: 0 !important; overflow: visible !important; width: 100% !important; margin: 0 !important; padding: 0 !important; display: block !important; background: #fff !important; color: #000 !important; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
      .search-floater, .actions-area, .btn, .modal-overlay, .select-wrapper, #filterCol { display: none !important; }
      table { width: 100% !important; border-collapse: collapse !important; table-layout: auto !important; }
      th, td { white-space: normal !important; word-wrap: break-word !important; padding: 6px 4px !important; border: 1px solid #000 !important; font-size: 10px !important; height: auto !important; color: #000 !important; background: #fff !important; }
      thead { display: table-header-group !important; }
      thead th { background: #eee !important; color: #000 !important; position: static !important; border-bottom: 2px solid #000 !important; }
      tr { break-inside: avoid; page-break-inside: avoid; background: #fff !important; }
      .row-total td { position: static !important; background: #f0f0f0 !important; color: #000 !important; border-top: 2px solid #000 !important; }
      .header-container { display: block !important; padding-bottom: 20px; border-bottom: 2px solid #000; margin-bottom: 20px; }
      .title-area h1 { color: #000 !important; -webkit-text-fill-color: initial !important; background: none !important; text-align: left !important; font-size: 18px !important; }
      .subtitle-area { color: #000 !important; }
      .author-pill { border: 1px solid #000 !important; color: #000 !important; background: none !important; }
      .table-card { box-shadow: none !important; border: none !important; }
      .badge { border: 1px solid #000; background: #fff; color: #000; }
      :root { --bg-body: #ffffff !important; --text-main: #000000 !important; --bg-card: #ffffff !important; }
    }
  </style>
  <script>
    var curSort = { col: -1, dir: 'asc' };
    function printRpt(){window.print()}
    function toggleTheme() { var h = document.documentElement; var c = h.getAttribute('data-theme'); var n = c === 'dark' ? 'light' : 'dark'; h.setAttribute('data-theme', n); localStorage.setItem('theme', n); }
    function updateDatalist() { var colIdx = document.getElementById('filterCol').value; var list = document.getElementById('search-options'); list.innerHTML = ''; var set = new Set(); var rows = document.querySelectorAll('tbody tr:not(.row-total)'); rows.forEach(function(row) { var td; if(colIdx === 'all') {} else { td = row.children[parseInt(colIdx)]; if(td) { var txt = td.innerText.trim(); if(txt && txt.length > 1 && txt.length < 50) set.add(txt); } } }); var arr = Array.from(set).sort().slice(0, 500); arr.forEach(function(val) { var opt = document.createElement('option'); opt.value = val; list.appendChild(opt); }); }
    function onSelectorChange() { document.getElementById('search').value = ''; filterTbl(); updateDatalist(); }
    function filterTbl(){ var term = document.getElementById('search').value.toLowerCase(); var colIdx = document.getElementById('filterCol').value; var tbody = document.querySelector('tbody'); var rows = tbody.querySelectorAll('tr:not(.row-total)'); var totalRow = tbody.querySelector('.row-total'); var sums = []; rows.forEach(function(row) { var visible = false; var tds = row.querySelectorAll('td'); if(colIdx === 'all') { for(var i=0; i<tds.length; i++) { if(tds[i].innerText.toLowerCase().indexOf(term) > -1) { visible = true; break; } } } else { var targetTd = tds[colIdx]; if(targetTd && targetTd.innerText.toLowerCase().indexOf(term) > -1) visible = true; } row.style.display = visible ? '' : 'none'; if(visible) { tds.forEach(function(td, idx) { var val = parseFloat(td.getAttribute('data-val')); if(!isNaN(val)) { if(!sums[idx]) sums[idx] = 0; sums[idx] += val; } }); } }); if(totalRow) { totalRow.querySelectorAll('td').forEach(function(td, n) { if(td.hasAttribute('data-sum')) { var fmt = td.getAttribute('data-fmt')||''; var isCur = fmt.length > 0 && !fmt.includes('%'); var isPct = fmt.includes('%'); var sum = sums[n] || 0; var txt = ''; if(isPct) txt = (sum * 100).toFixed(2) + '%'; else txt = sum.toLocaleString('es-PE', {minimumFractionDigits:2, maximumFractionDigits:2}); if(isCur) txt = fmt + ' ' + txt; td.innerText = txt; } }); } }
    function sortGrid(idx) { var tbody = document.querySelector('tbody'); var rows = Array.from(tbody.querySelectorAll('tr:not(.row-total)')); var totalRow = tbody.querySelector('.row-total'); if (curSort.col === idx) { curSort.dir = curSort.dir === 'asc' ? 'desc' : 'asc'; } else { curSort.col = idx; curSort.dir = 'asc'; } document.querySelectorAll('thead th').forEach(function(th) { th.className = th.className.replace(/ asc| desc/g, ''); }); document.querySelectorAll('thead th')[idx].className += ' ' + curSort.dir; rows.sort(function(a, b) { var cellA = a.children[idx]; var cellB = b.children[idx]; var valA = cellA.hasAttribute('data-val') ? parseFloat(cellA.getAttribute('data-val')) : null; var valB = cellB.hasAttribute('data-val') ? parseFloat(cellB.getAttribute('data-val')) : null; if (valA !== null && valB !== null) return curSort.dir === 'asc' ? valA - valB : valB - valA; var inpA = cellA.querySelector('input'); var inpB = cellB.querySelector('input'); var txtA = inpA ? inpA.value : cellA.innerText.trim().toLowerCase(); var txtB = inpB ? inpB.value : cellB.innerText.trim().toLowerCase(); return curSort.dir === 'asc' ? txtA.localeCompare(txtB) : txtB.localeCompare(txtA); }); rows.forEach(function(r) { tbody.appendChild(r); }); if(totalRow) tbody.appendChild(totalRow); }
    function dlXLS() { var rows = document.querySelectorAll('table tr'); var csv = []; rows.forEach(function(row) { if(row.style.display !== 'none') { var cols = []; row.querySelectorAll('th, td').forEach(function(cell) { var txt = cell.querySelector('input') ? cell.querySelector('input').value : cell.innerText; cols.push('"' + txt.replace(/"/g, '""') + '"'); }); csv.push(cols.join(';')); } }); var blob = new Blob(['\\uFEFF' + csv.join('\\r\\n')], { type: 'text/csv;charset=utf-8;' }); var url = URL.createObjectURL(blob); var a = document.createElement('a'); a.href = url; a.download = 'reporte.csv'; a.click(); }
    document.addEventListener('DOMContentLoaded', function() { updateDatalist(); var saved = localStorage.getItem('theme') || 'light'; document.documentElement.setAttribute('data-theme', saved); var rows = document.querySelectorAll('tbody tr:not(.row-total)'); var modal = document.getElementById('detailModal'); var modalBody = modal.querySelector('.modal-body'); var modalTitle = modal.querySelector('.modal-title'); var headers = Array.from(document.querySelectorAll('thead th')).map(function(th) { return th.innerText; }); function showModal(row) { if (navigator.vibrate) navigator.vibrate(50); modalBody.innerHTML = ''; var cells = row.querySelectorAll('td'); modalTitle.innerText = cells[0].innerText || 'Detalle'; cells.forEach(function(cell, index) { var val = cell.innerText; if (cell.querySelector('a')) val = cell.innerHTML; if (cell.querySelector('.badge')) val = cell.querySelector('.badge').innerText; if (cell.querySelector('input')) val = cell.querySelector('input').value; var item = document.createElement('div'); item.className = 'detail-item'; item.innerHTML = '<div class="detail-label">' + headers[index] + '</div><div class="detail-value">' + val + '</div>'; modalBody.appendChild(item); }); modal.classList.add('active'); } rows.forEach(function(row) { row.addEventListener('dblclick', function() { showModal(row); }); }); });
    function closeModal() { document.getElementById('detailModal').classList.remove('active'); }
  </script>
</head>
<body>
  <div id='detailModal' class='modal-overlay' onclick='if(event.target === this) closeModal()'>
    <div class='modal-card'>
      <div class='modal-header'>
        <h3 class='modal-title'>Detalle</h3><button class='btn' style='border:none' onclick='closeModal()'><svg width='20' height='20' fill='none' stroke='currentColor' stroke-width='2' viewBox='0 0 24 24'><path d='M18 6L6 18M6 6l12 12'></path></svg></button>
      </div>
      <div class='modal-body'></div>
    </div>
  </div>
  <div class='page'>
    <div class='header-container'>
      <div class='title-area'>
        <h1>${esc(fname)}</h1>
        <div class='author-pill'><svg width='14' height='14' fill='none' stroke='currentColor' stroke-width='2' viewBox='0 0 24 24'><path d='M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2'></path><circle cx='12' cy='7' r='4'></circle></svg> ${esc(
          author
        )} <div class='subtitle-area'>From DataReport By MFPT: ${timestamp}</div></div>
      </div>
      <div class='meta-row'>
        <datalist id='search-options'></datalist>
        <div class='search-floater'>
          <div class='select-wrapper'>
            <select id='filterCol' class='filter-select' onchange='onSelectorChange()'>${filterOptions}</select>
            <svg class='select-arrow' fill='none' stroke='currentColor' stroke-width='2' viewBox='0 0 24 24'><polyline points='6 9 12 15 18 9'></polyline></svg>
          </div>
          <div class='search-input-wrapper'>
            <input type='text' id='search' list='search-options' autocomplete='on' class='search-input' onkeyup='filterTbl()' placeholder='Buscar...'>
          </div>
        </div>
        <div class='actions-area'>
          <div class='btn-group'>
            <button class='btn btn-primary' onclick='toggleTheme()' title='Cambiar Tema'><svg width='20' height='20' viewBox='0 0 48 48'><g data-name='Layer 2'><path fill='none' d='M0 0h48v48H0z' /><path d='M14 24a10 10 0 0 0 10 10V14a10 10 0 0 0-10 10' /><path d='M24 2a22 22 0 1 0 22 22A21.9 21.9 0 0 0 24 2M6 24A18.1 18.1 0 0 1 24 6v8a10 10 0 0 1 0 20v8A18.1 18.1 0 0 1 6 24' /></g></svg></button>
            <button class='btn btn-primary' onclick='dlXLS()' title='Exportar a CSV'><svg width='20' height='20' viewBox='0 0 32 32' fill='none'><rect x='8' y='2' width='24' height='28' rx='2' fill='#2FB776' /><path d='M8 23H32V28C32 29.1046 31.1046 30 30 30H10C8.89543 30 8 29.1046 8 28V23Z' fill='url(#paint0_linear_87_7712)' /><rect x='20' y='16' width='12' height='7' fill='#229C5B' /><rect x='20' y='9' width='12' height='7' fill='#27AE68' /><path d='M8 4C8 2.89543 8.89543 2 10 2H20V9H8V4Z' fill='#1D854F' /><rect x='8' y='9' width='12' height='7' fill='#197B43' /><rect x='8' y='16' width='12' height='7' fill='#1B5B38' /><path d='M8 12C8 10.3431 9.34315 9 11 9H17C18.6569 9 20 10.3431 20 12V24C20 25.6569 18.6569 27 17 27H8V12Z' fill='#000000' fill-opacity='0.3' /><rect y='7' width='18' height='18' rx='2' fill='url(#paint1_linear_87_7712)' /><path d='M13 21L10.1821 15.9L12.8763 11H10.677L9.01375 14.1286L7.37801 11H5.10997L7.81787 15.9L5 21H7.19931L8.97251 17.6857L10.732 21H13Z' fill='white' /><defs><linearGradient id='paint0_linear_87_7712' x1='8' y1='26.5' x2='32' y2='26.5' gradientUnits='userSpaceOnUse'><stop stop-color='#163C27' /><stop offset='1' stop-color='#2A6043' /></linearGradient><linearGradient id='paint1_linear_87_7712' x1='0' y1='16' x2='18' y2='16' gradientUnits='userSpaceOnUse'><stop stop-color='#185A30' /><stop offset='1' stop-color='#176F3D' /></linearGradient></defs></svg></button>
            <button class='btn btn-primary' onclick='printRpt()' title='Imprimir'><svg width='24' height='24' viewBox='0 0 48 48' version='1' enable-background='new 0 0 48 48'><rect x='9' y='11' fill='#424242' width='30' height='3' /><path fill='#616161' d='M4,25h40v-7c0-2.2-1.8-4-4-4H8c-2.2,0-4,1.8-4,4V25z' /><path fill='#424242' d='M8,36h32c2.2,0,4-1.8,4-4v-8H4v8C4,34.2,5.8,36,8,36z' /><circle fill='#00E676' cx='40' cy='18' r='1' /><rect x='11' y='4' fill='#90CAF9' width='26' height='10' /><path fill='#242424' d='M37.5,31h-27C9.7,31,9,30.3,9,29.5v0c0-0.8,0.7-1.5,1.5-1.5h27c0.8,0,1.5,0.7,1.5,1.5v0 C39,30.3,38.3,31,37.5,31z' /><rect x='11' y='31' fill='#90CAF9' width='26' height='11' /><rect x='11' y='29' fill='#42A5F5' width='26' height='2' /><g fill='#1976D2'><rect x='16' y='33' width='17' height='2' /><rect x='16' y='37' width='13' height='2' /></g></svg></button>
          </div>
        </div>
      </div>
    </div>
    <div class='table-card'>
      <div class='table-container'>
        <table><thead><tr>${headersHtml}</tr></thead><tbody>${rowsHtml}${totalRowHtml}</tbody></table>
      </div>
    </div>
  </div>
</body>
</html>`;

      const url = URL.createObjectURL(
        new Blob([fullHtml], { type: "text/html" })
      );
      const a = document.createElement("a");
      a.href = url;
      a.download = `${fname}.html`;
      a.click();
    } else if (format === "pdf") {
      const { jsPDF } = window.jspdf;
      const doc = new jsPDF({ orientation: "landscape" });
      doc.text(fname, 14, 15);
      doc.setFontSize(10);
      doc.setTextColor(100);
      doc.text(`Autor: ${author} | ${timestamp}`, 14, 22);
      doc.autoTable({
        head: [Object.keys(exportData[0])],
        body: exportData.map(Object.values),
        startY: 28,
        theme: "grid",
        styles: { fontSize: 8 },
        headStyles: { fillColor: [59, 130, 246] }
      });
      doc.save(`${fname}.pdf`);
    }
    this.showToast(`Exportado a ${format.toUpperCase()}`, "success");
  }

  // --- UTILS --- //
  toggleMenu(id) {
    document.getElementById(id).classList.toggle("show");
  }
  toggleTheme() {
    const html = document.documentElement;
    const isDark = html.getAttribute("data-theme") === "dark";
    html.setAttribute("data-theme", isDark ? "light" : "dark");
    document.getElementById("themeIcon").className = isDark
      ? "ph ph-moon"
      : "ph ph-sun";
  }

  setLoading(v) {
    if (v) {
      this.els.loadingState.classList.remove("hidden");
      this.els.emptyState.classList.add("hidden");
      this.els.tableWrapper.classList.add("hidden");
    } else {
      this.els.loadingState.classList.add("hidden");
    }
  }

  getAlignClass(type) {
    if (["number", "currency", "integer", "percent"].includes(type))
      return "text-right";
    if (["date", "datetime", "time"].includes(type)) return "text-center";
    return "";
  }

  formatValue(val, config) {
    const type = config.type;
    const decimals = config.decimals !== undefined ? config.decimals : 2;
    const curr = config.currency || "PEN";

    if (val === null || val === undefined || val === "") return "";
    if (type === "text") return String(val);

    if (typeof val === "string") {
      const vUpper = val.toUpperCase();
      if (["APROBADO", "OK", "SI", "SÍ", "ACTIVO", "PAGADO"].includes(vUpper))
        return `<span class="badge badge-success">${val}</span>`;
      if (["ERROR", "NO", "BAJA", "FALLO", "ANULADO"].includes(vUpper))
        return `<span class="badge badge-danger">${val}</span>`;
      if (["PENDIENTE", "ESPERA", "PROCESO"].includes(vUpper))
        return `<span class="badge badge-warning">${val}</span>`;
      if (val.startsWith("http"))
        return `<a href="${val}" target="_blank" style="color:var(--accent)">Link</a>`;
      return val;
    }

    if (
      type === "date" ||
      type === "datetime" ||
      type === "time" ||
      val instanceof Date
    ) {
      try {
        const d = val instanceof Date ? val : new Date(val);
        if (isNaN(d.getTime())) return val;

        if (type === "time")
          return d.toLocaleTimeString("es-PE", {
            hour: "2-digit",
            minute: "2-digit"
          });

        const opts = {};
        if (type === "datetime") {
          opts.hour = "2-digit";
          opts.minute = "2-digit";
        }

        const style = config.dateStyle || "short";
        if (style === "short") {
          opts.day = "2-digit";
          opts.month = "2-digit";
          opts.year = "numeric";
        } else if (style === "medium") {
          opts.day = "numeric";
          opts.month = "short";
          opts.year = "numeric";
        } else if (style === "long") {
          opts.day = "numeric";
          opts.month = "long";
          opts.year = "numeric";
        } else if (style === "full") {
          opts.weekday = "long";
          opts.day = "numeric";
          opts.month = "long";
          opts.year = "numeric";
        }

        return d.toLocaleDateString("es-PE", opts);
      } catch (e) {
        return val;
      }
    }

    if (typeof val === "number") {
      if (type === "currency")
        return val.toLocaleString("es-PE", {
          style: "currency",
          currency: curr,
          minimumFractionDigits: decimals,
          maximumFractionDigits: decimals
        });
      if (type === "percent")
        return val.toLocaleString("es-PE", {
          style: "percent",
          minimumFractionDigits: decimals,
          maximumFractionDigits: decimals
        });
      if (type === "integer") return parseInt(val).toLocaleString("es-PE");
      return val.toLocaleString("es-PE", {
        minimumFractionDigits: decimals,
        maximumFractionDigits: decimals
      });
    }
    return val;
  }

  showToast(msg, type = "info") {
    const c = document.getElementById("toastContainer");
    const t = document.createElement("div");
    t.className = `toast toast-${type}`;
    const icon =
      type === "success"
        ? "ph-check-circle"
        : type === "error"
        ? "ph-warning-circle"
        : "ph-info";
    t.innerHTML = `<i class="ph ${icon}" style="font-size:20px; color:${
      type === "success" ? "var(--success)" : "var(--danger)"
    }"></i><span>${msg}</span>`;
    c.appendChild(t);
    setTimeout(() => {
      t.style.opacity = "0";
      t.addEventListener("transitionend", () => t.remove());
    }, 3000);
  }

  // --- CUSTOM CSV EXPORT LOGIC --- //

  openCsvMapper() {
    this.els.exportMenu.classList.remove("show");

    if (this.columns.length === 0)
      return this.showToast("No hay datos cargados", "error");

    // AGREGAMOS this.els.mapOrdenCompra AL ARRAY
    const selects = [
      this.els.mapLocalidad,
      this.els.mapScanCode,
      this.els.mapProducto,
      this.els.mapPedido,
      this.els.mapOrdenCompra
    ];

    selects.forEach((sel) => {
      sel.innerHTML = '<option value="">-- Seleccionar Columna --</option>';
      this.columns.forEach((col) => {
        const opt = document.createElement("option");
        opt.value = col;
        opt.innerText = col;
        sel.appendChild(opt);
      });
    });

    // Auto-selección inteligente
    this.autoSelectOption(this.els.mapLocalidad, [
      "localidad",
      "nom_tienda",
      "rs comprador",
      "local",
      "ciudad",
      "sede",
      "ubicacion"  
    ]);
    this.autoSelectOption(this.els.mapScanCode, [
      "cod. empaque",
      "upc",
      "code",
      "codigo",
      "ean",
      "sku"
    ]);
    this.autoSelectOption(this.els.mapProducto, [
      "producto",
      "descripcion",
      "descripcion_larga",
      "sku_name",
      "item",
      "nombre"
    ]);
    this.autoSelectOption(this.els.mapPedido, [
      "pedido", 
      "unidades",
      "cant", 
      "solicitud"
    ]);
    this.autoSelectOption(this.els.mapOrdenCompra, [
      "orden",
      "num_oc",
      "compra",
      "oc",
      "po",
      "purchase",
      "numero"
    ]);

    // Resetear estado manual
    this.els.chkManualLocalidad.checked = false;
    this.toggleLocalidadInput();
    this.els.inputManualLocalidad.value = "";
    this.els.chkAutoOC.checked = false;
    this.toggleAutoOC(); // Esto reseteará la vista al select normal
    this.els.csvMapModal.classList.add("active");
  }

  autoSelectOption(selectElement, keywords) {
    const options = Array.from(selectElement.options);
    const found = options.find((opt) => {
      const txt = opt.text.toLowerCase();
      return keywords.some((k) => txt.includes(k));
    });
    if (found) selectElement.value = found.value;
  }

  toggleLocalidadInput() {
    const isManual = this.els.chkManualLocalidad.checked;
    if (isManual) {
      this.els.mapLocalidad.classList.add("hidden");
      this.els.inputManualLocalidad.classList.remove("hidden");
      this.els.inputManualLocalidad.focus();
    } else {
      this.els.mapLocalidad.classList.remove("hidden");
      this.els.inputManualLocalidad.classList.add("hidden");
    }
  }

  toggleAutoOC() {
    const isAuto = this.els.chkAutoOC.checked;

    if (isAuto) {
      // Ocultar select, mostrar preview
      this.els.mapOrdenCompra.classList.add("hidden");
      this.els.previewAutoOC.classList.remove("hidden");

      // Generar preview del formato
      const now = new Date();
      const pad = (n) => String(n).padStart(2, "0");
      const fmt = `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(
        now.getDate()
      )}_${pad(now.getHours())}${pad(now.getMinutes())}${pad(
        now.getSeconds()
      )}`;
      this.els.previewAutoOC.value = fmt;
    } else {
      // Mostrar select, ocultar preview
      this.els.mapOrdenCompra.classList.remove("hidden");
      this.els.previewAutoOC.classList.add("hidden");
    }
  }

  generateCustomCSV() {
    // 1. Preparar Auto-Generación de Fecha/Hora
    const isAutoOC = this.els.chkAutoOC.checked;
    let autoOCValue = "";

    if (isAutoOC) {
      const now = new Date();
      const pad = (n) => String(n).padStart(2, "0");
      // Formato: yyyymmdd_hhmmss
      autoOCValue = `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(
        now.getDate()
      )}_${pad(now.getHours())}${pad(now.getMinutes())}${pad(
        now.getSeconds()
      )}`;
    }

    // 2. Obtener Mapeo
    const map = {
      locCol: this.els.mapLocalidad.value,
      scanCol: this.els.mapScanCode.value,
      prodCol: this.els.mapProducto.value,
      pedCol: this.els.mapPedido.value,
      ocCol: this.els.mapOrdenCompra.value,
      isManualLoc: this.els.chkManualLocalidad.checked,
      manualLocVal: this.els.inputManualLocalidad.value.trim().toUpperCase()
    };

    // 3. Validaciones
    if (!map.isManualLoc && !map.locCol)
      return this.showToast("Falta definir columna Localidad", "error");
    if (map.isManualLoc && !map.manualLocVal)
      return this.showToast("Falta valor manual de Localidad", "error");
    if (!map.scanCol) return this.showToast("Falta columna Scan Code", "error");
    if (!map.prodCol) return this.showToast("Falta columna Producto", "error");
    if (!map.pedCol) return this.showToast("Falta columna Pedido", "error");
    if (!isAutoOC && !map.ocCol)
      return this.showToast("Falta columna Orden de Compra", "error");

    // 4. Generar CSV (ESTÁNDAR US/EN)
    const csvRows = [];

    // CAMBIO 1: Usamos comas en los encabezados
    csvRows.push("LOCALIDAD,SCAN_COD,PRODUCTO X,PEDIDO,ORDEN DE COMPRA");

    // Función de limpieza para formato US
    // Reemplazamos cualquier coma dentro del texto por un espacio para no romper las columnas
    const clean = (txt) => {
      if (txt === null || txt === undefined) return "";
      let str = String(txt);
      return str
        .replace(/,/g, " ")
        .replace(/[\r\n]+/g, " ")
        .trim();
    };

    this.visibleData.forEach((row) => {
      let valLoc = map.isManualLoc ? map.manualLocVal : row[map.locCol] || "";
      let valScan = row[map.scanCol] || "";
      let valProd = row[map.prodCol] || "";
      let valPed = row[map.pedCol] || "";
      let valOC = isAutoOC ? autoOCValue : row[map.ocCol] || "";

      // CAMBIO 2: Unimos los valores con comas
      csvRows.push(
        `${clean(valLoc)},${clean(valScan)},${clean(valProd)},${clean(
          valPed
        )},${clean(valOC)}`
      );
    });

    // 5. Descargar
    const csvContent = "\uFEFF" + csvRows.join("\r\n");
    const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");

    const fname = this.els.reportTitle.value.trim() || "Reporte";
    link.setAttribute("href", url);
    link.setAttribute("download", `${fname}_custom.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);

    this.els.csvMapModal.classList.remove("active");
    this.showToast("CSV (Formato US) generado correctamente", "success");
  }
}

const app = new DataViewerApp();




